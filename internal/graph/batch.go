// Copyright (C) 2026 Jack Miller
//
// This program is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
//
// This program is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU General Public License for more details.
//
// You should have received a copy of the GNU General Public License
// along with this program.  If not, see <https://www.gnu.org/licenses/>.
package graph

import (
	"context"
	"encoding/json"
	"fmt"
	"net/http"
	"strconv"
	"strings"
	"time"
)

const maxBatchSize = 20

type batchRequest struct {
	ID     string `json:"id"`
	Method string `json:"method"`
	URL    string `json:"url"`
}

type batchResponse struct {
	ID      string            `json:"id"`
	Status  int               `json:"status"`
	Headers map[string]string `json:"headers"`
	Body    json.RawMessage   `json:"body"`
}

type batchRequestEnvelope struct {
	Requests []batchRequest `json:"requests"`
}

type batchResponseEnvelope struct {
	Responses []batchResponse `json:"responses"`
}

type batchValueResult struct {
	Values   []map[string]any
	NextLink string
}

func chunkRequests(requests []batchRequest, size int) [][]batchRequest {
	if len(requests) == 0 {
		return nil
	}
	var chunks [][]batchRequest
	for i := 0; i < len(requests); i += size {
		end := i + size
		if end > len(requests) {
			end = len(requests)
		}
		chunks = append(chunks, requests[i:end])
	}
	return chunks
}

const maxRetryWait = 60 * time.Second

func maxRetryAfter(responses []batchResponse) time.Duration {
	var max time.Duration
	for _, r := range responses {
		for k, v := range r.Headers {
			if !strings.EqualFold(k, "Retry-After") {
				continue
			}
			if secs, err := strconv.Atoi(strings.TrimSpace(v)); err == nil && secs > 0 {
				d := time.Duration(secs) * time.Second
				if d > max {
					max = d
				}
			}
		}
	}
	if max > maxRetryWait {
		max = maxRetryWait
	}
	return max
}

func parseBatchValues(resp batchResponse) (*batchValueResult, error) {
	if resp.Status >= 400 {
		return nil, fmt.Errorf("batch sub-request %s failed with status %d", resp.ID, resp.Status)
	}
	var page struct {
		Value    []map[string]any `json:"value"`
		NextLink string           `json:"@odata.nextLink"`
	}
	if err := json.Unmarshal(resp.Body, &page); err != nil {
		return nil, fmt.Errorf("batch sub-request %s: %w", resp.ID, err)
	}
	return &batchValueResult{Values: page.Value, NextLink: page.NextLink}, nil
}

func (g *Client) batch(ctx context.Context, requests []batchRequest) ([]batchResponse, error) {
	return g.batchWithEndpoint(ctx, requests, graphBase)
}

func (g *Client) batchWithEndpoint(ctx context.Context, requests []batchRequest, baseURL string) ([]batchResponse, error) {
	if len(requests) == 0 {
		return nil, nil
	}

	// Build lookup from ID to original request for retries.
	reqByID := make(map[string]batchRequest, len(requests))
	for _, r := range requests {
		reqByID[r.ID] = r
	}

	all := make(map[string]batchResponse, len(requests))
	chunks := chunkRequests(requests, maxBatchSize)

	for ci, chunk := range chunks {
		pending := chunk

		const maxRounds = 3
		for round := 0; round < maxRounds; round++ {
			envelope := batchRequestEnvelope{Requests: pending}
			raw, err := g.do(ctx, http.MethodPost, baseURL+"/$batch", envelope)
			if err != nil {
				return nil, fmt.Errorf("batch chunk %d: %w", ci, err)
			}

			var respEnv batchResponseEnvelope
			if err := json.Unmarshal(raw, &respEnv); err != nil {
				return nil, fmt.Errorf("batch chunk %d: parse response: %w", ci, err)
			}

			// Separate successes from retryable failures.
			var failed []batchResponse
			for _, r := range respEnv.Responses {
				if r.Status == http.StatusTooManyRequests || r.Status >= 500 {
					failed = append(failed, r)
				} else {
					all[r.ID] = r
				}
			}

			if len(failed) == 0 {
				break
			}

			// Last round — accept failures as-is.
			if round == maxRounds-1 {
				for _, r := range failed {
					all[r.ID] = r
				}
				break
			}

			// Wait before retrying.
			wait := maxRetryAfter(failed)
			if wait < time.Second {
				wait = time.Duration(1<<round) * time.Second
			}
			g.emitProgress(fmt.Sprintf("Batch chunk %d: %d sub-requests throttled; retrying in %s", ci+1, len(failed), wait))
			time.Sleep(wait)

			// Rebuild pending from failed IDs.
			pending = make([]batchRequest, 0, len(failed))
			for _, r := range failed {
				if orig, ok := reqByID[r.ID]; ok {
					pending = append(pending, orig)
				}
			}
		}

		g.emitProgress(fmt.Sprintf("Batch: completed chunk %d/%d", ci+1, len(chunks)))
	}

	// Return responses ordered by original request order.
	result := make([]batchResponse, 0, len(requests))
	for _, r := range requests {
		if resp, ok := all[r.ID]; ok {
			result = append(result, resp)
		}
	}
	return result, nil
}
