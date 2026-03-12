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
	"encoding/json"
	"testing"
	"time"
)

func TestChunkRequests(t *testing.T) {
	t.Parallel()

	t.Run("empty", func(t *testing.T) {
		t.Parallel()
		chunks := chunkRequests(nil, 20)
		if len(chunks) != 0 {
			t.Fatalf("expected 0 chunks, got %d", len(chunks))
		}
	})

	t.Run("45 requests", func(t *testing.T) {
		t.Parallel()
		reqs := make([]batchRequest, 45)
		for i := range reqs {
			reqs[i] = batchRequest{ID: string(rune('0' + i))}
		}
		chunks := chunkRequests(reqs, 20)
		if len(chunks) != 3 {
			t.Fatalf("expected 3 chunks, got %d", len(chunks))
		}
		if len(chunks[0]) != 20 {
			t.Fatalf("chunk 0: expected 20, got %d", len(chunks[0]))
		}
		if len(chunks[1]) != 20 {
			t.Fatalf("chunk 1: expected 20, got %d", len(chunks[1]))
		}
		if len(chunks[2]) != 5 {
			t.Fatalf("chunk 2: expected 5, got %d", len(chunks[2]))
		}
	})

	t.Run("exact multiple", func(t *testing.T) {
		t.Parallel()
		reqs := make([]batchRequest, 20)
		chunks := chunkRequests(reqs, 20)
		if len(chunks) != 1 {
			t.Fatalf("expected 1 chunk, got %d", len(chunks))
		}
	})
}

func TestSelectRetryable(t *testing.T) {
	t.Parallel()

	responses := []batchResponse{
		{ID: "1", Status: 200},
		{ID: "2", Status: 429},
		{ID: "3", Status: 404},
		{ID: "4", Status: 500},
		{ID: "5", Status: 503},
		{ID: "6", Status: 403},
	}
	ids := selectRetryable(responses)
	if len(ids) != 3 {
		t.Fatalf("expected 3 retryable, got %d: %v", len(ids), ids)
	}
	want := map[string]bool{"2": true, "4": true, "5": true}
	for _, id := range ids {
		if !want[id] {
			t.Fatalf("unexpected retryable ID: %s", id)
		}
	}
}

func TestMaxRetryAfter(t *testing.T) {
	t.Parallel()

	responses := []batchResponse{
		{ID: "1", Status: 429, Headers: map[string]string{"Retry-After": "3"}},
		{ID: "2", Status: 429, Headers: map[string]string{"Retry-After": "7"}},
		{ID: "3", Status: 429, Headers: map[string]string{"Retry-After": "2"}},
	}
	got := maxRetryAfter(responses)
	if got != 7*time.Second {
		t.Fatalf("expected 7s, got %s", got)
	}
}

func TestMaxRetryAfterNoHeaders(t *testing.T) {
	t.Parallel()

	responses := []batchResponse{
		{ID: "1", Status: 429},
		{ID: "2", Status: 500},
	}
	got := maxRetryAfter(responses)
	if got != 0 {
		t.Fatalf("expected 0, got %s", got)
	}
}

func TestParseBatchValues(t *testing.T) {
	t.Parallel()

	body, _ := json.Marshal(map[string]any{
		"value": []map[string]any{
			{"id": "a1", "intent": "available"},
			{"id": "a2", "intent": "required"},
		},
		"@odata.nextLink": "https://graph.microsoft.com/v1.0/next",
	})
	resp := batchResponse{ID: "1", Status: 200, Body: body}
	result, err := parseBatchValues(resp)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if len(result.Values) != 2 {
		t.Fatalf("expected 2 values, got %d", len(result.Values))
	}
	if result.NextLink != "https://graph.microsoft.com/v1.0/next" {
		t.Fatalf("unexpected nextLink: %s", result.NextLink)
	}
}

func TestParseBatchValuesEmpty(t *testing.T) {
	t.Parallel()

	body, _ := json.Marshal(map[string]any{
		"value": []map[string]any{},
	})
	resp := batchResponse{ID: "1", Status: 200, Body: body}
	result, err := parseBatchValues(resp)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if len(result.Values) != 0 {
		t.Fatalf("expected 0 values, got %d", len(result.Values))
	}
	if result.NextLink != "" {
		t.Fatalf("expected empty nextLink, got %s", result.NextLink)
	}
}

func TestParseBatchValuesError(t *testing.T) {
	t.Parallel()

	t.Run("error status", func(t *testing.T) {
		t.Parallel()
		resp := batchResponse{ID: "1", Status: 404, Body: json.RawMessage(`{"error":{"code":"NotFound"}}`)}
		_, err := parseBatchValues(resp)
		if err == nil {
			t.Fatal("expected error for 404 status")
		}
	})

	t.Run("invalid json", func(t *testing.T) {
		t.Parallel()
		resp := batchResponse{ID: "1", Status: 200, Body: json.RawMessage(`not json`)}
		_, err := parseBatchValues(resp)
		if err == nil {
			t.Fatal("expected error for invalid json")
		}
	})
}
