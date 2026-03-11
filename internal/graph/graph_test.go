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
	"encoding/base64"
	"encoding/json"
	"errors"
	"strings"
	"testing"
	"time"
)

func TestShouldRetryStatus(t *testing.T) {
	t.Parallel()

	retryable := []int{429, 500, 502, 503, 504}
	for _, status := range retryable {
		if !shouldRetryStatus(status) {
			t.Fatalf("expected status %d to be retryable", status)
		}
	}
	if shouldRetryStatus(404) {
		t.Fatal("expected 404 to be non-retryable")
	}
}

func TestRetryDelayUsesRetryAfterAndCapsExponentialBackoff(t *testing.T) {
	t.Parallel()

	if got := retryDelay(0, "3"); got != 3*time.Second {
		t.Fatalf("retry-after override mismatch: got %s want %s", got, 3*time.Second)
	}
	if got := retryDelay(10, ""); got != 8*time.Second {
		t.Fatalf("expected capped exponential backoff, got %s", got)
	}
}

func TestFormatGraphErrorPrefersGraphEnvelope(t *testing.T) {
	t.Parallel()

	raw := []byte(`{"error":{"code":"Request_BadRequest","message":"Bad input"}}`)
	err := formatGraphError("GET", "https://graph.microsoft.com/v1.0/test", "400 Bad Request", raw)
	msg := err.Error()
	if !strings.Contains(msg, "Request_BadRequest") || !strings.Contains(msg, "Bad input") {
		t.Fatalf("unexpected formatted error: %s", msg)
	}
}

func TestIsGraphNotFound(t *testing.T) {
	t.Parallel()

	if !isGraphNotFound(formatGraphError("GET", "https://graph.microsoft.com/v1.0/groups/1", "404 Not Found", []byte(`{"error":{"code":"Request_ResourceNotFound","message":"Missing"}}`))) {
		t.Fatal("expected 404 graph error to be treated as not found")
	}
	if isGraphNotFound(formatGraphError("GET", "https://graph.microsoft.com/v1.0/groups/1", "403 Forbidden", []byte(`{"error":{"code":"Authorization_RequestDenied","message":"Denied"}}`))) {
		t.Fatal("expected 403 graph error to not be treated as not found")
	}
	if isGraphNotFound(errors.New("plain error")) {
		t.Fatal("expected generic error to not be treated as not found")
	}
	if !isGraphNotFound(formatGraphError("GET", "https://graph.microsoft.com/v1.0/groups/MyGroup", "400 Bad Request", []byte(`{"error":{"code":"Request_BadRequest","message":"Invalid object identifier 'MyGroup'."}}`))) {
		t.Fatal("expected 400 invalid object identifier to be treated as not found")
	}
	if isGraphNotFound(formatGraphError("GET", "https://graph.microsoft.com/v1.0/groups/1", "400 Bad Request", []byte(`{"error":{"code":"Request_BadRequest","message":"Some other bad request."}}`))) {
		t.Fatal("expected 400 non-identifier bad request to not be treated as not found")
	}
}

func TestIsGraphForbidden(t *testing.T) {
	t.Parallel()

	if !isGraphForbidden(formatGraphError("GET", "https://graph.microsoft.com/v1.0/users/1", "403 Forbidden", []byte(`{"error":{"code":"Authorization_RequestDenied","message":"Denied"}}`))) {
		t.Fatal("expected 403 graph error to be treated as forbidden")
	}
	if isGraphForbidden(formatGraphError("GET", "https://graph.microsoft.com/v1.0/users/1", "404 Not Found", []byte(`{"error":{"code":"Request_ResourceNotFound","message":"Missing"}}`))) {
		t.Fatal("expected 404 graph error to not be treated as forbidden")
	}
	if isGraphForbidden(errors.New("plain error")) {
		t.Fatal("expected generic error to not be treated as forbidden")
	}
}

func TestDecodeJWTClaims(t *testing.T) {
	t.Parallel()

	header := base64.RawURLEncoding.EncodeToString([]byte(`{"alg":"none"}`))
	payloadMap := map[string]any{
		"tid": "tenant-id",
		"scp": "User.Read.All",
	}
	payloadJSON, err := json.Marshal(payloadMap)
	if err != nil {
		t.Fatalf("json.Marshal failed: %v", err)
	}
	payload := base64.RawURLEncoding.EncodeToString(payloadJSON)
	token := header + "." + payload + ".signature"

	claims, err := decodeJWTClaims(token)
	if err != nil {
		t.Fatalf("decodeJWTClaims returned error: %v", err)
	}
	if claims["tid"] != "tenant-id" || claims["scp"] != "User.Read.All" {
		t.Fatalf("unexpected claims: %+v", claims)
	}
}

func TestSelectUniqueMatch(t *testing.T) {
	t.Parallel()

	formatter := func(item map[string]any) string { return asString(item["displayName"]) + " | " + asString(item["id"]) }
	if _, err := selectUniqueMatch("group", "Workstations", nil, formatter); !errors.Is(err, errNotFound) {
		t.Fatalf("expected errNotFound, got %v", err)
	}
	item, err := selectUniqueMatch("group", "Workstations", []map[string]any{{"displayName": "Workstations", "id": "1"}}, formatter)
	if err != nil {
		t.Fatalf("unexpected error for unique match: %v", err)
	}
	if asString(item["id"]) != "1" {
		t.Fatalf("unexpected selected item: %+v", item)
	}
	_, err = selectUniqueMatch("group", "Workstations", []map[string]any{
		{"displayName": "Workstations", "id": "1"},
		{"displayName": "Workstations", "id": "2"},
	}, formatter)
	var ambErr *ambiguousMatchError
	if !errors.As(err, &ambErr) {
		t.Fatalf("expected ambiguousMatchError, got %v", err)
	}
	if !strings.Contains(err.Error(), "use object ID instead") {
		t.Fatalf("unexpected ambiguity message: %v", err)
	}
}

func TestClassifyWindowsVersion(t *testing.T) {
	t.Parallel()

	tests := []struct {
		name            string
		operatingSystem string
		osVersion       string
		want            string
	}{
		{name: "windows 11 3-part", operatingSystem: "Windows", osVersion: "10.0.22631", want: "Windows 11"},
		{name: "windows 11 4-part", operatingSystem: "Windows", osVersion: "10.0.22631.3958", want: "Windows 11"},
		{name: "windows 11 low revision", operatingSystem: "Windows", osVersion: "10.0.22621.1234", want: "Windows 11"},
		{name: "windows 10 3-part", operatingSystem: "Windows", osVersion: "10.0.19045", want: "Windows 10"},
		{name: "windows 10 4-part", operatingSystem: "Windows", osVersion: "10.0.19045.3803", want: "Windows 10"},
		{name: "non windows", operatingSystem: "iOS", osVersion: "17.0", want: "Other/Unknown"},
		{name: "bad version", operatingSystem: "Windows", osVersion: "unknown", want: "Other/Unknown"},
	}

	for _, tt := range tests {
		tt := tt
		t.Run(tt.name, func(t *testing.T) {
			t.Parallel()
			if got := classifyWindowsVersion(tt.operatingSystem, tt.osVersion); got != tt.want {
				t.Fatalf("classifyWindowsVersion(%q, %q) = %q, want %q", tt.operatingSystem, tt.osVersion, got, tt.want)
			}
		})
	}
}

func TestRenderTopFailingAppsReportExcludesZeroFailuresAndShowsSkipped(t *testing.T) {
	t.Parallel()

	out := renderTopFailingAppsReport([]appFailureStat{
		{ID: "a1", Name: "Portal", Failed: 4, Total: 10},
		{ID: "a2", Name: "VPN", Failed: 0, Total: 8},
		{ID: "a3", Name: "Agent", Failed: 2, Total: 5},
	}, failingAppsSummary{Scanned: 3, WithFailures: 2, Skipped: 1})
	if !strings.Contains(out, "Apps skipped due to errors: 1") {
		t.Fatalf("expected skipped count in report:\n%s", out)
	}
	if strings.Contains(out, "VPN") {
		t.Fatalf("expected zero-failure app to be excluded:\n%s", out)
	}
	if !strings.Contains(out, "Portal") || !strings.Contains(out, "a1") {
		t.Fatalf("expected ranked app and ID in report:\n%s", out)
	}
}
