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
package render

import (
	"strings"
	"testing"
)

func TestRenderTableAndParseTableFromTextRoundTrip(t *testing.T) {
	t.Parallel()

	headers := []string{"Name", "Count"}
	rows := [][]string{
		{"Windows 11", "42"},
		{"Windows 10", "11"},
	}

	table := RenderTable(headers, rows)
	gotHeaders, gotRows, ok := ParseTableFromText(table)
	if !ok {
		t.Fatal("expected ParseTableFromText to detect a table")
	}
	if strings.Join(gotHeaders, "|") != strings.Join(headers, "|") {
		t.Fatalf("headers mismatch: got %v want %v", gotHeaders, headers)
	}
	if len(gotRows) != len(rows) {
		t.Fatalf("row count mismatch: got %d want %d", len(gotRows), len(rows))
	}
	for i := range rows {
		if strings.Join(gotRows[i], "|") != strings.Join(rows[i], "|") {
			t.Fatalf("row %d mismatch: got %v want %v", i, gotRows[i], rows[i])
		}
	}
}

func TestParseTableFromTextReturnsFalseForNonTable(t *testing.T) {
	t.Parallel()

	_, _, ok := ParseTableFromText("plain text only")
	if ok {
		t.Fatal("expected non-table text to return ok=false")
	}
}

func TestRenderInspectorProducesParsableFieldValueTable(t *testing.T) {
	t.Parallel()

	out := RenderInspector("App Inspector", [][2]string{
		{"Display Name", "Company Portal"},
		{"Assignment Count", "3"},
	})

	if !strings.Contains(out, "App Inspector") {
		t.Fatalf("expected inspector title in output:\n%s", out)
	}
	headers, rows, ok := ParseTableFromText(out)
	if !ok {
		t.Fatalf("expected inspector output to contain a table:\n%s", out)
	}
	if strings.Join(headers, "|") != "Field|Value" {
		t.Fatalf("unexpected headers: %v", headers)
	}
	if len(rows) != 2 || rows[0][0] != "Display Name" || rows[1][1] != "3" {
		t.Fatalf("unexpected rows: %v", rows)
	}
}
