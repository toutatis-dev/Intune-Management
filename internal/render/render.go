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
	"encoding/csv"
	"fmt"
	"os"
	"strings"

	"github.com/mattn/go-runewidth"
)

func padRight(s string, w int) string {
	sw := runewidth.StringWidth(s)
	if sw >= w {
		return s
	}
	return s + strings.Repeat(" ", w-sw)
}

func truncate(s string, w int) string {
	if w <= 0 || runewidth.StringWidth(s) <= w {
		return s
	}
	if w <= 3 {
		return runewidth.Truncate(s, w, "")
	}
	return runewidth.Truncate(s, w, "...")
}

func RenderTable(headers []string, rows [][]string) string {
	if len(headers) == 0 {
		return ""
	}
	widths := make([]int, len(headers))
	for i, h := range headers {
		widths[i] = runewidth.StringWidth(h)
	}
	for _, r := range rows {
		for i := range headers {
			if i < len(r) {
				if w := runewidth.StringWidth(r[i]); w > widths[i] {
					widths[i] = w
				}
			}
		}
	}
	for i := range widths {
		if widths[i] > 48 {
			widths[i] = 48
		}
	}
	var b strings.Builder
	for i, h := range headers {
		if i > 0 {
			b.WriteString(" | ")
		}
		b.WriteString(padRight(truncate(h, widths[i]), widths[i]))
	}
	b.WriteString("\n")
	for i, w := range widths {
		if i > 0 {
			b.WriteString("-+-")
		}
		b.WriteString(strings.Repeat("-", w))
	}
	b.WriteString("\n")
	for _, r := range rows {
		for i := range headers {
			if i > 0 {
				b.WriteString(" | ")
			}
			cell := ""
			if i < len(r) {
				cell = r[i]
			}
			b.WriteString(padRight(truncate(cell, widths[i]), widths[i]))
		}
		b.WriteString("\n")
	}
	return b.String()
}

func ParseTableFromText(s string) ([]string, [][]string, bool) {
	lines := strings.Split(s, "\n")
	for i := 0; i+1 < len(lines); i++ {
		h := strings.TrimSpace(lines[i])
		sep := strings.TrimSpace(lines[i+1])
		if h == "" || sep == "" {
			continue
		}
		if !strings.Contains(h, " | ") || !strings.Contains(sep, "-+-") {
			continue
		}
		headers := splitTableLine(h)
		if len(headers) == 0 {
			continue
		}
		rows := make([][]string, 0)
		for j := i + 2; j < len(lines); j++ {
			line := strings.TrimSpace(lines[j])
			if line == "" || !strings.Contains(line, " | ") {
				break
			}
			cells := splitTableLine(line)
			for len(cells) < len(headers) {
				cells = append(cells, "")
			}
			if len(cells) > len(headers) {
				cells = cells[:len(headers)]
			}
			rows = append(rows, cells)
		}
		if len(rows) > 0 {
			return headers, rows, true
		}
	}
	return nil, nil, false
}

func splitTableLine(line string) []string {
	parts := strings.Split(line, "|")
	out := make([]string, 0, len(parts))
	for _, p := range parts {
		out = append(out, strings.TrimSpace(p))
	}
	return out
}

func ExportCSV(path string, headers []string, rows [][]string) error {
	f, err := os.Create(path)
	if err != nil {
		return err
	}
	defer f.Close()

	w := csv.NewWriter(f)
	if err := w.Write(headers); err != nil {
		return err
	}
	for _, r := range rows {
		if err := w.Write(r); err != nil {
			return err
		}
	}
	w.Flush()
	return w.Error()
}

func RenderInspector(title string, values [][2]string) string {
	rows := make([][]string, 0, len(values))
	for _, pair := range values {
		rows = append(rows, []string{pair[0], pair[1]})
	}
	return fmt.Sprintf("%s\n\n%s", title, RenderTable([]string{"Field", "Value"}, rows))
}
