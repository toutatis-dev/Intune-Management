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
package csvutil

import (
	"encoding/csv"
	"errors"
	"fmt"
	"io"
	"os"
	"strings"

	"intune-management/internal/render"
)

var ErrNoDataRows = errors.New("csv has no data rows")

type Issue struct {
	Severity string
	Row      int
	Field    string
	Code     string
	Message  string
}

type ValidationResult struct {
	FilePath string
	Rows     int
	Errors   int
	Warnings int
	Issues   []Issue
	Pass     bool
}

type Dataset struct {
	Headers         []string
	OriginalHeaders []string
	Rows            []map[string]string
}

func countSeverity(issues []Issue, severity string) int {
	total := 0
	for _, issue := range issues {
		if issue.Severity == severity {
			total++
		}
	}
	return total
}

func normalizeHeaders(headers []string, requiredHeaders map[string]struct{}) ([]string, []Issue) {
	normalized := make([]string, len(headers))
	issues := make([]Issue, 0)
	seen := map[string]int{}
	for i, h := range headers {
		trimmed := strings.TrimSpace(h)
		normalized[i] = trimmed
		if trimmed != h {
			issues = append(issues, Issue{
				Severity: "Warning",
				Row:      1,
				Field:    h,
				Code:     "HEADER_WHITESPACE",
				Message:  "Header has leading/trailing whitespace",
			})
		}
		if first, exists := seen[trimmed]; exists {
			issues = append(issues, Issue{
				Severity: "Error",
				Row:      1,
				Field:    trimmed,
				Code:     "DUPLICATE_HEADER",
				Message:  fmt.Sprintf("Duplicate header after normalization; first seen in column %d", first+1),
			})
			continue
		}
		seen[trimmed] = i
		if len(requiredHeaders) > 0 {
			if _, ok := requiredHeaders[trimmed]; !ok && trimmed != "" {
				issues = append(issues, Issue{
					Severity: "Warning",
					Row:      1,
					Field:    trimmed,
					Code:     "EXTRA_HEADER",
					Message:  "Header is not used by this workflow",
				})
			}
		}
	}
	return normalized, issues
}

func hasError(issues []Issue) bool {
	for _, issue := range issues {
		if issue.Severity == "Error" {
			return true
		}
	}
	return false
}

func ReadNormalized(path string) (Dataset, error) {
	f, err := os.Open(path)
	if err != nil {
		return Dataset{}, err
	}
	defer f.Close()

	r := csv.NewReader(f)
	rows, err := r.ReadAll()
	if err != nil {
		return Dataset{}, err
	}
	if len(rows) == 0 {
		return Dataset{}, io.EOF
	}

	normalizedHeaders, issues := normalizeHeaders(rows[0], nil)
	if hasError(issues) {
		return Dataset{}, errors.New(FormatValidationReport("CSV Parse Failed", ValidationResult{
			FilePath: path,
			Errors:   countSeverity(issues, "Error"),
			Warnings: countSeverity(issues, "Warning"),
			Issues:   issues,
			Pass:     false,
		}))
	}
	if len(rows) < 2 {
		return Dataset{}, ErrNoDataRows
	}

	data := Dataset{
		Headers:         normalizedHeaders,
		OriginalHeaders: append([]string(nil), rows[0]...),
		Rows:            make([]map[string]string, 0, len(rows)-1),
	}
	for _, row := range rows[1:] {
		item := make(map[string]string, len(normalizedHeaders))
		for i, header := range normalizedHeaders {
			val := ""
			if i < len(row) {
				val = strings.TrimSpace(row[i])
			}
			item[header] = val
		}
		data.Rows = append(data.Rows, item)
	}
	return data, nil
}

func Read(path string) ([]map[string]string, error) {
	data, err := ReadNormalized(path)
	if err != nil {
		return nil, err
	}
	return data.Rows, nil
}

func ValidateStrict(path string, requiredHeaders, keyColumns []string) (ValidationResult, error) {
	res := ValidationResult{FilePath: path, Pass: true}
	f, err := os.Open(path)
	if err != nil {
		return res, err
	}
	defer f.Close()

	r := csv.NewReader(f)
	rows, err := r.ReadAll()
	if err != nil {
		return res, err
	}
	if len(rows) == 0 {
		return res, io.EOF
	}

	requiredSet := make(map[string]struct{}, len(requiredHeaders))
	for _, req := range requiredHeaders {
		requiredSet[req] = struct{}{}
	}

	normalizedHeaders, issues := normalizeHeaders(rows[0], requiredSet)
	res.Issues = append(res.Issues, issues...)
	headerMap := map[string]int{}
	for i, h := range normalizedHeaders {
		if _, exists := headerMap[h]; !exists {
			headerMap[h] = i
		}
	}
	for _, req := range requiredHeaders {
		if _, ok := headerMap[req]; !ok {
			res.Issues = append(res.Issues, Issue{
				Severity: "Error",
				Row:      1,
				Field:    req,
				Code:     "MISSING_HEADER",
				Message:  "Required header is missing",
			})
		}
	}
	if len(rows) < 2 {
		res.Issues = append(res.Issues, Issue{
			Severity: "Error",
			Row:      1,
			Code:     "NO_DATA_ROWS",
			Message:  "CSV contains headers but no data rows",
		})
	}
	if hasError(res.Issues) {
		res.Errors = countSeverity(res.Issues, "Error")
		res.Warnings = countSeverity(res.Issues, "Warning")
		res.Pass = false
		return res, nil
	}

	seen := map[string]int{}
	for i, row := range rows[1:] {
		rowNum := i + 2
		res.Rows++
		item := make(map[string]string, len(normalizedHeaders))
		for idx, header := range normalizedHeaders {
			val := ""
			if idx < len(row) {
				val = strings.TrimSpace(row[idx])
			}
			item[header] = val
		}
		keyParts := make([]string, 0, len(keyColumns))
		for _, req := range requiredHeaders {
			if item[req] == "" {
				res.Issues = append(res.Issues, Issue{
					Severity: "Error",
					Row:      rowNum,
					Field:    req,
					Code:     "MISSING_REQUIRED_VALUE",
					Message:  "Required value is empty",
				})
			}
		}
		for _, k := range keyColumns {
			keyParts = append(keyParts, strings.ToLower(item[k]))
		}
		key := strings.Join(keyParts, "|")
		if key != "" {
			if first, exists := seen[key]; exists {
				res.Issues = append(res.Issues, Issue{
					Severity: "Error",
					Row:      rowNum,
					Field:    strings.Join(keyColumns, "+"),
					Code:     "DUPLICATE_KEY",
					Message:  fmt.Sprintf("Duplicate key; first seen at row %d", first),
				})
			} else {
				seen[key] = rowNum
			}
		}
	}

	res.Errors = countSeverity(res.Issues, "Error")
	res.Warnings = countSeverity(res.Issues, "Warning")
	res.Pass = res.Errors == 0
	return res, nil
}

func FormatValidationReport(title string, res ValidationResult) string {
	status := "PASS"
	if !res.Pass {
		status = "FAIL"
	}
	var b strings.Builder
	fmt.Fprintf(&b, "%s\n\nFile: %s\nRows: %d\nErrors: %d\nWarnings: %d\nStatus: %s\n\n",
		title, res.FilePath, res.Rows, res.Errors, res.Warnings, status)
	if len(res.Issues) == 0 {
		b.WriteString("No validation issues found.\n")
		return b.String()
	}

	rows := make([][]string, 0, len(res.Issues))
	for _, issue := range res.Issues {
		rows = append(rows, []string{
			issue.Severity,
			fmt.Sprintf("%d", issue.Row),
			issue.Field,
			issue.Code,
			issue.Message,
		})
	}
	b.WriteString(render.RenderTable([]string{"Severity", "Row", "Field", "Code", "Message"}, rows))
	return b.String()
}
