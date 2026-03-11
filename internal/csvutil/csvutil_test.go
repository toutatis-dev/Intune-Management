package csvutil

import (
	"os"
	"path/filepath"
	"strings"
	"testing"

	"intune-management/internal/render"
)

func TestValidateStrictPassesValidUsersCSV(t *testing.T) {
	t.Parallel()

	path := writeTempFile(t, "users.csv", "User_Principal_Name\nalice@example.com\nbob@example.com\n")
	res, err := ValidateStrict(path, []string{"User_Principal_Name"}, []string{"User_Principal_Name"})
	if err != nil {
		t.Fatalf("ValidateStrict returned error: %v", err)
	}
	if !res.Pass {
		t.Fatalf("expected validation to pass, got %+v", res)
	}
	if res.Rows != 2 || res.Errors != 0 {
		t.Fatalf("unexpected summary: %+v", res)
	}
}

func TestValidateStrictFailsMissingHeader(t *testing.T) {
	t.Parallel()

	path := writeTempFile(t, "users-missing-header.csv", "Wrong_Header\nalice@example.com\n")
	res, err := ValidateStrict(path, []string{"User_Principal_Name"}, []string{"User_Principal_Name"})
	if err != nil {
		t.Fatalf("ValidateStrict returned error: %v", err)
	}
	if res.Pass {
		t.Fatalf("expected validation to fail, got %+v", res)
	}
	if res.Errors != 1 {
		t.Fatalf("expected exactly one error, got %+v", res)
	}
	if !hasIssueCode(res.Issues, "MISSING_HEADER") {
		t.Fatalf("unexpected issues: %+v", res.Issues)
	}
}

func TestValidateStrictFailsDuplicateKeyAndEmptyValue(t *testing.T) {
	t.Parallel()

	path := writeTempFile(t, "apps.csv", "Group_Name,App_Name\nGroupA,App1\nGroupA,App1\nGroupB,\n")
	res, err := ValidateStrict(path, []string{"Group_Name", "App_Name"}, []string{"Group_Name", "App_Name"})
	if err != nil {
		t.Fatalf("ValidateStrict returned error: %v", err)
	}
	if res.Pass {
		t.Fatalf("expected validation to fail, got %+v", res)
	}
	if res.Errors < 2 {
		t.Fatalf("expected at least two errors, got %+v", res)
	}
	codes := make([]string, 0, len(res.Issues))
	for _, issue := range res.Issues {
		codes = append(codes, issue.Code)
	}
	joined := strings.Join(codes, ",")
	if !strings.Contains(joined, "DUPLICATE_KEY") || !strings.Contains(joined, "MISSING_REQUIRED_VALUE") {
		t.Fatalf("expected duplicate and missing value errors, got %+v", res.Issues)
	}
}

func TestValidateStrictAddsWhitespaceWarning(t *testing.T) {
	t.Parallel()

	path := writeTempFile(t, "groups.csv", " Group_Name \nWorkstations\n")
	res, err := ValidateStrict(path, []string{"Group_Name"}, []string{"Group_Name"})
	if err != nil {
		t.Fatalf("ValidateStrict returned error: %v", err)
	}
	if !res.Pass {
		t.Fatalf("expected validation to pass, got %+v", res)
	}
	if res.Warnings != 1 {
		t.Fatalf("expected one warning, got %+v", res)
	}
	if res.Issues[0].Code != "HEADER_WHITESPACE" {
		t.Fatalf("unexpected issue code: %+v", res.Issues)
	}
}

func TestValidateStrictFailsDuplicateNormalizedHeaders(t *testing.T) {
	t.Parallel()

	path := writeTempFile(t, "dup-headers.csv", "Group_Name, Group_Name \nWorkstations,Servers\n")
	res, err := ValidateStrict(path, []string{"Group_Name"}, []string{"Group_Name"})
	if err != nil {
		t.Fatalf("ValidateStrict returned error: %v", err)
	}
	if res.Pass {
		t.Fatalf("expected duplicate normalized headers to fail, got %+v", res)
	}
	if !hasIssueCode(res.Issues, "DUPLICATE_HEADER") {
		t.Fatalf("expected DUPLICATE_HEADER issue, got %+v", res.Issues)
	}
}

func TestValidateStrictFailsNoDataRows(t *testing.T) {
	t.Parallel()

	path := writeTempFile(t, "headers-only.csv", "User_Principal_Name\n")
	res, err := ValidateStrict(path, []string{"User_Principal_Name"}, []string{"User_Principal_Name"})
	if err != nil {
		t.Fatalf("ValidateStrict returned error: %v", err)
	}
	if res.Pass {
		t.Fatalf("expected headers-only CSV to fail, got %+v", res)
	}
	if !hasIssueCode(res.Issues, "NO_DATA_ROWS") {
		t.Fatalf("expected NO_DATA_ROWS issue, got %+v", res.Issues)
	}
}

func TestReadNormalizedTrimsHeadersAndValues(t *testing.T) {
	t.Parallel()

	path := writeTempFile(t, "trimmed.csv", " Group_Name , App_Name \n Sales Team , Company Portal \n")
	data, err := ReadNormalized(path)
	if err != nil {
		t.Fatalf("ReadNormalized returned error: %v", err)
	}
	if strings.Join(data.Headers, "|") != "Group_Name|App_Name" {
		t.Fatalf("unexpected normalized headers: %v", data.Headers)
	}
	if got := data.Rows[0]["Group_Name"]; got != "Sales Team" {
		t.Fatalf("unexpected normalized group value: %q", got)
	}
	if got := data.Rows[0]["App_Name"]; got != "Company Portal" {
		t.Fatalf("unexpected normalized app value: %q", got)
	}
}

func TestFormatValidationReportIncludesSummaryAndIssueTable(t *testing.T) {
	t.Parallel()

	res := ValidationResult{
		FilePath: "sample.csv",
		Rows:     3,
		Errors:   1,
		Warnings: 1,
		Pass:     false,
		Issues: []Issue{
			{Severity: "Error", Row: 2, Field: "App_Name", Code: "MISSING_REQUIRED_VALUE", Message: "Required value is empty"},
			{Severity: "Warning", Row: 1, Field: " Group_Name ", Code: "HEADER_WHITESPACE", Message: "Header has leading/trailing whitespace"},
		},
	}

	report := FormatValidationReport("CSV Quality Report", res)
	if !strings.Contains(report, "Status: FAIL") {
		t.Fatalf("expected FAIL status in report:\n%s", report)
	}
	if !strings.Contains(report, "MISSING_REQUIRED_VALUE") {
		t.Fatalf("expected issue code in report:\n%s", report)
	}
	if _, _, ok := render.ParseTableFromText(report); !ok {
		t.Fatalf("expected issue table in report:\n%s", report)
	}
}

func hasIssueCode(issues []Issue, code string) bool {
	for _, issue := range issues {
		if issue.Code == code {
			return true
		}
	}
	return false
}

func writeTempFile(t *testing.T, name, contents string) string {
	t.Helper()
	dir := t.TempDir()
	path := filepath.Join(dir, name)
	if err := os.WriteFile(path, []byte(contents), 0644); err != nil {
		t.Fatalf("os.WriteFile failed: %v", err)
	}
	return path
}
