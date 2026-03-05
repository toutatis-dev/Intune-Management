package main

import (
	"encoding/base64"
	"encoding/json"
	"os"
	"path/filepath"
	"regexp"
	"strings"
	"testing"
	"time"
)

func TestRenderTableAndParseTableFromTextRoundTrip(t *testing.T) {
	t.Parallel()

	headers := []string{"Name", "Count"}
	rows := [][]string{
		{"Windows 11", "42"},
		{"Windows 10", "11"},
	}

	table := renderTable(headers, rows)
	gotHeaders, gotRows, ok := parseTableFromText(table)
	if !ok {
		t.Fatal("expected parseTableFromText to detect a table")
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

	_, _, ok := parseTableFromText("plain text only")
	if ok {
		t.Fatal("expected non-table text to return ok=false")
	}
}

func TestValidateCSVStrictPassesValidUsersCSV(t *testing.T) {
	t.Parallel()

	path := writeTempFile(t, "users.csv", "User_Principal_Name\nalice@example.com\nbob@example.com\n")
	res, err := validateCSVStrict(path, []string{"User_Principal_Name"}, []string{"User_Principal_Name"})
	if err != nil {
		t.Fatalf("validateCSVStrict returned error: %v", err)
	}
	if !res.Pass {
		t.Fatalf("expected validation to pass, got %+v", res)
	}
	if res.Rows != 2 || res.Errors != 0 {
		t.Fatalf("unexpected summary: %+v", res)
	}
}

func TestValidateCSVStrictFailsMissingHeader(t *testing.T) {
	t.Parallel()

	path := writeTempFile(t, "users-missing-header.csv", "Wrong_Header\nalice@example.com\n")
	res, err := validateCSVStrict(path, []string{"User_Principal_Name"}, []string{"User_Principal_Name"})
	if err != nil {
		t.Fatalf("validateCSVStrict returned error: %v", err)
	}
	if res.Pass {
		t.Fatalf("expected validation to fail, got %+v", res)
	}
	if res.Errors != 1 {
		t.Fatalf("expected exactly one error, got %+v", res)
	}
	if len(res.Issues) == 0 || res.Issues[0].Code != "MISSING_HEADER" {
		t.Fatalf("unexpected issues: %+v", res.Issues)
	}
}

func TestValidateCSVStrictFailsDuplicateKeyAndEmptyValue(t *testing.T) {
	t.Parallel()

	path := writeTempFile(t, "apps.csv", "Group_Name,App_Name\nGroupA,App1\nGroupA,App1\nGroupB,\n")
	res, err := validateCSVStrict(path, []string{"Group_Name", "App_Name"}, []string{"Group_Name", "App_Name"})
	if err != nil {
		t.Fatalf("validateCSVStrict returned error: %v", err)
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

func TestValidateCSVStrictAddsWhitespaceWarning(t *testing.T) {
	t.Parallel()

	path := writeTempFile(t, "groups.csv", " Group_Name \nWorkstations\n")
	res, err := validateCSVStrict(path, []string{"Group_Name"}, []string{"Group_Name"})
	if err != nil {
		t.Fatalf("validateCSVStrict returned error: %v", err)
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

func TestFormatValidationReportIncludesSummaryAndIssueTable(t *testing.T) {
	t.Parallel()

	res := csvValidationResult{
		FilePath: "sample.csv",
		Rows:     3,
		Errors:   1,
		Warnings: 1,
		Pass:     false,
		Issues: []csvIssue{
			{Severity: "Error", Row: 2, Field: "App_Name", Code: "MISSING_REQUIRED_VALUE", Message: "Required value is empty"},
			{Severity: "Warning", Row: 1, Field: " Group_Name ", Code: "HEADER_WHITESPACE", Message: "Header has leading/trailing whitespace"},
		},
	}

	report := formatValidationReport("CSV Quality Report", res)
	if !strings.Contains(report, "Status: FAIL") {
		t.Fatalf("expected FAIL status in report:\n%s", report)
	}
	if !strings.Contains(report, "MISSING_REQUIRED_VALUE") {
		t.Fatalf("expected issue code in report:\n%s", report)
	}
	if _, _, ok := parseTableFromText(report); !ok {
		t.Fatalf("expected issue table in report:\n%s", report)
	}
}

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

func TestValidateForActionMapsCSVActions(t *testing.T) {
	t.Parallel()

	path := writeTempFile(t, "users.csv", "User_Principal_Name\nalice@example.com\n")
	res, ok, err := validateForAction(actionSpec{id: actAddUsersCSV}, []string{path, "GroupA"})
	if err != nil {
		t.Fatalf("validateForAction returned error: %v", err)
	}
	if !ok || !res.Pass {
		t.Fatalf("expected CSV action validation to run and pass: ok=%v res=%+v", ok, res)
	}

	_, ok, err = validateForAction(actionSpec{id: actListGroups}, nil)
	if err != nil {
		t.Fatalf("unexpected error for non-CSV action: %v", err)
	}
	if ok {
		t.Fatal("expected non-CSV action to skip validation")
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
		{name: "windows 11", operatingSystem: "Windows", osVersion: "10.0.22631", want: "Windows 11"},
		{name: "windows 10", operatingSystem: "Windows", osVersion: "10.0.19045", want: "Windows 10"},
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

func TestRenderInspectorProducesParsableFieldValueTable(t *testing.T) {
	t.Parallel()

	out := renderInspector("App Inspector", [][2]string{
		{"Display Name", "Company Portal"},
		{"Assignment Count", "3"},
	})

	if !strings.Contains(out, "App Inspector") {
		t.Fatalf("expected inspector title in output:\n%s", out)
	}
	headers, rows, ok := parseTableFromText(out)
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

func TestActionLabelCoversNewActions(t *testing.T) {
	t.Parallel()

	cases := map[actionID]string{
		actReportWindowsBreakdown: "Windows OS Breakdown",
		actInspectApp:             "Inspect App",
		actAuthHealth:             "Auth Health",
	}
	for id, want := range cases {
		if got := actionLabel(id); got != want {
			t.Fatalf("actionLabel(%v) = %q, want %q", id, got, want)
		}
	}
}

func TestSlugifyNameAndDefaultExportPath(t *testing.T) {
	t.Parallel()

	if got := slugifyName(" Top 10 Failing App Deployments "); got != "top-10-failing-app-deployments" {
		t.Fatalf("unexpected slugifyName result: %q", got)
	}
	if got := slugifyName("!!!"); got != "report" {
		t.Fatalf("expected fallback slug, got %q", got)
	}

	m := model{lastActionLabel: "Windows OS Breakdown"}
	path := m.defaultExportPath()
	if filepath.Ext(path) != ".csv" {
		t.Fatalf("expected csv extension, got %q", path)
	}
	if filepath.Dir(path) != exportBaseDir() {
		t.Fatalf("expected export path under %q, got %q", exportBaseDir(), filepath.Dir(path))
	}
	base := filepath.Base(path)
	matched, err := regexp.MatchString(`^windows-os-breakdown-\d{8}-\d{6}\.csv$`, base)
	if err != nil {
		t.Fatalf("regexp.MatchString failed: %v", err)
	}
	if !matched {
		t.Fatalf("unexpected export filename shape: %q", base)
	}
}

func TestHelpTextForState(t *testing.T) {
	t.Parallel()

	menuHelp := helpTextForState(stateReports)
	if !strings.Contains(menuHelp, "Menu Help") || !strings.Contains(menuHelp, "Esc: Back") {
		t.Fatalf("unexpected menu help text:\n%s", menuHelp)
	}

	resultHelp := helpTextForState(stateOutput)
	if !strings.Contains(resultHelp, "Result Help") || !strings.Contains(resultHelp, "e: Export current table") {
		t.Fatalf("unexpected result help text:\n%s", resultHelp)
	}

	fallbackHelp := helpTextForState(stateHelp)
	if !strings.Contains(fallbackHelp, "Close help") {
		t.Fatalf("unexpected fallback help text:\n%s", fallbackHelp)
	}
}

func TestResultSummaryViewIncludesStickyContext(t *testing.T) {
	t.Parallel()

	m := model{
		client: &graphClient{cfg: authConfig{TenantID: "tenant-123", ClientID: "client-abc"}},
		styles: newUIStyles(),
		lastActionLabel: "Top 10 Failing App Deployments",
		lastHeaders: []string{"App", "Count"},
		lastRows: [][]string{
			{"Portal", "4"},
			{"VPN", "2"},
		},
		dryRun: true,
	}

	out := m.resultSummaryView()
	for _, want := range []string{
		"Action: Top 10 Failing App Deployments",
		"Tenant: tenant-123",
		"Client: client-abc",
		"Mode: DRY-RUN",
		"Rows: 2",
		"Export: available",
	} {
		if !strings.Contains(out, want) {
			t.Fatalf("expected summary to contain %q:\n%s", want, out)
		}
	}
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
