package main

import (
	"encoding/base64"
	"encoding/json"
	"errors"
	"os"
	"path/filepath"
	"regexp"
	"strings"
	"testing"
	"time"

	tea "github.com/charmbracelet/bubbletea"
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
	if !hasIssueCode(res.Issues, "MISSING_HEADER") {
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

func TestReadCSVNormalizedTrimsHeadersAndValues(t *testing.T) {
	t.Parallel()

	path := writeTempFile(t, "trimmed.csv", " Group_Name , App_Name \n Sales Team , Company Portal \n")
	data, err := readCSVNormalized(path)
	if err != nil {
		t.Fatalf("readCSVNormalized returned error: %v", err)
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

func TestValidateCSVStrictFailsDuplicateNormalizedHeaders(t *testing.T) {
	t.Parallel()

	path := writeTempFile(t, "dup-headers.csv", "Group_Name, Group_Name \nWorkstations,Servers\n")
	res, err := validateCSVStrict(path, []string{"Group_Name"}, []string{"Group_Name"})
	if err != nil {
		t.Fatalf("validateCSVStrict returned error: %v", err)
	}
	if res.Pass {
		t.Fatalf("expected duplicate normalized headers to fail, got %+v", res)
	}
	if !hasIssueCode(res.Issues, "DUPLICATE_HEADER") {
		t.Fatalf("expected DUPLICATE_HEADER issue, got %+v", res.Issues)
	}
}

func TestValidateCSVStrictFailsNoDataRows(t *testing.T) {
	t.Parallel()

	path := writeTempFile(t, "headers-only.csv", "User_Principal_Name\n")
	res, err := validateCSVStrict(path, []string{"User_Principal_Name"}, []string{"User_Principal_Name"})
	if err != nil {
		t.Fatalf("validateCSVStrict returned error: %v", err)
	}
	if res.Pass {
		t.Fatalf("expected headers-only CSV to fail, got %+v", res)
	}
	if !hasIssueCode(res.Issues, "NO_DATA_ROWS") {
		t.Fatalf("expected NO_DATA_ROWS issue, got %+v", res.Issues)
	}
}

func TestPreviewForActionUsesNormalizedHeaders(t *testing.T) {
	t.Parallel()

	path := writeTempFile(t, "apps-preview.csv", " Group_Name , App_Name \n Sales , Portal \n")
	out, ok, err := previewForAction(actionSpec{id: actAddAppsCSV}, []string{path})
	if err != nil {
		t.Fatalf("previewForAction returned error: %v", err)
	}
	if !ok {
		t.Fatal("expected previewForAction to handle actAddAppsCSV")
	}
	if !strings.Contains(out, "Sales") || !strings.Contains(out, "Portal") {
		t.Fatalf("expected preview output to contain normalized values:\n%s", out)
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
	if !isGraphNotFound(formatGraphError("GET", "https://graph.microsoft.com/v1.0/groups/MyGroup", "400 Bad Request", []byte(`{"error":{"code":"Request_BadRequest","message":"Invalid object identifier 'MyGroup'."}}`))){
		t.Fatal("expected 400 invalid object identifier to be treated as not found")
	}
	if isGraphNotFound(formatGraphError("GET", "https://graph.microsoft.com/v1.0/groups/1", "400 Bad Request", []byte(`{"error":{"code":"Request_BadRequest","message":"Some other bad request."}}`))){
		t.Fatal("expected 400 non-identifier bad request to not be treated as not found")
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

func TestWaitProgressCmdReturnsProgress(t *testing.T) {
	t.Parallel()

	ch := make(chan progressMsg, 1)
	done := make(chan struct{})
	ch <- progressMsg{text: "working"}
	msg := waitProgressCmd(ch, done)()
	got, ok := msg.(progressMsg)
	if !ok || got.text != "working" {
		t.Fatalf("unexpected progress message: %#v", msg)
	}
}

func TestWaitProgressCmdStopsOnDone(t *testing.T) {
	t.Parallel()

	ch := make(chan progressMsg)
	done := make(chan struct{})
	close(done)
	msg := waitProgressCmd(ch, done)()
	if _, ok := msg.(progressStopMsg); !ok {
		t.Fatalf("expected progressStopMsg, got %#v", msg)
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

func TestResolveTopFailingAppSelectionUsesAppIDAndRejectsAmbiguousName(t *testing.T) {
	t.Parallel()

	rows := [][]string{
		{"1", "Portal", "app-1", "4", "10", "40.0%"},
		{"2", "Portal", "app-2", "3", "6", "50.0%"},
	}
	got, err := resolveTopFailingAppSelection("1", rows)
	if err != nil {
		t.Fatalf("unexpected rank lookup error: %v", err)
	}
	if got != "app-1" {
		t.Fatalf("expected app ID from rank lookup, got %q", got)
	}
	if _, err := resolveTopFailingAppSelection("Portal", rows); err == nil || !strings.Contains(err.Error(), "use rank instead") {
		t.Fatalf("expected ambiguous app name error, got %v", err)
	}
}

func TestResultSummaryViewIncludesStickyContext(t *testing.T) {
	t.Parallel()

	m := model{
		client:          &graphClient{cfg: authConfig{TenantID: "tenant-123", ClientID: "client-abc"}},
		styles:          newUIStyles(),
		lastActionLabel: "Top 10 Failing App Deployments",
		lastHeaders:     []string{"App", "Count"},
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

func hasIssueCode(issues []csvIssue, code string) bool {
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

// --- Round 2 tests ---

func TestParentMenuState(t *testing.T) {
	t.Parallel()

	tests := []struct {
		state menuState
		want  menuState
	}{
		{stateUsersGroups, stateMain},
		{stateDevicesApps, stateMain},
		{stateReports, stateMain},
		{stateSettings, stateMain},
		{stateReportCsv, stateReports},
		{stateReportInspect, stateReports},
		{stateMain, stateMain},
	}
	for _, tt := range tests {
		if got := parentMenuState(tt.state); got != tt.want {
			t.Fatalf("parentMenuState(%d) = %d, want %d", tt.state, got, tt.want)
		}
	}
}

func TestNumberHotkeys(t *testing.T) {
	t.Parallel()

	m := newModel(nil)
	m.state = stateUsersGroups
	m.width = 120
	m.height = 40

	// Press "3" on a menu with 6 items — should select item 3.
	visible := m.visibleMenu()
	if len(visible) < 5 {
		t.Fatalf("expected at least 5 visible menu items, got %d", len(visible))
	}
	updated, _ := m.Update(tea.KeyMsg(tea.Key{Type: tea.KeyRunes, Runes: []rune{'3'}}))
	m2 := updated.(model)
	// Number hotkey re-sends Enter, so the model should have transitioned.
	// At minimum the cursor should have moved to index 2 before the Enter.
	// The exact state depends on the action; just verify no panic and state changed.
	if m2.state == stateUsersGroups && m2.cursor == 0 {
		t.Fatal("expected number hotkey to change cursor or state")
	}

	// Press "9" on a menu with 6 items — should be ignored.
	m3 := newModel(nil)
	m3.state = stateUsersGroups
	m3.width = 120
	m3.height = 40
	updated2, _ := m3.Update(tea.KeyMsg(tea.Key{Type: tea.KeyRunes, Runes: []rune{'9'}}))
	m4 := updated2.(model)
	if m4.state != stateUsersGroups || m4.cursor != 0 {
		t.Fatalf("expected out-of-range number hotkey to be ignored, state=%d cursor=%d", m4.state, m4.cursor)
	}
}

func TestInputBackNavigation(t *testing.T) {
	t.Parallel()

	m := newModel(nil)
	m.state = stateInput
	m.lastMenuState = stateUsersGroups
	m.currentSpec = specForAction(actAddUsersCSV) // 2-prompt action
	m.inputs = []string{"/tmp/users.csv"}
	m.input.SetValue("TestGroup")
	m.input.Prompt = m.currentSpec.prompts[1] + ": "

	// Press Esc on step 2 — should go back to step 1 with previous value restored.
	updated, _ := m.Update(tea.KeyMsg(tea.Key{Type: tea.KeyEscape}))
	m2 := updated.(model)
	if m2.state != stateInput {
		t.Fatalf("expected state to remain stateInput, got %d", m2.state)
	}
	if len(m2.inputs) != 0 {
		t.Fatalf("expected inputs to be empty after back, got %v", m2.inputs)
	}
	if m2.input.Value() != "/tmp/users.csv" {
		t.Fatalf("expected previous value restored, got %q", m2.input.Value())
	}

	// Press Esc again on step 1 — should return to menu.
	updated2, _ := m2.Update(tea.KeyMsg(tea.Key{Type: tea.KeyEscape}))
	m3 := updated2.(model)
	if m3.state != stateUsersGroups {
		t.Fatalf("expected return to menu state, got %d", m3.state)
	}
}

func TestVisibleMenuFiltering(t *testing.T) {
	t.Parallel()

	m := newModel(nil)
	m.state = stateUsersGroups

	// Empty filter returns all items.
	m.filterQuery = ""
	all := m.visibleMenu()
	if len(all) != len(m.userMenu) {
		t.Fatalf("expected %d items with empty filter, got %d", len(m.userMenu), len(all))
	}

	// No-match filter returns empty.
	m.filterQuery = "zzzznonexistent"
	none := m.visibleMenu()
	if len(none) != 0 {
		t.Fatalf("expected 0 items for non-matching filter, got %d", len(none))
	}

	// Case-insensitive match.
	m.filterQuery = "BULK"
	matched := m.visibleMenu()
	if len(matched) != 1 {
		t.Fatalf("expected 1 match for 'BULK', got %d", len(matched))
	}
	if !strings.Contains(matched[0].item.label, "Bulk") {
		t.Fatalf("expected match to contain 'Bulk', got %q", matched[0].item.label)
	}
}

func TestJumpToNextMatchWraps(t *testing.T) {
	t.Parallel()

	m := newModel(nil)
	m.output = "alpha\nbeta\ngamma\ndelta\nbeta again"
	m.searchQuery = "beta"
	m.searchMatchLine = -1
	m.viewport.Width = 80
	m.viewport.Height = 20
	m.viewport.SetContent(m.output)

	// Forward: should find first "beta" at line 1.
	m.jumpToNextMatch(true)
	if m.searchMatchLine != 1 {
		t.Fatalf("expected first forward match at line 1, got %d", m.searchMatchLine)
	}

	// Forward again: should wrap to line 4 ("beta again").
	m.jumpToNextMatch(true)
	if m.searchMatchLine != 4 {
		t.Fatalf("expected second forward match at line 4, got %d", m.searchMatchLine)
	}

	// Forward again: should wrap back to line 1.
	m.jumpToNextMatch(true)
	if m.searchMatchLine != 1 {
		t.Fatalf("expected wrap-around forward match at line 1, got %d", m.searchMatchLine)
	}

	// Backward from line 1: should wrap to line 4.
	m.jumpToNextMatch(false)
	if m.searchMatchLine != 4 {
		t.Fatalf("expected backward wrap match at line 4, got %d", m.searchMatchLine)
	}
}

func TestExportDirectoryValidation(t *testing.T) {
	t.Parallel()

	m := newModel(nil)
	m.state = stateExportPrompt
	m.lastHeaders = []string{"A"}
	m.lastRows = [][]string{{"1"}}
	m.styles = newUIStyles()
	m.width = 120
	m.height = 40
	m.vpReady = true

	// Enter a path with a nonexistent parent directory.
	m.exportInput.SetValue(filepath.Join(t.TempDir(), "nonexistent", "subdir", "file.csv"))
	updated, _ := m.Update(tea.KeyMsg(tea.Key{Type: tea.KeyEnter}))
	m2 := updated.(model)
	if m2.state != stateOutput {
		t.Fatalf("expected transition to stateOutput on bad dir, got %d", m2.state)
	}
	if !strings.Contains(m2.output, "Directory does not exist") {
		t.Fatalf("expected directory error message, got %q", m2.output)
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
