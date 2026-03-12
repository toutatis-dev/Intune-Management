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
package app

import (
	"os"
	"path/filepath"
	"regexp"
	"strings"
	"testing"

	tea "github.com/charmbracelet/bubbletea"

	"intune-management/internal/config"
	"intune-management/internal/graph"
)

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
	wantDir, err := exportBaseDir()
	if err != nil {
		// If exportBaseDir fails, defaultExportPath should return empty.
		if path != "" {
			t.Fatalf("expected empty path when exportBaseDir fails, got %q", path)
		}
		return
	}
	if filepath.Ext(path) != ".csv" {
		t.Fatalf("expected csv extension, got %q", path)
	}
	if filepath.Dir(path) != wantDir {
		t.Fatalf("expected export path under %q, got %q", wantDir, filepath.Dir(path))
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

	m := newModel(nil, "test")
	m.state = stateUsersGroups
	m.width = 120
	m.height = 40

	visible := m.visibleMenu()
	if len(visible) < 5 {
		t.Fatalf("expected at least 5 visible menu items, got %d", len(visible))
	}
	updated, _ := m.Update(tea.KeyMsg(tea.Key{Type: tea.KeyRunes, Runes: []rune{'3'}}))
	m2 := updated.(model)
	if m2.state == stateUsersGroups && m2.cursor == 0 {
		t.Fatal("expected number hotkey to change cursor or state")
	}

	m3 := newModel(nil, "test")
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

	m := newModel(nil, "test")
	m.state = stateInput
	m.lastMenuState = stateUsersGroups
	m.currentSpec = specForAction(actAddUsersCSV)
	m.inputs = []string{"/tmp/users.csv"}
	m.input.SetValue("TestGroup")
	m.input.Prompt = m.currentSpec.prompts[1] + ": "

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

	updated2, _ := m2.Update(tea.KeyMsg(tea.Key{Type: tea.KeyEscape}))
	m3 := updated2.(model)
	if m3.state != stateUsersGroups {
		t.Fatalf("expected return to menu state, got %d", m3.state)
	}
}

func TestVisibleMenuFiltering(t *testing.T) {
	t.Parallel()

	m := newModel(nil, "test")
	m.state = stateUsersGroups

	m.filterQuery = ""
	all := m.visibleMenu()
	if len(all) != len(m.userMenu) {
		t.Fatalf("expected %d items with empty filter, got %d", len(m.userMenu), len(all))
	}

	m.filterQuery = "zzzznonexistent"
	none := m.visibleMenu()
	if len(none) != 0 {
		t.Fatalf("expected 0 items for non-matching filter, got %d", len(none))
	}

	m.filterQuery = "BULK"
	matched := m.visibleMenu()
	if len(matched) != 1 {
		t.Fatalf("expected 1 match for 'BULK', got %d", len(matched))
	}
	if !strings.Contains(matched[0].item.label, "Bulk") {
		t.Fatalf("expected match to contain 'Bulk', got %q", matched[0].item.label)
	}
}

func TestResultSummaryViewIncludesStickyContext(t *testing.T) {
	t.Parallel()

	m := model{
		client:          graph.NewStubClient(config.AuthConfig{TenantID: "tenant-123", ClientID: "client-abc"}),
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

func TestJumpToNextMatchWraps(t *testing.T) {
	t.Parallel()

	m := newModel(nil, "test")
	m.output = "alpha\nbeta\ngamma\ndelta\nbeta again"
	m.searchQuery = "beta"
	m.searchMatchLine = -1
	m.viewport.Width = 80
	m.viewport.Height = 20
	m.viewport.SetContent(m.output)

	m.jumpToNextMatch(true)
	if m.searchMatchLine != 1 {
		t.Fatalf("expected first forward match at line 1, got %d", m.searchMatchLine)
	}

	m.jumpToNextMatch(true)
	if m.searchMatchLine != 4 {
		t.Fatalf("expected second forward match at line 4, got %d", m.searchMatchLine)
	}

	m.jumpToNextMatch(true)
	if m.searchMatchLine != 1 {
		t.Fatalf("expected wrap-around forward match at line 1, got %d", m.searchMatchLine)
	}

	m.jumpToNextMatch(false)
	if m.searchMatchLine != 4 {
		t.Fatalf("expected backward wrap match at line 4, got %d", m.searchMatchLine)
	}
}

func TestExportDirectoryValidation(t *testing.T) {
	t.Parallel()

	m := newModel(nil, "test")
	m.state = stateExportPrompt
	m.lastHeaders = []string{"A"}
	m.lastRows = [][]string{{"1"}}
	m.styles = newUIStyles()
	m.width = 120
	m.height = 40
	m.vpReady = true

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

func writeTempFile(t *testing.T, name, contents string) string {
	t.Helper()

	dir := t.TempDir()
	path := filepath.Join(dir, name)
	if err := os.WriteFile(path, []byte(contents), 0644); err != nil {
		t.Fatalf("os.WriteFile failed: %v", err)
	}
	return path
}
