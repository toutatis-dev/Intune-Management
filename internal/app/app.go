package app

import (
	"context"
	"errors"
	"fmt"
	"os"
	"path/filepath"
	"strconv"
	"strings"
	"time"

	"github.com/charmbracelet/bubbles/spinner"
	"github.com/charmbracelet/bubbles/textinput"
	"github.com/charmbracelet/bubbles/viewport"
	tea "github.com/charmbracelet/bubbletea"
	"github.com/charmbracelet/lipgloss"

	"intune-management/internal/config"
	"intune-management/internal/csvutil"
	"intune-management/internal/graph"
	"intune-management/internal/render"
)

// --- TUI helpers bridging UI actions to library packages ---

func validateForAction(spec actionSpec, inputs []string) (csvutil.ValidationResult, bool, error) {
	switch spec.id {
	case actAddUsersCSV, actReportCsvUsers:
		res, err := csvutil.ValidateStrict(inputs[0], []string{"User_Principal_Name"}, []string{"User_Principal_Name"})
		return res, true, err
	case actMakeGroupsCSV, actReportCsvGroups:
		res, err := csvutil.ValidateStrict(inputs[0], []string{"Group_Name"}, []string{"Group_Name"})
		return res, true, err
	case actAddAppsCSV, actReportCsvApps:
		res, err := csvutil.ValidateStrict(inputs[0], []string{"Group_Name", "App_Name"}, []string{"Group_Name", "App_Name"})
		return res, true, err
	default:
		return csvutil.ValidationResult{}, false, nil
	}
}

func previewForAction(spec actionSpec, inputs []string) (string, bool, error) {
	switch spec.id {
	case actAddUsersCSV:
		data, err := csvutil.ReadNormalized(inputs[0])
		if err != nil {
			return "", true, err
		}
		sample := make([][]string, 0, min(10, len(data.Rows)))
		for i, row := range data.Rows {
			if i >= 10 {
				break
			}
			sample = append(sample, []string{row["User_Principal_Name"]})
		}
		return fmt.Sprintf("Preview - Bulk Add Users\n\nCSV: %s\nTarget Group: %s\nRows: %d\nShowing first %d rows\n\n%s",
			inputs[0],
			inputs[1],
			len(data.Rows),
			len(sample),
			render.RenderTable([]string{"User Principal Name"}, sample),
		), true, nil
	case actMakeGroupsCSV:
		data, err := csvutil.ReadNormalized(inputs[0])
		if err != nil {
			return "", true, err
		}
		sample := make([][]string, 0, min(10, len(data.Rows)))
		for i, row := range data.Rows {
			if i >= 10 {
				break
			}
			sample = append(sample, []string{row["Group_Name"]})
		}
		return fmt.Sprintf("Preview - Create Groups\n\nCSV: %s\nRows: %d\nShowing first %d rows\n\n%s",
			inputs[0],
			len(data.Rows),
			len(sample),
			render.RenderTable([]string{"Group Name"}, sample),
		), true, nil
	case actAddAppsCSV:
		data, err := csvutil.ReadNormalized(inputs[0])
		if err != nil {
			return "", true, err
		}
		sample := make([][]string, 0, min(10, len(data.Rows)))
		for i, row := range data.Rows {
			if i >= 10 {
				break
			}
			sample = append(sample, []string{row["Group_Name"], row["App_Name"]})
		}
		return fmt.Sprintf("Preview - Assign Apps\n\nCSV: %s\nRows: %d\nShowing first %d rows\n\n%s",
			inputs[0],
			len(data.Rows),
			len(sample),
			render.RenderTable([]string{"Group Name", "App Name"}, sample),
		), true, nil
	default:
		return "", false, nil
	}
}

func resolveTopFailingAppSelection(input string, rows [][]string) (string, error) {
	trimmed := strings.TrimSpace(input)
	if trimmed == "" {
		return "", errors.New("selection cannot be empty")
	}
	if rank, err := strconv.Atoi(trimmed); err == nil {
		for _, row := range rows {
			if len(row) >= 3 && row[0] == fmt.Sprintf("%d", rank) {
				return row[2], nil
			}
		}
		return "", errors.New("rank not found in current report")
	}
	matches := make([][]string, 0)
	for _, row := range rows {
		if len(row) >= 3 && strings.EqualFold(strings.TrimSpace(row[1]), trimmed) {
			matches = append(matches, row)
		}
	}
	if len(matches) == 1 {
		return matches[0][2], nil
	}
	if len(matches) > 1 {
		return "", errors.New("app name matches multiple rows; use rank instead")
	}
	return "", errors.New("app not found in current report")
}

// --- Menu types and enums ---

type menuState int

const (
	stateMain menuState = iota
	stateUsersGroups
	stateDevicesApps
	stateReports
	stateReportCsv
	stateReportInspect
	stateSettings
	stateMenuFilter
	stateInput
	statePreview
	stateExportPrompt
	stateConfirm
	stateWorking
	stateDrillPrompt
	stateHelp
	stateOutput
	stateOutputSearch
)

type actionID int

const (
	actNone actionID = iota
	actListUsersGroup
	actListGroups
	actListUsers
	actSearchGroups
	actAddUsersCSV
	actListDevices
	actListDevicesGroup
	actMakeGroupsCSV
	actAddAppsCSV
	actListGroupApps
	actReportComplianceSnapshot
	actReportWindowsBreakdown
	actReportTopFailingApps
	actReportCsvUsers
	actReportCsvGroups
	actReportCsvApps
	actInspectUser
	actInspectGroup
	actInspectDevice
	actInspectApp
	actSetClientID
	actSetTenantID
	actViewAuth
	actAuthHealth
	actResetAuth
	actToggleDryRun
)

type menuItem struct {
	label       string
	description string
	action      actionID
	next        menuState
}

type actionSpec struct {
	id      actionID
	prompts []string
}

type resultMsg struct {
	text string
	err  error
}

type progressMsg struct {
	text string
}

type progressStopMsg struct{}

type confirmKind int

const (
	confirmNone confirmKind = iota
	confirmAction
	confirmExport
)

// --- Model ---

type model struct {
	client        *graph.Client
	state         menuState
	cursor        int
	menuTop       int
	width         int
	height        int
	lastMenuState menuState

	mainMenu       []menuItem
	userMenu       []menuItem
	devMenu        []menuItem
	repMenu        []menuItem
	repCSVMenu     []menuItem
	repInspectMenu []menuItem
	cfgMenu        []menuItem

	spin            spinner.Model
	viewport        viewport.Model
	vpReady         bool
	styles          uiStyles
	input           textinput.Model
	filterInput     textinput.Model
	exportInput     textinput.Model
	drillInput      textinput.Model
	filterQuery     string
	currentSpec     actionSpec
	inputs          []string
	output          string
	lastHeaders     []string
	lastRows        [][]string
	lastActionLabel string

	confirmKind        confirmKind
	confirmTitle       string
	confirmBody        string
	confirmCancelState menuState
	pendingSpec        actionSpec
	pendingInputs      []string
	pendingExportPath  string
	dryRun             bool
	progressCh         chan progressMsg
	progressDone       chan struct{}
	progressActive     bool
	progressText       string
	cancelCtx          context.Context
	cancelFunc         context.CancelFunc
	helpReturnState    menuState
	lastActionID       actionID
	searchInput        textinput.Model
	searchQuery        string
	searchMatchLine    int
}

type uiStyles struct {
	app        lipgloss.Style
	header     lipgloss.Style
	subHeader  lipgloss.Style
	panel      lipgloss.Style
	selected   lipgloss.Style
	normalItem lipgloss.Style
	desc       lipgloss.Style
	hint       lipgloss.Style
	inputLabel lipgloss.Style
	ok         lipgloss.Style
	err        lipgloss.Style
}

func newUIStyles() uiStyles {
	return uiStyles{
		app:        lipgloss.NewStyle().Padding(1, 2),
		header:     lipgloss.NewStyle().Bold(true).Foreground(lipgloss.Color("230")).Background(lipgloss.Color("24")).Padding(0, 1),
		subHeader:  lipgloss.NewStyle().Bold(true).Foreground(lipgloss.Color("117")),
		panel:      lipgloss.NewStyle().Border(lipgloss.RoundedBorder()).BorderForeground(lipgloss.Color("62")).Padding(1, 2),
		selected:   lipgloss.NewStyle().Foreground(lipgloss.Color("15")).Background(lipgloss.Color("31")).Bold(true).Padding(0, 1),
		normalItem: lipgloss.NewStyle().Foreground(lipgloss.Color("252")).Padding(0, 1),
		desc:       lipgloss.NewStyle().Foreground(lipgloss.Color("245")),
		hint:       lipgloss.NewStyle().Foreground(lipgloss.Color("244")),
		inputLabel: lipgloss.NewStyle().Foreground(lipgloss.Color("111")).Bold(true),
		ok:         lipgloss.NewStyle().Foreground(lipgloss.Color("120")),
		err:        lipgloss.NewStyle().Foreground(lipgloss.Color("203")).Bold(true),
	}
}

func newModel(client *graph.Client) model {
	ti := textinput.New()
	ti.CharLimit = 512
	ti.Width = 72
	ti.Focus()
	fi := textinput.New()
	fi.CharLimit = 120
	fi.Width = 48
	fi.Prompt = "Filter: "
	ei := textinput.New()
	ei.CharLimit = 300
	ei.Width = 72
	ei.Prompt = "Export CSV path: "
	di := textinput.New()
	di.CharLimit = 200
	di.Width = 72
	di.Prompt = "Enter rank or exact app name: "
	si := textinput.New()
	si.CharLimit = 120
	si.Width = 48
	si.Prompt = "Search: "
	sp := spinner.New()
	sp.Spinner = spinner.Dot

	return model{
		client:        client,
		state:         stateMain,
		lastMenuState: stateMain,
		mainMenu: []menuItem{
			{label: "Manage Users and Groups", description: "List users, search groups, and bulk add members from CSV", next: stateUsersGroups},
			{label: "Manage Devices and Groups", description: "List devices, create groups, and manage app assignments", next: stateDevicesApps},
			{label: "Reports", description: "Read-only compliance and CSV quality analytics", next: stateReports},
			{label: "Settings", description: "Set Graph client and tenant IDs", next: stateSettings},
			{label: "Quit", description: "Exit the tool", action: actNone, next: -1},
		},
		userMenu: []menuItem{
			{label: "List users in group", description: "Show user members for a single group", action: actListUsersGroup},
			{label: "List all groups", description: "Display all Azure AD groups", action: actListGroups},
			{label: "List all users", description: "Display all users with UPN and object ID", action: actListUsers},
			{label: "Search groups", description: "Find groups by partial display name", action: actSearchGroups},
			{label: "Bulk add users (CSV)", description: "CSV header required: User_Principal_Name", action: actAddUsersCSV},
			{label: "Back", description: "Return to main menu", next: stateMain},
		},
		devMenu: []menuItem{
			{label: "List all devices", description: "Show all Entra devices", action: actListDevices},
			{label: "List devices in group", description: "Show only device members for a group", action: actListDevicesGroup},
			{label: "Create groups from CSV", description: "CSV header required: Group_Name", action: actMakeGroupsCSV},
			{label: "Assign apps by CSV", description: "CSV headers required: Group_Name, App_Name", action: actAddAppsCSV},
			{label: "List app-group assignments", description: "Show deployments in a table (press e in results to export)", action: actListGroupApps},
			{label: "Back", description: "Return to main menu", next: stateMain},
		},
		repMenu: []menuItem{
			{label: "Device compliance snapshot", description: "Compliant/noncompliant totals from Intune managed devices", action: actReportComplianceSnapshot},
			{label: "Windows OS breakdown", description: "Windows 10 vs 11 vs unknown from managed devices", action: actReportWindowsBreakdown},
			{label: "Top 10 failing app deployments", description: "Rank apps by failed device install statuses", action: actReportTopFailingApps},
			{label: "Object inspector", description: "Lookup one user, group, device, or app by ID or name", next: stateReportInspect},
			{label: "CSV validation checks", description: "Run strict quality checks for CSV workflows", next: stateReportCsv},
			{label: "Back", description: "Return to main menu", next: stateMain},
		},
		repCSVMenu: []menuItem{
			{label: "Validate Users->Group CSV", description: "Strict quality checks for User_Principal_Name format", action: actReportCsvUsers},
			{label: "Validate Create-Groups CSV", description: "Strict quality checks for Group_Name format", action: actReportCsvGroups},
			{label: "Validate App-Assignment CSV", description: "Strict quality checks for Group_Name + App_Name format", action: actReportCsvApps},
			{label: "Back", description: "Return to Reports menu", next: stateReports},
		},
		repInspectMenu: []menuItem{
			{label: "Inspect user", description: "Lookup by object ID, UPN, or exact display name", action: actInspectUser},
			{label: "Inspect group", description: "Lookup by object ID or exact display name", action: actInspectGroup},
			{label: "Inspect device", description: "Lookup by object ID or exact display name", action: actInspectDevice},
			{label: "Inspect app", description: "Lookup by app ID or exact display name", action: actInspectApp},
			{label: "Back", description: "Return to Reports menu", next: stateReports},
		},
		cfgMenu: []menuItem{
			{label: "Set Graph Client ID", description: "App registration client ID used for sign-in", action: actSetClientID},
			{label: "Set Graph Tenant ID", description: "Tenant GUID/domain or 'common'", action: actSetTenantID},
			{label: "View Current Auth Config", description: "Display current client and tenant IDs", action: actViewAuth},
			{label: "Auth Health", description: "Show token tenant/client/scopes and expiry", action: actAuthHealth},
			{label: "Toggle Dry-Run Mode", description: "When enabled, write operations are simulated only", action: actToggleDryRun},
			{label: "Reset Auth Defaults", description: "Client ID: Graph PowerShell app, Tenant: common", action: actResetAuth},
			{label: "Back", description: "Return to main menu", next: stateMain},
		},
		spin:            sp,
		styles:          newUIStyles(),
		input:           ti,
		filterInput:     fi,
		exportInput:     ei,
		drillInput:      di,
		searchInput:     si,
		searchMatchLine: -1,
		viewport:        viewport.New(80, 20),
		progressCh:      make(chan progressMsg, 64),
	}
}

func (m model) Init() tea.Cmd { return m.spin.Tick }

func (m model) menu() []menuItem {
	switch m.state {
	case stateMain:
		return m.mainMenu
	case stateUsersGroups:
		return m.userMenu
	case stateDevicesApps:
		return m.devMenu
	case stateReports:
		return m.repMenu
	case stateReportCsv:
		return m.repCSVMenu
	case stateReportInspect:
		return m.repInspectMenu
	case stateSettings:
		return m.cfgMenu
	default:
		return nil
	}
}

type visibleMenuItem struct {
	index int
	item  menuItem
}

func (m model) visibleMenu() []visibleMenuItem {
	menu := m.menu()
	if strings.TrimSpace(m.filterQuery) == "" {
		out := make([]visibleMenuItem, 0, len(menu))
		for i, item := range menu {
			out = append(out, visibleMenuItem{index: i, item: item})
		}
		return out
	}
	q := strings.ToLower(strings.TrimSpace(m.filterQuery))
	out := make([]visibleMenuItem, 0, len(menu))
	for i, item := range menu {
		if strings.Contains(strings.ToLower(item.label), q) || strings.Contains(strings.ToLower(item.description), q) {
			out = append(out, visibleMenuItem{index: i, item: item})
		}
	}
	return out
}

func (m *model) menuPageSize() int {
	return max(4, m.height-14)
}

func (m *model) ensureMenuCursorVisible() {
	visible := m.visibleMenu()
	page := m.menuPageSize()
	if len(visible) == 0 {
		m.cursor = 0
		m.menuTop = 0
		return
	}
	if m.cursor > len(visible)-1 {
		m.cursor = len(visible) - 1
	}
	if m.cursor < m.menuTop {
		m.menuTop = m.cursor
	}
	if m.cursor >= m.menuTop+page {
		m.menuTop = m.cursor - page + 1
	}
	if m.menuTop < 0 {
		m.menuTop = 0
	}
}

func parentMenuState(s menuState) menuState {
	switch s {
	case stateUsersGroups, stateDevicesApps, stateReports, stateSettings:
		return stateMain
	case stateReportCsv, stateReportInspect:
		return stateReports
	default:
		return stateMain
	}
}

func (m *model) resetMenuPosition(state menuState) {
	m.state = state
	m.cursor = 0
	m.menuTop = 0
	m.filterQuery = ""
	m.filterInput.SetValue("")
}

func (m *model) returnToLastMenu() {
	m.state = m.lastMenuState
	m.ensureMenuCursorVisible()
}

func (m *model) jumpToNextMatch(forward bool) {
	lines := strings.Split(m.output, "\n")
	q := strings.ToLower(m.searchQuery)
	start := m.searchMatchLine + 1
	if !forward {
		start = m.searchMatchLine - 1
	}
	n := len(lines)
	for i := 0; i < n; i++ {
		var idx int
		if forward {
			idx = (start + i) % n
			if idx < 0 {
				idx += n
			}
		} else {
			idx = (start - i) % n
			if idx < 0 {
				idx += n
			}
		}
		if strings.Contains(strings.ToLower(lines[idx]), q) {
			m.searchMatchLine = idx
			m.applySearchHighlight()
			lineOffset := max(0, idx-m.viewport.Height/2)
			m.viewport.SetYOffset(lineOffset)
			return
		}
	}
}

func (m *model) applySearchHighlight() {
	lines := strings.Split(m.output, "\n")
	if m.searchMatchLine < 0 || m.searchMatchLine >= len(lines) || m.searchQuery == "" {
		m.viewport.SetContent(m.output)
		return
	}
	hl := lipgloss.NewStyle().Reverse(true)
	built := make([]string, len(lines))
	copy(built, lines)
	built[m.searchMatchLine] = hl.Render(lines[m.searchMatchLine])
	m.viewport.SetContent(strings.Join(built, "\n"))
}

func (m *model) setOutput(text string) {
	m.output = text
	m.viewport.SetContent(text)
	m.viewport.GotoTop()
	m.lastHeaders = nil
	m.lastRows = nil
	if h, r, ok := render.ParseTableFromText(text); ok {
		m.lastHeaders = h
		m.lastRows = r
	}
	m.state = stateOutput
}

func actionLabel(id actionID) string {
	switch id {
	case actListUsersGroup:
		return "List Users In Group"
	case actListGroups:
		return "List All Groups"
	case actListUsers:
		return "List All Users"
	case actSearchGroups:
		return "Search Groups"
	case actAddUsersCSV:
		return "Bulk Add Users (CSV)"
	case actListDevices:
		return "List All Devices"
	case actListDevicesGroup:
		return "List Devices In Group"
	case actMakeGroupsCSV:
		return "Create Groups From CSV"
	case actAddAppsCSV:
		return "Assign Apps By CSV"
	case actListGroupApps:
		return "List App-Group Assignments"
	case actReportComplianceSnapshot:
		return "Device Compliance Snapshot"
	case actReportWindowsBreakdown:
		return "Windows OS Breakdown"
	case actReportTopFailingApps:
		return "Top 10 Failing App Deployments"
	case actInspectUser:
		return "Inspect User"
	case actInspectGroup:
		return "Inspect Group"
	case actInspectDevice:
		return "Inspect Device"
	case actInspectApp:
		return "Inspect App"
	case actReportCsvUsers:
		return "Validate Users->Group CSV"
	case actReportCsvGroups:
		return "Validate Create-Groups CSV"
	case actReportCsvApps:
		return "Validate App-Assignment CSV"
	case actSetClientID:
		return "Set Graph Client ID"
	case actSetTenantID:
		return "Set Graph Tenant ID"
	case actViewAuth:
		return "View Current Auth Config"
	case actAuthHealth:
		return "Auth Health"
	case actResetAuth:
		return "Reset Auth Defaults"
	case actToggleDryRun:
		return "Toggle Dry-Run Mode"
	default:
		return "Result"
	}
}

func (m model) resultSummaryView() string {
	rowCount := "n/a"
	if len(m.lastRows) > 0 {
		rowCount = fmt.Sprintf("%d", len(m.lastRows))
	}
	exportState := "unavailable"
	if len(m.lastHeaders) > 0 && len(m.lastRows) > 0 {
		exportState = "available"
	}
	mode := "LIVE"
	if m.dryRun {
		mode = "DRY-RUN"
	}

	cfg := m.client.Config()
	lines := []string{
		fmt.Sprintf("Action: %s", m.lastActionLabel),
		fmt.Sprintf("Tenant: %s", cfg.TenantID),
		fmt.Sprintf("Client: %s", cfg.ClientID),
		fmt.Sprintf("Mode: %s", mode),
		fmt.Sprintf("Rows: %s", rowCount),
		fmt.Sprintf("Export: %s", exportState),
	}
	return m.styles.panel.Render(strings.Join(lines, "\n"))
}

func (m *model) setPreview(text string) {
	m.output = text
	m.viewport.SetContent(text)
	m.viewport.GotoTop()
	m.lastHeaders = nil
	m.lastRows = nil
	if h, r, ok := render.ParseTableFromText(text); ok {
		m.lastHeaders = h
		m.lastRows = r
	}
	m.state = statePreview
}

func exportBaseDir() string {
	exe, err := os.Executable()
	if err != nil {
		cwd, cwdErr := os.Getwd()
		if cwdErr != nil {
			return "."
		}
		return cwd
	}
	return filepath.Dir(exe)
}

func slugifyName(s string) string {
	s = strings.ToLower(strings.TrimSpace(s))
	var b strings.Builder
	lastDash := false
	for _, r := range s {
		if (r >= 'a' && r <= 'z') || (r >= '0' && r <= '9') {
			b.WriteRune(r)
			lastDash = false
			continue
		}
		if !lastDash {
			b.WriteRune('-')
			lastDash = true
		}
	}
	out := strings.Trim(b.String(), "-")
	if out == "" {
		return "report"
	}
	return out
}

func (m model) defaultExportPath() string {
	name := slugifyName(m.lastActionLabel)
	if name == "" {
		name = "report"
	}
	stamp := time.Now().Format("20060102-150405")
	return filepath.Join(exportBaseDir(), fmt.Sprintf("%s-%s.csv", name, stamp))
}

func helpTextForState(state menuState) string {
	switch state {
	case stateMain, stateUsersGroups, stateDevicesApps, stateReports, stateReportCsv, stateReportInspect, stateSettings:
		return strings.Join([]string{
			"Menu Help",
			"",
			"Up/Down or j/k: Move selection",
			"PgUp/PgDn: Move by page",
			"Home/End: Jump to top or bottom",
			"1-9: Jump to item by number",
			"Enter: Select item",
			"/: Filter menu options",
			"Esc: Back, or quit from main menu",
			"q: Quit from main menu",
			"?: Open this help",
		}, "\n")
	case stateOutput:
		return strings.Join([]string{
			"Result Help",
			"",
			"Up/Down PgUp/PgDn Home/End: Scroll result",
			"/: Search within results",
			"n/N: Next/previous search match",
			"e: Export current table when available",
			"d: Drill into failing apps (Top 10 report only)",
			"Enter/Esc: Return to previous menu",
			"?: Open this help",
		}, "\n")
	case statePreview:
		return strings.Join([]string{
			"Preview Help",
			"",
			"Up/Down PgUp/PgDn Home/End: Scroll preview",
			"Enter: Continue to confirm",
			"Esc: Cancel and return",
			"?: Open this help",
		}, "\n")
	case stateDrillPrompt:
		return strings.Join([]string{
			"Drill-Down Help",
			"",
			"Enter a rank from the current top-failing report or an exact app name.",
			"Enter: Run drill-down",
			"Esc: Cancel",
		}, "\n")
	case stateWorking:
		return strings.Join([]string{
			"Working Help",
			"",
			"Spinner and progress text show current Graph activity.",
			"Esc: Cancel operation and return to menu",
			"Ctrl+C: Quit application",
			"?: Open this help",
		}, "\n")
	case stateConfirm:
		return strings.Join([]string{
			"Confirm Help",
			"",
			"y or Enter: Confirm",
			"n or Esc: Cancel",
			"?: Open this help",
		}, "\n")
	default:
		return strings.Join([]string{
			"Help",
			"",
			"Esc or Enter: Close help",
		}, "\n")
	}
}

func isWriteAction(id actionID) bool {
	switch id {
	case actAddUsersCSV, actMakeGroupsCSV, actAddAppsCSV, actSetClientID, actSetTenantID, actResetAuth:
		return true
	default:
		return false
	}
}

func confirmBodyForAction(spec actionSpec, inputs []string) string {
	switch spec.id {
	case actAddUsersCSV:
		return fmt.Sprintf("This will add users from CSV to a group.\n\nCSV: %s\nGroup: %s", safeInput(inputs, 0), safeInput(inputs, 1))
	case actMakeGroupsCSV:
		return fmt.Sprintf("This will create groups from CSV.\n\nCSV: %s", safeInput(inputs, 0))
	case actAddAppsCSV:
		return fmt.Sprintf("This will assign apps to groups from CSV.\n\nCSV: %s", safeInput(inputs, 0))
	case actSetClientID:
		return fmt.Sprintf("This will update and persist Graph Client ID.\n\nNew Client ID: %s", safeInput(inputs, 0))
	case actSetTenantID:
		return fmt.Sprintf("This will update and persist Graph Tenant ID.\n\nNew Tenant ID: %s", safeInput(inputs, 0))
	case actResetAuth:
		return "This will reset and persist auth defaults.\n\nClient ID: Graph PowerShell app\nTenant ID: common"
	default:
		return "This operation will modify data."
	}
}

func safeInput(inputs []string, idx int) string {
	if idx < 0 || idx >= len(inputs) {
		return ""
	}
	return inputs[idx]
}

func (m *model) startConfirmAction(spec actionSpec, inputs []string, cancelState menuState) {
	m.confirmKind = confirmAction
	m.confirmTitle = "Confirm Write Operation"
	mode := "LIVE mode: this will perform real writes."
	if m.dryRun {
		mode = "DRY-RUN mode: this will be simulated (no writes)."
	}
	m.confirmBody = mode + "\n\n" + confirmBodyForAction(spec, inputs)
	m.confirmCancelState = cancelState
	m.pendingSpec = spec
	m.pendingInputs = append([]string(nil), inputs...)
	m.state = stateConfirm
}

func (m *model) startConfirmExport(path string, cancelState menuState) {
	m.confirmKind = confirmExport
	m.confirmTitle = "Confirm File Write"
	mode := "LIVE mode: this will write a CSV file."
	if m.dryRun {
		mode = "DRY-RUN mode: export will be simulated."
	}
	m.confirmBody = mode + "\n\nPath: " + path
	if _, err := os.Stat(path); err == nil {
		m.confirmBody += "\n\n⚠ File already exists and will be overwritten."
	}
	m.confirmCancelState = cancelState
	m.pendingExportPath = path
	m.state = stateConfirm
}

func (m *model) clearConfirm() {
	m.confirmKind = confirmNone
	m.confirmTitle = ""
	m.confirmBody = ""
	m.confirmCancelState = stateOutput
	m.pendingSpec = actionSpec{}
	m.pendingInputs = nil
	m.pendingExportPath = ""
}

func waitProgressCmd(ch <-chan progressMsg, done <-chan struct{}) tea.Cmd {
	return func() tea.Msg {
		select {
		case msg := <-ch:
			return msg
		case <-done:
			return progressStopMsg{}
		}
	}
}

func (m *model) startWorking() {
	if m.progressActive && m.progressDone != nil {
		close(m.progressDone)
	}
	if m.cancelFunc != nil {
		m.cancelFunc()
	}
	for {
		select {
		case <-m.progressCh:
		default:
			goto drained
		}
	}
drained:
	ctx, cancel := context.WithTimeout(context.Background(), 5*time.Minute)
	m.cancelCtx = ctx
	m.cancelFunc = cancel
	m.progressDone = make(chan struct{})
	m.progressActive = true
	m.state = stateWorking
	m.output = ""
	m.progressText = "Starting operation..."
}

func (m *model) stopWorking() {
	if m.progressActive && m.progressDone != nil {
		close(m.progressDone)
	}
	if m.cancelFunc != nil {
		m.cancelFunc()
		m.cancelFunc = nil
	}
	m.progressDone = nil
	m.progressActive = false
	m.progressText = ""
}

func specForAction(id actionID) actionSpec {
	switch id {
	case actListUsersGroup:
		return actionSpec{id: id, prompts: []string{"Enter group display name"}}
	case actSearchGroups:
		return actionSpec{id: id, prompts: []string{"Enter group search term"}}
	case actAddUsersCSV:
		return actionSpec{id: id, prompts: []string{"Enter CSV path (header: User_Principal_Name)", "Enter group display name"}}
	case actListDevicesGroup:
		return actionSpec{id: id, prompts: []string{"Enter group display name"}}
	case actMakeGroupsCSV:
		return actionSpec{id: id, prompts: []string{"Enter CSV path (header: Group_Name)"}}
	case actAddAppsCSV:
		return actionSpec{id: id, prompts: []string{"Enter CSV path (headers: Group_Name, App_Name)"}}
	case actReportCsvUsers:
		return actionSpec{id: id, prompts: []string{"Enter CSV path for Users->Group validation"}}
	case actReportCsvGroups:
		return actionSpec{id: id, prompts: []string{"Enter CSV path for Create-Groups validation"}}
	case actReportCsvApps:
		return actionSpec{id: id, prompts: []string{"Enter CSV path for App-Assignment validation"}}
	case actInspectUser:
		return actionSpec{id: id, prompts: []string{"Enter user object ID, UPN, or exact display name"}}
	case actInspectGroup:
		return actionSpec{id: id, prompts: []string{"Enter group object ID or exact display name"}}
	case actInspectDevice:
		return actionSpec{id: id, prompts: []string{"Enter device object ID or exact display name"}}
	case actInspectApp:
		return actionSpec{id: id, prompts: []string{"Enter app ID or exact display name"}}
	case actSetClientID:
		return actionSpec{id: id, prompts: []string{"Enter Graph client ID"}}
	case actSetTenantID:
		return actionSpec{id: id, prompts: []string{"Enter Graph tenant ID (GUID/domain/common)"}}
	default:
		return actionSpec{id: id}
	}
}

func (m model) authSummary() string {
	mode := "OFF"
	if m.dryRun {
		mode = "ON"
	}
	cfg := m.client.Config()
	return fmt.Sprintf("Client ID: %s\nTenant ID: %s\nDry-Run: %s", cfg.ClientID, cfg.TenantID, mode)
}

func (m *model) applyAuthConfig(cfg config.AuthConfig) error {
	client, err := graph.NewClientWithConfig(cfg)
	if err != nil {
		return err
	}
	if err := config.SaveToFile(cfg); err != nil {
		return err
	}
	m.client = client
	return nil
}

func (m *model) runLocalAction(id actionID, inputs []string) (string, error, bool) {
	switch id {
	case actViewAuth:
		return m.authSummary(), nil, true
	case actToggleDryRun:
		m.dryRun = !m.dryRun
		mode := "OFF"
		if m.dryRun {
			mode = "ON"
		}
		return "Dry-run mode is now " + mode + ".", nil, true
	case actResetAuth:
		if m.dryRun {
			return "Dry-run: would reset auth defaults.\n\n" + m.authSummary(), nil, true
		}
		cfg := config.AuthConfig{ClientID: config.DefaultClientID, TenantID: "common"}
		if err := m.applyAuthConfig(cfg); err != nil {
			return "", err, true
		}
		return "Auth config reset.\n\n" + m.authSummary(), nil, true
	case actSetClientID:
		clientID := strings.TrimSpace(inputs[0])
		if clientID == "" {
			return "", errors.New("client ID cannot be empty"), true
		}
		cfg := m.client.Config()
		cfg.ClientID = clientID
		if m.dryRun {
			return "Dry-run: would update Graph client ID to:\n" + clientID, nil, true
		}
		if err := m.applyAuthConfig(cfg); err != nil {
			return "", err, true
		}
		return "Updated Graph client ID.\n\n" + m.authSummary(), nil, true
	case actSetTenantID:
		tenantID := strings.TrimSpace(inputs[0])
		if tenantID == "" {
			return "", errors.New("tenant ID cannot be empty"), true
		}
		cfg := m.client.Config()
		cfg.TenantID = tenantID
		if m.dryRun {
			return "Dry-run: would update Graph tenant ID to:\n" + tenantID, nil, true
		}
		if err := m.applyAuthConfig(cfg); err != nil {
			return "", err, true
		}
		return "Updated Graph tenant ID.\n\n" + m.authSummary(), nil, true
	default:
		return "", nil, false
	}
}

func (m model) runActionCmd(spec actionSpec, inputs []string) tea.Cmd {
	return func() tea.Msg {
		ctx := m.cancelCtx
		m.client.SetProgressHook(func(text string) {
			select {
			case m.progressCh <- progressMsg{text: text}:
			default:
			}
		})
		defer m.client.SetProgressHook(nil)
		var (
			out string
			err error
		)
		switch spec.id {
		case actListUsersGroup:
			out, err = m.client.ListUsersInGroup(ctx, inputs[0])
		case actListGroups:
			out, err = m.client.ListGroups(ctx)
		case actListUsers:
			out, err = m.client.ListUsers(ctx)
		case actSearchGroups:
			out, err = m.client.SearchGroups(ctx, inputs[0])
		case actAuthHealth:
			out, err = m.client.AuthHealth(ctx)
		case actAddUsersCSV:
			out, err = m.client.AddUsersCSV(ctx, inputs[0], inputs[1], m.dryRun)
		case actListDevices:
			out, err = m.client.ListDevices(ctx)
		case actReportComplianceSnapshot:
			out, err = m.client.ReportComplianceSnapshot(ctx)
		case actReportWindowsBreakdown:
			out, err = m.client.ReportWindowsBreakdown(ctx)
		case actReportTopFailingApps:
			out, err = m.client.ReportTopFailingApps(ctx)
		case actInspectUser:
			out, err = m.client.InspectUser(ctx, inputs[0])
		case actInspectGroup:
			out, err = m.client.InspectGroup(ctx, inputs[0])
		case actInspectDevice:
			out, err = m.client.InspectDevice(ctx, inputs[0])
		case actInspectApp:
			out, err = m.client.InspectApp(ctx, inputs[0])
		case actListDevicesGroup:
			out, err = m.client.ListDevicesInGroup(ctx, inputs[0])
		case actMakeGroupsCSV:
			out, err = m.client.MakeGroupsCSV(ctx, inputs[0], m.dryRun)
		case actAddAppsCSV:
			out, err = m.client.AddAppsCSV(ctx, inputs[0], m.dryRun)
		case actListGroupApps:
			out, err = m.client.ListGroupApps(ctx)
		case actReportCsvUsers:
			if res, ok, verr := validateForAction(spec, inputs); verr != nil {
				err = verr
			} else if ok {
				out = csvutil.FormatValidationReport("CSV Quality Report - Users to Group", res)
			}
		case actReportCsvGroups:
			if res, ok, verr := validateForAction(spec, inputs); verr != nil {
				err = verr
			} else if ok {
				out = csvutil.FormatValidationReport("CSV Quality Report - Create Groups", res)
			}
		case actReportCsvApps:
			if res, ok, verr := validateForAction(spec, inputs); verr != nil {
				err = verr
			} else if ok {
				out = csvutil.FormatValidationReport("CSV Quality Report - App Assignments", res)
			}
		default:
			out = "No action."
		}
		return resultMsg{text: out, err: err}
	}
}

func (m model) Update(msg tea.Msg) (tea.Model, tea.Cmd) {
	switch msg := msg.(type) {
	case tea.WindowSizeMsg:
		m.width = msg.Width
		m.height = msg.Height
		m.input.Width = min(72, max(24, msg.Width-16))
		m.exportInput.Width = min(72, max(24, msg.Width-16))
		m.drillInput.Width = min(72, max(24, msg.Width-16))
		panelInnerW := max(20, msg.Width-12)
		panelInnerH := max(6, msg.Height-18)
		m.viewport.Width = panelInnerW
		m.viewport.Height = panelInnerH
		if m.output != "" {
			m.viewport.SetContent(m.output)
		}
		m.vpReady = true
		return m, nil
	case tea.KeyMsg:
		switch m.state {
		case stateMain, stateUsersGroups, stateDevicesApps, stateReports, stateReportCsv, stateReportInspect, stateSettings:
			visible := m.visibleMenu()
			switch msg.String() {
			case "?":
				m.helpReturnState = m.state
				m.state = stateHelp
				return m, nil
			case "ctrl+c":
				return m, tea.Quit
			case "q":
				if m.state == stateMain {
					return m, tea.Quit
				}
			case "esc":
				if m.state != stateMain {
					m.resetMenuPosition(parentMenuState(m.state))
				}
			case "/":
				m.lastMenuState = m.state
				m.filterInput.SetValue(m.filterQuery)
				m.filterInput.CursorEnd()
				m.filterInput.Focus()
				m.input.Blur()
				m.exportInput.Blur()
				m.state = stateMenuFilter
				return m, textinput.Blink
			case "up", "k":
				if m.cursor > 0 {
					m.cursor--
					m.ensureMenuCursorVisible()
				}
			case "down", "j":
				if m.cursor < len(visible)-1 {
					m.cursor++
					m.ensureMenuCursorVisible()
				}
			case "pgup":
				page := m.menuPageSize()
				m.cursor = max(0, m.cursor-page)
				m.ensureMenuCursorVisible()
			case "pgdown":
				page := m.menuPageSize()
				m.cursor = min(max(0, len(visible)-1), m.cursor+page)
				m.ensureMenuCursorVisible()
			case "home":
				m.cursor = 0
				m.menuTop = 0
			case "end":
				m.cursor = max(0, len(visible)-1)
				m.ensureMenuCursorVisible()
			case "enter":
				if len(visible) == 0 {
					return m, nil
				}
				item := visible[m.cursor].item
				if m.state == stateMain && item.label == "Quit" {
					return m, tea.Quit
				}
				if item.action != actNone {
					spec := specForAction(item.action)
					m.lastActionLabel = actionLabel(spec.id)
					m.lastActionID = spec.id
					m.lastMenuState = m.state
					if len(spec.prompts) == 0 {
						if isWriteAction(spec.id) {
							m.startConfirmAction(spec, nil, m.lastMenuState)
							return m, nil
						}
						if out, err, handled := m.runLocalAction(spec.id, nil); handled {
							if err != nil {
								m.setOutput("Error:\n" + err.Error())
							} else {
								m.setOutput(out)
							}
							return m, nil
						}
						m.startWorking()
						return m, tea.Batch(m.spin.Tick, waitProgressCmd(m.progressCh, m.progressDone), m.runActionCmd(spec, nil))
					}
					m.currentSpec = spec
					m.inputs = nil
					m.input.SetValue("")
					m.input.Prompt = spec.prompts[0] + ": "
					m.input.Focus()
					m.filterInput.Blur()
					m.exportInput.Blur()
					m.state = stateInput
					return m, textinput.Blink
				}
				if item.next == stateMain {
					m.resetMenuPosition(stateMain)
					return m, nil
				}
				if item.next == stateUsersGroups || item.next == stateDevicesApps || item.next == stateReports || item.next == stateReportCsv || item.next == stateReportInspect || item.next == stateSettings {
					m.resetMenuPosition(item.next)
					return m, nil
				}
			default:
				if k := msg.String(); len(k) == 1 && k[0] >= '1' && k[0] <= '9' {
					idx := int(k[0]-'0') - 1
					if idx < len(visible) {
						m.cursor = idx
						m.ensureMenuCursorVisible()
						return m.Update(tea.KeyMsg(tea.Key{Type: tea.KeyEnter}))
					}
				}
			}
		case stateMenuFilter:
			switch msg.String() {
			case "esc":
				m.filterInput.Blur()
				m.state = m.lastMenuState
				return m, nil
			case "enter":
				m.filterQuery = strings.TrimSpace(m.filterInput.Value())
				m.cursor = 0
				m.menuTop = 0
				m.filterInput.Blur()
				m.state = m.lastMenuState
				m.ensureMenuCursorVisible()
				return m, nil
			}
			var cmd tea.Cmd
			m.filterInput, cmd = m.filterInput.Update(msg)
			return m, cmd
		case stateInput:
			switch msg.String() {
			case "esc":
				if len(m.inputs) > 0 {
					prev := m.inputs[len(m.inputs)-1]
					m.inputs = m.inputs[:len(m.inputs)-1]
					m.input.SetValue(prev)
					m.input.Prompt = m.currentSpec.prompts[len(m.inputs)] + ": "
					m.input.CursorEnd()
					return m, nil
				}
				m.input.Blur()
				m.returnToLastMenu()
				return m, nil
			case "enter":
				val := strings.TrimSpace(m.input.Value())
				m.inputs = append(m.inputs, val)
				prevPrompt := m.currentSpec.prompts[len(m.inputs)-1]
				if strings.Contains(strings.ToLower(prevPrompt), "csv path") {
					if _, err := os.Stat(val); err != nil {
						m.inputs = m.inputs[:len(m.inputs)-1]
						m.setOutput("Error:\nFile not found: " + val)
						return m, nil
					}
				}
				if len(m.inputs) < len(m.currentSpec.prompts) {
					m.input.SetValue("")
					m.input.Prompt = m.currentSpec.prompts[len(m.inputs)] + ": "
					m.input.Focus()
					return m, nil
				}
				m.input.Blur()
				if res, ok, verr := validateForAction(m.currentSpec, m.inputs); ok {
					if verr != nil {
						m.setOutput("Error:\nCSV validation failed to run: " + verr.Error())
						return m, nil
					}
					if !res.Pass {
						m.setOutput(csvutil.FormatValidationReport("CSV Validation Failed (Write Blocked)", res))
						return m, nil
					}
				}
				if preview, ok, perr := previewForAction(m.currentSpec, m.inputs); ok {
					if perr != nil {
						m.setOutput("Error:\nPreview failed: " + perr.Error())
						return m, nil
					}
					m.pendingSpec = m.currentSpec
					m.pendingInputs = append([]string(nil), m.inputs...)
					m.setPreview(preview)
					return m, nil
				}
				if isWriteAction(m.currentSpec.id) {
					m.startConfirmAction(m.currentSpec, m.inputs, stateInput)
					return m, nil
				}
				if out, err, handled := m.runLocalAction(m.currentSpec.id, m.inputs); handled {
					if err != nil {
						m.setOutput("Error:\n" + err.Error())
					} else {
						m.setOutput(out)
					}
					return m, nil
				}
				m.startWorking()
				return m, tea.Batch(m.spin.Tick, waitProgressCmd(m.progressCh, m.progressDone), m.runActionCmd(m.currentSpec, m.inputs))
			}
			var cmd tea.Cmd
			m.input, cmd = m.input.Update(msg)
			return m, cmd
		case statePreview:
			switch msg.String() {
			case "?":
				m.helpReturnState = m.state
				m.state = stateHelp
				return m, nil
			case "esc":
				m.returnToLastMenu()
				m.output = ""
				m.viewport.SetContent("")
				m.viewport.GotoTop()
				return m, nil
			case "enter":
				spec := m.pendingSpec
				inputs := append([]string(nil), m.pendingInputs...)
				m.startConfirmAction(spec, inputs, statePreview)
				return m, nil
			}
			var cmd tea.Cmd
			m.viewport, cmd = m.viewport.Update(msg)
			return m, cmd
		case stateExportPrompt:
			switch msg.String() {
			case "esc":
				m.exportInput.Blur()
				m.state = stateOutput
				return m, nil
			case "enter":
				path := strings.TrimSpace(m.exportInput.Value())
				if path == "" {
					m.setOutput("Error:\nExport path cannot be empty.")
					return m, nil
				}
				if len(m.lastHeaders) == 0 || len(m.lastRows) == 0 {
					m.setOutput("Error:\nNo table data available to export.")
					return m, nil
				}
				if dir := filepath.Dir(path); dir != "." {
					if _, err := os.Stat(dir); err != nil {
						m.setOutput("Error:\nDirectory does not exist: " + dir)
						return m, nil
					}
				}
				m.exportInput.Blur()
				m.startConfirmExport(path, stateExportPrompt)
				return m, nil
			}
			var cmd tea.Cmd
			m.exportInput, cmd = m.exportInput.Update(msg)
			return m, cmd
		case stateDrillPrompt:
			switch msg.String() {
			case "esc":
				m.drillInput.Blur()
				m.state = stateOutput
				return m, nil
			case "enter":
				appName, err := resolveTopFailingAppSelection(m.drillInput.Value(), m.lastRows)
				if err != nil {
					m.setOutput("Error:\n" + err.Error())
					return m, nil
				}
				m.drillInput.Blur()
				m.lastActionLabel = "Failing App Drill-Down"
				m.lastActionID = actNone
				m.startWorking()
				ctx := m.cancelCtx
				return m, tea.Batch(m.spin.Tick, waitProgressCmd(m.progressCh, m.progressDone), func() tea.Msg {
					out, err := m.client.ReportAppFailureDetails(ctx, appName)
					return resultMsg{text: out, err: err}
				})
			}
			var cmd tea.Cmd
			m.drillInput, cmd = m.drillInput.Update(msg)
			return m, cmd
		case stateConfirm:
			switch msg.String() {
			case "?":
				m.helpReturnState = m.state
				m.state = stateHelp
				return m, nil
			case "n", "esc":
				cancelState := m.confirmCancelState
				m.clearConfirm()
				m.state = cancelState
				if cancelState == stateExportPrompt {
					m.exportInput.Focus()
					return m, textinput.Blink
				}
				return m, nil
			case "y", "enter":
				switch m.confirmKind {
				case confirmAction:
					spec := m.pendingSpec
					inputs := append([]string(nil), m.pendingInputs...)
					m.clearConfirm()
					if out, err, handled := m.runLocalAction(spec.id, inputs); handled {
						if err != nil {
							m.setOutput("Error:\n" + err.Error())
						} else {
							m.setOutput(out)
						}
						return m, nil
					}
					m.startWorking()
					return m, tea.Batch(m.spin.Tick, waitProgressCmd(m.progressCh, m.progressDone), m.runActionCmd(spec, inputs))
				case confirmExport:
					path := m.pendingExportPath
					m.clearConfirm()
					if m.dryRun {
						m.setOutput(m.output + "\n\nDry-run: would export CSV to " + path)
						return m, nil
					}
					if err := render.ExportCSV(path, m.lastHeaders, m.lastRows); err != nil {
						m.setOutput("Error:\nFailed to export CSV: " + err.Error())
						return m, nil
					}
					m.setOutput(m.output + "\n\nExported CSV: " + path)
					return m, nil
				default:
					m.clearConfirm()
					m.returnToLastMenu()
					return m, nil
				}
			}
			return m, nil
		case stateWorking:
			switch msg.String() {
			case "?":
				m.helpReturnState = m.state
				m.state = stateHelp
				return m, nil
			case "esc":
				m.stopWorking()
				m.setOutput("Operation cancelled.")
				return m, nil
			case "ctrl+c":
				return m, tea.Quit
			}
		case stateHelp:
			switch msg.String() {
			case "enter", "esc", "?":
				m.state = m.helpReturnState
				return m, nil
			}
		case stateOutput:
			switch msg.String() {
			case "?":
				m.helpReturnState = m.state
				m.state = stateHelp
				return m, nil
			case "enter", "esc":
				m.returnToLastMenu()
				m.output = ""
				m.searchQuery = ""
				m.searchMatchLine = -1
				m.lastHeaders = nil
				m.lastRows = nil
				m.viewport.SetContent("")
				m.viewport.GotoTop()
				return m, nil
			case "/":
				m.searchInput.SetValue(m.searchQuery)
				m.searchInput.CursorEnd()
				m.searchInput.Focus()
				m.state = stateOutputSearch
				return m, textinput.Blink
			case "n":
				if m.searchQuery != "" {
					m.jumpToNextMatch(true)
				}
				return m, nil
			case "N":
				if m.searchQuery != "" {
					m.jumpToNextMatch(false)
				}
				return m, nil
			case "e":
				if len(m.lastHeaders) == 0 || len(m.lastRows) == 0 {
					return m, nil
				}
				m.exportInput.SetValue(m.defaultExportPath())
				m.exportInput.CursorEnd()
				m.exportInput.Focus()
				m.filterInput.Blur()
				m.input.Blur()
				m.state = stateExportPrompt
				return m, textinput.Blink
			case "d":
				if m.lastActionID != actReportTopFailingApps || len(m.lastRows) == 0 {
					return m, nil
				}
				m.drillInput.SetValue("")
				m.drillInput.Focus()
				m.state = stateDrillPrompt
				return m, textinput.Blink
			}
			var cmd tea.Cmd
			m.viewport, cmd = m.viewport.Update(msg)
			return m, cmd
		case stateOutputSearch:
			switch msg.String() {
			case "esc":
				m.searchInput.Blur()
				m.state = stateOutput
				return m, nil
			case "enter":
				m.searchQuery = strings.TrimSpace(m.searchInput.Value())
				m.searchInput.Blur()
				m.searchMatchLine = -1
				if m.searchQuery != "" {
					m.jumpToNextMatch(true)
				} else {
					m.viewport.SetContent(m.output)
				}
				m.state = stateOutput
				return m, nil
			}
			var cmd tea.Cmd
			m.searchInput, cmd = m.searchInput.Update(msg)
			return m, cmd
		}
	case spinner.TickMsg:
		if m.state == stateWorking {
			var cmd tea.Cmd
			m.spin, cmd = m.spin.Update(msg)
			return m, cmd
		}
	case progressMsg:
		if m.state == stateWorking && m.progressActive {
			m.progressText = msg.text
			return m, waitProgressCmd(m.progressCh, m.progressDone)
		}
	case progressStopMsg:
		return m, nil
	case resultMsg:
		m.stopWorking()
		if msg.err != nil {
			m.setOutput("Error:\n" + msg.err.Error())
		} else {
			m.setOutput(msg.text)
		}
		return m, nil
	}
	return m, nil
}

func (m model) dryRunBanner() string {
	if !m.dryRun {
		return ""
	}
	return lipgloss.NewStyle().Bold(true).Foreground(lipgloss.Color("230")).Background(lipgloss.Color("208")).Padding(0, 1).Render("⚠ DRY-RUN MODE") + "\n\n"
}

func (m model) View() string {
	banner := m.dryRunBanner()
	switch m.state {
	case stateDrillPrompt:
		body := m.styles.panel.Render(fmt.Sprintf("%s\n\n%s\n\n%s",
			m.styles.subHeader.Render("Top Failing Apps Drill-Down"),
			m.drillInput.View(),
			m.styles.hint.Render("Enter rank or exact app name   Esc: cancel"),
		))
		return m.styles.app.Render(banner + m.styles.header.Render(" Intune Management Tool ") + "\n\n" + body)
	case stateHelp:
		body := m.styles.panel.Render(fmt.Sprintf("%s\n\n%s\n\n%s",
			m.styles.subHeader.Render("Keyboard Help"),
			helpTextForState(m.helpReturnState),
			m.styles.hint.Render("Enter/Esc/?: close help"),
		))
		return m.styles.app.Render(banner + m.styles.header.Render(" Intune Management Tool ") + "\n\n" + body)
	case stateConfirm:
		body := m.styles.panel.Render(fmt.Sprintf("%s\n\n%s\n\n%s",
			m.styles.subHeader.Render(m.confirmTitle),
			m.confirmBody,
			m.styles.hint.Render("y/Enter: confirm   n/Esc: cancel"),
		))
		return m.styles.app.Render(banner + m.styles.header.Render(" Intune Management Tool ") + "\n\n" + body)
	case stateMenuFilter:
		body := m.styles.panel.Render(fmt.Sprintf("%s\n\n%s\n\n%s",
			m.styles.subHeader.Render("Filter Menu Options"),
			m.filterInput.View(),
			m.styles.hint.Render("Enter: apply filter   Esc: cancel"),
		))
		return m.styles.app.Render(banner + m.styles.header.Render(" Intune Management Tool ") + "\n\n" + body)
	case stateExportPrompt:
		body := m.styles.panel.Render(fmt.Sprintf("%s\n\n%s\n\n%s",
			m.styles.subHeader.Render("Export Current Table to CSV"),
			m.exportInput.View(),
			m.styles.hint.Render("Enter: export   Esc: cancel"),
		))
		return m.styles.app.Render(banner + m.styles.header.Render(" Intune Management Tool ") + "\n\n" + body)
	case stateInput:
		step := len(m.inputs) + 1
		total := len(m.currentSpec.prompts)
		title := m.styles.subHeader.Render(fmt.Sprintf("Input %d/%d", step, total))
		escHint := "Esc: cancel"
		if len(m.inputs) > 0 {
			escHint = "Esc: back"
		}
		body := m.styles.panel.Render(fmt.Sprintf("%s\n\n%s\n\n%s",
			title,
			m.input.View(),
			m.styles.hint.Render("Enter: continue   "+escHint),
		))
		return m.styles.app.Render(banner + m.styles.header.Render(" Intune Management Tool ") + "\n\n" + body)
	case statePreview:
		prefix := m.styles.subHeader.Render("Preview Before Write")
		content := m.output
		if m.vpReady {
			content = m.viewport.View()
		}
		headerPanel := m.resultSummaryView()
		body := m.styles.panel.Render(fmt.Sprintf("%s\n\n%s\n\n%s",
			prefix,
			content,
			m.styles.hint.Render("Up/Down PgUp/PgDn Home/End: scroll   Enter: continue   Esc: cancel"),
		))
		return m.styles.app.Render(banner + m.styles.header.Render(" Intune Management Tool ") + "\n\n" + headerPanel + "\n\n" + body)
	case stateWorking:
		progress := m.progressText
		if strings.TrimSpace(progress) == "" {
			progress = "Starting operation..."
		}
		body := m.styles.panel.Render(fmt.Sprintf("%s Running Graph operation...\n\n%s\n\n%s",
			m.spin.View(),
			m.styles.hint.Render(progress),
			m.styles.hint.Render("Esc: cancel"),
		))
		return m.styles.app.Render(banner + m.styles.header.Render(" Intune Management Tool ") + "\n\n" + body)
	case stateOutput:
		prefix := m.styles.ok.Render("Result")
		if strings.HasPrefix(m.output, "Error:") {
			prefix = m.styles.err.Render("Result")
		}
		content := m.output
		if m.vpReady {
			content = m.viewport.View()
		}
		exportHint := "/: search   Enter/Esc: return"
		if len(m.lastHeaders) > 0 && len(m.lastRows) > 0 {
			exportHint = "/: search   e: export table   Enter/Esc: return"
		}
		if m.lastActionID == actReportTopFailingApps {
			exportHint = "/: search   d: drill down   e: export table   Enter/Esc: return"
		}
		if m.searchQuery != "" {
			exportHint = fmt.Sprintf("/%s (n/N: next/prev)   ", m.searchQuery) + exportHint
		}
		headerPanel := m.resultSummaryView()
		body := m.styles.panel.Render(fmt.Sprintf("%s\n\n%s\n\n%s",
			prefix,
			content,
			m.styles.hint.Render(exportHint),
		))
		return m.styles.app.Render(banner + m.styles.header.Render(" Intune Management Tool ") + "\n\n" + headerPanel + "\n\n" + body)
	case stateOutputSearch:
		content := m.output
		if m.vpReady {
			content = m.viewport.View()
		}
		body := m.styles.panel.Render(fmt.Sprintf("%s\n\n%s\n\n%s",
			m.styles.subHeader.Render("Search Results"),
			content,
			m.searchInput.View(),
		))
		return m.styles.app.Render(banner + m.styles.header.Render(" Intune Management Tool ") + "\n\n" + body)
	default:
		title := "Main Menu"
		sub := "Pick an operation area"
		if m.state == stateUsersGroups {
			title = "Users and Groups"
			sub = "Identity membership and group operations"
		} else if m.state == stateDevicesApps {
			title = "Devices and Applications"
			sub = "Device inventory and Intune app assignment workflows"
		} else if m.state == stateReports {
			title = "Reports"
			sub = "Read-only analytics and strict CSV quality checks"
		} else if m.state == stateReportCsv {
			title = "Reports - CSV Validation"
			sub = "Strict schema and data-quality validation for CSV workflows"
		} else if m.state == stateReportInspect {
			title = "Reports - Object Inspector"
			sub = "Read-only lookup for users, groups, devices, and apps"
		} else if m.state == stateSettings {
			title = "Settings"
			sub = "Authentication configuration for Microsoft Graph"
		}
		menuView := m.renderMenu()
		filterLine := ""
		if m.filterQuery != "" {
			filterLine = "\n" + m.styles.hint.Render("Active filter: "+m.filterQuery+" (press / to edit)")
		}
		screen := banner + m.styles.header.Render(" Intune Management Tool ") + "\n\n" +
			m.styles.subHeader.Render(title) + "\n" +
			m.styles.hint.Render(sub) + "\n\n" +
			m.styles.panel.Render(menuView) + filterLine + "\n\n" +
			m.styles.hint.Render("Arrows/jk 1-9: move   /: filter   Enter: select   Esc: back   q: quit")
		return m.styles.app.Render(screen)
	}
}

func (m model) renderMenu() string {
	var lines []string
	visible := m.visibleMenu()
	if len(visible) == 0 {
		return m.styles.hint.Render("No matching menu options.")
	}
	page := m.menuPageSize()
	start := m.menuTop
	if start > len(visible)-1 {
		start = max(0, len(visible)-page)
	}
	end := min(len(visible), start+page)

	if start > 0 {
		lines = append(lines, m.styles.hint.Render("... more above ..."))
	}
	for i := start; i < end; i++ {
		item := visible[i].item
		entry := fmt.Sprintf("%d. %s", i+1, item.label)
		if i == m.cursor {
			lines = append(lines, m.styles.selected.Render("> "+entry))
		} else {
			lines = append(lines, m.styles.normalItem.Render("  "+entry))
		}
		if item.description != "" {
			lines = append(lines, "   "+m.styles.desc.Render(item.description))
		}
	}
	if end < len(visible) {
		lines = append(lines, m.styles.hint.Render("... more below ..."))
	}
	lines = append(lines, m.styles.hint.Render(fmt.Sprintf("Showing %d-%d of %d", start+1, end, len(visible))))
	return strings.Join(lines, "\n")
}

// Run starts the TUI application.
func Run() error {
	client, err := graph.NewClient()
	if err != nil {
		return err
	}
	p := tea.NewProgram(newModel(client))
	if _, err := p.Run(); err != nil {
		return err
	}
	return nil
}
