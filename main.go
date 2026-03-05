package main

import (
	"bytes"
	"context"
	"encoding/base64"
	"encoding/csv"
	"encoding/json"
	"errors"
	"fmt"
	"io"
	"net/http"
	"net/url"
	"os"
	"path/filepath"
	"strconv"
	"strings"
	"time"

	"github.com/Azure/azure-sdk-for-go/sdk/azcore/policy"
	"github.com/Azure/azure-sdk-for-go/sdk/azidentity"
	"github.com/charmbracelet/bubbles/spinner"
	"github.com/charmbracelet/bubbles/textinput"
	"github.com/charmbracelet/bubbles/viewport"
	tea "github.com/charmbracelet/bubbletea"
	"github.com/charmbracelet/lipgloss"
)

var requiredScopes = []string{
	"https://graph.microsoft.com/User.Read.All",
	"https://graph.microsoft.com/Group.ReadWrite.All",
	"https://graph.microsoft.com/Device.Read.All",
	"https://graph.microsoft.com/DeviceManagementApps.ReadWrite.All",
	"https://graph.microsoft.com/DeviceManagementManagedDevices.Read.All",
}

const (
	defaultClientID = "14d82eec-204b-4c2f-b7e8-296a70dab67e" // Graph PowerShell public client
	graphBase       = "https://graph.microsoft.com/v1.0"
)

type graphClient struct {
	cred         *azidentity.DeviceCodeCredential
	http         *http.Client
	scope        []string
	cfg          authConfig
	progressHook func(string)
}

type authConfig struct {
	ClientID string
	TenantID string
}

func configFilePath() string {
	exe, err := os.Executable()
	if err != nil {
		return "intune-management.config.json"
	}
	return filepath.Join(filepath.Dir(exe), "intune-management.config.json")
}

func loadAuthConfigFromFile() (authConfig, error) {
	path := configFilePath()
	b, err := os.ReadFile(path)
	if err != nil {
		return authConfig{}, err
	}
	var cfg authConfig
	if err := json.Unmarshal(b, &cfg); err != nil {
		return authConfig{}, err
	}
	return cfg, nil
}

func saveAuthConfigToFile(cfg authConfig) error {
	path := configFilePath()
	b, err := json.MarshalIndent(cfg, "", "  ")
	if err != nil {
		return err
	}
	return os.WriteFile(path, b, 0644)
}

func resolveAuthConfig() authConfig {
	cfg := authConfig{}
	if fileCfg, err := loadAuthConfigFromFile(); err == nil {
		cfg = fileCfg
	}
	envClient := os.Getenv("GRAPH_CLIENT_ID")
	envTenant := os.Getenv("GRAPH_TENANT_ID")
	if envClient != "" {
		cfg.ClientID = envClient
	}
	if envTenant != "" {
		cfg.TenantID = envTenant
	}
	if cfg.ClientID == "" {
		cfg.ClientID = defaultClientID
	}
	if cfg.TenantID == "" {
		cfg.TenantID = "common"
	}
	return cfg
}

func newGraphClient() (*graphClient, error) {
	return newGraphClientWithConfig(resolveAuthConfig())
}

func newGraphClientWithConfig(cfg authConfig) (*graphClient, error) {
	cred, err := azidentity.NewDeviceCodeCredential(&azidentity.DeviceCodeCredentialOptions{
		ClientID: cfg.ClientID,
		TenantID: cfg.TenantID,
		UserPrompt: func(ctx context.Context, message azidentity.DeviceCodeMessage) error {
			fmt.Printf("\n%s\n\n", message.Message)
			return nil
		},
	})
	if err != nil {
		return nil, err
	}

	return &graphClient{
		cred:  cred,
		http:  &http.Client{Timeout: 60 * time.Second},
		scope: requiredScopes,
		cfg:   cfg,
	}, nil
}

func (g *graphClient) do(ctx context.Context, method, fullURL string, body any) ([]byte, error) {
	token, err := g.cred.GetToken(ctx, policy.TokenRequestOptions{Scopes: g.scope})
	if err != nil {
		return nil, err
	}

	var payload []byte
	if body != nil {
		payload, err = json.Marshal(body)
		if err != nil {
			return nil, err
		}
	}

	const maxRetries = 4
	for attempt := 0; attempt <= maxRetries; attempt++ {
		var reader io.Reader
		if len(payload) > 0 {
			reader = bytes.NewReader(payload)
		}
		req, err := http.NewRequestWithContext(ctx, method, fullURL, reader)
		if err != nil {
			return nil, err
		}
		req.Header.Set("Authorization", "Bearer "+token.Token)
		req.Header.Set("Accept", "application/json")
		if len(payload) > 0 {
			req.Header.Set("Content-Type", "application/json")
		}

		resp, err := g.http.Do(req)
		if err != nil {
			if attempt == maxRetries {
				return nil, err
			}
			wait := retryDelay(attempt, "")
			g.emitProgress(fmt.Sprintf("Transient network error; retrying in %s (%d/%d)", wait, attempt+1, maxRetries))
			time.Sleep(wait)
			continue
		}

		raw, readErr := io.ReadAll(resp.Body)
		resp.Body.Close()
		if readErr != nil {
			return nil, readErr
		}

		if resp.StatusCode < 400 {
			return raw, nil
		}

		if shouldRetryStatus(resp.StatusCode) && attempt < maxRetries {
			wait := retryDelay(attempt, resp.Header.Get("Retry-After"))
			g.emitProgress(fmt.Sprintf("Graph %d received; retrying in %s (%d/%d)", resp.StatusCode, wait, attempt+1, maxRetries))
			time.Sleep(wait)
			continue
		}
		return nil, formatGraphError(method, fullURL, resp.Status, raw)
	}

	return nil, errors.New("graph request failed after retries")
}

type pageResponse struct {
	Value    []map[string]any `json:"value"`
	NextLink string           `json:"@odata.nextLink"`
}

type graphErrorEnvelope struct {
	Error struct {
		Code    string `json:"code"`
		Message string `json:"message"`
	} `json:"error"`
}

func formatGraphError(method, fullURL, status string, raw []byte) error {
	var env graphErrorEnvelope
	if err := json.Unmarshal(raw, &env); err == nil && env.Error.Code != "" {
		return fmt.Errorf("graph %s %s failed: %s | %s: %s", method, fullURL, status, env.Error.Code, env.Error.Message)
	}
	msg := strings.TrimSpace(string(raw))
	if msg == "" {
		return fmt.Errorf("graph %s %s failed: %s", method, fullURL, status)
	}
	if len(msg) > 500 {
		msg = msg[:500] + "..."
	}
	return fmt.Errorf("graph %s %s failed: %s | %s", method, fullURL, status, msg)
}

func shouldRetryStatus(status int) bool {
	switch status {
	case http.StatusTooManyRequests, http.StatusInternalServerError, http.StatusBadGateway, http.StatusServiceUnavailable, http.StatusGatewayTimeout:
		return true
	default:
		return false
	}
}

func retryDelay(attempt int, retryAfter string) time.Duration {
	if retryAfter != "" {
		if secs, err := strconv.Atoi(strings.TrimSpace(retryAfter)); err == nil && secs > 0 {
			return time.Duration(secs) * time.Second
		}
	}
	base := 500 * time.Millisecond
	d := base * time.Duration(1<<attempt)
	if d > 8*time.Second {
		d = 8 * time.Second
	}
	return d
}

func decodeJWTClaims(token string) (map[string]any, error) {
	parts := strings.Split(token, ".")
	if len(parts) < 2 {
		return nil, errors.New("invalid JWT format")
	}
	payload, err := base64.RawURLEncoding.DecodeString(parts[1])
	if err != nil {
		return nil, err
	}
	var claims map[string]any
	if err := json.Unmarshal(payload, &claims); err != nil {
		return nil, err
	}
	return claims, nil
}

func (g *graphClient) authHealth(ctx context.Context) (string, error) {
	token, err := g.cred.GetToken(ctx, policy.TokenRequestOptions{Scopes: g.scope})
	if err != nil {
		return "", err
	}
	claims, err := decodeJWTClaims(token.Token)
	if err != nil {
		return "", err
	}
	exp := asString(claims["exp"])
	if exp != "" {
		if unix, convErr := strconv.ParseInt(exp, 10, 64); convErr == nil {
			exp = time.Unix(unix, 0).UTC().Format(time.RFC3339)
		}
	}
	effectiveClient := asString(claims["appid"])
	if effectiveClient == "" {
		effectiveClient = asString(claims["azp"])
	}
	effectiveTenant := asString(claims["tid"])
	tokenScopes := asString(claims["scp"])

	var b strings.Builder
	b.WriteString("Auth Health\n\n")
	fmt.Fprintf(&b, "Configured Client ID: %s\n", g.cfg.ClientID)
	fmt.Fprintf(&b, "Configured Tenant ID: %s\n", g.cfg.TenantID)
	fmt.Fprintf(&b, "Token Client ID: %s\n", effectiveClient)
	fmt.Fprintf(&b, "Token Tenant ID: %s\n", effectiveTenant)
	fmt.Fprintf(&b, "Token Expires (UTC): %s\n", exp)
	fmt.Fprintf(&b, "Token Scopes: %s\n\n", tokenScopes)
	b.WriteString("Requested Scopes:\n")
	for _, s := range g.scope {
		fmt.Fprintf(&b, "- %s\n", s)
	}
	return b.String(), nil
}

func (g *graphClient) list(ctx context.Context, path string) ([]map[string]any, error) {
	next := graphBase + path
	all := make([]map[string]any, 0)
	for next != "" {
		b, err := g.do(ctx, http.MethodGet, next, nil)
		if err != nil {
			return nil, err
		}
		var page pageResponse
		if err := json.Unmarshal(b, &page); err != nil {
			return nil, err
		}
		all = append(all, page.Value...)
		g.emitProgress(fmt.Sprintf("Fetched %d items from %s", len(all), path))
		next = page.NextLink
	}
	return all, nil
}

func (g *graphClient) setProgressHook(hook func(string)) {
	g.progressHook = hook
}

func (g *graphClient) emitProgress(text string) {
	if g.progressHook != nil {
		g.progressHook(text)
	}
}

func escapeOData(s string) string {
	return strings.ReplaceAll(s, "'", "''")
}

func asString(v any) string {
	if v == nil {
		return ""
	}
	switch t := v.(type) {
	case string:
		return t
	default:
		return fmt.Sprintf("%v", v)
	}
}

func padRight(s string, w int) string {
	if len(s) >= w {
		return s
	}
	return s + strings.Repeat(" ", w-len(s))
}

func truncate(s string, w int) string {
	if w <= 0 || len(s) <= w {
		return s
	}
	if w <= 3 {
		return s[:w]
	}
	return s[:w-3] + "..."
}

func renderTable(headers []string, rows [][]string) string {
	if len(headers) == 0 {
		return ""
	}
	widths := make([]int, len(headers))
	for i, h := range headers {
		widths[i] = len(h)
	}
	for _, r := range rows {
		for i := range headers {
			if i < len(r) && len(r[i]) > widths[i] {
				widths[i] = len(r[i])
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

func parseTableFromText(s string) ([]string, [][]string, bool) {
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

func exportCSV(path string, headers []string, rows [][]string) error {
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

func (g *graphClient) findGroupByDisplayName(ctx context.Context, name string) (map[string]any, error) {
	filter := url.QueryEscape(fmt.Sprintf("displayName eq '%s'", escapeOData(name)))
	items, err := g.list(ctx, "/groups?$select=id,displayName&$filter="+filter)
	if err != nil {
		return nil, err
	}
	if len(items) == 0 {
		return nil, errors.New("group not found")
	}
	return items[0], nil
}

func (g *graphClient) listUsersInGroup(ctx context.Context, groupName string) (string, error) {
	group, err := g.findGroupByDisplayName(ctx, groupName)
	if err != nil {
		return "", err
	}
	groupID := asString(group["id"])
	members, err := g.list(ctx, fmt.Sprintf("/groups/%s/members?$select=id,displayName,userPrincipalName", groupID))
	if err != nil {
		return "", err
	}
	rows := make([][]string, 0)
	for _, m := range members {
		if asString(m["@odata.type"]) != "#microsoft.graph.user" {
			continue
		}
		rows = append(rows, []string{
			asString(m["displayName"]),
			asString(m["userPrincipalName"]),
			asString(m["id"]),
		})
	}
	if len(rows) == 0 {
		return fmt.Sprintf("Users in group: %s\n(No user members)", asString(group["displayName"])), nil
	}
	return fmt.Sprintf("Users in group: %s\n\n%s",
		asString(group["displayName"]),
		renderTable([]string{"Display Name", "UPN", "Object ID"}, rows),
	), nil
}

func (g *graphClient) listUsers(ctx context.Context) (string, error) {
	users, err := g.list(ctx, "/users?$select=id,displayName,userPrincipalName")
	if err != nil {
		return "", err
	}
	rows := make([][]string, 0, len(users))
	for _, u := range users {
		rows = append(rows, []string{asString(u["displayName"]), asString(u["userPrincipalName"]), asString(u["id"])})
	}
	return fmt.Sprintf("Users: %d\n\n%s", len(users), renderTable([]string{"Display Name", "UPN", "Object ID"}, rows)), nil
}

func (g *graphClient) listGroups(ctx context.Context) (string, error) {
	groups, err := g.list(ctx, "/groups?$select=id,displayName")
	if err != nil {
		return "", err
	}
	rows := make([][]string, 0, len(groups))
	for _, grp := range groups {
		rows = append(rows, []string{asString(grp["displayName"]), asString(grp["id"])})
	}
	return fmt.Sprintf("Groups: %d\n\n%s", len(groups), renderTable([]string{"Display Name", "Object ID"}, rows)), nil
}

func (g *graphClient) searchGroups(ctx context.Context, term string) (string, error) {
	groups, err := g.list(ctx, "/groups?$select=id,displayName")
	if err != nil {
		return "", err
	}
	termLower := strings.ToLower(term)
	rows := make([][]string, 0)
	for _, grp := range groups {
		name := asString(grp["displayName"])
		if strings.Contains(strings.ToLower(name), termLower) {
			rows = append(rows, []string{name, asString(grp["id"])})
		}
	}
	if len(rows) == 0 {
		return fmt.Sprintf("Groups matching %q:\n(No matches)", term), nil
	}
	return fmt.Sprintf("Groups matching %q: %d\n\n%s", term, len(rows), renderTable([]string{"Display Name", "Object ID"}, rows)), nil
}

func (g *graphClient) listDevices(ctx context.Context) (string, error) {
	devices, err := g.list(ctx, "/devices?$select=id,deviceId,displayName")
	if err != nil {
		return "", err
	}
	rows := make([][]string, 0, len(devices))
	for _, d := range devices {
		rows = append(rows, []string{asString(d["displayName"]), asString(d["id"]), asString(d["deviceId"])})
	}
	return fmt.Sprintf("Devices: %d\n\n%s", len(devices), renderTable([]string{"Display Name", "Object ID", "Device ID"}, rows)), nil
}

func (g *graphClient) reportComplianceSnapshot(ctx context.Context) (string, error) {
	devices, err := g.list(ctx, "/deviceManagement/managedDevices?$select=id,deviceName,complianceState,operatingSystem,osVersion,lastSyncDateTime")
	if err != nil {
		return "", err
	}
	counts := map[string]int{
		"compliant":     0,
		"noncompliant":  0,
		"inGracePeriod": 0,
		"unknown/other": 0,
	}
	for _, d := range devices {
		state := strings.TrimSpace(asString(d["complianceState"]))
		switch strings.ToLower(state) {
		case "compliant":
			counts["compliant"]++
		case "noncompliant":
			counts["noncompliant"]++
		case "ingraceperiod":
			counts["inGracePeriod"]++
		default:
			counts["unknown/other"]++
		}
	}
	total := len(devices)
	rows := [][]string{}
	for _, k := range []string{"compliant", "noncompliant", "inGracePeriod", "unknown/other"} {
		pct := 0.0
		if total > 0 {
			pct = (float64(counts[k]) / float64(total)) * 100
		}
		rows = append(rows, []string{k, fmt.Sprintf("%d", counts[k]), fmt.Sprintf("%.1f%%", pct)})
	}
	return fmt.Sprintf("Device Compliance Snapshot\n\nManaged devices: %d\n\n%s",
		total,
		renderTable([]string{"Compliance State", "Count", "Percent"}, rows),
	), nil
}

func (g *graphClient) listDevicesInGroup(ctx context.Context, groupName string) (string, error) {
	group, err := g.findGroupByDisplayName(ctx, groupName)
	if err != nil {
		return "", err
	}
	groupID := asString(group["id"])
	members, err := g.list(ctx, fmt.Sprintf("/groups/%s/members?$select=id,displayName,deviceId", groupID))
	if err != nil {
		return "", err
	}
	rows := make([][]string, 0)
	for _, m := range members {
		if asString(m["@odata.type"]) != "#microsoft.graph.device" {
			continue
		}
		rows = append(rows, []string{asString(m["displayName"]), asString(m["id"]), asString(m["deviceId"])})
	}
	if len(rows) == 0 {
		return fmt.Sprintf("Devices in group: %s\n(No device members)", asString(group["displayName"])), nil
	}
	return fmt.Sprintf("Devices in group: %s\n\n%s",
		asString(group["displayName"]),
		renderTable([]string{"Display Name", "Object ID", "Device ID"}, rows),
	), nil
}

func readCSV(path string) ([]map[string]string, error) {
	f, err := os.Open(path)
	if err != nil {
		return nil, err
	}
	defer f.Close()
	r := csv.NewReader(f)
	rows, err := r.ReadAll()
	if err != nil {
		return nil, err
	}
	if len(rows) < 2 {
		return nil, errors.New("csv has no data rows")
	}
	headers := rows[0]
	out := make([]map[string]string, 0, len(rows)-1)
	for _, row := range rows[1:] {
		item := map[string]string{}
		for i := range headers {
			if i < len(row) {
				item[headers[i]] = row[i]
			} else {
				item[headers[i]] = ""
			}
		}
		out = append(out, item)
	}
	return out, nil
}

type csvIssue struct {
	Severity string
	Row      int
	Field    string
	Code     string
	Message  string
}

type csvValidationResult struct {
	FilePath string
	Rows     int
	Errors   int
	Warnings int
	Issues   []csvIssue
	Pass     bool
}

func validateCSVStrict(path string, requiredHeaders, keyColumns []string) (csvValidationResult, error) {
	res := csvValidationResult{FilePath: path, Pass: true}
	f, err := os.Open(path)
	if err != nil {
		return res, err
	}
	defer f.Close()

	r := csv.NewReader(f)
	headers, err := r.Read()
	if err != nil {
		return res, err
	}
	headerMap := map[string]int{}
	for i, h := range headers {
		headerMap[strings.TrimSpace(h)] = i
	}

	for _, req := range requiredHeaders {
		if _, ok := headerMap[req]; !ok {
			res.Issues = append(res.Issues, csvIssue{
				Severity: "Error", Row: 1, Field: req, Code: "MISSING_HEADER",
				Message: "Required header is missing",
			})
		}
	}
	if len(res.Issues) > 0 {
		res.Errors = len(res.Issues)
		res.Pass = false
		return res, nil
	}

	for _, h := range headers {
		trim := strings.TrimSpace(h)
		if trim != h {
			res.Issues = append(res.Issues, csvIssue{
				Severity: "Warning", Row: 1, Field: h, Code: "HEADER_WHITESPACE",
				Message: "Header has leading/trailing whitespace",
			})
		}
	}

	seen := map[string]int{}
	rowNum := 1
	for {
		row, readErr := r.Read()
		if readErr == io.EOF {
			break
		}
		rowNum++
		if readErr != nil {
			res.Issues = append(res.Issues, csvIssue{
				Severity: "Error", Row: rowNum, Field: "", Code: "MALFORMED_ROW",
				Message: readErr.Error(),
			})
			continue
		}
		res.Rows++
		keyParts := make([]string, 0, len(keyColumns))
		for _, req := range requiredHeaders {
			idx := headerMap[req]
			val := ""
			if idx < len(row) {
				val = strings.TrimSpace(row[idx])
			}
			if val == "" {
				res.Issues = append(res.Issues, csvIssue{
					Severity: "Error", Row: rowNum, Field: req, Code: "MISSING_REQUIRED_VALUE",
					Message: "Required value is empty",
				})
			}
		}
		for _, k := range keyColumns {
			idx := headerMap[k]
			val := ""
			if idx < len(row) {
				val = strings.ToLower(strings.TrimSpace(row[idx]))
			}
			keyParts = append(keyParts, val)
		}
		key := strings.Join(keyParts, "|")
		if key != "" {
			if first, exists := seen[key]; exists {
				res.Issues = append(res.Issues, csvIssue{
					Severity: "Error", Row: rowNum, Field: strings.Join(keyColumns, "+"), Code: "DUPLICATE_KEY",
					Message: fmt.Sprintf("Duplicate key; first seen at row %d", first),
				})
			} else {
				seen[key] = rowNum
			}
		}
	}

	for _, i := range res.Issues {
		if i.Severity == "Error" {
			res.Errors++
		} else if i.Severity == "Warning" {
			res.Warnings++
		}
	}
	res.Pass = res.Errors == 0
	return res, nil
}

func formatValidationReport(title string, res csvValidationResult) string {
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
	b.WriteString(renderTable([]string{"Severity", "Row", "Field", "Code", "Message"}, rows))
	return b.String()
}

func validateForAction(spec actionSpec, inputs []string) (csvValidationResult, bool, error) {
	switch spec.id {
	case actAddUsersCSV, actReportCsvUsers:
		res, err := validateCSVStrict(inputs[0], []string{"User_Principal_Name"}, []string{"User_Principal_Name"})
		return res, true, err
	case actMakeGroupsCSV, actReportCsvGroups:
		res, err := validateCSVStrict(inputs[0], []string{"Group_Name"}, []string{"Group_Name"})
		return res, true, err
	case actAddAppsCSV, actReportCsvApps:
		res, err := validateCSVStrict(inputs[0], []string{"Group_Name", "App_Name"}, []string{"Group_Name", "App_Name"})
		return res, true, err
	default:
		return csvValidationResult{}, false, nil
	}
}

func (g *graphClient) addUsersCSV(ctx context.Context, csvPath, groupName string, dryRun bool) (string, error) {
	rows, err := readCSV(csvPath)
	if err != nil {
		return "", err
	}
	group, err := g.findGroupByDisplayName(ctx, groupName)
	if err != nil {
		return "", err
	}
	groupID := asString(group["id"])

	var b strings.Builder
	for _, row := range rows {
		upn := row["User_Principal_Name"]
		if upn == "" {
			fmt.Fprintf(&b, "Skipped row: missing User_Principal_Name\n")
			continue
		}
		filter := url.QueryEscape(fmt.Sprintf("userPrincipalName eq '%s'", escapeOData(upn)))
		users, err := g.list(ctx, "/users?$select=id,userPrincipalName&$filter="+filter)
		if err != nil || len(users) == 0 {
			fmt.Fprintf(&b, "User not found: %s\n", upn)
			continue
		}
		userID := asString(users[0]["id"])
		body := map[string]string{
			"@odata.id": fmt.Sprintf("https://graph.microsoft.com/v1.0/directoryObjects/%s", userID),
		}
		if dryRun {
			fmt.Fprintf(&b, "Would add %s\n", upn)
			continue
		}
		_, err = g.do(ctx, http.MethodPost, fmt.Sprintf("%s/groups/%s/members/$ref", graphBase, groupID), body)
		if err != nil {
			fmt.Fprintf(&b, "Failed to add %s: %v\n", upn, err)
			continue
		}
		fmt.Fprintf(&b, "Added %s\n", upn)
	}
	return b.String(), nil
}

func (g *graphClient) makeGroupsCSV(ctx context.Context, csvPath string, dryRun bool) (string, error) {
	rows, err := readCSV(csvPath)
	if err != nil {
		return "", err
	}
	var b strings.Builder
	for _, row := range rows {
		groupName := row["Group_Name"]
		if groupName == "" {
			fmt.Fprintf(&b, "Skipped row: missing Group_Name\n")
			continue
		}
		_, err := g.findGroupByDisplayName(ctx, groupName)
		if err == nil {
			fmt.Fprintf(&b, "Exists: %s\n", groupName)
			continue
		}
		body := map[string]any{
			"displayName":     groupName,
			"mailNickname":    strings.ReplaceAll(groupName, " ", "_"),
			"description":     groupName,
			"mailEnabled":     false,
			"securityEnabled": true,
		}
		if dryRun {
			fmt.Fprintf(&b, "Would create: %s\n", groupName)
			continue
		}
		_, err = g.do(ctx, http.MethodPost, graphBase+"/groups", body)
		if err != nil {
			fmt.Fprintf(&b, "Failed to create %s: %v\n", groupName, err)
			continue
		}
		fmt.Fprintf(&b, "Created: %s\n", groupName)
	}
	return b.String(), nil
}

func (g *graphClient) addAppsCSV(ctx context.Context, csvPath string, dryRun bool) (string, error) {
	rows, err := readCSV(csvPath)
	if err != nil {
		return "", err
	}
	var b strings.Builder
	for _, row := range rows {
		appName := row["App_Name"]
		groupName := row["Group_Name"]
		if appName == "" || groupName == "" {
			fmt.Fprintf(&b, "Skipped row: missing App_Name or Group_Name\n")
			continue
		}

		appFilter := url.QueryEscape(fmt.Sprintf("displayName eq '%s'", escapeOData(appName)))
		apps, err := g.list(ctx, "/deviceAppManagement/mobileApps?$select=id,displayName&$filter="+appFilter)
		if err != nil || len(apps) == 0 {
			fmt.Fprintf(&b, "App not found: %s\n", appName)
			continue
		}
		appID := asString(apps[0]["id"])

		group, err := g.findGroupByDisplayName(ctx, groupName)
		if err != nil {
			fmt.Fprintf(&b, "Group not found: %s\n", groupName)
			continue
		}
		groupID := asString(group["id"])

		body := map[string]any{
			"intent": "available",
			"target": map[string]any{
				"@odata.type": "#microsoft.graph.groupAssignmentTarget",
				"groupId":     groupID,
			},
		}
		if dryRun {
			fmt.Fprintf(&b, "Would assign %s -> %s\n", appName, groupName)
			continue
		}
		_, err = g.do(ctx, http.MethodPost, fmt.Sprintf("%s/deviceAppManagement/mobileApps/%s/assignments", graphBase, appID), body)
		if err != nil {
			fmt.Fprintf(&b, "Failed assignment app=%s group=%s: %v\n", appName, groupName, err)
			continue
		}
		fmt.Fprintf(&b, "Assigned %s -> %s\n", appName, groupName)
	}
	return b.String(), nil
}

func (g *graphClient) listGroupApps(ctx context.Context) (string, error) {
	apps, err := g.list(ctx, "/deviceAppManagement/mobileApps?$select=id,displayName")
	if err != nil {
		return "", err
	}
	type row struct {
		AppName      string
		GroupName    string
		AssignmentID string
		Intent       string
	}
	var rows []row
	for _, app := range apps {
		appID := asString(app["id"])
		assignments, err := g.list(ctx, fmt.Sprintf("/deviceAppManagement/mobileApps/%s/assignments?$select=id,intent,target", appID))
		if err != nil {
			continue
		}
		for _, a := range assignments {
			target, ok := a["target"].(map[string]any)
			if !ok {
				continue
			}
			groupID := asString(target["groupId"])
			if groupID == "" {
				continue
			}
			b, err := g.do(ctx, http.MethodGet, graphBase+"/groups/"+groupID+"?$select=displayName", nil)
			if err != nil {
				continue
			}
			var grp map[string]any
			if json.Unmarshal(b, &grp) != nil {
				continue
			}
			rows = append(rows, row{
				AppName:      asString(app["displayName"]),
				GroupName:    asString(grp["displayName"]),
				AssignmentID: asString(a["id"]),
				Intent:       asString(a["intent"]),
			})
		}
	}

	var b strings.Builder
	if len(rows) == 0 {
		return "No group app assignments found.", nil
	}
	tabRows := make([][]string, 0, len(rows))
	for _, r := range rows {
		tabRows = append(tabRows, []string{r.AppName, r.GroupName, r.AssignmentID, r.Intent})
	}
	fmt.Fprintf(&b, "App-group assignments: %d\n\n%s", len(rows), renderTable([]string{"App", "Group", "Assignment ID", "Intent"}, tabRows))
	return b.String(), nil
}

type menuState int

const (
	stateMain menuState = iota
	stateUsersGroups
	stateDevicesApps
	stateReports
	stateSettings
	stateMenuFilter
	stateInput
	stateExportPrompt
	stateConfirm
	stateWorking
	stateOutput
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
	actReportCsvUsers
	actReportCsvGroups
	actReportCsvApps
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

type confirmKind int

const (
	confirmNone confirmKind = iota
	confirmAction
	confirmExport
)

type model struct {
	client        *graphClient
	state         menuState
	cursor        int
	menuTop       int
	width         int
	height        int
	lastMenuState menuState

	mainMenu []menuItem
	userMenu []menuItem
	devMenu  []menuItem
	repMenu  []menuItem
	cfgMenu  []menuItem

	spin        spinner.Model
	viewport    viewport.Model
	vpReady     bool
	styles      uiStyles
	input       textinput.Model
	filterInput textinput.Model
	exportInput textinput.Model
	filterQuery string
	currentSpec actionSpec
	inputs      []string
	output      string
	lastHeaders []string
	lastRows    [][]string

	confirmKind        confirmKind
	confirmTitle       string
	confirmBody        string
	confirmCancelState menuState
	pendingSpec        actionSpec
	pendingInputs      []string
	pendingExportPath  string
	dryRun             bool
	progressCh         chan progressMsg
	progressText       string
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

func newModel(client *graphClient) model {
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
			{label: "Validate Users->Group CSV", description: "Strict quality checks for User_Principal_Name format", action: actReportCsvUsers},
			{label: "Validate Create-Groups CSV", description: "Strict quality checks for Group_Name format", action: actReportCsvGroups},
			{label: "Validate App-Assignment CSV", description: "Strict quality checks for Group_Name + App_Name format", action: actReportCsvApps},
			{label: "Back", description: "Return to main menu", next: stateMain},
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
		spin:        sp,
		styles:      newUIStyles(),
		input:       ti,
		filterInput: fi,
		exportInput: ei,
		viewport:    viewport.New(80, 20),
		progressCh:  make(chan progressMsg, 64),
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
	return maxInt(4, m.height-14)
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

func (m *model) setOutput(text string) {
	m.output = text
	m.viewport.SetContent(text)
	m.viewport.GotoTop()
	m.lastHeaders = nil
	m.lastRows = nil
	if h, r, ok := parseTableFromText(text); ok {
		m.lastHeaders = h
		m.lastRows = r
	}
	m.state = stateOutput
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

func waitProgressCmd(ch <-chan progressMsg) tea.Cmd {
	return func() tea.Msg {
		msg := <-ch
		return msg
	}
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
	return fmt.Sprintf("Client ID: %s\nTenant ID: %s\nDry-Run: %s", m.client.cfg.ClientID, m.client.cfg.TenantID, mode)
}

func (m *model) applyAuthConfig(cfg authConfig) error {
	client, err := newGraphClientWithConfig(cfg)
	if err != nil {
		return err
	}
	if err := saveAuthConfigToFile(cfg); err != nil {
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
		cfg := authConfig{ClientID: defaultClientID, TenantID: "common"}
		if err := m.applyAuthConfig(cfg); err != nil {
			return "", err, true
		}
		return "Auth config reset.\n\n" + m.authSummary(), nil, true
	case actSetClientID:
		clientID := strings.TrimSpace(inputs[0])
		if clientID == "" {
			return "", errors.New("client ID cannot be empty"), true
		}
		cfg := m.client.cfg
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
		cfg := m.client.cfg
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
		ctx := context.Background()
		m.client.setProgressHook(func(text string) {
			select {
			case m.progressCh <- progressMsg{text: text}:
			default:
			}
		})
		defer m.client.setProgressHook(nil)
		var (
			out string
			err error
		)
		switch spec.id {
		case actListUsersGroup:
			out, err = m.client.listUsersInGroup(ctx, inputs[0])
		case actListGroups:
			out, err = m.client.listGroups(ctx)
		case actListUsers:
			out, err = m.client.listUsers(ctx)
		case actSearchGroups:
			out, err = m.client.searchGroups(ctx, inputs[0])
		case actAuthHealth:
			out, err = m.client.authHealth(ctx)
		case actAddUsersCSV:
			out, err = m.client.addUsersCSV(ctx, inputs[0], inputs[1], m.dryRun)
		case actListDevices:
			out, err = m.client.listDevices(ctx)
		case actReportComplianceSnapshot:
			out, err = m.client.reportComplianceSnapshot(ctx)
		case actListDevicesGroup:
			out, err = m.client.listDevicesInGroup(ctx, inputs[0])
		case actMakeGroupsCSV:
			out, err = m.client.makeGroupsCSV(ctx, inputs[0], m.dryRun)
		case actAddAppsCSV:
			out, err = m.client.addAppsCSV(ctx, inputs[0], m.dryRun)
		case actListGroupApps:
			out, err = m.client.listGroupApps(ctx)
		case actReportCsvUsers:
			if res, ok, verr := validateForAction(spec, inputs); verr != nil {
				err = verr
			} else if ok {
				out = formatValidationReport("CSV Quality Report - Users to Group", res)
			}
		case actReportCsvGroups:
			if res, ok, verr := validateForAction(spec, inputs); verr != nil {
				err = verr
			} else if ok {
				out = formatValidationReport("CSV Quality Report - Create Groups", res)
			}
		case actReportCsvApps:
			if res, ok, verr := validateForAction(spec, inputs); verr != nil {
				err = verr
			} else if ok {
				out = formatValidationReport("CSV Quality Report - App Assignments", res)
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
		m.input.Width = minInt(72, maxInt(24, msg.Width-16))
		m.exportInput.Width = minInt(72, maxInt(24, msg.Width-16))
		panelInnerW := maxInt(20, msg.Width-12)
		panelInnerH := maxInt(6, msg.Height-10)
		m.viewport.Width = panelInnerW
		m.viewport.Height = panelInnerH
		if m.output != "" {
			m.viewport.SetContent(m.output)
		}
		m.vpReady = true
		return m, nil
	case tea.KeyMsg:
		switch m.state {
		case stateMain, stateUsersGroups, stateDevicesApps, stateReports, stateSettings:
			visible := m.visibleMenu()
			switch msg.String() {
			case "ctrl+c", "q":
				if m.state == stateMain {
					return m, tea.Quit
				}
				m.resetMenuPosition(stateMain)
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
				m.cursor = maxInt(0, m.cursor-page)
				m.ensureMenuCursorVisible()
			case "pgdown":
				page := m.menuPageSize()
				m.cursor = minInt(maxInt(0, len(visible)-1), m.cursor+page)
				m.ensureMenuCursorVisible()
			case "home":
				m.cursor = 0
				m.menuTop = 0
			case "end":
				m.cursor = maxInt(0, len(visible)-1)
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
						m.state = stateWorking
						m.output = ""
						m.progressText = "Starting operation..."
						return m, tea.Batch(m.spin.Tick, waitProgressCmd(m.progressCh), m.runActionCmd(spec, nil))
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
				if item.next == stateUsersGroups || item.next == stateDevicesApps || item.next == stateReports || item.next == stateSettings {
					m.resetMenuPosition(item.next)
					return m, nil
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
				m.input.Blur()
				m.returnToLastMenu()
				return m, nil
			case "enter":
				val := strings.TrimSpace(m.input.Value())
				m.inputs = append(m.inputs, val)
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
						m.setOutput(formatValidationReport("CSV Validation Failed (Write Blocked)", res))
						return m, nil
					}
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
				m.state = stateWorking
				m.output = ""
				m.progressText = "Starting operation..."
				return m, tea.Batch(m.spin.Tick, waitProgressCmd(m.progressCh), m.runActionCmd(m.currentSpec, m.inputs))
			}
			var cmd tea.Cmd
			m.input, cmd = m.input.Update(msg)
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
				m.exportInput.Blur()
				m.startConfirmExport(path, stateExportPrompt)
				return m, nil
			}
			var cmd tea.Cmd
			m.exportInput, cmd = m.exportInput.Update(msg)
			return m, cmd
		case stateConfirm:
			switch msg.String() {
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
					m.state = stateWorking
					m.output = ""
					m.progressText = "Starting operation..."
					return m, tea.Batch(m.spin.Tick, waitProgressCmd(m.progressCh), m.runActionCmd(spec, inputs))
				case confirmExport:
					path := m.pendingExportPath
					m.clearConfirm()
					if m.dryRun {
						m.setOutput(m.output + "\n\nDry-run: would export CSV to " + path)
						return m, nil
					}
					if err := exportCSV(path, m.lastHeaders, m.lastRows); err != nil {
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
			if msg.String() == "ctrl+c" {
				return m, tea.Quit
			}
		case stateOutput:
			switch msg.String() {
			case "enter", "esc":
				m.returnToLastMenu()
				m.output = ""
				m.lastHeaders = nil
				m.lastRows = nil
				m.viewport.SetContent("")
				m.viewport.GotoTop()
				return m, nil
			case "e":
				if len(m.lastHeaders) == 0 || len(m.lastRows) == 0 {
					return m, nil
				}
				m.exportInput.SetValue("")
				m.exportInput.Focus()
				m.filterInput.Blur()
				m.input.Blur()
				m.state = stateExportPrompt
				return m, textinput.Blink
			}
			var cmd tea.Cmd
			m.viewport, cmd = m.viewport.Update(msg)
			return m, cmd
		}
	case spinner.TickMsg:
		if m.state == stateWorking {
			var cmd tea.Cmd
			m.spin, cmd = m.spin.Update(msg)
			return m, cmd
		}
	case progressMsg:
		if m.state == stateWorking {
			m.progressText = msg.text
			return m, waitProgressCmd(m.progressCh)
		}
	case resultMsg:
		if msg.err != nil {
			m.setOutput("Error:\n" + msg.err.Error())
		} else {
			m.setOutput(msg.text)
		}
		m.progressText = ""
		return m, nil
	}
	return m, nil
}

func (m model) View() string {
	switch m.state {
	case stateConfirm:
		body := m.styles.panel.Render(fmt.Sprintf("%s\n\n%s\n\n%s",
			m.styles.subHeader.Render(m.confirmTitle),
			m.confirmBody,
			m.styles.hint.Render("y/Enter: confirm   n/Esc: cancel"),
		))
		return m.styles.app.Render(m.styles.header.Render(" Intune Management Tool ") + "\n\n" + body)
	case stateMenuFilter:
		body := m.styles.panel.Render(fmt.Sprintf("%s\n\n%s\n\n%s",
			m.styles.subHeader.Render("Filter Menu Options"),
			m.filterInput.View(),
			m.styles.hint.Render("Enter: apply filter   Esc: cancel"),
		))
		return m.styles.app.Render(m.styles.header.Render(" Intune Management Tool ") + "\n\n" + body)
	case stateExportPrompt:
		body := m.styles.panel.Render(fmt.Sprintf("%s\n\n%s\n\n%s",
			m.styles.subHeader.Render("Export Current Table to CSV"),
			m.exportInput.View(),
			m.styles.hint.Render("Enter: export   Esc: cancel"),
		))
		return m.styles.app.Render(m.styles.header.Render(" Intune Management Tool ") + "\n\n" + body)
	case stateInput:
		step := len(m.inputs) + 1
		total := len(m.currentSpec.prompts)
		title := m.styles.subHeader.Render(fmt.Sprintf("Input %d/%d", step, total))
		body := m.styles.panel.Render(fmt.Sprintf("%s\n\n%s\n\n%s",
			title,
			m.input.View(),
			m.styles.hint.Render("Enter: continue   Esc: cancel"),
		))
		return m.styles.app.Render(m.styles.header.Render(" Intune Management Tool ") + "\n\n" + body)
	case stateWorking:
		progress := m.progressText
		if strings.TrimSpace(progress) == "" {
			progress = "Starting operation..."
		}
		body := m.styles.panel.Render(fmt.Sprintf("%s Running Graph operation...\n\n%s",
			m.spin.View(),
			m.styles.hint.Render(progress),
		))
		return m.styles.app.Render(m.styles.header.Render(" Intune Management Tool ") + "\n\n" + body)
	case stateOutput:
		prefix := m.styles.ok.Render("Result")
		if strings.HasPrefix(m.output, "Error:") {
			prefix = m.styles.err.Render("Result")
		}
		content := m.output
		if m.vpReady {
			content = m.viewport.View()
		}
		exportHint := "Up/Down PgUp/PgDn Home/End: scroll   Enter/Esc: return"
		if len(m.lastHeaders) > 0 && len(m.lastRows) > 0 {
			exportHint = "Up/Down PgUp/PgDn Home/End: scroll   e: export table   Enter/Esc: return"
		}
		body := m.styles.panel.Render(fmt.Sprintf("%s\n\n%s\n\n%s",
			prefix,
			content,
			m.styles.hint.Render(exportHint),
		))
		return m.styles.app.Render(m.styles.header.Render(" Intune Management Tool ") + "\n\n" + body)
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
		} else if m.state == stateSettings {
			title = "Settings"
			sub = "Authentication configuration for Microsoft Graph"
		}
		menuView := m.renderMenu()
		filterLine := ""
		if m.filterQuery != "" {
			filterLine = "\n" + m.styles.hint.Render("Active filter: "+m.filterQuery+" (press / to edit)")
		}
		screen := m.styles.header.Render(" Intune Management Tool ") + "\n\n" +
			m.styles.subHeader.Render(title) + "\n" +
			m.styles.hint.Render(sub) + "\n\n" +
			m.styles.panel.Render(menuView) + filterLine + "\n\n" +
			m.styles.hint.Render("Arrows/jk PgUp/PgDn Home/End: move   /: filter   Enter: select   q: back/quit")
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
		start = maxInt(0, len(visible)-page)
	}
	end := minInt(len(visible), start+page)

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

func minInt(a, b int) int {
	if a < b {
		return a
	}
	return b
}

func maxInt(a, b int) int {
	if a > b {
		return a
	}
	return b
}

func main() {
	client, err := newGraphClient()
	if err != nil {
		fmt.Println("Failed to initialize Graph auth:", err)
		os.Exit(1)
	}
	p := tea.NewProgram(newModel(client))
	if _, err := p.Run(); err != nil {
		fmt.Println("Program error:", err)
		os.Exit(1)
	}
}
