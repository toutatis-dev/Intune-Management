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
	"sort"
	"strconv"
	"strings"
	"time"

	"github.com/Azure/azure-sdk-for-go/sdk/azcore/policy"
	"github.com/Azure/azure-sdk-for-go/sdk/azidentity"
	"github.com/atotto/clipboard"
	"github.com/charmbracelet/bubbles/spinner"
	"github.com/charmbracelet/bubbles/textinput"
	"github.com/charmbracelet/bubbles/viewport"
	tea "github.com/charmbracelet/bubbletea"
	"github.com/charmbracelet/lipgloss"
	"github.com/mattn/go-runewidth"
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
	return os.WriteFile(path, b, 0600)
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
			if strings.TrimSpace(message.UserCode) != "" {
				fmt.Printf("Device code: %s\n", message.UserCode)
				if err := clipboard.WriteAll(message.UserCode); err == nil {
					fmt.Println("Copied device code to clipboard.")
				} else {
					fmt.Println("Could not copy device code to clipboard; copy it manually from above.")
				}
			}
			fmt.Println()
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
	var payload []byte
	if body != nil {
		var err error
		payload, err = json.Marshal(body)
		if err != nil {
			return nil, err
		}
	}

	const maxRetries = 4
	hadAuthRetry := false
	for attempt := 0; attempt <= maxRetries; attempt++ {
		token, err := g.cred.GetToken(ctx, policy.TokenRequestOptions{Scopes: g.scope})
		if err != nil {
			return nil, err
		}

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

		if resp.StatusCode == http.StatusUnauthorized && !hadAuthRetry {
			hadAuthRetry = true
			g.emitProgress("Token expired; refreshing credentials...")
			continue
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

type graphRequestError struct {
	Method     string
	URL        string
	Status     string
	StatusCode int
	Code       string
	Message    string
	Raw        string
}

func (e *graphRequestError) Error() string {
	if e == nil {
		return ""
	}
	if e.Code != "" {
		return fmt.Sprintf("graph %s %s failed: %s | %s: %s", e.Method, e.URL, e.Status, e.Code, e.Message)
	}
	if e.Raw == "" {
		return fmt.Sprintf("graph %s %s failed: %s", e.Method, e.URL, e.Status)
	}
	return fmt.Sprintf("graph %s %s failed: %s | %s", e.Method, e.URL, e.Status, e.Raw)
}

var (
	errNotFound      = errors.New("not found")
	errNoCSVDataRows = errors.New("csv has no data rows")
)

func formatGraphError(method, fullURL, status string, raw []byte) error {
	statusCode := 0
	if parts := strings.Fields(status); len(parts) > 0 {
		statusCode, _ = strconv.Atoi(parts[0])
	}
	var env graphErrorEnvelope
	if err := json.Unmarshal(raw, &env); err == nil && env.Error.Code != "" {
		return &graphRequestError{
			Method:     method,
			URL:        fullURL,
			Status:     status,
			StatusCode: statusCode,
			Code:       env.Error.Code,
			Message:    env.Error.Message,
		}
	}
	msg := strings.TrimSpace(string(raw))
	if len(msg) > 500 {
		msg = msg[:500] + "..."
	}
	return &graphRequestError{
		Method:     method,
		URL:        fullURL,
		Status:     status,
		StatusCode: statusCode,
		Raw:        msg,
	}
}

func isGraphNotFound(err error) bool {
	var reqErr *graphRequestError
	if !errors.As(err, &reqErr) {
		return false
	}
	if reqErr.StatusCode == http.StatusNotFound {
		return true
	}
	switch strings.ToLower(strings.TrimSpace(reqErr.Code)) {
	case "itemnotfound", "request_resourcenotfound", "resourcenotfound", "directoryobjectnotfound":
		return true
	case "request_badrequest":
		return strings.Contains(strings.ToLower(reqErr.Message), "invalid object identifier")
	default:
		return false
	}
}

func isGraphForbidden(err error) bool {
	var reqErr *graphRequestError
	if !errors.As(err, &reqErr) {
		return false
	}
	if reqErr.StatusCode == http.StatusForbidden {
		return true
	}
	switch strings.ToLower(strings.TrimSpace(reqErr.Code)) {
	case "authorization_requestdenied", "forbidden":
		return true
	default:
		return false
	}
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

	return renderInspector("Auth Health", [][2]string{
		{"Configured Client ID", g.cfg.ClientID},
		{"Configured Tenant ID", g.cfg.TenantID},
		{"Token Client ID", effectiveClient},
		{"Token Tenant ID", effectiveTenant},
		{"Token Expires (UTC)", exp},
		{"Token Scopes", tokenScopes},
		{"Requested Scopes", strings.Join(g.scope, ", ")},
	}), nil
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

func renderTable(headers []string, rows [][]string) string {
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

type ambiguousMatchError struct {
	ObjectType string
	Name       string
	Candidates []string
}

func (e *ambiguousMatchError) Error() string {
	if e == nil {
		return ""
	}
	msg := fmt.Sprintf("multiple %ss found with display name %q; use object ID instead", e.ObjectType, e.Name)
	if len(e.Candidates) == 0 {
		return msg
	}
	return msg + "\n\n" + strings.Join(e.Candidates, "\n")
}

func selectUniqueMatch(objectType, identifier string, items []map[string]any, formatter func(map[string]any) string) (map[string]any, error) {
	switch len(items) {
	case 0:
		return nil, fmt.Errorf("%s %w", objectType, errNotFound)
	case 1:
		return items[0], nil
	default:
		candidates := make([]string, 0, minInt(10, len(items)))
		for i, item := range items {
			if i >= 10 {
				break
			}
			candidates = append(candidates, formatter(item))
		}
		if len(items) > len(candidates) {
			candidates = append(candidates, fmt.Sprintf("...and %d more", len(items)-len(candidates)))
		}
		return nil, &ambiguousMatchError{
			ObjectType: objectType,
			Name:       identifier,
			Candidates: candidates,
		}
	}
}

func formatGroupCandidate(item map[string]any) string {
	return fmt.Sprintf("%s | %s", asString(item["displayName"]), asString(item["id"]))
}

func formatUserCandidate(item map[string]any) string {
	return fmt.Sprintf("%s | %s | %s", asString(item["displayName"]), asString(item["userPrincipalName"]), asString(item["id"]))
}

func formatDeviceCandidate(item map[string]any) string {
	return fmt.Sprintf("%s | %s | %s", asString(item["displayName"]), asString(item["deviceId"]), asString(item["id"]))
}

func formatAppCandidate(item map[string]any) string {
	return fmt.Sprintf("%s | %s", asString(item["displayName"]), asString(item["id"]))
}

func (g *graphClient) findUniqueByDisplayName(ctx context.Context, path, selectFields, objectType, name string, formatter func(map[string]any) string) (map[string]any, error) {
	filter := url.QueryEscape(fmt.Sprintf("displayName eq '%s'", escapeOData(name)))
	items, err := g.list(ctx, fmt.Sprintf("%s?$select=%s&$filter=%s", path, selectFields, filter))
	if err != nil {
		return nil, err
	}
	return selectUniqueMatch(objectType, name, items, formatter)
}

func (g *graphClient) findGroupByDisplayName(ctx context.Context, name string) (map[string]any, error) {
	group, err := g.findUniqueByDisplayName(ctx, "/groups", "id,displayName", "group", name, formatGroupCandidate)
	if errors.Is(err, errNotFound) {
		return nil, fmt.Errorf("group %w", errNotFound)
	}
	return group, err
}

func (g *graphClient) findUserByDisplayName(ctx context.Context, name string) (map[string]any, error) {
	user, err := g.findUniqueByDisplayName(ctx, "/users", "id,displayName,userPrincipalName,accountEnabled", "user", name, formatUserCandidate)
	if errors.Is(err, errNotFound) {
		return nil, fmt.Errorf("user %w", errNotFound)
	}
	return user, err
}

func (g *graphClient) findDeviceByDisplayName(ctx context.Context, name string) (map[string]any, error) {
	device, err := g.findUniqueByDisplayName(ctx, "/devices", "id,displayName,deviceId,operatingSystem,accountEnabled", "device", name, formatDeviceCandidate)
	if errors.Is(err, errNotFound) {
		return nil, fmt.Errorf("device %w", errNotFound)
	}
	return device, err
}

func (g *graphClient) findAppByDisplayName(ctx context.Context, name string) (map[string]any, error) {
	app, err := g.findUniqueByDisplayName(ctx, "/deviceAppManagement/mobileApps", "id,displayName,publisher", "app", name, formatAppCandidate)
	if errors.Is(err, errNotFound) {
		return nil, fmt.Errorf("app %w", errNotFound)
	}
	return app, err
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
	filter := url.QueryEscape(fmt.Sprintf("startswith(displayName,'%s')", escapeOData(term)))
	groups, err := g.list(ctx, fmt.Sprintf("/groups?$select=id,displayName&$filter=%s", filter))
	if err != nil {
		return "", err
	}
	rows := make([][]string, 0, len(groups))
	for _, grp := range groups {
		rows = append(rows, []string{asString(grp["displayName"]), asString(grp["id"])})
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

func classifyWindowsVersion(operatingSystem, osVersion string) string {
	if !strings.Contains(strings.ToLower(operatingSystem), "windows") {
		return "Other/Unknown"
	}
	parts := strings.Split(osVersion, ".")
	if len(parts) < 3 {
		return "Other/Unknown"
	}
	build, err := strconv.Atoi(parts[2])
	if err != nil {
		return "Other/Unknown"
	}
	if build >= 22000 {
		return "Windows 11"
	}
	return "Windows 10"
}

func (g *graphClient) reportWindowsBreakdown(ctx context.Context) (string, error) {
	devices, err := g.list(ctx, "/deviceManagement/managedDevices?$select=id,deviceName,operatingSystem,osVersion")
	if err != nil {
		return "", err
	}
	counts := map[string]int{
		"Windows 10":    0,
		"Windows 11":    0,
		"Other/Unknown": 0,
	}
	for _, d := range devices {
		classification := classifyWindowsVersion(asString(d["operatingSystem"]), asString(d["osVersion"]))
		counts[classification]++
	}
	total := len(devices)
	rows := make([][]string, 0, 3)
	for _, name := range []string{"Windows 10", "Windows 11", "Other/Unknown"} {
		pct := 0.0
		if total > 0 {
			pct = (float64(counts[name]) / float64(total)) * 100
		}
		rows = append(rows, []string{name, fmt.Sprintf("%d", counts[name]), fmt.Sprintf("%.1f%%", pct)})
	}
	return fmt.Sprintf("Windows OS Breakdown\n\nManaged devices: %d\n\n%s",
		total,
		renderTable([]string{"Category", "Count", "Percent"}, rows),
	), nil
}

type appFailureStat struct {
	ID     string
	Name   string
	Failed int
	Total  int
}

type failingAppsSummary struct {
	Scanned      int
	WithFailures int
	Skipped      int
}

func rankFailingApps(stats []appFailureStat) []appFailureStat {
	filtered := make([]appFailureStat, 0, len(stats))
	for _, stat := range stats {
		if stat.Failed > 0 {
			filtered = append(filtered, stat)
		}
	}
	sort.Slice(filtered, func(i, j int) bool {
		if filtered[i].Failed == filtered[j].Failed {
			return strings.ToLower(filtered[i].Name) < strings.ToLower(filtered[j].Name)
		}
		return filtered[i].Failed > filtered[j].Failed
	})
	if len(filtered) > 10 {
		filtered = filtered[:10]
	}
	return filtered
}

func renderTopFailingAppsReport(stats []appFailureStat, summary failingAppsSummary) string {
	ranked := rankFailingApps(stats)
	if len(ranked) == 0 {
		if summary.Skipped > 0 {
			return fmt.Sprintf("Top 10 Failing App Deployments\n\nApps scanned: %d\nApps with failures: 0\nApps skipped due to errors: %d\n\nNo failing app deployments found from successful app status queries.", summary.Scanned, summary.Skipped)
		}
		return fmt.Sprintf("Top 10 Failing App Deployments\n\nApps scanned: %d\nApps with failures: 0\nApps skipped due to errors: 0\n\nNo failing app deployments found.", summary.Scanned)
	}
	rows := make([][]string, 0, len(ranked))
	for i, stat := range ranked {
		rate := "0.0%"
		if stat.Total > 0 {
			rate = fmt.Sprintf("%.1f%%", (float64(stat.Failed)/float64(stat.Total))*100)
		}
		rows = append(rows, []string{
			fmt.Sprintf("%d", i+1),
			stat.Name,
			stat.ID,
			fmt.Sprintf("%d", stat.Failed),
			fmt.Sprintf("%d", stat.Total),
			rate,
		})
	}
	return fmt.Sprintf("Top 10 Failing App Deployments\n\nApps scanned: %d\nApps with failures: %d\nApps skipped due to errors: %d\n\n%s",
		summary.Scanned,
		summary.WithFailures,
		summary.Skipped,
		renderTable([]string{"Rank", "App", "App ID", "Failed Devices", "Total Statuses", "Failure Rate"}, rows),
	)
}

func (g *graphClient) reportTopFailingApps(ctx context.Context) (string, error) {
	apps, err := g.list(ctx, "/deviceAppManagement/mobileApps?$select=id,displayName")
	if err != nil {
		return "", err
	}
	stats := make([]appFailureStat, 0, len(apps))
	summary := failingAppsSummary{Scanned: len(apps)}
	for i, app := range apps {
		if (i+1)%20 == 0 {
			g.emitProgress(fmt.Sprintf("Processed %d/%d apps...", i+1, len(apps)))
		}
		appID := asString(app["id"])
		statuses, err := g.list(ctx, fmt.Sprintf("/deviceAppManagement/mobileApps/%s/deviceStatuses?$select=installState,deviceId", appID))
		if err != nil {
			summary.Skipped++
			continue
		}
		stat := appFailureStat{ID: appID, Name: asString(app["displayName"]), Total: len(statuses)}
		for _, s := range statuses {
			if strings.EqualFold(strings.TrimSpace(asString(s["installState"])), "failed") {
				stat.Failed++
			}
		}
		if stat.Failed > 0 {
			summary.WithFailures++
			stats = append(stats, stat)
		}
	}
	return renderTopFailingAppsReport(stats, summary), nil
}

func (g *graphClient) reportAppFailureDetails(ctx context.Context, identifier string) (string, error) {
	var app map[string]any
	body, err := g.do(ctx, http.MethodGet, graphBase+"/deviceAppManagement/mobileApps/"+url.PathEscape(identifier)+"?$select=id,displayName", nil)
	if err == nil {
		if err := json.Unmarshal(body, &app); err != nil {
			return "", err
		}
	} else if isGraphForbidden(err) {
		return "", errors.New("access denied: insufficient permissions to read this app")
	} else if isGraphNotFound(err) {
		app, err = g.findUniqueByDisplayName(ctx, "/deviceAppManagement/mobileApps", "id,displayName", "app", identifier, formatAppCandidate)
		if errors.Is(err, errNotFound) {
			return "", errors.New("app not found")
		}
		if err != nil {
			return "", err
		}
	} else {
		return "", err
	}

	statuses, err := g.list(ctx, fmt.Sprintf("/deviceAppManagement/mobileApps/%s/deviceStatuses", url.PathEscape(asString(app["id"]))))
	if err != nil {
		return "", err
	}
	rows := make([][]string, 0)
	for _, s := range statuses {
		if !strings.EqualFold(strings.TrimSpace(asString(s["installState"])), "failed") {
			continue
		}
		deviceName := asString(s["deviceDisplayName"])
		if deviceName == "" {
			deviceName = asString(s["deviceName"])
		}
		rows = append(rows, []string{
			deviceName,
			asString(s["deviceId"]),
			asString(s["installState"]),
			asString(s["lastSyncDateTime"]),
		})
	}
	if len(rows) == 0 {
		return fmt.Sprintf("Failing App Drill-Down\n\nApp: %s\n\nNo failed device statuses found.", asString(app["displayName"])), nil
	}
	return fmt.Sprintf("Failing App Drill-Down\n\nApp: %s\nFailed devices: %d\n\n%s",
		asString(app["displayName"]),
		len(rows),
		renderTable([]string{"Device Name", "Device ID", "State", "Last Sync"}, rows),
	), nil
}

func renderInspector(title string, values [][2]string) string {
	rows := make([][]string, 0, len(values))
	for _, pair := range values {
		rows = append(rows, []string{pair[0], pair[1]})
	}
	return fmt.Sprintf("%s\n\n%s", title, renderTable([]string{"Field", "Value"}, rows))
}

func (g *graphClient) inspectUser(ctx context.Context, identifier string) (string, error) {
	var user map[string]any
	if strings.Contains(identifier, "@") {
		filter := url.QueryEscape(fmt.Sprintf("userPrincipalName eq '%s'", escapeOData(identifier)))
		items, err := g.list(ctx, "/users?$select=id,displayName,userPrincipalName,accountEnabled&$filter="+filter)
		if err != nil {
			return "", err
		}
		user, err = selectUniqueMatch("user", identifier, items, formatUserCandidate)
		if errors.Is(err, errNotFound) {
			return "", errors.New("user not found")
		}
		if err != nil {
			return "", err
		}
	} else {
		body, err := g.do(ctx, http.MethodGet, graphBase+"/users/"+url.PathEscape(identifier)+"?$select=id,displayName,userPrincipalName,accountEnabled", nil)
		if err == nil {
			if err := json.Unmarshal(body, &user); err != nil {
				return "", err
			}
		} else if isGraphForbidden(err) {
			return "", errors.New("access denied: insufficient permissions to read this user")
		} else if isGraphNotFound(err) {
			user, err = g.findUserByDisplayName(ctx, identifier)
			if err != nil {
				return "", err
			}
		} else {
			return "", err
		}
	}
	return renderInspector("User Inspector", [][2]string{
		{"Display Name", asString(user["displayName"])},
		{"UPN", asString(user["userPrincipalName"])},
		{"Object ID", asString(user["id"])},
		{"Enabled", asString(user["accountEnabled"])},
	}), nil
}

func (g *graphClient) inspectGroup(ctx context.Context, identifier string) (string, error) {
	var group map[string]any
	body, err := g.do(ctx, http.MethodGet, graphBase+"/groups/"+url.PathEscape(identifier)+"?$select=id,displayName,description,mailNickname,securityEnabled,mailEnabled", nil)
	if err == nil {
		if err := json.Unmarshal(body, &group); err != nil {
			return "", err
		}
	} else if isGraphForbidden(err) {
		return "", errors.New("access denied: insufficient permissions to read this group")
	} else if isGraphNotFound(err) {
		group, err = g.findUniqueByDisplayName(ctx, "/groups", "id,displayName,description,mailNickname,securityEnabled,mailEnabled", "group", identifier, formatGroupCandidate)
		if errors.Is(err, errNotFound) {
			return "", errors.New("group not found")
		}
		if err != nil {
			return "", err
		}
	} else {
		return "", err
	}
	return renderInspector("Group Inspector", [][2]string{
		{"Display Name", asString(group["displayName"])},
		{"Description", asString(group["description"])},
		{"Object ID", asString(group["id"])},
		{"Mail Nickname", asString(group["mailNickname"])},
		{"Security Enabled", asString(group["securityEnabled"])},
		{"Mail Enabled", asString(group["mailEnabled"])},
	}), nil
}

func (g *graphClient) inspectDevice(ctx context.Context, identifier string) (string, error) {
	var device map[string]any
	body, err := g.do(ctx, http.MethodGet, graphBase+"/devices/"+url.PathEscape(identifier)+"?$select=id,displayName,deviceId,operatingSystem,accountEnabled", nil)
	if err == nil {
		if err := json.Unmarshal(body, &device); err != nil {
			return "", err
		}
	} else if isGraphForbidden(err) {
		return "", errors.New("access denied: insufficient permissions to read this device")
	} else if isGraphNotFound(err) {
		device, err = g.findDeviceByDisplayName(ctx, identifier)
		if err != nil {
			return "", err
		}
	} else {
		return "", err
	}
	return renderInspector("Device Inspector", [][2]string{
		{"Display Name", asString(device["displayName"])},
		{"Object ID", asString(device["id"])},
		{"Device ID", asString(device["deviceId"])},
		{"Operating System", asString(device["operatingSystem"])},
		{"Enabled", asString(device["accountEnabled"])},
	}), nil
}

func (g *graphClient) inspectApp(ctx context.Context, identifier string) (string, error) {
	var app map[string]any
	body, err := g.do(ctx, http.MethodGet, graphBase+"/deviceAppManagement/mobileApps/"+url.PathEscape(identifier)+"?$select=id,displayName,publisher", nil)
	if err == nil {
		if err := json.Unmarshal(body, &app); err != nil {
			return "", err
		}
	} else if isGraphForbidden(err) {
		return "", errors.New("access denied: insufficient permissions to read this app")
	} else if isGraphNotFound(err) {
		app, err = g.findAppByDisplayName(ctx, identifier)
		if err != nil {
			return "", err
		}
	} else {
		return "", err
	}

	assignments, assignErr := g.list(ctx, fmt.Sprintf("/deviceAppManagement/mobileApps/%s/assignments?$select=id", url.PathEscape(asString(app["id"]))))
	assignmentCount := fmt.Sprintf("N/A (%v)", assignErr)
	if assignErr == nil {
		assignmentCount = fmt.Sprintf("%d", len(assignments))
	}

	return renderInspector("App Inspector", [][2]string{
		{"Display Name", asString(app["displayName"])},
		{"Publisher", asString(app["publisher"])},
		{"Object ID", asString(app["id"])},
		{"Assignment Count", assignmentCount},
	}), nil
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

type csvDataset struct {
	Headers         []string
	OriginalHeaders []string
	Rows            []map[string]string
}

func countCSVSeverity(issues []csvIssue, severity string) int {
	total := 0
	for _, issue := range issues {
		if issue.Severity == severity {
			total++
		}
	}
	return total
}

func normalizeCSVHeaders(headers []string, requiredHeaders map[string]struct{}) ([]string, []csvIssue) {
	normalized := make([]string, len(headers))
	issues := make([]csvIssue, 0)
	seen := map[string]int{}
	for i, h := range headers {
		trimmed := strings.TrimSpace(h)
		normalized[i] = trimmed
		if trimmed != h {
			issues = append(issues, csvIssue{
				Severity: "Warning",
				Row:      1,
				Field:    h,
				Code:     "HEADER_WHITESPACE",
				Message:  "Header has leading/trailing whitespace",
			})
		}
		if first, exists := seen[trimmed]; exists {
			issues = append(issues, csvIssue{
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
				issues = append(issues, csvIssue{
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

func hasCSVError(issues []csvIssue) bool {
	for _, issue := range issues {
		if issue.Severity == "Error" {
			return true
		}
	}
	return false
}

func readCSVNormalized(path string) (csvDataset, error) {
	f, err := os.Open(path)
	if err != nil {
		return csvDataset{}, err
	}
	defer f.Close()
	r := csv.NewReader(f)
	rows, err := r.ReadAll()
	if err != nil {
		return csvDataset{}, err
	}
	if len(rows) == 0 {
		return csvDataset{}, io.EOF
	}
	normalizedHeaders, issues := normalizeCSVHeaders(rows[0], nil)
	if hasCSVError(issues) {
		return csvDataset{}, errors.New(formatValidationReport("CSV Parse Failed", csvValidationResult{
			FilePath: path,
			Errors:   countCSVSeverity(issues, "Error"),
			Warnings: countCSVSeverity(issues, "Warning"),
			Issues:   issues,
			Pass:     false,
		}))
	}
	if len(rows) < 2 {
		return csvDataset{}, errNoCSVDataRows
	}
	data := csvDataset{
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

func readCSV(path string) ([]map[string]string, error) {
	data, err := readCSVNormalized(path)
	if err != nil {
		return nil, err
	}
	return data.Rows, nil
}

func validateCSVStrict(path string, requiredHeaders, keyColumns []string) (csvValidationResult, error) {
	res := csvValidationResult{FilePath: path, Pass: true}
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
	normalizedHeaders, issues := normalizeCSVHeaders(rows[0], requiredSet)
	res.Issues = append(res.Issues, issues...)
	headerMap := map[string]int{}
	for i, h := range normalizedHeaders {
		if _, exists := headerMap[h]; !exists {
			headerMap[h] = i
		}
	}
	for _, req := range requiredHeaders {
		if _, ok := headerMap[req]; !ok {
			res.Issues = append(res.Issues, csvIssue{
				Severity: "Error", Row: 1, Field: req, Code: "MISSING_HEADER",
				Message: "Required header is missing",
			})
		}
	}
	if len(rows) < 2 {
		res.Issues = append(res.Issues, csvIssue{
			Severity: "Error",
			Row:      1,
			Field:    "",
			Code:     "NO_DATA_ROWS",
			Message:  "CSV contains headers but no data rows",
		})
	}
	if hasCSVError(res.Issues) {
		res.Errors = countCSVSeverity(res.Issues, "Error")
		res.Warnings = countCSVSeverity(res.Issues, "Warning")
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
				res.Issues = append(res.Issues, csvIssue{
					Severity: "Error", Row: rowNum, Field: req, Code: "MISSING_REQUIRED_VALUE",
					Message: "Required value is empty",
				})
			}
		}
		for _, k := range keyColumns {
			keyParts = append(keyParts, strings.ToLower(item[k]))
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
	res.Errors = countCSVSeverity(res.Issues, "Error")
	res.Warnings = countCSVSeverity(res.Issues, "Warning")
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

func previewForAction(spec actionSpec, inputs []string) (string, bool, error) {
	switch spec.id {
	case actAddUsersCSV:
		data, err := readCSVNormalized(inputs[0])
		if err != nil {
			return "", true, err
		}
		sample := make([][]string, 0, minInt(10, len(data.Rows)))
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
			renderTable([]string{"User Principal Name"}, sample),
		), true, nil
	case actMakeGroupsCSV:
		data, err := readCSVNormalized(inputs[0])
		if err != nil {
			return "", true, err
		}
		sample := make([][]string, 0, minInt(10, len(data.Rows)))
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
			renderTable([]string{"Group Name"}, sample),
		), true, nil
	case actAddAppsCSV:
		data, err := readCSVNormalized(inputs[0])
		if err != nil {
			return "", true, err
		}
		sample := make([][]string, 0, minInt(10, len(data.Rows)))
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
			renderTable([]string{"Group Name", "App Name"}, sample),
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

func (g *graphClient) addUsersCSV(ctx context.Context, csvPath, groupName string, dryRun bool) (string, error) {
	data, err := readCSVNormalized(csvPath)
	if err != nil {
		return "", err
	}
	group, err := g.findGroupByDisplayName(ctx, groupName)
	if err != nil {
		return "", err
	}
	groupID := asString(group["id"])

	var b strings.Builder
	var added, failed, skipped int
	for _, row := range data.Rows {
		upn := row["User_Principal_Name"]
		if upn == "" {
			fmt.Fprintf(&b, "Skipped row: missing User_Principal_Name\n")
			skipped++
			continue
		}
		filter := url.QueryEscape(fmt.Sprintf("userPrincipalName eq '%s'", escapeOData(upn)))
		users, err := g.list(ctx, "/users?$select=id,userPrincipalName&$filter="+filter)
		if err != nil {
			fmt.Fprintf(&b, "Failed to look up user %s: %v\n", upn, err)
			failed++
			continue
		}
		if len(users) == 0 {
			fmt.Fprintf(&b, "User not found: %s\n", upn)
			failed++
			continue
		}
		userID := asString(users[0]["id"])
		body := map[string]string{
			"@odata.id": fmt.Sprintf("https://graph.microsoft.com/v1.0/directoryObjects/%s", userID),
		}
		if dryRun {
			fmt.Fprintf(&b, "Would add %s\n", upn)
			added++
			continue
		}
		_, err = g.do(ctx, http.MethodPost, fmt.Sprintf("%s/groups/%s/members/$ref", graphBase, groupID), body)
		if err != nil {
			fmt.Fprintf(&b, "Failed to add %s: %v\n", upn, err)
			failed++
			continue
		}
		fmt.Fprintf(&b, "Added %s\n", upn)
		added++
	}
	fmt.Fprintf(&b, "\nSummary: %d added, %d failed, %d skipped", added, failed, skipped)
	return b.String(), nil
}

func (g *graphClient) makeGroupsCSV(ctx context.Context, csvPath string, dryRun bool) (string, error) {
	data, err := readCSVNormalized(csvPath)
	if err != nil {
		return "", err
	}
	var b strings.Builder
	var created, failed, skipped int
	for _, row := range data.Rows {
		groupName := row["Group_Name"]
		if groupName == "" {
			fmt.Fprintf(&b, "Skipped row: missing Group_Name\n")
			skipped++
			continue
		}
		_, err := g.findGroupByDisplayName(ctx, groupName)
		if err == nil {
			fmt.Fprintf(&b, "Exists: %s\n", groupName)
			skipped++
			continue
		}
		if !errors.Is(err, errNotFound) {
			fmt.Fprintf(&b, "Failed to check existing group %s: %v\n", groupName, err)
			failed++
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
			created++
			continue
		}
		_, err = g.do(ctx, http.MethodPost, graphBase+"/groups", body)
		if err != nil {
			fmt.Fprintf(&b, "Failed to create %s: %v\n", groupName, err)
			failed++
			continue
		}
		fmt.Fprintf(&b, "Created: %s\n", groupName)
		created++
	}
	fmt.Fprintf(&b, "\nSummary: %d created, %d failed, %d skipped", created, failed, skipped)
	return b.String(), nil
}

func (g *graphClient) addAppsCSV(ctx context.Context, csvPath string, dryRun bool) (string, error) {
	data, err := readCSVNormalized(csvPath)
	if err != nil {
		return "", err
	}
	var b strings.Builder
	var assigned, failed, skipped int
	for _, row := range data.Rows {
		appName := row["App_Name"]
		groupName := row["Group_Name"]
		if appName == "" || groupName == "" {
			fmt.Fprintf(&b, "Skipped row: missing App_Name or Group_Name\n")
			skipped++
			continue
		}
		app, err := g.findAppByDisplayName(ctx, appName)
		if err != nil {
			fmt.Fprintf(&b, "App lookup failed for %s: %v\n", appName, err)
			failed++
			continue
		}
		appID := asString(app["id"])

		group, err := g.findGroupByDisplayName(ctx, groupName)
		if err != nil {
			fmt.Fprintf(&b, "Group lookup failed for %s: %v\n", groupName, err)
			failed++
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
			assigned++
			continue
		}
		_, err = g.do(ctx, http.MethodPost, fmt.Sprintf("%s/deviceAppManagement/mobileApps/%s/assignments", graphBase, appID), body)
		if err != nil {
			fmt.Fprintf(&b, "Failed assignment app=%s group=%s: %v\n", appName, groupName, err)
			failed++
			continue
		}
		fmt.Fprintf(&b, "Assigned %s -> %s\n", appName, groupName)
		assigned++
	}
	fmt.Fprintf(&b, "\nSummary: %d assigned, %d failed, %d skipped", assigned, failed, skipped)
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
	groupNameCache := map[string]string{}
	resolveGroup := func(groupID string) string {
		if name, ok := groupNameCache[groupID]; ok {
			return name
		}
		b, err := g.do(ctx, http.MethodGet, graphBase+"/groups/"+groupID+"?$select=displayName", nil)
		if err != nil {
			groupNameCache[groupID] = ""
			return ""
		}
		var grp map[string]any
		if json.Unmarshal(b, &grp) != nil {
			groupNameCache[groupID] = ""
			return ""
		}
		name := asString(grp["displayName"])
		groupNameCache[groupID] = name
		return name
	}
	for i, app := range apps {
		if (i+1)%20 == 0 {
			g.emitProgress(fmt.Sprintf("Processed %d/%d apps...", i+1, len(apps)))
		}
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
			groupName := resolveGroup(groupID)
			if groupName == "" {
				continue
			}
			rows = append(rows, row{
				AppName:      asString(app["displayName"]),
				GroupName:    groupName,
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

type model struct {
	client        *graphClient
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
		spin:           sp,
		styles:         newUIStyles(),
		input:          ti,
		filterInput:    fi,
		exportInput:    ei,
		drillInput:     di,
		searchInput:    si,
		searchMatchLine: -1,
		viewport:       viewport.New(80, 20),
		progressCh:     make(chan progressMsg, 64),
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
			// Scroll viewport so the matched line is visible.
			lineOffset := maxInt(0, idx-m.viewport.Height/2)
			m.viewport.SetYOffset(lineOffset)
			return
		}
	}
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

	lines := []string{
		fmt.Sprintf("Action: %s", m.lastActionLabel),
		fmt.Sprintf("Tenant: %s", m.client.cfg.TenantID),
		fmt.Sprintf("Client: %s", m.client.cfg.ClientID),
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
	if h, r, ok := parseTableFromText(text); ok {
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
			"d: Drill into top failing apps when available",
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
	// Drain stale progress messages from previous operations.
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
		ctx := m.cancelCtx
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
		case actReportWindowsBreakdown:
			out, err = m.client.reportWindowsBreakdown(ctx)
		case actReportTopFailingApps:
			out, err = m.client.reportTopFailingApps(ctx)
		case actInspectUser:
			out, err = m.client.inspectUser(ctx, inputs[0])
		case actInspectGroup:
			out, err = m.client.inspectGroup(ctx, inputs[0])
		case actInspectDevice:
			out, err = m.client.inspectDevice(ctx, inputs[0])
		case actInspectApp:
			out, err = m.client.inspectApp(ctx, inputs[0])
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
		m.drillInput.Width = minInt(72, maxInt(24, msg.Width-16))
		panelInnerW := maxInt(20, msg.Width-12)
		panelInnerH := maxInt(6, msg.Height-18)
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
				if m.state == stateMain {
					return m, tea.Quit
				}
				m.resetMenuPosition(parentMenuState(m.state))
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
						// Re-send Enter so the selection logic runs.
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
				// Early validation: check CSV file exists before asking remaining prompts.
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
						m.setOutput(formatValidationReport("CSV Validation Failed (Write Blocked)", res))
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
					out, err := m.client.reportAppFailureDetails(ctx, appName)
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

func (m model) View() string {
	switch m.state {
	case stateDrillPrompt:
		body := m.styles.panel.Render(fmt.Sprintf("%s\n\n%s\n\n%s",
			m.styles.subHeader.Render("Top Failing Apps Drill-Down"),
			m.drillInput.View(),
			m.styles.hint.Render("Enter rank or exact app name   Esc: cancel"),
		))
		return m.styles.app.Render(m.styles.header.Render(" Intune Management Tool ") + "\n\n" + body)
	case stateHelp:
		body := m.styles.panel.Render(fmt.Sprintf("%s\n\n%s\n\n%s",
			m.styles.subHeader.Render("Keyboard Help"),
			helpTextForState(m.helpReturnState),
			m.styles.hint.Render("Enter/Esc/?: close help"),
		))
		return m.styles.app.Render(m.styles.header.Render(" Intune Management Tool ") + "\n\n" + body)
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
		escHint := "Esc: cancel"
		if len(m.inputs) > 0 {
			escHint = "Esc: back"
		}
		body := m.styles.panel.Render(fmt.Sprintf("%s\n\n%s\n\n%s",
			title,
			m.input.View(),
			m.styles.hint.Render("Enter: continue   "+escHint),
		))
		return m.styles.app.Render(m.styles.header.Render(" Intune Management Tool ") + "\n\n" + body)
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
		return m.styles.app.Render(m.styles.header.Render(" Intune Management Tool ") + "\n\n" + headerPanel + "\n\n" + body)
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
		return m.styles.app.Render(m.styles.header.Render(" Intune Management Tool ") + "\n\n" + headerPanel + "\n\n" + body)
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
		screen := m.styles.header.Render(" Intune Management Tool ") + "\n\n" +
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
