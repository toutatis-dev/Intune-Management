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
package graph

import (
	"bytes"
	"context"
	"encoding/base64"
	"encoding/json"
	"errors"
	"fmt"
	"io"
	"net/http"
	"os"
	"os/exec"
	"path/filepath"
	"strconv"
	"strings"
	"time"

	"github.com/Azure/azure-sdk-for-go/sdk/azcore"
	"github.com/Azure/azure-sdk-for-go/sdk/azcore/policy"
	"github.com/Azure/azure-sdk-for-go/sdk/azidentity"
	"github.com/Azure/azure-sdk-for-go/sdk/azidentity/cache"
	"github.com/atotto/clipboard"

	"intune-management/internal/config"
	"intune-management/internal/render"
)

var requiredScopes = []string{
	"https://graph.microsoft.com/User.Read.All",
	"https://graph.microsoft.com/Group.ReadWrite.All",
	"https://graph.microsoft.com/Device.Read.All",
	"https://graph.microsoft.com/DeviceManagementApps.ReadWrite.All",
	"https://graph.microsoft.com/DeviceManagementManagedDevices.Read.All",
}

const graphBase = "https://graph.microsoft.com/v1.0"

// authenticator extends azcore.TokenCredential with the Authenticate method
// shared by InteractiveBrowserCredential and DeviceCodeCredential.
type authenticator interface {
	azcore.TokenCredential
	Authenticate(ctx context.Context, opts *policy.TokenRequestOptions) (azidentity.AuthenticationRecord, error)
}

type Client struct {
	cred         authenticator
	http         *http.Client
	scope        []string
	cfg          config.AuthConfig
	authMethod   string // "browser" or "device_code"
	progressHook func(string)
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

var errNotFound = errors.New("not found")

var authRecordFilePath = func() string {
	exe, err := os.Executable()
	if err != nil {
		return "intune-management.auth.json"
	}
	return filepath.Join(filepath.Dir(exe), "intune-management.auth.json")
}

func loadAuthRecord(cfg config.AuthConfig) (azidentity.AuthenticationRecord, bool) {
	var record azidentity.AuthenticationRecord
	b, err := os.ReadFile(authRecordFilePath())
	if err != nil {
		return record, false
	}
	if err := json.Unmarshal(b, &record); err != nil {
		return record, false
	}
	// Discard cached record if it doesn't match the current config.
	// Using a stale record from a different tenant or app registration
	// would silently authenticate against the wrong environment.
	if record.TenantID != "" && cfg.TenantID != "common" && !strings.EqualFold(record.TenantID, cfg.TenantID) {
		return azidentity.AuthenticationRecord{}, false
	}
	if record.ClientID != "" && !strings.EqualFold(record.ClientID, cfg.ClientID) {
		return azidentity.AuthenticationRecord{}, false
	}
	return record, true
}

func saveAuthRecord(record azidentity.AuthenticationRecord) error {
	b, err := json.Marshal(record)
	if err != nil {
		return err
	}
	return os.WriteFile(authRecordFilePath(), b, 0600)
}

func NewClient() (*Client, error) {
	return NewClientWithConfig(config.Resolve())
}

func NewClientWithConfig(cfg config.AuthConfig) (*Client, error) {
	if err := cfg.Validate(); err != nil {
		return nil, err
	}

	var tokenCache azidentity.Cache
	if c, err := cache.New(&cache.Options{Name: "intune-management"}); err == nil {
		tokenCache = c
	}

	record, hasRecord := loadAuthRecord(cfg)

	// Try interactive browser auth first — supports passkeys, conditional
	// access, and all MFA methods.  Falls back to device code for headless
	// or SSH sessions where a browser isn't available.
	cred, method, err := newBrowserCredential(cfg, tokenCache, record, hasRecord)
	if err != nil {
		cred, method, err = newDeviceCodeCredential(cfg, tokenCache, record, hasRecord)
		if err != nil {
			return nil, err
		}
	}

	return &Client{
		cred:       cred,
		http:       &http.Client{Timeout: 60 * time.Second},
		scope:      requiredScopes,
		cfg:        cfg,
		authMethod: method,
	}, nil
}

const redirectURL = "http://localhost"

func newBrowserCredential(cfg config.AuthConfig, tokenCache azidentity.Cache, record azidentity.AuthenticationRecord, hasRecord bool) (authenticator, string, error) {
	opts := &azidentity.InteractiveBrowserCredentialOptions{
		ClientID:                       cfg.ClientID,
		TenantID:                       cfg.TenantID,
		Cache:                          tokenCache,
		DisableAutomaticAuthentication: true,
		RedirectURL:                    redirectURL,
	}
	if hasRecord {
		opts.AuthenticationRecord = record
	}
	cred, err := azidentity.NewInteractiveBrowserCredential(opts)
	if err != nil {
		return nil, "", err
	}
	return cred, "browser", nil
}

func newDeviceCodeCredential(cfg config.AuthConfig, tokenCache azidentity.Cache, record azidentity.AuthenticationRecord, hasRecord bool) (authenticator, string, error) {
	opts := &azidentity.DeviceCodeCredentialOptions{
		ClientID:                       cfg.ClientID,
		TenantID:                       cfg.TenantID,
		Cache:                          tokenCache,
		DisableAutomaticAuthentication: true,
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
	}
	if hasRecord {
		opts.AuthenticationRecord = record
	}
	cred, err := azidentity.NewDeviceCodeCredential(opts)
	if err != nil {
		return nil, "", err
	}
	return cred, "device_code", nil
}

func (g *Client) Config() config.AuthConfig {
	return g.cfg
}

func (g *Client) SetProgressHook(hook func(string)) {
	g.progressHook = hook
}

func (g *Client) emitProgress(text string) {
	if g.progressHook != nil {
		g.progressHook(text)
	}
}

func (g *Client) getToken(ctx context.Context) (string, error) {
	opts := policy.TokenRequestOptions{Scopes: g.scope}
	token, err := g.cred.GetToken(ctx, opts)
	if err != nil {
		if !needsInteraction(err) {
			return "", err
		}
		record, authErr := g.cred.Authenticate(ctx, &opts)
		if authErr != nil {
			// Only fall back to device code when the browser could not be
			// launched (headless/SSH).  User cancellation, CA blocks, and
			// transient errors should surface directly so the caller can
			// retry browser auth rather than silently switching flows.
			if g.authMethod == "browser" && isBrowserLaunchFailure(authErr) {
				return g.fallbackToDeviceCode(ctx)
			}
			return "", authErr
		}
		if saveErr := saveAuthRecord(record); saveErr != nil {
			fmt.Fprintf(os.Stderr, "\u26a0 Warning: could not persist auth record: %v\n", saveErr)
		}
		token, err = g.cred.GetToken(ctx, opts)
		if err != nil {
			return "", err
		}
		return token.Token, nil
	}
	return token.Token, nil
}

func needsInteraction(err error) bool {
	var authErr *azidentity.AuthenticationRequiredError
	if errors.As(err, &authErr) {
		return true
	}
	msg := err.Error()
	return strings.Contains(msg, "AuthenticationRequiredError") || strings.Contains(msg, "interaction is required")
}

// isBrowserLaunchFailure returns true when the error indicates the system
// browser could not be opened (e.g. headless server, no $DISPLAY, missing
// xdg-open).  It intentionally excludes user cancellation, CA policy blocks,
// and transient network errors so those surface to the caller for retry.
func isBrowserLaunchFailure(err error) bool {
	var execErr *exec.Error
	if errors.As(err, &execErr) {
		return true
	}
	msg := strings.ToLower(err.Error())
	return strings.Contains(msg, "could not open browser") ||
		strings.Contains(msg, "no browser") ||
		strings.Contains(msg, "cannot open") ||
		strings.Contains(msg, "display is not set")
}

func (g *Client) fallbackToDeviceCode(ctx context.Context) (string, error) {
	var tokenCache azidentity.Cache
	if c, err := cache.New(&cache.Options{Name: "intune-management"}); err == nil {
		tokenCache = c
	}
	record, hasRecord := loadAuthRecord(g.cfg)
	cred, _, err := newDeviceCodeCredential(g.cfg, tokenCache, record, hasRecord)
	if err != nil {
		return "", err
	}
	g.cred = cred
	g.authMethod = "device_code"

	opts := policy.TokenRequestOptions{Scopes: g.scope}
	authRecord, err := g.cred.Authenticate(ctx, &opts)
	if err != nil {
		return "", err
	}
	if saveErr := saveAuthRecord(authRecord); saveErr != nil {
		fmt.Fprintf(os.Stderr, "\u26a0 Warning: could not persist auth record: %v\n", saveErr)
	}
	tok, err := g.cred.GetToken(ctx, opts)
	if err != nil {
		return "", err
	}
	return tok.Token, nil
}

func (g *Client) do(ctx context.Context, method, fullURL string, body any) ([]byte, error) {
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
		token, err := g.getToken(ctx)
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
		req.Header.Set("Authorization", "Bearer "+token)
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

func minInt(a, b int) int {
	if a < b {
		return a
	}
	return b
}

func (g *Client) list(ctx context.Context, path string) ([]map[string]any, error) {
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

// NewStubClient creates a Client with the given config but no credentials.
// Useful for testing UI components that don't make API calls.
func NewStubClient(cfg config.AuthConfig) *Client {
	return &Client{cfg: cfg}
}

func (g *Client) AuthHealth(ctx context.Context) (string, error) {
	token, err := g.getToken(ctx)
	if err != nil {
		return "", err
	}
	claims, err := decodeJWTClaims(token)
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

	return render.RenderInspector("Auth Health", [][2]string{
		{"Auth Method", g.authMethod},
		{"Configured Client ID", g.cfg.ClientID},
		{"Configured Tenant ID", g.cfg.TenantID},
		{"Token Client ID", effectiveClient},
		{"Token Tenant ID", effectiveTenant},
		{"Token Expires (UTC)", exp},
		{"Token Scopes", tokenScopes},
		{"Requested Scopes", strings.Join(g.scope, ", ")},
	}), nil
}
