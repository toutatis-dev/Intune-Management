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
	"context"
	"encoding/json"
	"errors"
	"fmt"
	"net/http"
	"net/url"
	"sort"
	"strconv"
	"strings"

	"intune-management/internal/render"
)

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

func (g *Client) ReportComplianceSnapshot(ctx context.Context) (string, error) {
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
		render.RenderTable([]string{"Compliance State", "Count", "Percent"}, rows),
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

func (g *Client) ReportWindowsBreakdown(ctx context.Context) (string, error) {
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
		render.RenderTable([]string{"Category", "Count", "Percent"}, rows),
	), nil
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
		render.RenderTable([]string{"Rank", "App", "App ID", "Failed Devices", "Total Statuses", "Failure Rate"}, rows),
	)
}

func (g *Client) ReportTopFailingApps(ctx context.Context) (string, error) {
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

func (g *Client) ReportAppFailureDetails(ctx context.Context, identifier string) (string, error) {
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
		render.RenderTable([]string{"Device Name", "Device ID", "State", "Last Sync"}, rows),
	), nil
}
