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
	"strings"

	"intune-management/internal/csvutil"
	"intune-management/internal/render"
)

func (g *Client) ListUsersInGroup(ctx context.Context, groupName string) (string, error) {
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
		render.RenderTable([]string{"Display Name", "UPN", "Object ID"}, rows),
	), nil
}

func (g *Client) ListUsers(ctx context.Context) (string, error) {
	users, err := g.list(ctx, "/users?$select=id,displayName,userPrincipalName")
	if err != nil {
		return "", err
	}
	rows := make([][]string, 0, len(users))
	for _, u := range users {
		rows = append(rows, []string{asString(u["displayName"]), asString(u["userPrincipalName"]), asString(u["id"])})
	}
	return fmt.Sprintf("Users: %d\n\n%s", len(users), render.RenderTable([]string{"Display Name", "UPN", "Object ID"}, rows)), nil
}

func (g *Client) ListGroups(ctx context.Context) (string, error) {
	groups, err := g.list(ctx, "/groups?$select=id,displayName")
	if err != nil {
		return "", err
	}
	rows := make([][]string, 0, len(groups))
	for _, grp := range groups {
		rows = append(rows, []string{asString(grp["displayName"]), asString(grp["id"])})
	}
	return fmt.Sprintf("Groups: %d\n\n%s", len(groups), render.RenderTable([]string{"Display Name", "Object ID"}, rows)), nil
}

func (g *Client) SearchGroups(ctx context.Context, term string) (string, error) {
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
	return fmt.Sprintf("Groups matching %q: %d\n\n%s", term, len(rows), render.RenderTable([]string{"Display Name", "Object ID"}, rows)), nil
}

func (g *Client) ListDevices(ctx context.Context) (string, error) {
	devices, err := g.list(ctx, "/devices?$select=id,deviceId,displayName")
	if err != nil {
		return "", err
	}
	rows := make([][]string, 0, len(devices))
	for _, d := range devices {
		rows = append(rows, []string{asString(d["displayName"]), asString(d["id"]), asString(d["deviceId"])})
	}
	return fmt.Sprintf("Devices: %d\n\n%s", len(devices), render.RenderTable([]string{"Display Name", "Object ID", "Device ID"}, rows)), nil
}

func (g *Client) InspectUser(ctx context.Context, identifier string) (string, error) {
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
	return render.RenderInspector("User Inspector", [][2]string{
		{"Display Name", asString(user["displayName"])},
		{"UPN", asString(user["userPrincipalName"])},
		{"Object ID", asString(user["id"])},
		{"Enabled", asString(user["accountEnabled"])},
	}), nil
}

func (g *Client) InspectGroup(ctx context.Context, identifier string) (string, error) {
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
	return render.RenderInspector("Group Inspector", [][2]string{
		{"Display Name", asString(group["displayName"])},
		{"Description", asString(group["description"])},
		{"Object ID", asString(group["id"])},
		{"Mail Nickname", asString(group["mailNickname"])},
		{"Security Enabled", asString(group["securityEnabled"])},
		{"Mail Enabled", asString(group["mailEnabled"])},
	}), nil
}

func (g *Client) InspectDevice(ctx context.Context, identifier string) (string, error) {
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
	return render.RenderInspector("Device Inspector", [][2]string{
		{"Display Name", asString(device["displayName"])},
		{"Object ID", asString(device["id"])},
		{"Device ID", asString(device["deviceId"])},
		{"Operating System", asString(device["operatingSystem"])},
		{"Enabled", asString(device["accountEnabled"])},
	}), nil
}

func (g *Client) InspectApp(ctx context.Context, identifier string) (string, error) {
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

	return render.RenderInspector("App Inspector", [][2]string{
		{"Display Name", asString(app["displayName"])},
		{"Publisher", asString(app["publisher"])},
		{"Object ID", asString(app["id"])},
		{"Assignment Count", assignmentCount},
	}), nil
}

func (g *Client) ListDevicesInGroup(ctx context.Context, groupName string) (string, error) {
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
		render.RenderTable([]string{"Display Name", "Object ID", "Device ID"}, rows),
	), nil
}

func (g *Client) AddUsersCSV(ctx context.Context, csvPath, groupName string, dryRun bool) (string, error) {
	data, err := csvutil.ReadNormalized(csvPath)
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

func (g *Client) MakeGroupsCSV(ctx context.Context, csvPath string, dryRun bool) (string, error) {
	data, err := csvutil.ReadNormalized(csvPath)
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

func (g *Client) AddAppsCSV(ctx context.Context, csvPath string, dryRun bool) (string, error) {
	data, err := csvutil.ReadNormalized(csvPath)
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

func valueOrDash(s string) string {
	if s == "" {
		return "-"
	}
	return s
}

func friendlyAppType(odataType string) string {
	suffix := strings.TrimPrefix(odataType, "#microsoft.graph.")
	if suffix == odataType {
		return odataType // no prefix found
	}
	switch suffix {
	case "win32LobApp":
		return "Win32"
	case "windowsMicrosoftEdgeApp":
		return "Edge"
	case "winGetApp":
		return "WinGet"
	case "microsoftStoreForBusinessApp":
		return "Store for Business"
	case "officeSuiteApp":
		return "Microsoft 365"
	case "windowsUniversalAppX":
		return "UWP/APPX"
	case "windowsMobileMSI":
		return "MSI"
	case "windowsWebApp":
		return "Windows Web"
	case "webApp":
		return "Web"
	case "iosStoreApp":
		return "iOS Store"
	case "iosVppApp":
		return "iOS VPP"
	case "managedIOSStoreApp":
		return "Managed iOS"
	case "androidStoreApp":
		return "Android Store"
	case "androidManagedStoreApp":
		return "Android Managed"
	case "androidForWorkApp":
		return "Android for Work"
	case "managedAndroidStoreApp":
		return "Managed Android"
	case "macOSDmgApp":
		return "macOS DMG"
	case "macOSLobApp":
		return "macOS LOB"
	case "macOSMicrosoftEdgeApp":
		return "macOS Edge"
	case "macOSMicrosoftDefenderApp":
		return "macOS Defender"
	case "macOSOfficeSuiteApp":
		return "macOS Office"
	default:
		return suffix
	}
}

func (g *Client) ListGroupApps(ctx context.Context) (string, error) {
	apps, err := g.list(ctx, "/deviceAppManagement/mobileApps?$select=id,@odata.type,displayName,isAssigned,publisher,displayVersion")
	if err != nil {
		return "", err
	}
	type row struct {
		AppName   string
		AppType   string
		Publisher string
		Version   string
		GroupName string
		Intent    string
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
	// Build batch requests — one per app for assignments.
	reqs := make([]batchRequest, len(apps))
	for i, app := range apps {
		reqs[i] = batchRequest{
			ID:     fmt.Sprintf("%d", i),
			Method: "GET",
			URL:    fmt.Sprintf("/deviceAppManagement/mobileApps/%s/assignments?$select=id,intent,target", asString(app["id"])),
		}
	}
	responses, err := g.batch(ctx, reqs)
	if err != nil {
		return "", err
	}
	if len(responses) != len(apps) {
		return "", fmt.Errorf("batch response count %d does not match app count %d", len(responses), len(apps))
	}
	skipped := 0
	hasAssignment := make([]bool, len(apps))
	for i, resp := range responses {
		result, err := parseBatchValues(resp)
		if err != nil {
			skipped++
			continue
		}
		assignments := result.Values
		// Fall back to sequential pagination for the rare paginated response.
		if result.NextLink != "" {
			extra, err := g.listFromURL(ctx, result.NextLink)
			if err != nil {
				skipped++
				continue
			}
			assignments = append(result.Values, extra...)
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
			hasAssignment[i] = true
			rows = append(rows, row{
				AppName:   asString(apps[i]["displayName"]),
				AppType:   friendlyAppType(asString(apps[i]["@odata.type"])),
				Publisher: valueOrDash(asString(apps[i]["publisher"])),
				Version:   valueOrDash(asString(apps[i]["displayVersion"])),
				GroupName: groupName,
				Intent:    asString(a["intent"]),
			})
		}
	}

	// Add orphaned/unassigned apps. Apps with isAssigned=true but no group
	// rows (e.g. all-users/all-devices targeting) are intentionally excluded —
	// they aren't orphaned, they just don't target a specific group.
	for i, app := range apps {
		if hasAssignment[i] {
			continue
		}
		assigned, _ := app["isAssigned"].(bool)
		if assigned {
			continue
		}
		rows = append(rows, row{
			AppName:   asString(app["displayName"]),
			AppType:   friendlyAppType(asString(app["@odata.type"])),
			Publisher: valueOrDash(asString(app["publisher"])),
			Version:   valueOrDash(asString(app["displayVersion"])),
			GroupName: "(none)",
			Intent:    "(unassigned)",
		})
	}

	var b strings.Builder
	if len(rows) == 0 {
		return "No group app assignments found.", nil
	}
	tabRows := make([][]string, 0, len(rows))
	for _, r := range rows {
		tabRows = append(tabRows, []string{r.AppName, r.AppType, r.Publisher, r.Version, r.GroupName, r.Intent})
	}
	fmt.Fprintf(&b, "App-group assignments: %d\n\n%s", len(rows), render.RenderTable([]string{"App", "Type", "Publisher", "Version", "Group", "Intent"}, tabRows))
	if skipped > 0 {
		fmt.Fprintf(&b, "\n(%d apps skipped due to errors)", skipped)
	}
	return b.String(), nil
}
