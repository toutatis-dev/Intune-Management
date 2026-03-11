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
	"errors"
	"fmt"
	"net/url"
	"strings"
)

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

func (g *Client) findUniqueByDisplayName(ctx context.Context, path, selectFields, objectType, name string, formatter func(map[string]any) string) (map[string]any, error) {
	filter := url.QueryEscape(fmt.Sprintf("displayName eq '%s'", escapeOData(name)))
	items, err := g.list(ctx, fmt.Sprintf("%s?$select=%s&$filter=%s", path, selectFields, filter))
	if err != nil {
		return nil, err
	}
	return selectUniqueMatch(objectType, name, items, formatter)
}

func (g *Client) findGroupByDisplayName(ctx context.Context, name string) (map[string]any, error) {
	group, err := g.findUniqueByDisplayName(ctx, "/groups", "id,displayName", "group", name, formatGroupCandidate)
	if errors.Is(err, errNotFound) {
		return nil, fmt.Errorf("group %w", errNotFound)
	}
	return group, err
}

func (g *Client) findUserByDisplayName(ctx context.Context, name string) (map[string]any, error) {
	user, err := g.findUniqueByDisplayName(ctx, "/users", "id,displayName,userPrincipalName,accountEnabled", "user", name, formatUserCandidate)
	if errors.Is(err, errNotFound) {
		return nil, fmt.Errorf("user %w", errNotFound)
	}
	return user, err
}

func (g *Client) findDeviceByDisplayName(ctx context.Context, name string) (map[string]any, error) {
	device, err := g.findUniqueByDisplayName(ctx, "/devices", "id,displayName,deviceId,operatingSystem,accountEnabled", "device", name, formatDeviceCandidate)
	if errors.Is(err, errNotFound) {
		return nil, fmt.Errorf("device %w", errNotFound)
	}
	return device, err
}

func (g *Client) findAppByDisplayName(ctx context.Context, name string) (map[string]any, error) {
	app, err := g.findUniqueByDisplayName(ctx, "/deviceAppManagement/mobileApps", "id,displayName,publisher", "app", name, formatAppCandidate)
	if errors.Is(err, errNotFound) {
		return nil, fmt.Errorf("app %w", errNotFound)
	}
	return app, err
}
