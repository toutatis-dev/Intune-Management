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
package config

import (
	"encoding/json"
	"fmt"
	"os"
	"path/filepath"
	"regexp"
)

const DefaultClientID = "14d82eec-204b-4c2f-b7e8-296a70dab67e"

type AuthConfig struct {
	ClientID string
	TenantID string
}

var uuidPattern = regexp.MustCompile(`^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$`)

// domainPattern matches a valid DNS hostname with at least one dot, e.g. contoso.onmicrosoft.com.
var domainPattern = regexp.MustCompile(`^[a-zA-Z0-9]([a-zA-Z0-9\-]{0,61}[a-zA-Z0-9])?(\.[a-zA-Z0-9]([a-zA-Z0-9\-]{0,61}[a-zA-Z0-9])?)+$`)

// Validate checks that ClientID and TenantID are well-formed.
// ClientID must be a UUID. TenantID must be a UUID, a domain name
// (e.g. contoso.onmicrosoft.com), or one of the well-known Azure AD
// aliases ("common", "organizations", "consumers").
func (c AuthConfig) Validate() error {
	if !uuidPattern.MatchString(c.ClientID) {
		return fmt.Errorf("invalid client ID %q: must be a UUID (xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx)", c.ClientID)
	}
	switch c.TenantID {
	case "common", "organizations", "consumers":
		return nil
	default:
		if !uuidPattern.MatchString(c.TenantID) && !domainPattern.MatchString(c.TenantID) {
			return fmt.Errorf("invalid tenant ID %q: must be a UUID, a domain (e.g. contoso.onmicrosoft.com), or one of \"common\", \"organizations\", \"consumers\"", c.TenantID)
		}
	}
	return nil
}

func FilePath() string {
	exe, err := os.Executable()
	if err != nil {
		return "intune-management.config.json"
	}
	return filepath.Join(filepath.Dir(exe), "intune-management.config.json")
}

func LoadFromFile() (AuthConfig, error) {
	path := FilePath()
	b, err := os.ReadFile(path)
	if err != nil {
		return AuthConfig{}, err
	}
	var cfg AuthConfig
	if err := json.Unmarshal(b, &cfg); err != nil {
		return AuthConfig{}, err
	}
	return cfg, nil
}

func SaveToFile(cfg AuthConfig) error {
	path := FilePath()
	b, err := json.MarshalIndent(cfg, "", "  ")
	if err != nil {
		return err
	}
	return os.WriteFile(path, b, 0600)
}

func Resolve() AuthConfig {
	cfg := AuthConfig{}
	if fileCfg, err := LoadFromFile(); err == nil {
		cfg = fileCfg
	}
	if envClient := os.Getenv("GRAPH_CLIENT_ID"); envClient != "" {
		cfg.ClientID = envClient
	}
	if envTenant := os.Getenv("GRAPH_TENANT_ID"); envTenant != "" {
		cfg.TenantID = envTenant
	}
	if cfg.ClientID == "" {
		cfg.ClientID = DefaultClientID
	}
	if cfg.TenantID == "" {
		cfg.TenantID = "common"
	}
	return cfg
}
