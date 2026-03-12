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
	"errors"
	"fmt"
	"os"
	"path/filepath"
	"regexp"
)

const DefaultClientID = "14d82eec-204b-4c2f-b7e8-296a70dab67e"

const configFileName = "intune-management.config.json"

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

var userConfigDirFunc = os.UserConfigDir

// UserConfigDir returns the user-scoped configuration directory for
// intune-management, creating it with mode 0700 if it does not exist.
func UserConfigDir() (string, error) {
	base, err := userConfigDirFunc()
	if err != nil {
		return "", fmt.Errorf("cannot determine user config directory: %w", err)
	}
	dir := filepath.Join(base, "intune-management")
	if err := os.MkdirAll(dir, 0700); err != nil {
		return "", fmt.Errorf("cannot create config directory: %w", err)
	}
	return dir, nil
}

// SafeReadFile reads a file after verifying it is not a symlink.
func SafeReadFile(path string) ([]byte, error) {
	fi, err := os.Lstat(path)
	if err != nil {
		return nil, err
	}
	if fi.Mode()&os.ModeSymlink != 0 {
		return nil, errors.New("refusing to read symlink: " + path)
	}
	return os.ReadFile(path)
}

// FilePath returns the path to the config file to read from. It searches:
// 1. The exe directory (IT-managed defaults)
// 2. The user config directory (user overrides)
// If neither exists, returns the user config dir path (where saves will go).
func FilePath() (string, error) {
	// Check exe directory first
	if exe, err := os.Executable(); err == nil {
		exePath := filepath.Join(filepath.Dir(exe), configFileName)
		if fi, err := os.Lstat(exePath); err == nil && fi.Mode()&os.ModeSymlink == 0 {
			return exePath, nil
		}
	}

	// Check / default to user config directory
	userDir, err := UserConfigDir()
	if err != nil {
		return "", err
	}
	return filepath.Join(userDir, configFileName), nil
}

func LoadFromFile() (AuthConfig, error) {
	path, err := FilePath()
	if err != nil {
		return AuthConfig{}, err
	}
	b, err := SafeReadFile(path)
	if err != nil {
		return AuthConfig{}, err
	}
	var cfg AuthConfig
	if err := json.Unmarshal(b, &cfg); err != nil {
		return AuthConfig{}, err
	}
	return cfg, nil
}

// SaveToFile always writes to the user config directory.
func SaveToFile(cfg AuthConfig) error {
	userDir, err := UserConfigDir()
	if err != nil {
		return err
	}
	path := filepath.Join(userDir, configFileName)
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
