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
	"os"
	"path/filepath"
)

const DefaultClientID = "14d82eec-204b-4c2f-b7e8-296a70dab67e"

type AuthConfig struct {
	ClientID string
	TenantID string
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
