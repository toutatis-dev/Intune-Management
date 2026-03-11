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
