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
	"os"
	"path/filepath"
	"runtime"
	"testing"
)

func TestAuthConfigValidate(t *testing.T) {
	t.Parallel()

	tests := []struct {
		name    string
		cfg     AuthConfig
		wantErr bool
	}{
		{
			name:    "valid UUID client and tenant",
			cfg:     AuthConfig{ClientID: "14d82eec-204b-4c2f-b7e8-296a70dab67e", TenantID: "a1b2c3d4-e5f6-7890-abcd-ef1234567890"},
			wantErr: false,
		},
		{
			name:    "valid with common tenant",
			cfg:     AuthConfig{ClientID: "14d82eec-204b-4c2f-b7e8-296a70dab67e", TenantID: "common"},
			wantErr: false,
		},
		{
			name:    "valid with organizations tenant",
			cfg:     AuthConfig{ClientID: "14d82eec-204b-4c2f-b7e8-296a70dab67e", TenantID: "organizations"},
			wantErr: false,
		},
		{
			name:    "valid with consumers tenant",
			cfg:     AuthConfig{ClientID: "14d82eec-204b-4c2f-b7e8-296a70dab67e", TenantID: "consumers"},
			wantErr: false,
		},
		{
			name:    "empty client ID",
			cfg:     AuthConfig{ClientID: "", TenantID: "common"},
			wantErr: true,
		},
		{
			name:    "malformed client ID",
			cfg:     AuthConfig{ClientID: "not-a-uuid", TenantID: "common"},
			wantErr: true,
		},
		{
			name:    "empty tenant ID",
			cfg:     AuthConfig{ClientID: "14d82eec-204b-4c2f-b7e8-296a70dab67e", TenantID: ""},
			wantErr: true,
		},
		{
			name:    "valid onmicrosoft.com domain tenant",
			cfg:     AuthConfig{ClientID: "14d82eec-204b-4c2f-b7e8-296a70dab67e", TenantID: "contoso.onmicrosoft.com"},
			wantErr: false,
		},
		{
			name:    "valid custom domain tenant",
			cfg:     AuthConfig{ClientID: "14d82eec-204b-4c2f-b7e8-296a70dab67e", TenantID: "contoso.com"},
			wantErr: false,
		},
		{
			name:    "malformed tenant ID - bare label without dot",
			cfg:     AuthConfig{ClientID: "14d82eec-204b-4c2f-b7e8-296a70dab67e", TenantID: "my-tenant"},
			wantErr: true,
		},
	}

	for _, tt := range tests {
		tt := tt
		t.Run(tt.name, func(t *testing.T) {
			t.Parallel()
			err := tt.cfg.Validate()
			if (err != nil) != tt.wantErr {
				t.Fatalf("Validate() error = %v, wantErr %v", err, tt.wantErr)
			}
		})
	}
}

func TestUserConfigDirCreatesDirectory(t *testing.T) {
	t.Parallel()

	dir, err := UserConfigDir()
	if err != nil {
		t.Fatalf("UserConfigDir() error: %v", err)
	}
	if dir == "" {
		t.Fatal("UserConfigDir() returned empty string")
	}
	fi, err := os.Stat(dir)
	if err != nil {
		t.Fatalf("directory does not exist: %v", err)
	}
	if !fi.IsDir() {
		t.Fatalf("expected directory, got %v", fi.Mode())
	}
}

func TestFilePathFallsBackToUserDir(t *testing.T) {
	// No t.Parallel() — mutates package-level userConfigDirFunc.

	// Point UserConfigDir at a temp directory so the test doesn't
	// depend on real %AppData% state.
	tmp := t.TempDir()
	origFunc := userConfigDirFunc
	userConfigDirFunc = func() (string, error) { return tmp, nil }
	t.Cleanup(func() { userConfigDirFunc = origFunc })

	// When no config file exists next to the exe, FilePath should
	// return a path inside the user config directory.
	path, err := FilePath()
	if err != nil {
		t.Fatalf("FilePath() error: %v", err)
	}
	expected := filepath.Join(tmp, "intune-management", configFileName)
	if path != expected {
		t.Fatalf("FilePath() = %q, want %q", path, expected)
	}
}

func TestSaveToFileWritesToUserDir(t *testing.T) {
	// No t.Parallel() — mutates package-level userConfigDirFunc.

	// Point UserConfigDir at a temp directory to avoid writing to real %AppData%.
	tmp := t.TempDir()
	origFunc := userConfigDirFunc
	userConfigDirFunc = func() (string, error) { return tmp, nil }
	t.Cleanup(func() { userConfigDirFunc = origFunc })

	cfg := AuthConfig{
		ClientID: "14d82eec-204b-4c2f-b7e8-296a70dab67e",
		TenantID: "common",
	}
	if err := SaveToFile(cfg); err != nil {
		t.Fatalf("SaveToFile() error: %v", err)
	}

	userDir, err := UserConfigDir()
	if err != nil {
		t.Fatalf("UserConfigDir() error: %v", err)
	}
	path := filepath.Join(userDir, configFileName)
	if _, err := os.Stat(path); err != nil {
		t.Fatalf("config file not found at user dir: %v", err)
	}
}

func TestSafeReadFileRejectsSymlinks(t *testing.T) {
	if runtime.GOOS == "windows" {
		t.Skip("symlink creation requires elevated privileges on Windows")
	}
	t.Parallel()

	dir := t.TempDir()
	target := filepath.Join(dir, "real.txt")
	if err := os.WriteFile(target, []byte("secret"), 0600); err != nil {
		t.Fatal(err)
	}
	link := filepath.Join(dir, "link.txt")
	if err := os.Symlink(target, link); err != nil {
		t.Fatal(err)
	}

	// Reading the real file should work
	if _, err := SafeReadFile(target); err != nil {
		t.Fatalf("SafeReadFile(real file) error: %v", err)
	}

	// Reading the symlink should fail
	if _, err := SafeReadFile(link); err == nil {
		t.Fatal("SafeReadFile(symlink) should have returned an error")
	}
}

func TestSafeReadFileReadsRegularFile(t *testing.T) {
	t.Parallel()

	dir := t.TempDir()
	path := filepath.Join(dir, "test.txt")
	want := []byte("hello world")
	if err := os.WriteFile(path, want, 0600); err != nil {
		t.Fatal(err)
	}
	got, err := SafeReadFile(path)
	if err != nil {
		t.Fatalf("SafeReadFile() error: %v", err)
	}
	if string(got) != string(want) {
		t.Fatalf("SafeReadFile() = %q, want %q", got, want)
	}
}
