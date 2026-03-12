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

import "testing"

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
