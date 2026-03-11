# Intune Management Tool

A fast, interactive Terminal User Interface (TUI) application written in Go for managing Microsoft Intune and Azure Active Directory (Entra ID) via the Microsoft Graph API.

## Features

### Users and Groups
* **List Users & Groups:** View all users and groups in the directory.
* **Search Groups:** Find groups by partial display name.
* **Group Members:** List all user or device members of a specific group.
* **Bulk Add Users (CSV):** Add users to a group in bulk by their `User_Principal_Name`.

### Devices and Apps
* **List Devices:** Show all managed devices across the tenant.
* **Bulk Create Groups (CSV):** Quickly create multiple security groups by `Group_Name`.
* **Assign Apps (CSV):** Assign applications to specific groups by `App_Name` and `Group_Name`.
* **View Assignments:** Review existing app-group assignments directly in the TUI.

### Reports and Analytics
* **Compliance Snapshot:** Get a quick view of compliant vs. non-compliant devices.
* **Windows OS Breakdown:** See the distribution of Windows 10, Windows 11, and other operating systems.
* **Failing App Deployments:** View the top 10 applications with the highest failure rates, with drill-down capabilities for troubleshooting specific devices.
* **Object Inspector:** Deep dive into specific users, groups, devices, or applications.
* **CSV Validation:** Run strict validation and linting on CSV data *before* performing operations, preventing syntax and schema errors.

### Configuration & Safety
* **Custom Auth:** Configure your own Microsoft Graph Client ID and Tenant ID.
* **Dry-Run Mode:** Simulate operations and view exactly what API calls would be executed, without actually modifying your directory.

## Installation

### Download Pre-compiled Binary (Recommended)
The latest stable releases for Windows, Linux, and macOS are available in the [Releases](../../releases) section of this repository.

1. Go to the [Releases](../../releases) page.
2. Download the executable for your operating system.
3. Extract (if necessary) and run the executable directly.

### Building from Source
If you prefer to build from source, ensure you have [Go 1.21+](https://golang.org/doc/install) installed.

1. Clone the repository:
   ```bash
   git clone https://github.com/toutatis-dev/Intune-Management.git
   cd intune-management
   ```
2. Build the application:
   ```bash
   go build -o intune-management.exe ./cmd/intune-management
   ```

## Usage

Start the tool by running the compiled executable:

```bash
./intune-management.exe
```

### Authentication
The tool uses Microsoft Graph's **Device Code Flow**. When an API call is made, the tool will present a device code. Visit [microsoft.com/devicelogin](https://microsoft.com/devicelogin), enter the code, and sign in with an account that has appropriate Intune/Azure AD permissions.

#### Custom App Registration (Recommended)
By default, the tool uses the well-known Microsoft Graph PowerShell Client ID. While this works out-of-the-box for many tenants, it is highly recommended to create your own Azure AD (Entra ID) App Registration to adhere to security best practices and avoid potential conditional access blocks.

1. Go to the **Azure Portal** > **App registrations** > **New registration**.
2. Name the application (e.g., "Intune Management TUI").
3. Under **Supported account types**, select **Accounts in this organizational directory only**.
4. Click **Register**.
5. In the left menu, go to **Authentication**. Under **Advanced settings**, set **Allow public client flows** to **Yes** and save.
6. Go to **API permissions** and add the following **Delegated** Microsoft Graph permissions:
   - `User.Read.All`
   - `Group.ReadWrite.All`
   - `Device.Read.All`
   - `DeviceManagementApps.ReadWrite.All`
   - `DeviceManagementManagedDevices.Read.All`
7. Click **Grant admin consent** for your tenant.

Once created, launch the tool, navigate to **Settings**, and update the **Graph Client ID** and **Graph Tenant ID** with the values from your new App Registration.

### CSV Formats

When running bulk operations, ensure your CSV files contain the required headers:

* **Bulk Add Users:** Requires `User_Principal_Name`
* **Create Groups:** Requires `Group_Name`
* **Assign Apps:** Requires `Group_Name` and `App_Name`

*Tip: Use the "Reports > CSV validation checks" menu to test your CSV files before executing real operations.*

## Keyboard Controls

* **Up/Down** or **j/k**: Move selection
* **PgUp/PgDn**: Move by page
* **Home/End**: Jump to top or bottom
* **1-9**: Jump to a menu item by number
* **Enter**: Select item or confirm input
* **/**: Filter menu items, or search within output results
* **e**: Export the current data table to a CSV file
* **d**: Drill down into an app (Top 10 Failing Apps report only)
* **Esc**: Go back to the previous screen or cancel the current operation
* **q**: Quit the application
* **?**: Open the help menu
