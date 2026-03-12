# Contributing to Intune Management Tool

First off, thank you for considering contributing to this project! We welcome all contributions, from bug reports and feature requests to code contributions and documentation improvements.

## Development Setup

1. **Prerequisites**: Ensure you have [Go](https://golang.org/doc/install) (version 1.22 or later) installed.
2. **Fork and Clone**: Fork the repository on GitHub, then clone your fork locally:
   ```bash
   git clone https://github.com/<your-username>/Intune-Management.git
   cd Intune-Management
   ```
3. **Run vet and tests**:
   ```bash
   go vet ./...
   go test -race ./...
   ```
4. **Build the application**:
   ```bash
   go build -o intune-management.exe ./cmd/intune-management
   ```

## Project Structure

This project follows standard Go project layouts:
- `cmd/intune-management/`: Contains the main entry point for the CLI application.
- `internal/app/`: Contains the TUI (Terminal User Interface) logic using the Bubble Tea framework.
- `internal/graph/`: Contains Microsoft Graph API interaction logic (auth, requests, lookups, and reports).
- `internal/csvutil/`: Contains CSV validation, linting, and parsing utilities.
- `internal/render/`: Contains UI rendering and text-based output formatting logic.
- `internal/config/`: Contains configuration handling for user and tenant IDs.

## Pull Request Guidelines

1. **Fork the repository** and create your branch from `main`.
2. **Format and vet your code**: We enforce `go fmt` and `go vet` in our CI pipeline. Before committing, always run:
   ```bash
   go fmt ./...
   go vet ./...
   ```
3. **Write tests**: If you are adding a new feature or fixing a bug, please include tests. Run `go test -race ./...` to ensure your changes don't break existing functionality or introduce data races.
4. **Use Conventional Commits**: Try to use standard commit prefixes like `feat:`, `fix:`, `docs:`, `style:`, `refactor:`, etc.
5. **Open a Pull Request**: Provide a clear description of the problem you are solving or the feature you are adding, and attach a screenshot if it involves a UI change.

## Reporting Issues

If you find a bug or have a feature request, please use the GitHub Issues tab. Provide as much context as possible, including your Operating System, exact error messages, and steps to reproduce the issue.
