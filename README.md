<!-- AUTO-GENERATED README — DO NOT EDIT. Changes will be overwritten on next publish. -->
# claude-code-plugin-outlook

Personal Outlook/MS365 email and calendar operations

![Version](https://img.shields.io/badge/version-1.2.6-blue) ![License: MIT](https://img.shields.io/badge/License-MIT-green) ![Node >= 18](https://img.shields.io/badge/node-%3E%3D18-brightgreen)

## Features

- **login** — Authenticate with Microsoft
- **verify-login** — Check authentication status

## Prerequisites

- [Node.js](https://nodejs.org/) >= 18
- [Claude Code](https://docs.anthropic.com/en/docs/claude-code) CLI
- MCP server binary for the target service (configured via `config.json`)

## Quick Start

```bash
git clone https://github.com/YOUR_GITHUB_USER/claude-code-plugin-outlook.git
cd claude-code-plugin-outlook
cp config.template.json config.json  # fill in your credentials
cd scripts && npm install
```

```bash
node scripts/dist/cli.js login
```

## Installation

1. Clone this repository
2. Copy `config.template.json` to `config.json` and fill in your credentials
3. Install dependencies:
   ```bash
   cd scripts && npm install
   ```
4. Ensure the MCP server binary is available on your system (see the service's documentation)

## Available Commands

| Command        | Description                 | Required Options |
| -------------- | --------------------------- | ---------------- |
| `login`        | Authenticate with Microsoft | (none)           |
| `verify-login` | Check authentication status | (none)           |

## Usage Examples

```bash
# List recent inbox messages
node scripts/dist/cli.js list-messages --top 10

# Get a specific email
node scripts/dist/cli.js get-message --id "AAMkAG..."

# Send an email
node scripts/dist/cli.js send-mail --to "friend@example.com" --subject "Hello" --body "How are you?"

# Get calendar events for a date range
node scripts/dist/cli.js get-calendar-view --start "2024-01-01T00:00:00Z" --end "2024-01-07T23:59:59Z"

# Search for emails
node scripts/dist/cli.js search --query "invoice"
```

## How It Works

This plugin wraps an MCP (Model Context Protocol) server, providing a CLI interface that communicates with the service's MCP binary. The CLI translates commands into MCP tool calls and returns structured JSON responses.

## Troubleshooting

| Issue | Solution |
|-------|----------|
| Authentication errors | Verify credentials in `config.json` |
| `ERR_MODULE_NOT_FOUND` | Run `cd scripts && npm install` |
| MCP connection timeout | Ensure the MCP server binary is installed and accessible |
| Rate limiting | The CLI handles retries automatically; wait and retry if persistent |
| Unexpected JSON output | Check API credentials haven't expired |

## Contributing

Issues and pull requests are welcome.

## License

MIT
