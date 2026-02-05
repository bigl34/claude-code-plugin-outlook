# claude-code-plugin-outlook

Personal Outlook/MS365 email and calendar operations for Claude Code.

![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)

## Quick Start

```bash
git clone https://github.com/your-username/claude-code-plugin-outlook.git
cd claude-code-plugin-outlook && npm install && npm run build
node scripts/dist/cli.js login
node scripts/dist/cli.js list-messages --top 5
```

## Features

*   **Authentication**: Secure login and session verification with Microsoft 365.
*   **Email Management**: List inbox, search, read, move, and delete messages.
*   **Composition**: Send emails or create drafts directly from the CLI.
*   **Folder Access**: Browse and manage email folders.
*   **Calendar Operations**: View, create, update, and delete calendar events.
*   **Contact & Tasks**: Access your address book and task lists.
*   **Search**: Global search capabilities across the MS365 account.
*   **JSON Output**: Structured data responses for easy integration.

## Prerequisites

*   Node.js >= 18
*   Claude Code CLI
*   Microsoft 365 / Outlook Account
*   Azure App Registration (Client ID & Secret) with appropriate Graph API scopes

## Installation

1.  Clone the repository:
    ```bash
    git clone https://github.com/your-username/claude-code-plugin-outlook.git
    cd claude-code-plugin-outlook
    ```

2.  Install dependencies:
    ```bash
    npm install
    ```

3.  Build the TypeScript CLI:
    ```bash
    npm run build
    ```

## Configuration

1.  Copy the template configuration file:
    ```bash
    cp config.template.json config.json
    ```

2.  Edit `config.json` and add your Azure App credentials:
    ```json
    {
      "clientId": "YOUR_CLIENT_ID",
      "tenantId": "YOUR_TENANT_ID",
      "clientSecret": "YOUR_CLIENT_SECRET"
    }
    ```

## Available Commands

### Authentication Commands
| Command | Description | Required Options |
|---------|-------------|------------------|
| `login` | Authenticate with Microsoft | (none) |
| `verify-login` | Check authentication status | (none) |

### Mail Commands
| Command | Description | Required Options |
|---------|-------------|------------------|
| `list-messages` | List inbox messages | (optional: `--top`, `--skip`, `--filter`) |
| `list-folders` | List mail folders | (none) |
| `list-folder-messages` | List messages in folder | `--folder-id` |
| `get-message` | Get message details | `--id` |
| `send-mail` | Send email | `--to`, `--subject`, `--body` |
| `create-draft` | Create draft email | `--to`, `--subject`, `--body` |
| `move-message` | Move message to folder | `--id`, `--folder-id` |
| `delete-message` | Delete message | `--id` |

### Calendar Commands
| Command | Description | Required Options |
|---------|-------------|------------------|
| `list-calendars` | List calendars | (none) |
| `list-events` | List calendar events | (optional: `--calendar-id`, `--top`) |
| `get-event` | Get event details | `--id` |
| `get-calendar-view` | Get events in date range | `--start`, `--end` |
| `create-event` | Create event | `--subject`, `--start`, `--end` |
| `update-event` | Update event | `--id` |
| `delete-event` | Delete event | `--id` |

### Other Commands
| Command | Description | Required Options |
|---------|-------------|------------------|
| `list-contacts` | List contacts | (optional: `--top`, `--skip`) |
| `list-tasks` | List tasks | (optional: `--list-id`, `--top`) |
| `search` | Search across MS365 | `--query` |

## Usage Examples

**List recent inbox messages:**
```bash
node scripts/dist/cli.js list-messages --top 10
```

**Send an email:**
```bash
node scripts/dist/cli.js send-mail --to "friend@example.com" --subject "Hello" --body "How are you?"
```

**Get calendar events for a specific week:**
```bash
node scripts/dist/cli.js get-calendar-view --start "2024-01-01T00:00:00Z" --end "2024-01-07T23:59:59Z"
```

## How it Works

This plugin acts as a bridge between Claude Code and the Microsoft Graph API. It uses a local Node.js CLI wrapper to perform authenticated HTTP requests against your Outlook account, returning JSON data that the agent parses to perform actions.

## Troubleshooting

*   **"Not authenticated" error**: Run `node scripts/dist/cli.js login` and follow the device code prompt.
*   **"Client secret missing"**: Ensure you have created `config.json` with valid Azure credentials.
*   **Build errors**: Make sure you have installed all dependencies with `npm install` and that you are using Node.js 18+.
*   **Permission denied**: Check your Azure App Registration scopes to ensure the app has permission to read/write mail and calendars.
*   **Command not found**: Verify you are running the command from the root directory and that `scripts/dist/cli.js` exists.

## Contributing

Issues and pull requests are welcome.

## License

MIT
