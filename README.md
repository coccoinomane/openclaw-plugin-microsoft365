# OpenClaw Microsoft 365 Plugin

Connect a Microsoft 365 account to OpenClaw via the [CLI for Microsoft 365](https://github.com/pnp/cli-microsoft365) (`m365`).

This plugin intentionally does **not** implement a custom Microsoft Graph SDK wrapper. It provides:

- an OpenClaw plugin manifest and runtime entry;
- a skill and bootstrap guide for agents;
- `resolve_exec_env` defaults for `m365` authentication (`CLIMICROSOFT365_ENTRAAPPID`, `CLIMICROSOFT365_TENANT`);
- convenient `M365_*` environment variables for agent snippets;
- command patterns for Outlook, Calendar, SharePoint, OneDrive/Office files, OneNote, To Do, Planner, Teams and raw Microsoft Graph.

## What this plugin does and does not do

This plugin helps OpenClaw agents use Microsoft 365 safely and consistently through `m365` and Microsoft Graph.

It does:

- load an OpenClaw skill for Microsoft 365 workflows;
- contribute Microsoft 365 environment defaults to `exec` calls;
- document the Entra app, OAuth/device-code, permission and smoke-test setup;
- prefer official CLI commands and `m365 request --url` for raw Graph endpoints.

It does **not**:

- ask for or store a Microsoft 365 password;
- bypass Microsoft Entra consent;
- create a custom Graph token cache outside `m365`;
- send email by default. Agents should draft first and send only after explicit user request plus confirmation.

## Requirements

- OpenClaw `>= 2026.6.6`.
- Node.js available to run `npx`, or a globally installed `m365` binary.
- A Microsoft Entra app registration configured as a public client.
- Delegated Microsoft Graph permissions granted to that app.
- A user account that can complete device-code login and consent/admin-consent where required.

## Backend command

Use the official CLI for Microsoft 365. A persistent install is optional:

```bash
npm install -g @pnp/cli-microsoft365
m365 status --output json
```

Agents and scripts should tolerate systems where `m365` is not installed globally. Use a shell function rather than storing a command with spaces in a variable; this works in both bash and zsh:

```bash
m365_cli() {
  if command -v m365 >/dev/null 2>&1; then
    m365 "$@"
  else
    npx -y -p @pnp/cli-microsoft365 m365 "$@"
  fi
}

m365_cli status --output json
```

## Install or load the plugin in OpenClaw

Preferred install, matching the other Git-managed OpenClaw Casa plugins:

```bash
openclaw plugins install git:github.com/coccoinomane/openclaw-plugin-microsoft365@v0.1.1
```

Then enable/configure the plugin in OpenClaw config:

```json5
{
  plugins: {
    entries: {
      microsoft365: {
        enabled: true,
        config: {
          tenantId: "${M365_TENANT_ID}",
          clientId: "${M365_CLIENT_ID}",
          account: "${M365_ACCOUNT}",
          connectionName: "${M365_CONNECTION_NAME}"
        }
      }
    }
  }
}
```

For local plugin development, use a checkout outside the main OpenClaw Casa repo and load that path temporarily. Do not add this plugin as a git submodule of OpenClaw Casa.

Restart OpenClaw after first install/enable if the running Gateway has not hot-reloaded the new plugin registry. The hook only contributes `exec` environment variables after the plugin has been loaded by the Gateway.

The runtime hook contributes these variables to `exec` calls when configured:

- `CLIMICROSOFT365_TENANT`
- `CLIMICROSOFT365_ENTRAAPPID`
- `M365_TENANT_ID`
- `M365_CLIENT_ID`
- `M365_ACCOUNT`
- `M365_CONNECTION_NAME`

The actual Graph permissions in an access token come from the delegated permissions configured and consented on the Entra app, not from an OpenClaw environment variable.

## Connect a Microsoft 365 account

### 1. Create or configure the Entra app

In Microsoft Entra admin center:

1. Open **App registrations** → **New registration**.
2. Choose a clear name, for example `OpenClaw Microsoft 365`.
3. Select the account type appropriate for your tenant. For a single organization, use single-tenant.
4. After creation, copy:
   - **Application (client) ID** → `M365_CLIENT_ID`
   - **Directory (tenant) ID** → `M365_TENANT_ID`
5. Open **Authentication**.
6. Enable **Allow public client flows** / **Treat application as a public client**.
7. Add a platform if needed: **Mobile and desktop applications**.
8. Add redirect URI `http://localhost/openclaw-m365` if it is not already present.

If Entra says the redirect URI must have distinct values, that usually means the URI already exists on another platform entry for the same app. Do not add it twice; continue with the existing URI.

Do not create a client secret for the device-code flow. The app is used as a public client.

### 2. Add Microsoft Graph delegated permissions

In the app registration, open **API permissions** → **Add a permission** → **Microsoft Graph** → **Delegated permissions**.

For broad Google-Workspace-like coverage, start from:

```text
openid
profile
email
offline_access
User.Read
Mail.ReadWrite
Mail.Send
MailboxSettings.ReadWrite
Calendars.ReadWrite
Contacts.ReadWrite
People.Read.All
Files.ReadWrite.All
Sites.ReadWrite.All
Notes.ReadWrite
Notes.ReadWrite.All
Tasks.ReadWrite
Tasks.ReadWrite.Shared
Group.ReadWrite.All
Directory.Read.All
Team.ReadBasic.All
Channel.ReadBasic.All
ChannelMessage.Send
Chat.ReadWrite
OnlineMeetings.ReadWrite
```

Then click **Grant admin consent** for the tenant when required.

Many broad permissions require tenant admin consent. If the token later contains only `User.Read profile openid email`, the broader permissions have not been granted/consented yet, or the user needs to log out and log in again after consent.

PowerPoint, Word and Excel are normally handled as Office files in OneDrive/SharePoint through `Files.*` and `Sites.*` permissions; there is no single Graph delegated permission equivalent to full Google Slides-style editing.

Optional CLI equivalent, for admins with sufficient privileges:

```bash
m365_cli entra app permission add \
  --appId "$M365_CLIENT_ID" \
  --delegatedPermissions "https://graph.microsoft.com/Mail.ReadWrite https://graph.microsoft.com/Mail.Send https://graph.microsoft.com/MailboxSettings.ReadWrite https://graph.microsoft.com/Calendars.ReadWrite https://graph.microsoft.com/Contacts.ReadWrite https://graph.microsoft.com/People.Read.All https://graph.microsoft.com/Files.ReadWrite.All https://graph.microsoft.com/Sites.ReadWrite.All https://graph.microsoft.com/Notes.ReadWrite https://graph.microsoft.com/Notes.ReadWrite.All https://graph.microsoft.com/Tasks.ReadWrite https://graph.microsoft.com/Tasks.ReadWrite.Shared https://graph.microsoft.com/Group.ReadWrite.All https://graph.microsoft.com/Directory.Read.All https://graph.microsoft.com/Team.ReadBasic.All https://graph.microsoft.com/Channel.ReadBasic.All https://graph.microsoft.com/ChannelMessage.Send https://graph.microsoft.com/Chat.ReadWrite https://graph.microsoft.com/OnlineMeetings.ReadWrite" \
  --grantAdminConsent
```

This command changes Entra app permissions. Agents should not run it unless the user explicitly asks and approves the external configuration change.

### 3. Set environment variables

Set these in the environment that runs OpenClaw, for example `~/.openclaw/.env`:

```bash
M365_TENANT_ID=<tenant-guid-or-verified-domain>
M365_CLIENT_ID=<application-client-id>
M365_ACCOUNT=user@example.com
M365_CONNECTION_NAME=openclaw-microsoft365
```

The official CLI recognizes:

```bash
export CLIMICROSOFT365_TENANT="$M365_TENANT_ID"
export CLIMICROSOFT365_ENTRAAPPID="$M365_CLIENT_ID"
```

The plugin contributes these automatically to OpenClaw `exec` calls after it is loaded.

### 4. Login with device code

Run:

```bash
m365_cli login \
  --authType deviceCode \
  --appId "$M365_CLIENT_ID" \
  --tenant "$M365_TENANT_ID" \
  --connectionName "${M365_CONNECTION_NAME:-openclaw-microsoft365}"
```

Open the URL shown by the CLI, enter the device code, and authenticate with the Microsoft 365 account.

If permissions were changed after a previous login, refresh the connection:

```bash
m365_cli logout
m365_cli login \
  --authType deviceCode \
  --appId "$M365_CLIENT_ID" \
  --tenant "$M365_TENANT_ID" \
  --connectionName "${M365_CONNECTION_NAME:-openclaw-microsoft365}"
```

### 5. Verify the connection

Status:

```bash
m365_cli status --output json
```

Current user:

```bash
m365_cli request \
  --url "https://graph.microsoft.com/v1.0/me?\$select=id,displayName,userPrincipalName,mail" \
  --output json
```

Token scopes, useful for debugging consent:

```bash
m365_cli util accesstoken get \
  --resource https://graph.microsoft.com \
  --decoded \
  --output json | grep '"scp"'
```

Read-only smoke tests:

```bash
m365_cli request \
  --url "https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages?\$top=3&\$select=id,receivedDateTime,from,subject,isRead,webLink&\$orderby=receivedDateTime desc" \
  --output json

m365_cli request \
  --url "https://graph.microsoft.com/v1.0/me/calendarView?startDateTime=2026-06-16T00:00:00%2B02:00&endDateTime=2026-06-23T00:00:00%2B02:00&\$top=3&\$select=id,subject,start,end,webLink&\$orderby=start/dateTime" \
  --output json

m365_cli request \
  --url "https://graph.microsoft.com/v1.0/me/drive/root/children?\$top=3&\$select=id,name,webUrl,size,lastModifiedDateTime,file,folder" \
  --output json
```

## Raw Graph through `m365 request`

Use high-level `m365` commands when they are expressive enough. Use `m365 request --url` for exact Microsoft Graph endpoints, richer filters/projections, attachments, or unsupported resources.

Example inbox query:

```bash
m365_cli request --url "https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages?\$top=10&\$select=id,receivedDateTime,from,subject,bodyPreview,isRead,importance,webLink&\$orderby=receivedDateTime desc" --output json
```

Example calendar view:

```bash
m365_cli request --url "https://graph.microsoft.com/v1.0/me/calendarView?startDateTime=2026-06-15T00:00:00%2B02:00&endDateTime=2026-06-22T00:00:00%2B02:00&\$select=id,subject,start,end,location,organizer,attendees,isOnlineMeeting,onlineMeeting,webLink&\$orderby=start/dateTime" --output json
```

Example OneDrive root listing:

```bash
m365_cli request --url "https://graph.microsoft.com/v1.0/me/drive/root/children?\$select=id,name,webUrl,size,lastModifiedDateTime,file,folder" --output json
```

## Troubleshooting

### `invalid_client` during device-code login

Usually the Entra app is not configured as a public client. Enable **Authentication** → **Allow public client flows**.

### Token only contains `User.Read profile openid email`

The broad delegated permissions are not yet granted/admin-consented, or the user is still using an old token. Add permissions, grant admin consent when required, then run `m365_cli logout` and login again.

### Graph mail or calendar returns `403`

The app/user token lacks the required delegated permission, or admin consent has not been granted. Check the token `scp` claim and Entra **API permissions**.

### OneDrive root returns `404`

The user's OneDrive may not be provisioned yet, or the token lacks file permissions. Verify `Files.ReadWrite.All`/`Sites.ReadWrite.All` consent and open OneDrive once in the browser if needed.

### `Insufficient privileges` when listing or changing app permissions

The logged-in user does not have enough Entra privileges for app administration. Use the Entra admin center or an account with the required role.

## Safety policy

Even if `Mail.Send` is granted, OpenClaw agents should draft by default and send only after an explicit user request plus confirmation. External writes, destructive SharePoint/file changes, Teams messages, and calendar mutations require explicit user intent.
