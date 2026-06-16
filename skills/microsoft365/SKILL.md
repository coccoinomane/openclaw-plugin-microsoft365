---
name: microsoft365
description: Use Microsoft 365 via the official CLI for Microsoft 365 (`m365`), especially for Outlook email, calendars, and OneDrive/SharePoint files. Use this whenever the user asks to check unread/recent mail, read or search email, draft or send email, inspect agenda/calendar events, create calendar events, list/search/download/upload files, or run raw Microsoft Graph requests.
---

# Microsoft 365 with CLI for Microsoft 365

Use the official [CLI for Microsoft 365](https://github.com/pnp/cli-microsoft365), exposed as `m365`, as the primary backend.

Do **not** use IMAP/SMTP password flows for Microsoft 365. Prefer Microsoft Graph through `m365` OAuth/device-code auth.

## Backend command resolution

Prefer an installed `m365` binary when available; otherwise use the npm package without requiring a global install. Use a shell function rather than storing a command with spaces in a variable:

```bash
m365_cli() {
  if command -v m365 >/dev/null 2>&1; then
    m365 "$@"
  else
    npx -y -p @pnp/cli-microsoft365 m365 "$@"
  fi
}
```

For one-off `exec` calls, it is fine to inline:

```bash
npx -y -p @pnp/cli-microsoft365 m365 status --output json
```

Always use `--output json` for agent-readable commands.

## Authentication model

Required operator configuration:

- `M365_TENANT_ID` or plugin config `tenantId`
- `M365_CLIENT_ID` or plugin config `clientId`
- optional `M365_ACCOUNT` / plugin config `account`
- optional `M365_CONNECTION_NAME` / plugin config `connectionName`
- optional `M365_SCOPES` / plugin config `scopes`

The plugin contributes these exec environment variables when configured:

- `CLIMICROSOFT365_TENANT`
- `CLIMICROSOFT365_ENTRAAPPID`
- `M365_TENANT_ID`
- `M365_CLIENT_ID`
- `M365_ACCOUNT`
- `M365_CONNECTION_NAME`
- `M365_SCOPES`

Login:

```bash
m365_cli login \
  --authType deviceCode \
  --appId "$M365_CLIENT_ID" \
  --tenant "$M365_TENANT_ID" \
  --connectionName "${M365_CONNECTION_NAME:-openclaw-microsoft365}"
```

Status check:

```bash
m365_cli status --output json
m365_cli util accesstoken get --resource https://graph.microsoft.com --decoded --output json
```

If Graph calls return `403`, inspect the decoded token `scp` claim. After permissions/admin consent change, run `m365_cli logout` and login again to refresh scopes.

## Recommended delegated Graph scopes

For broad email/calendar/files coverage, request delegated permissions such as:

```text
openid profile email offline_access
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

Use the minimum set that fits the operator's risk tolerance. Many broad scopes require tenant admin consent.

PowerPoint, Word, and Excel are usually handled as Office files in OneDrive/SharePoint via `Files.*` and `Sites.*` permissions; there is no general Graph scope equivalent to Google Slides editing.

## What the CLI docs imply for common tasks

Prefer these defaults:

- **Email list/search/read/mark read/draft**: use `m365 request --url @graph/...` raw Graph. It gives better `$select`, `$filter`, `$search`, `$top`, and PATCH support than high-level commands.
- **Direct email send**: high-level `m365 outlook mail send` exists, but it sends immediately. Use only after explicit user send request + confirmation.
- **Draft email**: no high-level draft command in CLI v11.8.0. Use Graph `POST /me/messages`.
- **Calendar agenda**: use raw Graph `calendarView`; it is faster than high-level `outlook event list` and does not require discovering calendar IDs.
- **Create/update calendar events**: use raw Graph `POST/PATCH /me/events` after explicit user intent.
- **Current user's OneDrive**: use raw Graph `/me/drive/...`.
- **SharePoint document libraries**: use `m365 file list/add/copy/move` or `m365 spo file/folder *` when you already know the site URL and folder path.

`m365 request` details that matter:

- HTTP methods must be lowercase: `get`, `post`, `patch`, `delete`. `PATCH` is rejected.
- When using `--body`, include `--content-type "application/json"` for JSON payloads.
- Use body files (`--body @file.json`) for non-trivial JSON to avoid shell quoting errors.
- Quote URLs so `$select`, `$filter`, `$orderby`, `$search` are not eaten by the shell.

## Quick recipes: email

### Recent inbox messages

```bash
m365_cli request \
  --url '@graph/me/mailFolders/inbox/messages?$top=20&$select=id,receivedDateTime,from,subject,bodyPreview,isRead,importance,webLink&$orderby=receivedDateTime desc' \
  --output json
```

### Unread email

Graph can be picky when combining `$filter` and `$orderby`. Fetch unread messages, then sort locally:

```bash
m365_cli request \
  --url '@graph/me/messages?$filter=isRead eq false&$top=50&$select=id,receivedDateTime,from,subject,bodyPreview,isRead,importance,webLink,parentFolderId' \
  --output json > /tmp/m365-unread.json

jq '.value | sort_by(.receivedDateTime) | reverse | map({receivedDateTime, from: .from.emailAddress, subject, isRead, importance, bodyPreview, webLink})' /tmp/m365-unread.json
```

### Folder unread counts

```bash
m365_cli request \
  --url '@graph/me/mailFolders?$top=100&$select=id,displayName,unreadItemCount,totalItemCount' \
  --output json
```

### Search email

Use Graph `$search` for broad mailbox search. Keep the query quoted inside the URL:

```bash
m365_cli request \
  --url '@graph/me/messages?$search="anritsu"&$top=20&$select=id,receivedDateTime,from,subject,bodyPreview,isRead,webLink' \
  --output json
```

If `$search` fails because of tenant/search limitations, fall back to date/folder filters plus local `jq`/Python matching.

### Read a specific message

```bash
m365_cli request \
  --url "@graph/me/messages/$MESSAGE_ID?\$select=id,receivedDateTime,from,toRecipients,ccRecipients,subject,body,bodyPreview,hasAttachments,isRead,webLink" \
  --output json
```

### Mark message read/unread

```bash
cat > /tmp/m365-message-read.json <<'JSON'
{"isRead": true}
JSON
m365_cli request \
  --method patch \
  --url "@graph/me/messages/$MESSAGE_ID" \
  --body @/tmp/m365-message-read.json \
  --content-type 'application/json' \
  --output json
```

Use `{"isRead": false}` to mark unread.

### Create an email draft by default

Default to drafts, not direct sends:

```bash
cat > /tmp/m365-draft.json <<'JSON'
{
  "subject": "Subject here",
  "body": { "contentType": "Text", "content": "Message body here" },
  "toRecipients": [
    { "emailAddress": { "address": "person@example.com" } }
  ],
  "ccRecipients": []
}
JSON

m365_cli request \
  --method post \
  --url '@graph/me/messages' \
  --body @/tmp/m365-draft.json \
  --content-type 'application/json' \
  --output json
```

The response contains the draft `id` and `webLink`. Share the Outlook web link with the user when useful.

### Send a draft only after explicit confirmation

```bash
m365_cli request \
  --method post \
  --url "@graph/me/messages/$DRAFT_ID/send" \
  --body '{}' \
  --content-type 'application/json' \
  --output json
```

### Direct send command, only when explicitly requested

The CLI docs expose direct send as `outlook mail send`:

```bash
m365_cli outlook mail send \
  --to 'person@example.com' \
  --subject 'Subject here' \
  --bodyContents @/tmp/email-body.txt \
  --bodyContentType Text \
  --output json
```

Use this only when the user explicitly asks to send and confirms. Attachments are supported with repeated `--attachment` flags, but the CLI docs note a 3 MB total attachment limit for this command.

## Quick recipes: calendar

### Agenda / calendar view

Use raw Graph `calendarView` for the signed-in user's calendar:

```bash
m365_cli request \
  --url '@graph/me/calendarView?startDateTime=2026-06-16T00:00:00%2B02:00&endDateTime=2026-06-23T00:00:00%2B02:00&$top=50&$select=id,subject,start,end,location,organizer,attendees,isOnlineMeeting,onlineMeeting,webLink&$orderby=start/dateTime' \
  --output json
```

Generate `startDateTime`/`endDateTime` with the user's timezone. For Italy, use `Europe/Rome` and URL-encode `+` as `%2B`.

### Search calendar events

```bash
m365_cli request \
  --url '@graph/me/events?$filter=contains(subject, '\''Anritsu'\'')&$top=20&$select=id,subject,start,end,location,webLink&$orderby=start/dateTime' \
  --output json
```

If Graph rejects the filter/order combination, fetch a bounded calendarView and filter locally.

### Create a calendar event, only after explicit user intent

```bash
cat > /tmp/m365-event.json <<'JSON'
{
  "subject": "Meeting title",
  "body": { "contentType": "Text", "content": "Agenda or notes" },
  "start": { "dateTime": "2026-06-16T15:00:00", "timeZone": "Europe/Rome" },
  "end": { "dateTime": "2026-06-16T15:30:00", "timeZone": "Europe/Rome" },
  "attendees": [
    {
      "emailAddress": { "address": "person@example.com", "name": "Person" },
      "type": "required"
    }
  ]
}
JSON

m365_cli request \
  --method post \
  --url '@graph/me/events' \
  --body @/tmp/m365-event.json \
  --content-type 'application/json' \
  --output json
```

Calendar writes and invites are external actions. Confirm intent before creating/updating/deleting events.

## Quick recipes: files

### List OneDrive root or a folder

Current user's OneDrive root:

```bash
m365_cli request \
  --url '@graph/me/drive/root/children?$top=50&$select=id,name,webUrl,size,lastModifiedDateTime,file,folder' \
  --output json
```

Folder by path:

```bash
m365_cli request \
  --url "@graph/me/drive/root:/path/to/folder:/children?\$top=50&\$select=id,name,webUrl,size,lastModifiedDateTime,file,folder" \
  --output json
```

### Search OneDrive files

```bash
m365_cli request \
  --url "@graph/me/drive/root/search(q='proposal')?\$top=25&\$select=id,name,webUrl,size,lastModifiedDateTime,file,folder" \
  --output json
```

### Download a OneDrive file by item id

```bash
m365_cli request \
  --url "@graph/me/drive/items/$ITEM_ID/content" \
  --filePath ./downloaded-file.ext \
  --output none
```

### Upload a small file to OneDrive by path

For simple uploads, Graph supports `PUT /content`. Confirm before overwriting an existing path.

```bash
m365_cli request \
  --method put \
  --url "@graph/me/drive/root:/target/path/file.ext:/content" \
  --body @./local-file.ext \
  --content-type 'application/octet-stream' \
  --output json
```

For larger files, use a Graph upload session rather than a single PUT.

### SharePoint document libraries

When you know the site URL and library/folder path, high-level file commands are concise:

```bash
m365_cli file list \
  --webUrl 'https://tenant.sharepoint.com/sites/project-x' \
  --folderUrl 'Shared Documents' \
  --output json

m365_cli file add \
  --filePath ./file.pdf \
  --folderUrl 'https://tenant.sharepoint.com/sites/project-x/Shared Documents' \
  --output json
```

For SharePoint-specific metadata and folders, use `m365 spo file list`, `m365 spo file get`, and `m365 spo folder list`.

## Output handling

- Always use `--output json` for agent-readable commands.
- Pipe through `jq`/Python locally to reduce context size before reading results.
- Store large raw outputs under `out/` and read only summaries/shortlists.
- Never store OAuth tokens or secrets in repository-tracked files.
- Do not include private message bodies or file contents in public/group replies unless the user explicitly asks.

## Operational safety

- External writes (sending mail, modifying calendar events, changing SharePoint/OneDrive files, Teams messages) require explicit user intent.
- For email, draft by default. Sending requires an explicit send request plus confirmation.
- For destructive changes, inspect the target first and prefer reversible operations.
- In group chats, summarize private Microsoft 365 data minimally and only when it is clearly appropriate for the current channel.
