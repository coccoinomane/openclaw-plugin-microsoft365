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

The plugin contributes these exec environment variables when configured:

- `CLIMICROSOFT365_TENANT`
- `CLIMICROSOFT365_ENTRAAPPID`
- `M365_TENANT_ID`
- `M365_CLIENT_ID`
- `M365_ACCOUNT`
- `M365_CONNECTION_NAME`

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

If Graph calls return `403`, inspect the decoded token `scp` claim. After permissions/admin consent change, run `m365_cli logout` and login again to refresh the token.

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
- **Unread/recent Outlook mail for briefs or monitoring**: use Graph server-side filtering via `m365 request --url` with `$filter=isRead eq false` plus the date/folder constraint. Do **not** use `m365 outlook message list` with a `--query` unread filter as the primary unread filter for recurring jobs: `--query` is JMESPath client-side after retrieval, so it can be slow and can miss relevant messages if a bounded/page-limited result is filtered locally.
- **High-level Outlook message list**: `m365 outlook message list` is a standard CLI command and is fine for quick/manual checks, especially with `--folderName`, `--startTime`, and `--endTime`; treat `--query` as local post-filtering, not Graph server-side filtering.
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

### Unread Inbox email with server-side filtering

For unread mail, especially recurring briefs/monitoring, filter server-side with Graph through the CLI. Keep both `isRead eq false` and the date/folder constraint in the URL; only sort locally if Graph rejects `$orderby`.

```bash
m365_cli request \
  --url '@graph/me/mailFolders/inbox/messages?$filter=isRead eq false and receivedDateTime ge 2026-06-18T00:00:00Z&$top=50&$select=id,receivedDateTime,from,subject,bodyPreview,isRead,importance,webLink&$orderby=receivedDateTime desc' \
  --output json
```

If Graph rejects the `$filter` + `$orderby` combination, retry without `$orderby` but keep the server-side unread/date filter, then sort locally:

```bash
m365_cli request \
  --url '@graph/me/mailFolders/inbox/messages?$filter=isRead eq false and receivedDateTime ge 2026-06-18T00:00:00Z&$top=50&$select=id,receivedDateTime,from,subject,bodyPreview,isRead,importance,webLink' \
  --output json > /tmp/m365-unread.json

jq '.value | sort_by(.receivedDateTime) | reverse | map({receivedDateTime, from: .from.emailAddress, subject, isRead, importance, bodyPreview, webLink})' /tmp/m365-unread.json
```

Avoid this as the primary implementation for briefs:

```bash
m365_cli outlook message list --folderName Inbox --query '[?isRead==`false`]' --output json
```

That command is standard CLI, but `--query` is client-side JMESPath filtering after retrieval. Use it only as a manual/small fallback when server-side Graph filtering is unavailable.

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

### Create and edit formatted email drafts

Draft by default; sending requires an explicit request plus confirmation. Read the original message or current draft first and confirm the account, recipients, subject and conversation. Use HTML unless the user asks for plain text.

Use `<p>`, `<ul><li>`, `<strong>`, `<a>` and `<br>` for formatting, with simple inline styles for fonts and spacing when needed. Escape inserted text and attribute values (for example, Python `html.escape`); validate link URLs before inserting them. Do not escape the completed HTML markup. Serialize payloads as UTF-8 JSON files rather than interpolating JSON in shell commands. Markdown, literal newlines, CSS classes and external stylesheets are not substitutes for email HTML.

#### New message

`POST /me/messages` accepts `body` directly, without a `message` wrapper:

```bash
cat > /tmp/m365-draft.json <<'JSON'
{
  "subject": "Subject here",
  "body": {
    "contentType": "HTML",
    "content": "<p>Hello,</p><ul><li><strong>Topic:</strong> detail</li></ul>"
  },
  "toRecipients": [{"emailAddress": {"address": "person@example.com"}}]
}
JSON
m365_cli request \
  --method post \
  --url '@graph/me/messages' \
  --body @/tmp/m365-draft.json \
  --content-type 'application/json' \
  --output json
```

Use a private, per-task directory for payloads and readbacks containing mailbox data; the `/tmp` filenames in these examples are placeholders.

#### Reply draft

For a formatted reply, `POST /me/messages/{originalId}/createReply` uses a `message` wrapper. Do not use a plain-text `comment` for rich-text replies; line breaks can collapse. Do not submit both `comment` and `message.body`.

```bash
cat > /tmp/m365-reply.json <<'JSON'
{
  "message": {
    "body": {
      "contentType": "HTML",
      "content": "<p>Hello,</p><ul><li><strong>Topic:</strong> detail</li></ul>"
    }
  }
}
JSON
m365_cli request \
  --method post \
  --url "@graph/me/messages/$ORIGINAL_ID/createReply" \
  --body @/tmp/m365-reply.json \
  --content-type 'application/json' \
  --output json
```

Preserve the intended recipients, subject and quoted thread. Do not assume that supplying `message.body` preserves the quoted message or that reply creation copies attachments. Inspect the generated draft and attachment collection, including inline images, and retain or restore required material before declaring it complete. Use a reply operation rather than a new unrelated message when conversation continuity matters.

#### Existing draft: PATCH first

Read the latest draft immediately before editing. Confirm `isDraft: true`, preserve user edits, and PATCH only fields the user asked to change. Omitting recipients or subject preserves them; replacing `body` replaces the entire body, so merge the requested edits with the current HTML, including signatures and quoted content. If a fresh read shows intervening changes, reconcile them before writing; do not overwrite from a stale snapshot. Attachment changes require their own Graph operations and explicit intent.

```bash
cat > /tmp/m365-draft-patch.json <<'JSON'
{
  "body": {
    "contentType": "HTML",
    "content": "<p>Hello,</p><ul><li><strong>Updated topic:</strong> detail</li></ul>"
  }
}
JSON
m365_cli request \
  --method patch \
  --url "@graph/me/messages/$DRAFT_ID" \
  --body @/tmp/m365-draft-patch.json \
  --content-type 'application/json' \
  --output json
```

The body above is illustrative; construct the actual payload from the current draft. Add `subject` only when it needs changing. Methods must be lowercase and JSON requests must include `--content-type 'application/json'`.

A controlled test on 2026-09-27 confirmed that PATCH updated a new draft's subject and HTML body while keeping the same ID, `isDraft: true`, and `<ul>`, `<li>` and `<strong>` elements. That test did not cover reply quotations or attachments. There is no blanket restriction against editing existing drafts.

Replacement is a fallback only when a specific, diagnosed failure prevents an in-place update. Explain the reason; do not replace drafts merely to fix formatting. Preserve user edits, recipients, subject, quoted content and required attachments, create the replacement, then verify it before deleting the old draft. Re-read the old draft before deletion to avoid losing intervening edits. If preservation or verification is uncertain, keep the original and report the issue.

#### Read back and return the saved draft

After creation or PATCH, read the saved draft rather than trusting the write response alone:

```bash
m365_cli request \
  --method get \
  --url "@graph/me/messages/$DRAFT_ID?\$select=id,isDraft,subject,body,toRecipients,ccRecipients,bccRecipients,conversationId,hasAttachments,webLink,lastModifiedDateTime" \
  --output json
```

Check `isDraft`, expected ID (unchanged for PATCH), subject, recipients, conversation, HTML content type and the expected paragraph/list/bold tags. Inspect attachments separately when relevant; `hasAttachments` alone does not account for inline-only attachments. Inspect rendering when possible, and distinguish HTML/readback verification from a visual check. Return the saved draft's Outlook `webLink`; after replacement return the new link, never the deleted draft's link. Never send as part of verification.

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
