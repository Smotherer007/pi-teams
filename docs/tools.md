# Tools, commands and prompts

Every tool takes optional `account` and `tenant` parameters; omitting them uses
the default account. See [configuration.md](configuration.md).

## Session and configuration

| Tool | Description |
|------|-------------|
| `teams_setup` | Write an account or tenant to the config file |
| `teams_login` | Sign in as yourself (browser, or device code where there is none) |
| `teams_logout` | Remove a saved session |
| `teams_accounts` | List accounts, switch the default, change a safety level, delete |
| `teams_status` | Which account is active, who pi acts as, what it may do |
| `teams_permissions` | Show, test or change the allow/deny rules |
| `teams_doctor` | Diagnose config, sign-in, consent and connectivity |
| `teams_watch` | Listen mode: status, enable, disable, and the filters for who and where |

## Reading

| Tool | Description |
|------|-------------|
| `teams_inbox` | Chats with new messages plus recent mentions of you |
| `teams_history` | What pi sent across chats (audit log), or what the per-chat workers were asked and answered (dispatch journal). Read-only; see [listen mode](listen-mode.md#what-pi-said-in-other-chats) |
| `teams_list_chats` | Recent chats with a preview of the last message |
| `teams_read_chat` | Messages of one chat |
| `teams_list_teams` | Teams you belong to |
| `teams_list_channels` | Channels of a team |
| `teams_read_channel` | Posts in a channel |
| `teams_read_thread` | One thread: opener plus every reply |
| `teams_search_messages` | Full-text search across chats and channels |
| `teams_list_members` | Members of a team or channel |
| `teams_find_user` | Directory lookup by name, e-mail or UPN |
| `teams_list_files` | Files in a channel's SharePoint folder |
| `teams_download_files` | Download the images and files of messages to a local directory |
| `teams_get_presence` | Your status, a colleague's, or a whole chat's |
| `teams_availability` | When several people are all free, from calendar free/busy |
| `teams_list_meetings` | Calendar events and meetings in a range |
| `teams_get_meeting` | One meeting in full |

## Writing

| Tool | Description |
|------|-------------|
| `teams_send_chat_message` | Send a chat message as you, optionally with images |
| `teams_create_chat` | Start a 1:1 or group chat, optionally with the first message |
| `teams_send_channel_message` | Post in a channel as you, optionally with images |
| `teams_reply_channel_message` | Reply inside an existing thread, optionally with images |
| `teams_react` | Add or remove an emoji reaction |
| `teams_mark_read` | Mark a chat as read (or unread again) — the natural end of a catch-up |
| `teams_update_message` | Edit a chat message you already sent |
| `teams_chat_members` | Add someone to a group chat, or remove them |
| `teams_delete_message` | Soft-delete (or restore) a message you sent |
| `teams_set_presence` | Set your Teams status |
| `teams_set_status_message` | Set the note under your name |
| `teams_create_meeting` | Schedule a Teams meeting with invitations |
| `teams_update_meeting` | Change a meeting you organize |
| `teams_respond_invite` | Accept, decline or tentatively accept an invitation |
| `teams_cancel_meeting` | Cancel a meeting |

### What channels cannot do

Editing and deleting apply to **chat** messages. A channel post cannot be
created, edited or deleted through Graph with the permissions pi asks for —
those need `Channel.Create` and `ChannelMessage.ReadWrite`, both admin-consent
scopes, and neither is worth making every setup wait for an administrator.
Posting and replying in channels, and reacting to channel messages, work as
usual; correcting a channel post stays in the Teams client.

## Commands

| Command | Description |
|---------|-------------|
| `/teams-status` | Show the connection and permissions |
| `/teams-login` | Sign in to Teams as yourself |
| `/teams-inbox` | What needs your attention right now |
| `/teams-listen` | Listen mode: `on`, `off`, or `status` |
| `/teams-dispatch` | Listen mode with one pi process per chat: who is being worked on, who is idle, who waits |
| `/teams-permissions` | What pi may and may not do |
| `/teams-doctor` | Diagnose the setup |

## Prompts

| Prompt | Description |
|--------|-------------|
| `/teams-catch-up` | Summarize what happened while you were away |
| `/teams-triage` | Work through unanswered messages one by one |
| `/teams-channel-digest` | Decisions, open questions and action items in a channel |
| `/teams-standup` | Draft and post a standup update |
| `/teams-meeting-prep` | Context and talking points for your next meeting |
| `/teams-schedule` | Find a time everyone is free and book it |
| `/teams-doctor` | Guided troubleshooting |

---

## Addressing things

**Channels** — a `Team/Channel` path, or `team` and `channel` separately. Names
resolve case-insensitively; an ambiguous name is reported, never guessed.

```yaml
teams_read_channel:
  channel: "Engineering/General"
  limit: 20
```

**Chats** — an ID, a group chat topic, or a participant's name or e-mail. If
more than one chat matches, the tool lists the candidates instead of picking
one.

```yaml
teams_send_chat_message:
  chat: "anna@contoso.com"
  body: "Running five minutes late."
```

**Mentions** — pass `mentions` and write `@Display Name` in the body where the
mention belongs:

```yaml
teams_send_channel_message:
  channel: "Engineering/General"
  body: "@Anna Schmidt the pipeline is green again."
  mentions: ["anna@contoso.com"]
```

---

## How messages are formatted

A message body is treated as **lightweight markdown** and converted to the HTML
subset Teams renders in a chat bubble:

| You write | Teams shows |
|-----------|-------------|
| `**bold**`, `*italic*`, `~~struck~~` | bold, italic, struck through |
| `` `code` ``, fenced blocks | inline code, code block |
| `- item`, `1. item` | bullet list, numbered list |
| `[label](https://…)` | link |
| `# Heading` | a bold line — Teams has no headings in a chat |

A markdown image and a table are deliberately **not** translated: Teams renders
neither in a chat bubble, and an image URL would only work if the recipient's
client could reach it. A picture that is already on disk goes out through the
`images` parameter instead — see below. For a table, put it in a file and link
to it.

Pass `html: true` on a send tool to bypass the conversion and supply raw HTML
instead. Everything else is escaped, so a message body cannot inject markup.

This is also why pi writes the way it does — short, answer-first, bullets for
anything enumerable. The medium is a chat bubble, not a document. The rules the
model follows are in `skills/teams-collaboration/SKILL.md`, under *How to format
a Teams message*.

## Sending images

`teams_send_chat_message`, `teams_send_channel_message` and
`teams_reply_channel_message` take `images: ["/pfad/bild.png"]` — local paths of
pictures to attach. They appear under the text, in the order given, and a
message may consist of images alone.

Teams keeps an inline picture inside the message itself: the bytes travel in the
same request as the text, and the body points at them (`hostedContents`).
Nothing is uploaded to SharePoint first, which is why this works with the
permissions the package already asks for and costs no admin consent.

- Accepted: png, jpeg, gif, webp, bmp. Anything else is refused before the
  message is built, because Teams would show it as a broken picture.
- **4 MB per image**, the limit Graph enforces on hosted content.
- The path must be readable from the process that sends. For a pi running in a
  container that is a path *inside* the container.

This is attachments-for-images only. A PDF, a DOCX or a ZIP cannot be sent: a
real file attachment needs the file in SharePoint or OneDrive and a `contentUrl`
pointing at it, which means asking for a write scope on the user's drive. The
package deliberately does not ask for `Files.ReadWrite`, so a non-image is
refused rather than half-supported — link to the file instead.

What was sent, pictures included, is recorded in the audit entry and shown in
the confirmation prompt at `safetyLevel: confirm`.
