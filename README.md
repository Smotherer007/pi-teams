# pi-teams

Microsoft Teams integration for the [pi coding agent](https://github.com/earendil-works/pi).

pi acts **as you**: it reads what you can read, and anything it sends comes
from your account, in your name, in the right thread. Multiple Teams accounts
and multiple company tenants live side by side, and a configuration file
decides exactly which teams, channels, chats and people pi may touch.

![pi-teams in the pi coding agent](screenshot.png)

## Installation

```bash
# From npm
pi install npm:@patimweb/pi-teams

# From a local checkout during development
pi install /path/to/pi-teams
```

## Quick start

1. **Register an app in Entra ID** (once per organization) — the step-by-step
   walkthrough, including the redirect URI and public-client settings that sign-in
   needs, is in [docs/entra-app-registration.md](docs/entra-app-registration.md).
2. **Configure the account:**
   ```
   teams_setup:
     name: work
     tenantId: contoso.onmicrosoft.com
     clientId: <application (client) ID>
     setDefault: true
   ```
3. **Sign in:** run `/teams-login`. Your browser opens at the Microsoft
   sign-in page and closes itself when you are done — nothing to type.
4. **Check it works:** `/teams-status`, then `/teams-inbox`.

Full walkthrough with the exact portal paths: **[docs/entra-app-registration.md](docs/entra-app-registration.md)**.

---

## Listen mode

Off by default. When it is on, pi polls your chats and turns an incoming message
into a prompt it answers — as you, in the chat it arrived in.

```
/teams-listen on        # switch it on for the session's account
/teams-listen status    # what is configured, and what is running right now
/teams-listen off
```

```json
"watch": {
  "enabled": true,
  "intervalSeconds": 60,
  "chats": ["Anna*", "Vertrieb*"],
  "from": ["anna.schmidt@contoso.com"],
  "mentionOnly": false,
  "cooldownSeconds": 300,
  "maxTriggersPerHour": 10
}
```

| Key | Default | Meaning |
|-----|---------|---------|
| `enabled` | `false` | whether the watcher runs |
| `intervalSeconds` | `60` | seconds between polls (minimum 15) |
| `chats` | `[]` | **where** pi listens: glob patterns matched against topic, label, chat ID and participant names. Empty means every recent chat. |
| `from` | `[]` | **who** pi listens to: glob patterns matched against display name, UPN and e-mail. Empty means any sender. |
| `mentionOnly` | `false` | only wake where you are mentioned |
| `cooldownSeconds` | `300` | stay quiet in a chat after waking pi for it |
| `maxTriggersPerHour` | `10` | hard cap on wakes per hour |

Set it globally, or per account (the account level wins field by field).
`teams_watch` writes the same keys; `/teams-listen` writes to the session's
account.

What it does, and what it deliberately does not:

- **Costs a model turn per wake.** That is what the filters and the hourly cap
  are for — narrow `chats` and `from`, and enable it for a handful of
  conversations rather than the whole company.
- **Answers what is still unread.** A chat wakes pi when two things are true: it
  moved since pi last looked (a cursor per account, kept on disk), and its newest
  message is still unread for you in Teams. So switching listen mode back on
  answers what you missed while pi was not running — and a chat you already read
  in Teams stays quiet. A backlog is drained a few chats per tick and capped by
  `maxTriggersPerHour`, so a long absence does not produce a burst.
- **Never wakes for its own messages.** pi posts as you, so its own reply comes
  back as "my own message" and stops the loop.
- **Respects the read rules.** A chat excluded by `permissions.read.chats` is
  not even polled.
- **Does not bypass the safety model.** A reply it decides to send goes through
  the same safety level and scope rules as one you asked for. At `confirm` you
  approve each one; for hands-free answers, narrow `permissions.write.chats` and
  set `safetyLevel: open` — deliberately, and on an account where you mean it.
- **Chats only, not channels.** Channel polling needs `ChannelMessage.Read.All`
  (admin consent) and multiplies the calls; answering in channels stays a
  requested action.

Why polling and not a webhook: Graph change notifications for chats need a
publicly reachable HTTPS endpoint plus subscription renewal, which a laptop on a
home network cannot offer. Polling costs one `/me/chats` call per interval and
works anywhere.

---

## Entra ID app registration

Full walkthrough: **[docs/entra-app-registration.md](docs/entra-app-registration.md)**.
The short version, with the two settings that sign-in actually depends on:

pi signs in as you with a delegated OAuth 2.0 flow — the browser opens, you sign
in with your own credentials and MFA, and pi never sees your password.

### 1. Create the registration

Entra admin center → **App registrations** → **New registration**

| Field | Value |
|-------|-------|
| Name | `pi-teams` (anything you like) |
| Supported account types | **Accounts in any organizational directory (multitenant)** — required if pi should also reach customer or guest tenants. Single tenant is fine for one company. |
| Redirect URI | leave empty |

Note the **Application (client) ID** and **Directory (tenant) ID**.

### 2. Enable public client flows

**Authentication** → *Advanced settings* → **Allow public client flows** →
**Yes**.

Without this, sign-in fails with `AADSTS7000218`.

### 3. Add delegated permissions

**API permissions** → **Add a permission** → **Microsoft Graph** →
**Delegated permissions**:

| Permission | Needed for |
|-----------|-----------|
| `User.Read`, `User.ReadBasic.All` | identity, looking colleagues up |
| `Team.ReadBasic.All`, `Channel.ReadBasic.All` | listing teams and channels |
| `ChannelMessage.Read.All` | reading channel messages |
| `ChannelMessage.Send` | posting and replying in channels |
| `ChannelMember.Read.All` | channel member lists |
| `Chat.ReadWrite` | reading chats, creating chats |
| `ChatMessage.Send` | sending chat messages |
| `Presence.ReadWrite`, `Presence.Read.All` | your status, colleagues' availability |
| `Calendars.ReadWrite` | reading and creating calendar events |
| `OnlineMeetings.ReadWrite` | Teams meeting links |
| `Files.Read.All`, `Sites.Read.All` | channel files |

`openid`, `profile` and `offline_access` are not listed: MSAL always requests
them and rejects them in an explicit scope list, so the lasting sign-in they
buy is automatic.

### 4. Add the redirect URI for the browser sign-in

**Authentication** → **Add a platform** → **Mobile and desktop applications** →
tick `http://localhost`.

MSAL picks a free loopback port at sign-in time, and Entra accepts any port on
localhost for a public client, so this one entry covers every run. If your
tenant insists on an exact URI, register e.g. `http://localhost:3000` and set
`loopbackPort: 3000` on the account.

---

## Who has to approve what

Two separate hurdles, and only one of them normally needs an administrator.

### Creating the registration

By default any member of the tenant may register an application (Entra ID →
**User settings** → *Users can register applications*). Where that is switched
off, the least-privileged role that can do it is **Application Developer**;
Cloud Application Administrator and Application Administrator also work.
Changing *Allow public client flows* and adding the redirect URI need no extra
role once the registration is yours.

### Consenting to the permissions

Only two of the default scopes require **admin consent**:

| Scope | Consent | What you lose without it |
|-------|---------|--------------------------|
| `ChannelMessage.Read.All` | **admin** | reading channel messages and threads |
| `ChannelMember.Read.All` | **admin** | channel member lists |
| `ChannelMessage.Send` | user | — |
| `Chat.ReadWrite`, `ChatMessage.Send` | user | — |
| `Team.ReadBasic.All`, `Channel.ReadBasic.All` | user | — |
| `User.Read`, `User.ReadBasic.All` | user | — |
| `Presence.ReadWrite`, `Presence.Read.All` | user | — |
| `Calendars.ReadWrite`, `OnlineMeetings.ReadWrite` | user | — |
| `Files.Read.All`, `Sites.Read.All` | user | — |

Posting in a channel needs no admin; *reading* the channel does. Microsoft
treats bulk reading of channel content as the more sensitive right.

**Without an administrator**, drop the two admin-consent scopes and everything
else still works — chats in full, posting and replying in channels, presence,
calendar, meetings and directory lookup:

```json
"scopes": [
  "User.Read", "User.ReadBasic.All",
  "Team.ReadBasic.All", "Channel.ReadBasic.All", "ChannelMessage.Send",
  "Chat.ReadWrite", "ChatMessage.Send",
  "Presence.ReadWrite", "Presence.Read.All",
  "Calendars.ReadWrite", "OnlineMeetings.ReadWrite",
  "Files.Read.All", "Sites.Read.All"
]
```

One caveat: if the tenant has *Users can consent to applications* switched off,
even the user-consentable scopes need an administrator — usually through the
admin consent request workflow.

`teams_doctor` reports exactly which scopes were actually granted, so a missing
consent shows up as a named permission rather than a 403.

### Guest and customer tenants

For a tenant that is not your own:

- the registration must be **multitenant**,
- your account must be a guest in that tenant,
- an administrator **of that tenant** consents to the app once.

In practice this is the bigger hurdle, not your own tenant. Once it is done,
add the tenant beneath your account (`parentAccount` in `teams_setup`) and sign
in against it.

---

## How the sign-in works

`teams_login` uses MSAL (`@azure/msal-node`) and picks the flow that fits the
machine:

| | `interactive` (default) | `device-code` | `client-credentials` |
|---|---|---|---|
| What you do | browser opens, you sign in | type a short code elsewhere | nothing |
| Identity | you | you | the application |
| Needs a browser | yes | no | no |
| Reading chats/channels | yes | yes | only via protected APIs |
| **Sending messages** | **yes** | **yes** | **no** — refused with an explanation |
| Presence, calendar | yes | yes | partially |

**Interactive** is chosen automatically wherever a browser can actually be
reached. MSAL opens the system browser, listens on a loopback port for the
redirect, and handles PKCE — nothing to copy or type.

**Device code** takes over automatically over SSH, in containers, and on any
machine without a display server, because a browser opened there would appear
where nobody is looking. Force it anywhere with `PI_TEAMS_NO_BROWSER=1`, per
account with `"authMode": "device-code"`, or per call with
`teams_login mode: device-code`.

Either way the session is cached on disk and renewed silently, so signing in is
a once-per-account affair.

Microsoft Graph only permits app-only writes to chats and channels for
migration scenarios (`Teamwork.Migrate.All`), so pi-teams refuses those calls
up front instead of letting them fail deep inside a Graph request. Use
`client-credentials` for unattended read-only jobs, and a delegated flow for
everything that speaks in your name.

---

## Configuration

Everything lives in `~/.pi/agent/pi-teams.json` (mode `0600`, created from a
template on first run). `teams_setup` writes the same file — use whichever you
prefer.

```json
{
  "accounts": [
    {
      "name": "work",
      "displayName": "Contoso",
      "tenantId": "contoso.onmicrosoft.com",
      "clientId": "00000000-0000-0000-0000-000000000000",
      "authMode": "interactive",
      "safetyLevel": "confirm",
      "permissions": {
        "read": {
          "teams": { "allow": ["*"] },
          "chats": { "allow": ["*"] }
        },
        "write": {
          "channels": {
            "allow": ["Engineering/*", "Project Alpha/General"],
            "deny": ["*/Announcements"]
          },
          "chats": { "allow": ["*"] },
          "people": { "deny": ["ceo@contoso.com"] }
        }
      },
      "tenants": [
        {
          "name": "customer-alpha",
          "tenantId": "customeralpha.onmicrosoft.com",
          "safetyLevel": "readonly",
          "permissions": {
            "read": { "channels": { "allow": ["Project Alpha/*"] } }
          }
        }
      ]
    },
    {
      "name": "private",
      "tenantId": "othercompany.onmicrosoft.com",
      "clientId": "11111111-1111-1111-1111-111111111111",
      "safetyLevel": "open"
    }
  ],
  "defaultAccount": "work",
  "safetyLevel": "confirm",
  "maxMessages": 25,
  "audit": true
}
```

### Accounts and tenants

- An **account** is a Teams identity you sign in as. Each has its own saved
  session.
- A **tenant** beneath an account is another directory that identity reaches:
  a customer, a subsidiary, a guest tenant. It gets its own sign-in, its own
  safety level and its own rules.

Every tool takes `account` and `tenant`:

```yaml
teams_send_channel_message:
  account: work
  tenant: customer-alpha
  channel: "Project Alpha/General"
  body: "Sprint review moved to Thursday."
```

### Safety levels

| Level | Effect |
|-------|--------|
| `open` | writes go through without asking |
| `confirm` | every write shows a confirmation with the destination and the exact text (**default**) |
| `readonly` | writes are blocked entirely |

Set globally, per account, or per tenant — the most specific setting wins. So
your own company can be `confirm` while a customer tenant stays `readonly`.

### AI disclosure footer

On by default. Every message, reply and edit pi sends in your name ends with a
short note saying it was written by an AI:

```json
"aiFooter": { "enabled": true, "text": "🤖 Erstellt mit pi (KI-Assistent)" }
```

| Key | Default | Meaning |
|-----|---------|---------|
| `enabled` | `true` | append the note to everything pi sends |
| `text` | `🤖 Generated with pi (an AI agent)` | the wording; empty means the default |

Turn it off with `"enabled": false` if the disclosure does not fit the people
you write to. Set it globally or per account (the account level wins field by
field, so one account can use its own wording while another keeps the global
one).

Details worth knowing:

- It is added by the tool that sends, not asked for in a prompt, so it cannot be
  forgotten by the model. The confirmation dialog at `safetyLevel: confirm`
  shows the text **with** the footer — you confirm what actually goes out.
- It is idempotent: editing an already-footered message does not stack a second
  one, and a model that copied the footer from the chat it is answering does not
  produce two.
- It applies to chat messages, channel posts and replies, the first message of a
  newly created chat, and edits. It does **not** apply to things pi sends on
  your behalf that are not messages — meeting invitations, presence changes.
- It says nothing about the *content*: the note is a disclosure, not a
  disclaimer, and it is the only part of an outgoing message pi adds on its own.

### Scope rules

Rules exist per **category** (`teams`, `channels`, `chats`, `people`) and per
**mode** (`read`, `write`). Category keys written directly in the block apply
to both modes; a `read` or `write` block refines one of them.

```json
"permissions": {
  "channels": { "deny": ["*/Announcements"] },
  "write": { "channels": { "allow": ["Engineering/*"] } }
}
```

- **`deny` accumulates** across global → account → tenant and always wins.
- **`allow` narrows**: the most specific level that defines one wins. Nothing
  configured means everything the signed-in user can reach.
- Patterns are case-insensitive globs: `*` matches any run of characters
  (including `/`), `?` matches one.
- A target matches if **any** of its identifiers does: channels match on
  `Team/Channel`, their name or their ID; chats match on topic, participant
  names and e-mail addresses; people match on name, UPN and e-mail.

Check a rule without touching anything:

```yaml
teams_permissions:
  action: test
  category: channels
  mode: write
  target: "Engineering/Announcements"
```

### Other settings

| Key | Default | Meaning |
|-----|---------|---------|
| `defaultAccount` | first account | account used when a tool omits `account` |
| `maxMessages` | `25` | default page size for message listings |
| `audit` | `true` | append every write to `~/.pi/agent/pi-teams-audit.jsonl` |
| `scopes` | see above | override the requested Graph scopes per account/tenant |
| `aiFooter` | on | append an AI disclosure to every message pi sends — see above |
| `graphBaseUrl` / `authorityHost` | Microsoft public cloud | sovereign cloud endpoints |

---

## Tools

### Session and configuration

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

### Reading

| Tool | Description |
|------|-------------|
| `teams_inbox` | Chats with new messages plus recent mentions of you |
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
| `teams_get_presence` | Your status, a colleague's, or a whole chat's |
| `teams_availability` | When several people are all free, from calendar free/busy |
| `teams_list_meetings` | Calendar events and meetings in a range |
| `teams_get_meeting` | One meeting in full |

### Writing

| Tool | Description |
|------|-------------|
| `teams_send_chat_message` | Send a chat message as you |
| `teams_create_chat` | Start a 1:1 or group chat, optionally with the first message |
| `teams_send_channel_message` | Post in a channel as you |
| `teams_reply_channel_message` | Reply inside an existing thread |
| `teams_react` | Add or remove an emoji reaction |
| `teams_mark_read` | Mark a chat as read (or unread again) — the natural end of a catch-up |
| `teams_update_message` | Edit a message you already sent |
| `teams_chat_members` | Add someone to a group chat, or remove them |
| `teams_delete_message` | Soft-delete (or restore) a message you sent |
| `teams_set_presence` | Set your Teams status |
| `teams_set_status_message` | Set the note under your name |
| `teams_create_meeting` | Schedule a Teams meeting with invitations |
| `teams_update_meeting` | Change a meeting you organize |
| `teams_respond_invite` | Accept, decline or tentatively accept an invitation |
| `teams_cancel_meeting` | Cancel a meeting |

Editing and deleting apply to **chat** messages. A channel post cannot be
created, edited or deleted through Graph with the permissions pi asks for
(`Channel.Create`, `ChannelMessage.ReadWrite`) — that stays in the Teams client.
Posting and replying in channels, and reacting to channel messages, work as
usual.

## Commands

| Command | Description |
|---------|-------------|
| `/teams-status` | Show the connection and permissions |
| `/teams-login` | Sign in to Teams as yourself |
| `/teams-inbox` | What needs your attention right now |
| `/teams-listen` | Listen mode: `on`, `off`, or `status` |
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

Tables and images are deliberately **not** translated: Teams renders neither in
a chat, so passing them through would look worse than plain text. Put a table in
a file and link it instead.

Pass `html: true` on a send tool to bypass the conversion and supply raw HTML
instead. Everything else is escaped, so a message body cannot inject markup.

This is also why pi writes the way it does: see *How to format a Teams message*
in the bundled skill. Short, answer-first, bullets for anything enumerable — the
medium is a chat bubble, not a document.

---

## Addressing things

**Channels** — a `Team/Channel` path, or `team` and `channel` separately.
Names resolve case-insensitively; an ambiguous name is reported, never guessed.

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

## Files pi writes

| File | Contents | Mode |
|------|----------|------|
| `~/.pi/agent/pi-teams.json` | accounts, tenants, safety levels, scope rules | `0600` |
| `~/.pi/agent/pi-teams-tokens/` | one MSAL token cache per account+tenant | `0600` |
| `~/.pi/agent/pi-teams-audit.jsonl` | one line per write: when, who, where, what | `0600` |
| `~/.pi/agent/pi-teams-watch/` | listen mode's cursor: the newest message it has looked at per chat | `0600` |

Both the config and the token cache are written through a private temp file and
an atomic rename, so an interrupted write cannot truncate them. The watch
cursor is what makes listen mode survive a restart, and deleting it only costs
one re-examination of the chats in the list — it is not silently re-answering
anything, because the decision also asks Teams whether the message is still
unread. To revoke pi's access entirely, run `teams_logout` with `all: true` and
remove the app's consent in Entra ID.

## Architecture

Data-oriented: plain immutable data, I/O at the edges, pure functions in the
middle.

- **`src/types.ts`** — domain data as plain interfaces. No behavior.
- **`src/config/index.ts`** — accounts, tenants, the safety cascade, listen-mode settings, persistence.
- **`src/config/scope.ts`** — allow/deny matching. Pure, and the most heavily
  tested module in the package.
- **`src/auth/`** — MSAL applications, the file-backed token cache, JWT claims.
- **`src/graph/client.ts`** — fetch wrapper: bearer token, paging, throttling.
- **`src/graph/*.ts`** — one module per resource area, returning domain types.
- **`src/graph/mappers.ts`** — Graph JSON → domain types. Pure.
- **`src/watch/`** — listen mode: the decision rules (`index.ts`, pure), the
  polling loop (`loop.ts`), and the prompt an incoming message turns into
  (`prompt.ts`).
- **`src/utils/formatting.ts`** — domain types → display strings. Pure.
- **`src/utils/richtext.ts`** — markdown → the HTML subset Teams renders. Pure.
- **`src/utils/slots.ts`** — availability view → the free slots everybody
  shares. Pure, and the part of scheduling that can be wrong without looking
  wrong.
- **`src/safety/`** — the three gates and the audit log.
- **`src/tools/`** — one module per tool.
- **`src/extension/index.ts`** — registration, commands, the `tool_call`
  interceptor, and the listen-mode lifecycle.

## Development

```bash
npm install
npm run typecheck
npm test
```

Tests cover the pure logic — scope matching, the config cascade, message body
construction, markdown rendering, HTML flattening, free-slot arithmetic, the
listen-mode decision rules, the safety gates — and run without a tenant or a
network connection.

## Troubleshooting

| Symptom | Cause and fix |
|---------|---------------|
| `AADSTS500113` (no reply address) | The registration has no redirect URI. Add `http://localhost` under *Mobile and desktop applications* — see [docs/entra-app-registration.md](docs/entra-app-registration.md) |
| `AADSTS7000218` | "Allow public client flows" is off in the app registration |
| `AADSTS65001` | Consent missing — sign in again, or have an admin consent |
| `AADSTS50011` (redirect URI mismatch) | Add `http://localhost` under *Mobile and desktop applications*, or pin `loopbackPort` |
| `AADSTS50020` | The account is not a member/guest of that tenant, or the app is single-tenant |
| Signed out after about an hour | The tenant grants no `offline_access`, so there is no refresh token. `teams_doctor` names the missing scope; otherwise sign in again when it expires |
| Browser never opens | No browser is reachable here; `teams_doctor` says why. `teams_login mode: device-code` always works |
| Browser opens, page never returns | A firewall is blocking the loopback port. Pin `loopbackPort` and allow it, or use the device code flow |
| "Not signed in" | Run `teams_login`; the saved session may have been revoked |
| Status line says "not signed in" although sign-in worked | The line is written at session start and after every auth tool; if it lags, `/teams-status` repaints it |
| "Blocked by configuration" | Your own scope rules. `teams_permissions action: test` shows which rule |
| "cannot run on an app-only token" | Set `authMode: "interactive"` and run `teams_login` |
| 403 reading channel messages | `ChannelMessage.Read.All` needs **admin** consent — see "Who has to approve what" |
| Listen mode does nothing | `teams_watch action: status` says whether it is enabled and running. Only messages arriving *after* it was switched on wake pi |
| Listen mode answers too much | Narrow `watch.chats` and `watch.from`, raise `cooldownSeconds`, lower `maxTriggersPerHour` |
| Listen mode stopped by itself | `teams_watch action: status` shows the last polling error; a revoked session or a sleeping laptop is the usual cause, and it resumes on its own |
| Throttled | Graph rate limit; the client retries with back-off, then reports it |

Run `teams_doctor` first — it checks all of the above at once.

## License

MIT
