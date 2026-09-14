# pi-teams

Microsoft Teams integration for the [pi coding agent](https://github.com/earendil-works/pi).

pi acts **as you**: it reads what you can read, and anything it sends comes
from your account, in your name, in the right thread. Multiple Teams accounts
and multiple company tenants live side by side, and a configuration file
decides exactly which teams, channels, chats and people pi may touch.

## Installation

```bash
# From npm
pi install npm:@patimweb/pi-teams

# From a local checkout during development
pi install /path/to/pi-teams
```

## Quick start

1. **Register an app in Entra ID** (once per organization — see below).
2. **Configure the account:**
   ```
   teams_setup:
     name: work
     tenantId: contoso.onmicrosoft.com
     clientId: <application (client) ID>
     setDefault: true
   ```
3. **Sign in:** run `/teams-login`. pi shows a code and a URL; open it, enter
   the code, sign in as yourself.
4. **Check it works:** `/teams-status`, then `/teams-inbox`.

---

## Entra ID app registration

pi signs in with the OAuth 2.0 **device authorization grant**, so it never sees
your password and your MFA works normally.

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
| `offline_access` | staying signed in |

Some of these require **admin consent**. Click *Grant admin consent* if you
can, or ask an administrator. `teams_doctor` reports exactly which scopes were
granted and which were not, so you never have to guess from a 403.

To request fewer permissions, set `scopes` on the account — see below.

### 4. Guest and customer tenants

For a tenant that is not your own:

- the registration must be **multitenant**,
- your account must be a guest in that tenant,
- an administrator there consents to the app once.

Then add it as a tenant beneath your account (`parentAccount`) and sign in
again against it.

---

## Acting as a user vs. app-only

| | `device-code` (default) | `client-credentials` |
|---|---|---|
| Identity | you | the application |
| Sign-in | once, interactive | none |
| Reading chats/channels | yes | only via protected APIs |
| **Sending messages** | **yes** | **no** — refused with an explanation |
| Presence, calendar | yes | partially |

Microsoft Graph only permits app-only writes to chats and channels for
migration scenarios (`Teamwork.Migrate.All`), so pi-teams refuses those calls
up front instead of letting them fail deep inside a Graph request. Use
`client-credentials` for unattended read-only jobs, `device-code` for
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
      "displayName": "NeoImpulse",
      "tenantId": "neoimpulse.onmicrosoft.com",
      "clientId": "00000000-0000-0000-0000-000000000000",
      "authMode": "device-code",
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
          "people": { "deny": ["ceo@neoimpulse.de"] }
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
| `graphBaseUrl` / `authorityHost` | Microsoft public cloud | sovereign cloud endpoints |

---

## Tools

### Session and configuration

| Tool | Description |
|------|-------------|
| `teams_setup` | Write an account or tenant to the config file |
| `teams_login` | Sign in as yourself (device code flow) |
| `teams_logout` | Remove a saved session |
| `teams_accounts` | List accounts, switch the default, change a safety level, delete |
| `teams_status` | Which account is active, who pi acts as, what it may do |
| `teams_permissions` | Show, test or change the allow/deny rules |
| `teams_doctor` | Diagnose config, sign-in, consent and connectivity |

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
| `teams_delete_message` | Soft-delete (or restore) a message you sent |
| `teams_create_channel` | Add a channel to a team |
| `teams_set_presence` | Set your Teams status |
| `teams_set_status_message` | Set the note under your name |
| `teams_create_meeting` | Schedule a Teams meeting with invitations |
| `teams_update_meeting` | Change a meeting you organize |
| `teams_cancel_meeting` | Cancel a meeting |

## Commands

| Command | Description |
|---------|-------------|
| `/teams-status` | Show the connection and permissions |
| `/teams-login` | Start the device code sign-in |
| `/teams-inbox` | What needs your attention right now |
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
| `/teams-doctor` | Guided troubleshooting |

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
| `~/.pi/agent/pi-teams-tokens.json` | access and refresh tokens per account+tenant | `0600` |
| `~/.pi/agent/pi-teams-audit.jsonl` | one line per write: when, who, where, what | `0600` |

Both the config and the token cache are written through a private temp file and
an atomic rename, so an interrupted write cannot truncate them. To revoke
pi's access entirely, run `teams_logout` with `all: true` and remove the app's
consent in Entra ID.

## Architecture

Data-oriented: plain immutable data, I/O at the edges, pure functions in the
middle.

- **`src/types.ts`** — domain data as plain interfaces. No behavior.
- **`src/config/index.ts`** — accounts, tenants, the safety cascade, persistence.
- **`src/config/scope.ts`** — allow/deny matching. Pure, and the most heavily
  tested module in the package.
- **`src/auth/`** — device code flow, client credentials, token cache, JWT claims.
- **`src/graph/client.ts`** — fetch wrapper: bearer token, paging, throttling.
- **`src/graph/*.ts`** — one module per resource area, returning domain types.
- **`src/graph/mappers.ts`** — Graph JSON → domain types. Pure.
- **`src/utils/formatting.ts`** — domain types → display strings. Pure.
- **`src/safety/`** — the three gates and the audit log.
- **`src/tools/`** — one module per tool.
- **`src/extension/index.ts`** — registration, commands, the `tool_call` interceptor.

## Development

```bash
npm install
npm run typecheck
npm test
```

Tests cover the pure logic — scope matching, the config cascade, message body
construction, HTML flattening, the safety gates — and run without a tenant or
a network connection.

## Troubleshooting

| Symptom | Cause and fix |
|---------|---------------|
| `AADSTS7000218` | "Allow public client flows" is off in the app registration |
| `AADSTS65001` | Consent missing — sign in again, or have an admin consent |
| `AADSTS50020` | The account is not a member/guest of that tenant, or the app is single-tenant |
| "Not signed in" | Run `teams_login`; the refresh token may have been revoked |
| "Blocked by configuration" | Your own scope rules. `teams_permissions action: test` shows which rule |
| "cannot run on an app-only token" | Set `authMode: "device-code"` and run `teams_login` |
| 403 on channel messages | `ChannelMessage.Read.All` / `ChannelMessage.Send` not consented |
| Throttled | Graph rate limit; the client retries with back-off, then reports it |

Run `teams_doctor` first — it checks all of the above at once.

## License

MIT
