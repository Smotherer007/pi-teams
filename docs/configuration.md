# Configuration

Everything lives in `~/.pi/agent/pi-teams.json` (mode `0600`, created from a
template on first run). `teams_setup`, `teams_accounts`, `teams_permissions` and
`teams_watch` write the same file — use whichever you prefer.

## A complete example

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

## Accounts and tenants

- An **account** is a Teams identity you sign in as. Each has its own saved
  session.
- A **tenant** beneath an account is another directory that identity reaches: a
  customer, a subsidiary, a guest tenant. It gets its own sign-in, its own
  safety level and its own rules.

Every tool takes `account` and `tenant`:

```yaml
teams_send_channel_message:
  account: work
  tenant: customer-alpha
  channel: "Project Alpha/General"
  body: "Sprint review moved to Thursday."
```

### Per-account keys

| Key | Meaning |
|-----|---------|
| `name` | short label used in every tool's `account` parameter |
| `displayName` | friendly name for status output |
| `tenantId` | tenant ID or domain |
| `clientId` | application (client) ID of the Entra registration |
| `authMode` | `interactive`, `device-code`, `client-credentials` or `auto` |
| `clientSecret` | only for `client-credentials` |
| `scopes` | override the requested Graph scopes |
| `loopbackPort` | pin the browser redirect port |
| `safetyLevel` | overrides the global level |
| `permissions` | scope rules layered on the global ones |
| `aiFooter` | overrides the global footer, field by field |
| `watch` | listen-mode settings for this account |
| `tenants` | additional tenants beneath this identity |

A tenant accepts the same keys except `tenants` and `displayName`.

## Safety levels

| Level | Effect |
|-------|--------|
| `open` | writes go through without asking |
| `confirm` | every write shows a confirmation with the destination and the exact text (**default**) |
| `readonly` | writes are blocked entirely |

Set globally, per account, or per tenant — the most specific setting wins. So
your own company can be `confirm` while a customer tenant stays `readonly`.

## Scope rules

Rules exist per **category** (`teams`, `channels`, `chats`, `people`) and per
**mode** (`read`, `write`). Category keys written directly in the block apply to
both modes; a `read` or `write` block refines one of them.

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

## AI disclosure footer

On by default. Every message, reply and edit pi sends in your name ends with a
short note saying it was written by an AI:

```json
"aiFooter": { "enabled": true, "text": "🤖 Erstellt mit pi (KI-Assistent)" }
```

| Key | Default | Meaning |
|-----|---------|---------|
| `enabled` | `true` | append the note to everything pi sends |
| `text` | `🤖 Generated with pi (an AI agent)` | the wording; empty means the default |

Set it globally or per account — the account level wins field by field, so one
account can use its own wording while another keeps the global one.

Details worth knowing:

- It is added by the tool that sends, not asked for in a prompt, so it cannot be
  forgotten by the model. The confirmation dialog at `safetyLevel: confirm` shows
  the text **with** the footer — you confirm what actually goes out.
- It is idempotent: editing an already-footered message does not stack a second
  one, and a model that copied the footer from the chat it is answering does not
  produce two.
- It applies to chat messages, channel posts and replies, the first message of a
  newly created chat, and edits. It does **not** apply to things pi sends on your
  behalf that are not messages — meeting invitations, presence changes.
- It says nothing about the *content*: the note is a disclosure, not a
  disclaimer, and it is the only part of an outgoing message pi adds on its own.

## Listen mode

The `watch` block is documented with the feature it configures, in
[listen-mode.md](listen-mode.md).

## Everything else

| Key | Default | Meaning |
|-----|---------|---------|
| `defaultAccount` | first account | account used when a tool omits `account` |
| `defaultTenant` | the account's home tenant | tenant used when a tool omits `tenant` |
| `maxMessages` | `25` | default page size for message listings |
| `audit` | `true` | append every write to `~/.pi/agent/pi-teams-audit.jsonl` |
| `downloadDir` | — | an additional directory `teams_download_files` may write to; see [security.md](security.md#where-downloads-may-land) |
| `graphBaseUrl` / `authorityHost` | Microsoft public cloud | sovereign cloud endpoints |

`scopes` and `aiFooter` are listed under the account keys above; both can also
be set globally.
