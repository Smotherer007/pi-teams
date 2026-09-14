---
name: teams-collaboration
description: Microsoft Teams chats, channels, meetings and presence, acting as the signed-in user. Use when the user asks to read, search, summarize, send or reply to Teams messages; to catch up on what they missed; to post in a channel or a thread; to start a chat; to react to a message; to check or set their Teams status; to see who is available; or to schedule, change or cancel a Teams meeting. Also covers multiple Teams accounts and multiple company tenants, and the allow/deny rules that decide what pi may touch.
---

# Microsoft Teams

pi acts **as the signed-in user**. A message it sends is indistinguishable from
one the user typed: same name, same avatar, same thread. Treat every write as
something the user is personally saying.

## Before anything else

Run `teams_status` if you are unsure which account is active or whether pi is
signed in. If it reports "not signed in", run `teams_login` — it prints a short
code and a URL that the user must open themselves.

If a tool fails with an authentication, permission or configuration error, run
`teams_doctor` and report what it says. Do not retry the failing call.

## Multiple accounts and companies

Every tool accepts `account` and `tenant`:

- `account` — a configured Teams identity (e.g. `work`, `private`)
- `tenant` — a tenant beneath that account: a customer, a guest tenant, a
  subsidiary

Omit both to use the default. Use `teams_accounts` to list what is configured.
When the user says "in the customer's Teams" or names a company, pass the
matching `tenant` rather than assuming the default is right.

## What pi may do

Three independent gates apply:

1. **Auth mode** — an app-only account cannot post as a person at all.
2. **Safety level** — `readonly` blocks writes, `confirm` asks the user first,
   `open` proceeds.
3. **Scope rules** — allow/deny lists per category (teams, channels, chats,
   people) and per mode (read, write).

When a call is refused with "Blocked by configuration", that is the user's own
rule set. Show them the rule, do not work around it, and do not try a different
tool to achieve the same thing. `teams_permissions` with `action: "test"`
checks a target without touching it.

## Tools

### Session

| Tool | Use when |
|------|----------|
| `teams_status` | Which account is active, who pi acts as, what it may do |
| `teams_login` | Not signed in, or adding an account/tenant |
| `teams_logout` | Removing a saved session |
| `teams_accounts` | Listing accounts, switching the default, changing a safety level |
| `teams_permissions` | Showing, testing or changing the allow/deny rules |
| `teams_doctor` | Anything failed and you need to know why |
| `teams_setup` | Adding a new account or tenant to the configuration |

### Reading

| Tool | Use when |
|------|----------|
| `teams_inbox` | "What did I miss?", "anything new in Teams?" |
| `teams_list_chats` | Finding a chat, seeing recent conversations |
| `teams_read_chat` | Reading one conversation |
| `teams_list_teams` / `teams_list_channels` | Discovering team and channel names |
| `teams_read_channel` | The posts in a channel (openers only) |
| `teams_read_thread` | One thread, opener plus every reply |
| `teams_search_messages` | Finding something said, when the location is unknown |
| `teams_list_members` | Who is in a team or channel |
| `teams_find_user` | Resolving a first name to a real person |
| `teams_list_files` | Files shared in a channel |
| `teams_get_presence` | Is someone available right now |
| `teams_list_meetings` / `teams_get_meeting` | Calendar and meeting details |

### Writing

| Tool | Use when |
|------|----------|
| `teams_send_chat_message` | Sending a chat message |
| `teams_create_chat` | Starting a new conversation |
| `teams_send_channel_message` | Posting a new channel message |
| `teams_reply_channel_message` | Answering inside an existing thread |
| `teams_react` | Adding or removing an emoji reaction |
| `teams_delete_message` | Taking back something the user sent |
| `teams_create_channel` | Adding a channel to a team |
| `teams_set_presence` / `teams_set_status_message` | Changing the user's status |
| `teams_create_meeting` / `teams_update_meeting` / `teams_cancel_meeting` | Scheduling work |

## How to write as the user

- Write in **their** voice. No "As an AI…", no signature, no note that a model
  composed it, unless they ask for one.
- Match the register of the conversation you just read. A channel post and a
  quick reply to a colleague are not the same thing.
- Show the exact text and the destination before sending whenever the request
  leaves any room for interpretation. Once it is sent, it is sent.
- Never invent facts, commitments, dates or agreement on the user's behalf. If
  a detail is missing, ask.
- Prefer `teams_reply_channel_message` over a new post when answering something
  — starting a second thread fragments the conversation.

## Channel addressing

Channels are addressed as `Team/Channel` or via separate `team` and `channel`
parameters. Names are resolved case-insensitively; an ambiguous name is
reported rather than guessed, and the answer is to pass the ID.

```yaml
teams_send_channel_message:
  channel: "Engineering/General"
  body: "Deployment finished, release notes are in the wiki."
```

## Mentions

Pass `mentions` with names, UPNs or e-mail addresses, and write `@Display Name`
in the body where the mention belongs:

```yaml
teams_send_chat_message:
  chat: "Project Alpha"
  body: "@Anna Schmidt could you take a look before the standup?"
  mentions: ["anna@contoso.com"]
```

Resolve names with `teams_find_user` first when only a first name is known.

## Threads

Channel conversations are threaded. `teams_read_channel` returns the openers
with their IDs; `teams_read_thread` returns one thread in full. To summarize a
channel properly, read the openers, then the threads that matter — do not
summarize from openers alone.

## Reading before writing

Before replying anywhere, read the recent messages first. It prevents
answering a question that was already answered, and it is what tells you the
tone to match.
