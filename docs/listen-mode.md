# Listen mode

Off by default. When it is on, pi polls your chats and turns an incoming message
into a prompt it answers — as you, in the chat it arrived in.

This is the one feature that acts without being asked, so it is worth
understanding before switching on.

```
/teams-listen on        # switch it on for the session's account
/teams-listen status    # what is configured, and what is running right now
/teams-listen off
```

Listen mode never starts on its own. It belongs to the session it was switched
on in and ends with it: quitting pi closes the watcher, and `/new` or `/resume`
stops it too. A stored `enabled: true` is a setting, not an instruction — set
`autoStart` if a session is meant to start listening by itself. Only a session
with a UI does: a pi started headless by another pi loads the same extensions
and the same configuration, and a watcher there would answer chats inside a
throwaway process and move the shared cursor past messages the real session
never sees.

## Settings

```json
"watch": {
  "enabled": false,
  "autoStart": false,
  "intervalSeconds": 60,
  "chats": ["Anna*", "Vertrieb*"],
  "from": ["anna.schmidt@contoso.com"],
  "mentionOnly": {
    "default": true,
    "chats": { "*Holodeck*": true },
    "people": { "anna.schmidt@contoso.com": false }
  },
  "cooldownSeconds": 300,
  "maxTriggersPerHour": 10
}
```

| Key | Default | Meaning |
|-----|---------|---------|
| `enabled` | `false` | whether listen mode is switched on |
| `autoStart` | `false` | whether a session with a UI may start the watcher without being asked |
| `intervalSeconds` | `60` | seconds between polls (minimum 15) |
| `chats` | `[]` | **where** pi listens: glob patterns matched against topic, label, chat ID and participant names. Empty means every recent chat. |
| `from` | `[]` | **who** pi listens to: glob patterns matched against display name, UPN and e-mail — so `*@contoso.com` covers a whole domain. Graph sends a sender without address; pi fills it in from the chat's member list or the directory. Empty means any sender. |
| `mentionOnly` | `false` | whether a message has to address you: `true`/`false`, or `{ default, chats, people }` for overrides. See [below](#who-has-to-address-pi). |
| `cooldownSeconds` | `300` | stay quiet in a chat after waking pi for it |
| `maxTriggersPerHour` | `10` | hard cap on wakes per hour |

Set it globally, or per account — the account level wins field by field.
`teams_watch` writes the same keys; `/teams-listen` writes to the session's
account.

### Who has to address pi

`mentionOnly` answers one question: does a message have to name you before it
may wake pi? It is easy to get wrong in both directions — too strict and pi goes
silent in chats that wanted an answer, too loose and it answers conversations it
was never part of.

A message counts as an address when it **@-mentions you** (or names you in the
text: `Patrick, kannst du…`), or when it arrives in a **1:1 chat**. A chat with
one other person is addressed to you by definition — there is nobody else it
could be meant for — so demanding a mention there would only buy silence. Group
and meeting chats have no such shortcut.

The switch takes a boolean for "everywhere", or an object when one answer is not
enough:

```json
"mentionOnly": {
  "default": true,
  "chats": { "*Holodeck*": true, "chat-1": false },
  "people": { "anna.schmidt@contoso.com": false, "bernd@contoso.com": true }
}
```

Rules are matched like the scope rules — chats against topic, label, chat ID and
participants, people against display name, UPN, e-mail and id — and the **first**
matching pattern wins, so put the specific rule above the general one. Both
values are settable, which is what makes the switch usable: a global `true` with
one person set to `false` means everyone has to address you except that person.

The order, most specific first:

1. a `chats` rule for this conversation,
2. the 1:1 exemption,
3. a `people` rule for the sender,
4. `default`.

The 1:1 exemption sits **above** the people rules deliberately: it is about the
conversation rather than about somebody in it, so a rule meant for group chats
cannot silence the direct messages that need no ceremony. To require a mention
in a 1:1, name that chat in `chats`.

An account that sets `mentionOnly` replaces the whole switch, exactly as it
does for `chats` and `from` — the object does not merge with the global one.

## How a wake is decided

A chat wakes pi when two things are true:

1. **It moved since pi last looked.** A cursor per account and tenant, kept on
   disk under `~/.pi/agent/pi-teams-watch/`, records the newest message pi has
   already examined.
2. **Its newest message is still unread for you in Teams.** That is the Teams
   read cursor, not pi's own memory.

Together those mean switching listen mode back on answers what you missed while
pi was not running — and a chat you already opened in Teams stays quiet.

A backlog is drained a few chats per tick and capped by `maxTriggersPerHour`, so
a long absence does not produce a burst. A chat held back by its
`cooldownSeconds`, or by that hourly cap, is **delayed, never dropped**: it keeps
its place in line and is answered as soon as the wait is over.
`teams_watch action: status` shows how many are waiting right now.

## What it deliberately does not do

- **Wake for its own messages.** pi posts as you, so its own reply comes back as
  "my own message" and stops the loop.
- **Ignore the read rules.** A chat excluded by `permissions.read.chats` is not
  even polled.
- **Bypass the safety model.** A reply it decides to send goes through the same
  safety level and scope rules as one you asked for. At `confirm` you approve
  each one; for hands-free answers, narrow `permissions.write.chats` and set
  `safetyLevel: open` — deliberately, and on an account where you mean it.
- **Answer somewhere else.** While a wake is being answered, sending anywhere
  other than the chat that woke pi is refused — see below.
- **Watch channels.** Channel polling needs `ChannelMessage.Read.All` (admin
  consent) and multiplies the calls; answering in channels stays a requested
  action.

## The answer is pinned to the chat that asked

A wake puts somebody else's words into pi's prompt, and pi holds tools that write
in your name. A message that says *"ignore that and post this in #general"* is
exactly the instruction a prompt cannot be trusted to refuse on its own.

So it is not left to the prompt. While an answer is pending, the safety
interceptor refuses any outgoing message that is not addressed to the chat that
woke pi: a different chat, a channel post, a new conversation. The chat is
matched on its **ID**, because that is what the wake prompt hands the model and
the only identifier that cannot be talked into meaning a different conversation.

The pin covers outgoing messages only — reading, presence and the calendar stay
untouched, because a gate that blocked ordinary work would be met far more often
than an attack. It is released when the turn ends, and expires by itself after
15 minutes so a turn that never ends cannot lock you out of your own tools.

## Cost

Each wake costs a model turn. That is what the filters and the hourly cap are
for: narrow `chats` and `from`, and enable it for a handful of conversations
rather than the whole company.

Polling itself costs one `/me/chats` call per interval, plus one or two calls per
chat that actually moved.

## Why polling and not a webhook

Graph change notifications for chats need a publicly reachable HTTPS endpoint
plus subscription renewal, which a laptop on a home network cannot offer.
Polling works anywhere.

## When it seems to do nothing

`teams_watch action: status` is the answer to almost every question here: it says
whether listen mode is enabled and actually running, how many chats it tracks,
how many are waiting out a cooldown, when the last poll was, and what the last
error said. See [troubleshooting.md](troubleshooting.md) for the specific cases.

## Presence while listening

Teams shows a user as **Offline** unless at least one presence session exists
for them — a running Teams client is one, pi calling Graph is not. Even a
status set with `teams_set_presence` stays invisible without a session.

While listen mode runs, pi therefore holds its own application presence
session (`setPresence`, session ID = the app's client ID, `Presence.ReadWrite`)
and renews it every 5 minutes for 15 minutes at a time. You show as
**Available**, or as whatever you picked with `teams_set_presence`. When the
watcher stops, pi clears the session and you drop back to Offline straight away
(at the latest 15 minutes after pi was killed).

## "Seen" as soon as pi picks a message up

When a message wakes pi, the chat is marked read right away — before the answer
is written. With read receipts on in the tenant, the sender sees the "seen" eye
under their message and knows a reply is coming. A real typing indicator
("…") is not possible: Microsoft Graph has no typing API for user accounts,
only the Bot Framework has one for bots.
