# Security

pi sends messages that are indistinguishable from ones you typed. Teams shows no
"sent by a bot" marker. This document is about what that means, what stops it
going wrong, and what ends up on your disk.

## What pi can do in your name

Everything the signed-in user can do, within the granted scopes: read every chat
and channel you can read, send messages as you, edit and delete your own
messages, change your presence, and create, change or cancel meetings.

There is no lesser mode. A delegated token *is* your identity — that is what
makes the integration useful, and it is why the gates below exist.

## The three gates

Every write passes all three.

**1. Auth mode.** An app-only (`client-credentials`) token cannot post as a
person at all. Microsoft Graph permits app-only writes to chats and channels
only for migration scenarios, so pi refuses those tools up front rather than
letting them fail inside a Graph call.

**2. Safety level.** `readonly` blocks every write; `confirm` (the default) shows
the destination and the exact text — including the AI footer — and waits for you;
`open` proceeds. Set globally, per account or per tenant, most specific wins.

**3. Scope rules.** Allow/deny lists per category (teams, channels, chats,
people) and per mode (read, write) decide which conversations exist for pi at
all. `deny` accumulates across levels and always wins. See
[configuration.md](configuration.md#scope-rules).

On top of those, listen mode adds a fourth constraint while it is answering: the
reply is pinned to the chat that woke pi. See
[listen-mode.md](listen-mode.md#the-answer-is-pinned-to-the-chat-that-asked).

## Message content is untrusted input

Anyone who can send you a Teams message — including external and federated users
— controls text that pi will read, and in listen mode, text that lands directly
in pi's prompt. Two places take that seriously rather than assuming goodwill.

### The access token never leaves the Graph origin

A message carries URLs: a file attachment's `contentUrl`, and every `<img src>`
in its body. Those are written by the sender, not by Microsoft. Attaching your
delegated token to one of them would hand whoever sent the message a token with
`Chat.ReadWrite`, `ChannelMessage.Send` and `Calendars.ReadWrite` — your Teams
account, for the lifetime of the token.

So downloads check the origin first. The `Authorization` header is only ever sent
to the Graph endpoint the account is configured for (derived from
`graphBaseUrl`, so sovereign clouds keep working). A file hosted anywhere else is
reported as skipped, with its URL, so you can open it yourself.

Legitimate SharePoint and OneDrive attachments still download: their URL is
re-encoded as a Graph sharing token, so Graph fetches the file on your behalf and
your own permissions still decide what comes back. The request stays on the Graph
origin either way.

### Where downloads may land

`teams_download_files` writes to disk, and the directory is a tool parameter —
chosen by the model, while it is reading messages other people wrote. It is
confined to:

- the working directory,
- the pi agent directory (`~/.pi/agent/pi-teams-files/` by default),
- and whatever `downloadDir` names in `pi-teams.json`.

Anything else is refused with the list of allowed roots. Widening it is a
decision the configuration makes, not the model. File names are sanitised
separately, so a name cannot escape the directory either.

## The audit log

With `audit: true` (the default) every write appends one JSON line to
`~/.pi/agent/pi-teams-audit.jsonl`: when, which tool, which account, who pi acted
as, the target, and a summary of what was sent. Failures are recorded too, with
the error.

This is the answer to "did you really send that at 23:40?" — a question a chat
history can only answer from one side.

## Files pi writes

| File | Contents | Mode |
|------|----------|------|
| `~/.pi/agent/pi-teams.json` | accounts, tenants, safety levels, scope rules | `0600` |
| `~/.pi/agent/pi-teams-tokens/` | one MSAL token cache per account+tenant | `0600` |
| `~/.pi/agent/pi-teams-audit.jsonl` | one line per write: when, who, where, what | `0600` |
| `~/.pi/agent/pi-teams-watch/` | listen mode's cursor: the newest message it has looked at per chat | `0600` |
| `~/.pi/agent/pi-teams-files/` | images and files downloaded from messages, one folder per chat or channel | — |

The config and the token cache are written through a private temp file and an
atomic rename, so an interrupted write cannot truncate them.

Deleting the watch cursor costs one re-examination of the chats in the list, not
a flood of re-answers: the decision also asks Teams whether each message is still
unread.

## Revoking access

1. `teams_logout all: true` removes every saved session from this machine.
2. Remove the app's consent in Entra ID — *Enterprise applications* → the app →
   *Permissions* / *Users and groups*. The token cache is a copy of a grant, not
   the grant itself.
3. For a stolen refresh token specifically, revoke the user's sessions in Entra
   ID; MSAL's cache cannot outlive that.

## Reporting a problem

If you find something here that is wrong, the repository's issue tracker is in
`package.json` under `bugs`. For anything that would let one Teams user reach
another user's token or account, please report it privately first.
