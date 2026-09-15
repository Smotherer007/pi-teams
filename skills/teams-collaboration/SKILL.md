---
name: teams-collaboration
description: Microsoft Teams chats, channels, meetings and presence, acting as the signed-in user. Use when the user asks to read, search, summarize, send, edit, retract or reply to Teams messages; to add or remove someone from a chat; to catch up on what they missed and clear it; to post in a channel or a thread; to react to a message; to mark a chat read or unread; to check or set their Teams status; to see who is available; to find a time everyone is free; to schedule, change, answer, or cancel a Teams meeting; or to make pi notice incoming Teams messages on its own (listen mode). Also covers how to format a message for a chat rather than a document, multiple Teams accounts and multiple company tenants, and the allow/deny rules that decide what pi may touch.
---

# Microsoft Teams

pi acts **as the signed-in user**. A message it sends is indistinguishable from
one the user typed: same name, same avatar, same thread. Treat every write as
something the user is personally saying.

## Before anything else

Run `teams_status` if you are unsure which account is active or whether pi is
signed in. If it reports "not signed in", run `teams_login`: it opens the
user's browser and completes on its own. On a machine with no browser it falls
back to a device code — then show the code and the URL to the user verbatim,
because they cannot finish signing in without them.

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

## When a tool says a scope is missing

One optional action needs a permission the account does not request by default:
removing someone from a **chat** (`ChatMember.ReadWrite`). The tool names the
exact scope. Pass that sentence on as the fix — add it to `scopes` in
`pi-teams.json`, sign in again — or do it in Teams. Do not look for a roundabout
way to achieve the same thing, and do not report it as a permission bug.

A few things cannot be done at all with the permissions pi asks for: creating a
channel, and editing or deleting a **channel** post. Those need
`Channel.Create` / `ChannelMessage.ReadWrite`, they are not requested, and they
are not assumed to be grantable. Say so plainly and point at the Teams client
instead of trying anyway.

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
| `teams_availability` | "When are we all free?" — before booking, not after |
| `teams_list_meetings` / `teams_get_meeting` | Calendar and meeting details |

### Writing

| Tool | Use when |
|------|----------|
| `teams_send_chat_message` | Sending a chat message |
| `teams_create_chat` | Starting a new conversation |
| `teams_send_channel_message` | Posting a new channel message |
| `teams_reply_channel_message` | Answering inside an existing thread |
| `teams_react` | Adding or removing an emoji reaction |
| `teams_update_message` | Fixing a message the user already sent |
| `teams_chat_members` | Bringing someone into a group chat, or taking them out |
| `teams_mark_read` | Clearing the unread marker after a catch-up |
| `teams_delete_message` | Taking back something the user sent |
| `teams_set_presence` / `teams_set_status_message` | Changing the user's status |
| `teams_create_meeting` / `teams_update_meeting` / `teams_cancel_meeting` | Scheduling work |
| `teams_respond_invite` | Accepting, declining or tentatively answering an invitation |
| `teams_watch` | Letting pi notice incoming messages on its own |

## How to write as the user

- Write in **their** voice. No "As an AI…", no signature, no note that a model
  composed it — that disclosure is configuration, not judgement: the sending
  tools append it themselves (`aiFooter`, on by default), so do not write it
  into the body or it arrives twice (the tools detect the duplicate, but the
  user sees the text in the confirmation twice as well, which looks like a
  mistake).
- Match the register of the conversation you just read. A channel post and a
  quick reply to a colleague are not the same thing.
- Show the exact text and the destination before sending whenever the request
  leaves any room for interpretation. Once it is sent, it is sent.
- Never invent facts, commitments, dates or agreement on the user's behalf. If
  a detail is missing, ask.
- Prefer `teams_reply_channel_message` over a new post when answering something
  — starting a second thread fragments the conversation.

## How to format a Teams message

A Teams chat is a chat bubble, not a document. What arrives badly is long
prose: the answer is buried, the reader has to hunt for it, and it reads as if a
machine wrote it. "Keep it short" does not fix that on its own — a shape does.
The body is rendered as HTML, and markdown is converted for you, so write
markdown and let the tool do the rest.

**The shape, which is nearly always this:**

```
QA läuft wieder.

- Ursache: abgelaufenes Zertifikat
- Fix: erneuert und neu deployt
- Offen: Monitoring für 90 Tage
```

Line one is the answer. Then a blank line, then only what the reader needs in
order to act on it. Five lines is already a long chat message.

**Before and after.** Both versions say the same thing; only one is a chat:

> **So nicht:** Ich habe mir die QA-Umgebung jetzt angeschaut und konnte den
> Fehler reproduzieren. Das Problem liegt darin, dass das TLS-Zertifikat
> abgelaufen war, wodurch alle Requests fehlgeschlagen sind. Ich habe es dann
> erneuert und die Umgebung neu deployt, jetzt läuft wieder alles. Ich würde
> allerdings noch empfehlen, das Monitoring einzurichten, damit das nicht
> wieder passiert.

> **So:** QA läuft wieder.
>
> - Ursache: abgelaufenes TLS-Zertifikat
> - Fix: erneuert und neu deployt
> - Empfehlung: Monitoring auf die Ablaufdaten

**Three skeletons that cover most messages:**

*Antwort auf eine Frage* — the answer, then at most one line of reasoning:

```
Ja, Freitag passt.

Ich bin bis 14:00 im Workshop, danach jederzeit.
```

*Status* — outcome first, then what is still open:

```
Deployment ist durch, Release Notes stehen im Wiki: <Link>

Offen: Smoke-Test auf Produktion, mache ich heute Nachmittag.
```

*Entscheidung nötig* — the decision, the options, the deadline:

```
Ich brauche bis morgen eine Entscheidung zum Sync-Intervall.

- 5 Min: schnell, ~3× Last auf dem SAP-System
- 15 Min: entspannt, dafür Verzögerung im Reporting
```

**Rules**

- **The first line carries the message.** Someone who reads only that line
  should know the outcome.
- **Bullets for anything enumerable** — causes, options, next steps. Three items
  and up are a list; two are a sentence.
- **One idea per line.** A line break is free; a clause chain held together by
  commas is not.
- **Bold the one thing that must not be missed**, and only that. Bold used
  everywhere is bold nowhere.
- **Numbers, names and dates exactly as they are.** Do not round a figure or
  paraphrase a deadline into something friendlier.
- **One question per message.** Two questions get one answer.
- **No greeting and no closing.** "Hallo zusammen", "Ich habe mir das mal
  angeschaut", "Kurz zusammengefasst:" and "Sag Bescheid, wenn du Fragen hast"
  are all noise in a chat. The message starts with the answer and stops when it
  is done.
- **Over five lines, it is a document.** Put it in a wiki page or a file and send
  the link with one sentence around it.

**What Teams renders:** `**bold**`, `*italic*`, `` `code` ``, `~~struck~~`,
`[label](url)`, `- bullets`, `1. numbered`, and code fences. A `# heading`
arrives as a bold line — fine for structure, not for hierarchy. **Tables are not
supported in a chat**: convert them to bullets, or put the table in a file and
link it.

**Language and tone:** answer in the language of the conversation, in the
register of the people in it. Short forms are normal in a chat ("passt", "ok,
5 min") — but never shorten a fact into ambiguity.

When pi listens on its own (see *Listen mode*), this matters most: an answer
that arrives unrequested must be readable at a glance, because the recipient is
not expecting it.

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

## Read state

`teams_mark_read` moves the user's read cursor on a **chat** — the same thing
as opening it, or re-marking it unread, in the app. Two rules:

- It is a **write**, and it goes through the safety gates like any other. Mark
  a chat read when the user asks, or when they clearly want the list cleared —
  not on your own initiative after reading it.
- Where the tenant shows **read receipts**, the sender can see that the chat was
  opened. That is not a reason to refuse, only a reason not to do it silently.

Channels have no equivalent: channel posts cannot be marked read or unread
through Graph, so "mark the channel as read" is done in the Teams client.

pi reads this same cursor when it decides what to answer in listen mode: a chat
whose newest message is already read does not wake it. So marking a chat read is
also how the user says "I have handled this, do not answer it for me".

## Correcting a message

Both actions work on **chat** messages the user themselves sent, which is also
the only case that makes sense: pi wrote them in the user's name. A channel post
cannot be edited or deleted with the permissions pi asks for — for that, point
at the Teams client.

- **`teams_update_message` replaces the text and keeps the message in place.**
  Teams shows it as edited in the same bubble. This is the right answer to a
  typo, a wrong time, a missing detail.
- **`teams_delete_message` takes it back.** Right when the message should not
  exist at all: wrong chat, wrong content, sent too early.

Never correct by deleting and resending. The replacement arrives as a brand-new
notification, everyone who already read the first message has to notice the gap,
and the thread loses its order. Editing is the polite version.

An edit sends the text only: mentions and attachments from the original are not
re-created. That is deliberate — re-sending mention entities would ping people a
second time for a message they were already notified about.

## Chat membership

`teams_chat_members` adds a person to a group chat or removes one. Adding
someone hands them the conversation, so:

- Resolve the person first (`teams_find_user`) when only a first name was given,
  and say the full name in your reply. "Tom" is not an identity.
- A **1:1** chat cannot be extended — Graph refuses. Use `teams_create_chat`
  with everyone who belongs in it.
- Removing someone needs an extra permission (`ChatMember.ReadWrite`); the tool
  says so. Adding works with the default scopes.

## Finding a time before booking one

`teams_availability` reads **free/busy**, never calendars: "is Tuesday at two
possible" is answerable, "what is Anna doing on Tuesday" is not, and must never
be presented as if it were.

1. Run it with everybody who has to be there and the length of the meeting.
2. Offer the user two or three of the slots it returns — not all of them.
3. Only once they pick one, book it with `teams_create_meeting`.

Never book a slot because it looked free, and never dress the busy blocks up as
anything more than times: knowing that somebody is busy is the point, knowing why
is their business.

## Answering an invitation

`teams_respond_invite` accepts, declines or marks an invitation as tentative,
and tells the organizer unless `sendResponse: false`. This fills the user's
calendar and answers another person, so it needs their decision — "say yes to
whatever looks relevant" is not one.

A meeting the user **organizes** has no invitation to answer: change it with
`teams_update_meeting`, cancel it with `teams_cancel_meeting`.

## Listen mode

Listen mode is the only case where pi acts without being asked: it polls the
user's chats, and an incoming message is turned into a prompt. Off by default,
and it costs one model turn per message that wakes it.

```yaml
teams_watch:
  action: enable
  chats: ["Anna*", "Vertrieb*"]     # where to listen
  from: ["anna.schmidt@contoso.com"] # who to listen to
  mentionOnly: false
  intervalSeconds: 60
```

- `chats` decides **where** pi listens (topic, label, chat ID, participants),
  `from` decides **to whom** (name, UPN, e-mail). Narrow both: a watcher that
  watches everything answers everything.
- `/teams-listen on|off|status` switches it in the session; `teams_watch
  action: status` reports what is configured and what is actually running.
- pi answers what is still **unread**: a chat wakes it when it has moved since
  pi last looked (a cursor kept on disk per account) and its newest message is
  still unread for the user in Teams. Switching listen mode back on therefore
  answers what was missed while pi was not running; a chat the user has already
  read in Teams stays quiet. A backlog is drained a few chats per tick, capped
  by `maxTriggersPerHour` — so on a fresh start, say how many chats are waiting
  rather than assuming silence means nothing happened.
- pi never wakes for its own messages, so it cannot answer itself.
- Listen mode does **not** bypass anything: a reply it decides to send is
  subject to the same safety level and scope rules as a reply you asked for. At
  `safetyLevel: confirm` every one of them is confirmed with the user first.

**When a listen prompt arrives**, do this:

1. Judge the message before answering. "Danke, passt" needs no reply; a
   question, a blocker or a decision does.
2. If an answer is warranted, send it with `teams_send_chat_message` to the
   **chat ID from the prompt** — never re-resolve the chat by name. pi posts as
   the user, so the wrong chat is a wrong statement in the user's name.
3. Keep it to the shape in *How to format a Teams message*: the answer on the
   first line, at most five short lines, bullets for anything enumerable. An
   unrequested message has to be readable at a glance.
4. If no answer is warranted, say so in one line and stop. Do not react, do not
   touch other chats, do not post to a channel.

Never let listen mode turn into a conversation with itself: if the incoming
message is already an answer to something pi sent, the loop ends there.
