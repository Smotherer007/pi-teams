---
description: Summarize what happened in Microsoft Teams while you were away
argument-hint: "[hours, e.g. 24]"
---
Catch me up on Microsoft Teams.

1. Run `teams_inbox` with `hours` set to $@ (default 24 if nothing was given).

2. For every chat with new messages, run `teams_read_chat` to read the actual
   messages — the preview alone is not enough to judge whether something needs
   an answer.

3. For every mention, run `teams_read_thread` (channel) or `teams_read_chat`
   (chat) to get the surrounding context.

4. Summarize as:
   - **Needs a reply from me** — who is waiting, on what, and since when
   - **FYI** — decisions, announcements and progress I should know about
   - **Noise** — one line, just the count

   Keep each item to one or two sentences and name the people involved.

5. For anything in "Needs a reply", propose a short draft reply, but do not
   send anything. Ask me which drafts to send.
