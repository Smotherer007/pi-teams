---
description: Summarize what happened in a Teams channel
argument-hint: "<Team/Channel> [days]"
---
Summarize a Microsoft Teams channel: $@

1. Resolve the channel with `teams_read_channel` (a `Team/Channel` path works
   directly). If no channel was named, ask which one — do not guess.

2. Read the recent posts. For every post that has replies, run
   `teams_read_thread` so the summary reflects the discussion, not just the
   opening message.

3. Produce:
   - **Decisions** — what was decided, by whom
   - **Open questions** — what is still unanswered, and who asked
   - **Action items** — who owes what, with any dates mentioned
   - **Notable** — anything else worth knowing

4. Reference the people by name and include the `messageId` for anything I
   might want to reply to.

Do not post anything in the channel.
