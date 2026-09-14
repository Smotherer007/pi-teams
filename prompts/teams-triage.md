---
description: Work through unanswered Teams messages one by one
argument-hint: "[hours]"
---
Triage my Microsoft Teams messages from the last $@ hours (default 24).

1. Run `teams_inbox` to find what is waiting.

2. For each item, read enough context (`teams_read_chat` or
   `teams_read_thread`) to judge it, then classify it as:
   - **Answer now** — a short reply unblocks someone
   - **Needs me** — requires a decision or information only I have
   - **Delegate** — someone else should handle it; say who
   - **Ignore** — no action needed

3. Present the list as a table: who, what, classification, and the chat or
   channel it lives in.

4. For everything in "Answer now", draft a reply in my voice. Show me all the
   drafts together, then send only the ones I approve, one tool call each.

5. Do not react, reply or change anything before I approve it.
