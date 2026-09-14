---
description: Find a time everyone is free and book the meeting
argument-hint: "[who, what, how long]"
---
Organize a Teams meeting: $@

1. Work out who has to be there. Resolve unclear names with `teams_find_user` —
   never guess an e-mail address. Ask me for anything you cannot derive from the
   request.

2. Run `teams_availability` with those people and the meeting length (30 minutes
   unless I said otherwise). If it finds nothing, widen the window with
   `start`/`end` before giving up, and tell me what you tried.

3. Show me two or three slots that fit, as weekday, date and time. **Do not book
   anything yet.**

4. Once I pick one, create it with `teams_create_meeting` in my time zone, with a
   subject that says what the meeting is about. Confirm the slot, the people and
   the subject in one line. Offer to add two or three lines of context from
   `teams_search_messages` to the invitation body — only if I say yes.

Do not invite anyone who was not named, and do not answer any invitation.
