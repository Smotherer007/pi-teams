---
description: Diagnose and fix the Microsoft Teams connection
---
Something is wrong with my Teams connection.

1. Run `teams_doctor` for every configured account.

2. Explain each problem in plain language, and say who can fix it — me or a
   tenant administrator. In particular:
   - missing consent → name the exact Graph scope and what it is needed for
   - "Allow public client flows" disabled → the setting is in Entra ID →
     App registrations → Authentication
   - expired or revoked sign-in → I need to run `teams_login` again
   - a guest tenant that rejects the sign-in → the app registration must be
     multi-tenant and I must be invited to that tenant

3. Run `teams_permissions` with `action: "show"` and tell me, in one short
   paragraph, what you are currently allowed to do and what you are not.

4. Propose the concrete next step. Do not change my configuration without
   asking first.
