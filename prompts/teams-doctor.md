---
description: Diagnose and fix the Microsoft Teams connection
---
Something is wrong with my Teams connection.

1. Run `teams_doctor` for every configured account.

2. Explain each problem in plain language, and say who can fix it — me or a
   tenant administrator. In particular:
   - missing consent → name the exact Graph scope and what it is needed for,
     and say whether it needs an admin (`ChannelMessage.Read.All` and
     `ChannelMember.Read.All` do; the rest I can consent to myself)
   - "Allow public client flows" disabled → the setting is in Entra ID →
     App registrations → Authentication
   - a redirect URI mismatch → `http://localhost` must be registered under
     "Mobile and desktop applications"
   - the browser sign-in unavailable → say why, and that
     `teams_login mode: device-code` works anywhere
   - expired or revoked sign-in → I need to run `teams_login` again
   - a guest tenant that rejects the sign-in → the app registration must be
     multi-tenant, I must be invited there, and an admin of that tenant must
     have consented

3. Run `teams_permissions` with `action: "show"` and tell me, in one short
   paragraph, what you are currently allowed to do and what you are not.

4. Propose the concrete next step. Do not change my configuration without
   asking first.
