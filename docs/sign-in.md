# Signing in

How pi gets a token, where the session lives, and what happens on a machine
with no browser. For the portal side — registering the app and consenting to
permissions — see [entra-app-registration.md](entra-app-registration.md).

## The three flows

`teams_login` uses MSAL (`@azure/msal-node`) and picks the flow that fits the
machine.

| | `interactive` (default) | `device-code` | `client-credentials` |
|---|---|---|---|
| What you do | browser opens, you sign in | type a short code elsewhere | nothing |
| Identity | you | you | the application |
| Needs a browser | yes | no | no |
| Reading chats/channels | yes | yes | only via protected APIs |
| **Sending messages** | **yes** | **yes** | **no** — refused with an explanation |
| Presence, calendar | yes | yes | partially |

**Interactive** is chosen wherever a browser can actually be reached. MSAL opens
the system browser, listens on a loopback port for the redirect, and handles
PKCE. Nothing to copy, nothing to type.

**Device code** takes over automatically over SSH, in containers, and on any
machine without a display server — a browser opened there would appear where
nobody is looking. Force it anywhere with `PI_TEAMS_NO_BROWSER=1`, per account
with `"authMode": "device-code"`, or per call with
`teams_login mode: device-code`.

**Client credentials** is app-only: no user, no browser, and no ability to speak
as a person. Microsoft Graph permits app-only writes to chats and channels only
for migration scenarios (`Teamwork.Migrate.All`), so pi refuses those calls up
front instead of letting them fail deep inside a Graph request. Use it for
unattended read-only jobs, and a delegated flow for everything that speaks in
your name.

## Staying signed in

The session is cached on disk and renewed silently, so signing in is a
once-per-account affair. The cache lives in `~/.pi/agent/pi-teams-tokens/`, one
file per account **and** tenant, mode `0600`.

Renewal depends on `offline_access`, which MSAL always requests — you never list
it in `scopes`, and listing it there is an error MSAL rejects. If a tenant
declines it, there is no refresh token and the session ends after about an hour;
`teams_doctor` names the missing scope.

To end a session deliberately: `teams_logout`, or `teams_logout all: true` for
every account. To revoke pi's access entirely, also remove the app's consent in
Entra ID — the token cache is a copy, not the grant.

## Browsers

On macOS and Windows the system default browser is opened. On Linux and the BSDs
`xdg-open` is used, and a display server has to be present.

In **WSL** the URL is handed to Windows PowerShell so the browser opens on the
Windows side: no X server, no `DISPLAY`. PowerShell rather than `cmd.exe`,
because the authorize URL carries several `&` and `cmd /c start` reads the first
of them as a command separator and opens a truncated URL.

If none of the built-in launchers fit, set `PI_TEAMS_BROWSER` to a command line
that opens a URL — `{}` is replaced by the URL, otherwise the URL is appended:

```bash
export PI_TEAMS_BROWSER="firefox --new-window {}"
```

Every launch failure carries the URL in its message, so it can always be pasted
into a browser by hand. That is the one route that needs nothing installed.

## The loopback port

MSAL picks a free port at sign-in time, and Entra accepts any port on
`http://localhost` for a public client — which is why one redirect URI entry
covers every run.

If your tenant insists on an exact redirect URI, register e.g.
`http://localhost:3000` and pin it:

```json
{ "accounts": [{ "name": "work", "loopbackPort": 3000 }] }
```

A pinned port that is already taken makes sign-in fail rather than silently
moving, which is the point.

## Guest and customer tenants

A second tenant beneath an account is signed into separately and gets its own
cache entry:

```
teams_login:
  account: work
  tenant: customer-alpha
```

It works only if the app registration is **multitenant**, you are a guest in
that tenant, and an administrator *of that tenant* has consented to the app
once. In practice that last point is the real hurdle — see
[entra-app-registration.md](entra-app-registration.md).

## Checking the result

`teams_status` says which account is active and who pi is acting as.
`teams_doctor` goes further: it acquires a token (proving silent renewal works),
reports which scopes the token actually carries, and names the ones that are
missing, so a consent problem shows up as a permission name rather than a 403.
