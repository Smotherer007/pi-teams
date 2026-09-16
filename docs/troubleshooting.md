# Troubleshooting

**Run `teams_doctor` first.** It checks the configuration, the sign-in, which
Graph scopes were actually granted, and whether live calls succeed — for every
configured account and tenant. Most of the table below is something it already
names for you.

## Sign-in

| Symptom | Cause and fix |
|---------|---------------|
| `AADSTS500113` (no reply address) | The registration has no redirect URI. Add `http://localhost` under *Mobile and desktop applications* — see [entra-app-registration.md](entra-app-registration.md) |
| `AADSTS7000218` | *Allow public client flows* is off in the app registration |
| `AADSTS50011` (redirect URI mismatch) | Add `http://localhost` under *Mobile and desktop applications*, or pin `loopbackPort` to the port the registration lists |
| `AADSTS65001` | Consent missing — sign in again, or have an admin consent |
| `AADSTS50020` | The account is not a member or guest of that tenant, or the app is single-tenant |
| Signed out after about an hour | The tenant grants no `offline_access`, so there is no refresh token. `teams_doctor` names the missing scope; otherwise sign in again when it expires |
| "Not signed in" | Run `teams_login`; the saved session may have been revoked |
| "cannot run on an app-only token" | Set `authMode: "interactive"` and run `teams_login` |
| Status line says "not signed in" although sign-in worked | The line is written at session start and after every auth tool; if it lags, `/teams-status` repaints it |

## The browser

| Symptom | Cause and fix |
|---------|---------------|
| Browser never opens | The launcher failed: the error names it, and carries the URL to open by hand. `teams_doctor` says whether a browser counts as reachable here, and `PI_TEAMS_BROWSER` points pi at your own command |
| Browser opens, page never returns | Read what the browser showed before suspecting a firewall: a pending admin consent (`AADSTS65004`), a Conditional Access policy (`AADSTS53003`), or a redirect URI that does not match the port (`AADSTS50011`). The timeout message names all three. A blocked loopback port is the fourth possibility — pin `loopbackPort` and allow it |
| Device code appears although a browser exists | pi decided no browser is reachable — an SSH session, no display server, or `PI_TEAMS_NO_BROWSER` is set. `teams_doctor` says which |

## Permissions

| Symptom | Cause and fix |
|---------|---------------|
| 403 reading channel messages | `ChannelMessage.Read.All` needs **admin** consent — see [entra-app-registration.md](entra-app-registration.md) |
| "Blocked by configuration" | Your own scope rules, not Microsoft's. `teams_permissions action: test` shows which rule matched |
| A channel post cannot be edited or deleted | Not supported with the scopes pi requests — see [tools.md](tools.md#what-channels-cannot-do) |
| Throttled | Graph rate limit; the client retries with back-off, then reports it |

## Messages and files

| Symptom | Cause and fix |
|---------|---------------|
| A screenshot in a message is missing from the output | The text of a message cannot carry a picture. `teams_read_chat` says how many images a message has; `teams_download_files` fetches them so they can be opened |
| A file is reported as skipped, "not hosted by Microsoft Graph" | Deliberate: pi only sends your token to Graph. Open the URL in the report yourself — see [security.md](security.md#the-access-token-never-leaves-the-graph-origin) |
| "Refusing to write downloads to …" | The target directory is outside the allowed roots. Set `downloadDir` in `pi-teams.json` — see [security.md](security.md#where-downloads-may-land) |

## Listen mode

| Symptom | Cause and fix |
|---------|---------------|
| Listen mode does nothing | `teams_watch action: status` says whether it is enabled and running, how many chats it tracks, and where the cursor lives. A chat you have already read in Teams stays quiet by design, and one held back by its cooldown or the hourly cap shows up as *waiting* |
| Never answers one specific chat | `waiting` in `teams_watch action: status` says whether it is sitting out a `cooldownSeconds` window; chats it has already decided about are in the cursor under `~/.pi/agent/pi-teams-watch/` |
| Answers too much | Narrow `watch.chats` and `watch.from`, raise `cooldownSeconds`, lower `maxTriggersPerHour` |
| Stopped by itself | `teams_watch action: status` shows the last polling error; a revoked session or a sleeping laptop is the usual cause, and it resumes on its own |
| A reply to another chat was refused during a wake | Deliberate: an answer is pinned to the chat that woke pi — see [listen-mode.md](listen-mode.md#the-answer-is-pinned-to-the-chat-that-asked) |
