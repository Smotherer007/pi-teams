# Entra ID app registration for pi-teams

The one-time setup that lets pi sign in as you. Fifteen minutes, once per
organization, and nothing in it needs a build or a server.

Everything below uses placeholders. Wherever you see

| Placeholder | Where it comes from |
|---|---|
| `<APPLICATION_CLIENT_ID>` | *Overview* → **Application (client) ID** |
| `<DIRECTORY_TENANT_ID>` | *Overview* → **Directory (tenant) ID** |
| `<TENANT_DOMAIN>` | e.g. `contoso.onmicrosoft.com`, or your verified domain |

…substitute your own values locally. Do not paste real tenant or client IDs into
a README, a ticket, a screenshot or a chat — they identify your organization,
and the client ID is half of what an attacker needs to start a sign-in flow
against your tenant.

---

## 1. Create the registration

**Entra admin center** → [entra.microsoft.com](https://entra.microsoft.com) →
**App registrations** → **New registration**

| Field | Value |
|---|---|
| Name | `pi-teams` (anything recognizable) |
| Supported account types | **Accounts in any organizational directory (multitenant)** if pi should also reach guest or customer tenants; **single tenant** is fine for one company |
| Redirect URI | leave empty here — it goes in step 2, on the right platform |

Two things to copy out of **Overview** before moving on:

- **Application (client) ID** → `<APPLICATION_CLIENT_ID>`
- **Directory (tenant) ID** → `<DIRECTORY_TENANT_ID>`

## 2. Authentication — the part that actually decides whether sign-in works

**App registrations** → your app → **Authentication**. Two settings matter, and
they are needed for two different reasons. Both are in this one blade.

Direct link, with your own application ID substituted:

```
https://entra.microsoft.com/#view/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/~/Authentication/appId/<APPLICATION_CLIENT_ID>
```

### 2a. A reply address, for the browser sign-in

**Authentication** → **Add a platform** → **Mobile and desktop applications** →
tick **`http://localhost`** → **Configure**.

MSAL picks a free loopback port when the browser sign-in starts, and Entra
treats the port-less `http://localhost` entry as matching *any* port for a
public client. One entry therefore covers every run — no per-port registration
and no fixed port needed.

Skipping this is the single most common setup mistake, and it fails *after* the
user has typed their password, with:

```
AADSTS500113: No reply address is registered for the application.
```

If your tenant insists on an exact URI, register e.g. `http://localhost:3000`
instead and pin the same number in pi (`loopbackPort` under the account, or
`teams_login port: 3000`).

### 2b. Public client flows, for the device code sign-in

Same blade → **Advanced settings** → **Allow public client flows** → **Yes** →
**Save**.

pi uses delegated OAuth flows and never handles a client secret for sign-in, so
Entra must treat the app as a public client. Without this, sign-in fails with:

```
AADSTS7000218: The request body must contain the following parameter:
'client_assertion' or 'client_secret'.
```

## 3. API permissions

**Entra admin center** → **Entra ID** → **App registrations** → your app →
**API permissions**. Four clicks, and the third one is the one people get
wrong:

1. **Add a permission** — the button above the permission list.
2. **Microsoft Graph** — the first tile under *Commonly used Microsoft APIs*.
   Not *Azure DevOps*, not *Azure Storage*: those are different APIs and their
   permissions will not help.
3. **Delegated permissions**, *not* **Application permissions**. pi acts as you.
   An application permission would give the app its own identity and a
   completely different consent model — and the two tiles sit directly under
   each other, which is why this is the usual mistake.
4. Type each name into **Start typing a permission to filter these results**,
   tick the box, and press **Add permissions** once at the end. Nothing is
   saved until that button is pressed, so a half-filled list can be closed.

What you should end up ticking:

| Permission | Needed for |
|---|---|
| `User.Read`, `User.ReadBasic.All` | identity, looking colleagues up |
| `Team.ReadBasic.All`, `Channel.ReadBasic.All` | listing teams and channels |
| `ChannelMessage.Read.All` | reading channel messages and threads |
| `ChannelMessage.Send` | posting and replying in channels |
| `ChannelMember.Read.All` | channel member lists |
| `Chat.ReadWrite` | reading chats, creating chats |
| `ChatMessage.Send` | sending chat messages |
| `Presence.ReadWrite`, `Presence.Read.All` | your status, colleagues' availability |
| `Calendars.ReadWrite` | reading and creating calendar events |
| `OnlineMeetings.ReadWrite` | Teams meeting links |
| `Files.Read.All`, `Sites.Read.All` | channel files |

`openid`, `profile` and `offline_access` are deliberately absent: MSAL always
requests them and rejects them in an explicit scope list. (You will still see
them in the picker, under **OpenId permissions**. Leave them alone.)

### The Admin consent required column

After adding them, the list has a column that decides whether step 4 is a
five-second job or a ticket to your administrator:

| Permission | Admin consent required |
|---|---|
| `ChannelMessage.Read.All` | **Yes** — reading channels and threads |
| `ChannelMember.Read.All` | **Yes** — channel member lists |
| `ChatMember.ReadWrite` | **Yes** — only if you added it, see below |
| everything else in the list above | No |

The column shows the *default* for an organization, and a tenant can relax it
per permission, per user or per app — read it as "expect to need an admin", not
as a property of the permission itself. The scopes showing **Yes** are the ones
step 4 is about.

**What is not needed:** no client secret, no certificate, no redirect URI beyond
step 2, no application (as opposed to delegated) permissions.

### Optional: one capability that needs one more permission

Everything above is what pi asks for by default. One action stays off until you
add the scope — to this list **and** to `scopes` in `pi-teams.json` — and sign in
again:

| Permission | Unlocks | Consent |
|---|---|---|
| `ChatMember.ReadWrite` | removing someone from a chat (`teams_chat_members`) | **admin** — the portal shows it as *Admin consent required: Yes*, and the tool says which scope is missing until it is granted |

Without it the tool answers with the exact scope it is missing instead of
failing at Graph with a consent error. Editing and deleting chat messages,
adding chat members and answering invitations all work with the default set.

**Deliberately not requested:** `Channel.Create` and
`ChannelMessage.ReadWrite`. They would unlock creating channels and editing or
deleting a **channel** post, and pi is built not to depend on permissions that
may never be granted. Channel posts are created, edited and deleted in the Teams
client; posting, replying and reacting in channels work without them.

**Also not covered:** `findMeetingTimes` and anything that reads another person's
*appointments* need `Calendars.Read.Shared`. `teams_availability` uses
`getSchedule` instead, which is free/busy only — no extra consent, and no view
of what people are actually doing.

## 4. Consent

**API permissions** → **Grant admin consent** — only for the two scopes that
require it:

| Scope | Consent | Without it |
|---|---|---|
| `ChannelMessage.Read.All` | **admin** | cannot read channel messages or threads |
| `ChannelMember.Read.All` | **admin** | cannot list channel members |
| everything else above | user | a normal sign-in consent prompt covers it |

Without an administrator, drop those two scopes from the configuration and
everything else still works: chats in full, posting and replying in channels,
presence, calendar, meetings, directory lookup. See the README, *Who has to
approve what*, for the reduced scope list.

If the tenant has *Users can consent to applications* switched off, even the
user-consentable scopes need an administrator — usually through the admin
consent request workflow.

## 5. Give it to pi

```yaml
teams_setup:
  name: work
  displayName: Contoso
  tenantId: "<DIRECTORY_TENANT_ID or <TENANT_DOMAIN>>"
  clientId: "<APPLICATION_CLIENT_ID>"
  setDefault: true
```

Or edit `~/.pi/agent/pi-teams.json` directly — the same file, mode `0600`:

```json
{
  "accounts": [
    {
      "name": "work",
      "tenantId": "<TENANT_DOMAIN>",
      "clientId": "<APPLICATION_CLIENT_ID>",
      "authMode": "interactive",
      "safetyLevel": "confirm"
    }
  ],
  "defaultAccount": "work"
}
```

## 6. Sign in and verify

1. `teams_login` — the browser opens at the Microsoft sign-in page, you sign in
   with your own credentials and MFA, and the page closes itself. Nothing to
   copy or type. Over SSH or in a container, pi falls back to the device code
   flow and shows a short code to enter on any device.

   In **WSL** the URL is handed to Windows PowerShell, so the browser on the
   Windows side opens — no X server and no `DISPLAY` needed. A launcher that
   cannot be started is reported as an error; if you are upgrading from a
   version that hung here instead, see the last row of the troubleshooting
   table.
2. `teams_status` — who pi acts as, which token is cached, what it may do.
3. `teams_doctor` — checks configuration, sign-in, granted scopes and
   connectivity in one go. It reports a missing consent as a **named
   permission**, which is the quickest way to see whether the admin consent of
   step 4 actually landed.

A successful sign-in is cached on disk and renewed silently, so this is a
once-per-account affair — as long as the tenant grants the session a refresh
token (see below).

## 7. Guest and customer tenants

For a tenant that is not your own:

1. the registration must be **multitenant** (step 1),
2. your account must be a **guest in that tenant**,
3. an administrator **of that tenant** consents to the app once.

In practice the third step is the hurdle, not your own tenant. Once it is done:

```yaml
teams_setup:
  parentAccount: work
  name: customer-alpha
  tenantId: "customeralpha.onmicrosoft.com"
```

then `teams_login` with `tenant: customer-alpha`, and `teams_doctor` to confirm
that tenant on its own.

---

## Troubleshooting the registration

| Error | Cause | Fix |
|---|---|---|
| `AADSTS500113` — *No reply address is registered* | no redirect URI on the registration | step 2a: add `http://localhost` under **Mobile and desktop applications** |
| `AADSTS50011` — *redirect URI mismatch* | an exact URI is registered that does not match the chosen port | register the port-less `http://localhost`, or pin the port in the same file |
| `AADSTS7000218` — *client_assertion or client_secret* | **Allow public client flows** is off | step 2b |
| `AADSTS65001` — *consent required* | permissions added but never consented | sign in again; for the two admin scopes, step 4 |
| `AADSTS65004` — *request pending* | an **admin consent request** is waiting for approval, or consent was declined once | step 4: an administrator presses **Grant admin consent**. A pending request is under **Enterprise applications** → your app → **Permissions**. Until it is approved the browser comes back with this code instead of a token |
| `AADSTS53003` — *blocked by Conditional Access* | a policy requires a managed or compliant device, which the **device code** flow cannot satisfy | this is why the browser sign-in has to work rather than quietly fall back: on such a tenant the device code route is a dead end, not a fallback. Use a machine that satisfies the policy, or add pi's loopback redirect to it |
| `AADSTS50020` — *not a member of the tenant* | wrong tenant, no guest account, or the app is single-tenant | steps 1 and 7 |
| `AADSTS700016` — *application not found in directory* | wrong tenant ID or client ID | recheck `<DIRECTORY_TENANT_ID>` and `<APPLICATION_CLIENT_ID>` |
| signed out after about an hour | the tenant grants no `offline_access` | the app needs a refresh token to renew silently; without it, run `teams_login` when the session expires. `teams_doctor` names the missing scope |
| browser never opens from WSL | older versions called `xdg-open`, which a plain WSL distro does not have — and the failure was swallowed | update pi-teams: current versions hand the URL to Windows PowerShell. If `/mnt/c` is not mounted, set `PI_TEAMS_NO_BROWSER=1` and use the device code flow — see `AADSTS53003` if Conditional Access blocks that |

Run `teams_doctor` first for anything in this table: it validates the
configuration, the cached session, the granted scopes and the connection to
Graph, and it says which of them failed rather than only that something did.

## Rotating or revoking access

- **To sign out of pi:** `teams_logout` (per account), or `all: true` for every
  account. This deletes the cached session, not the app.
- **To revoke pi's access for everyone:** Entra admin center → **Enterprise
  applications** → your app → **Properties** → *User assignment required*, or
  delete the registration. Deleting it invalidates every cached session at once.
- **To change the registration:** editing scopes or redirect URIs does not touch
  the local configuration — but it does invalidate a *new* consent, so sign in
  again afterwards.
