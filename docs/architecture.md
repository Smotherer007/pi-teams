# Architecture

Data-oriented: plain immutable data, I/O at the edges, pure functions in the
middle. The rule that shapes most of it — anything that decides something is
pure and tested; anything that talks to Graph is thin and boring.

## Modules

| Module | Responsibility |
|--------|----------------|
| `src/types.ts` | Domain data as plain interfaces. No behavior. |
| `src/config/index.ts` | Accounts, tenants, the safety cascade, listen-mode settings, persistence. |
| `src/config/scope.ts` | Allow/deny matching. Pure, and the most heavily tested module in the package. |
| `src/auth/` | MSAL applications, the file-backed token cache, JWT claims. |
| `src/graph/client.ts` | fetch wrapper: bearer token, paging, throttling, and the origin check that keeps the token on Graph. |
| `src/graph/*.ts` | One module per resource area, returning domain types. |
| `src/graph/mappers.ts` | Graph JSON → domain types. Pure. |
| `src/watch/index.ts` | Listen mode's decision rules. Pure. |
| `src/watch/loop.ts` | The polling loop and its state. |
| `src/watch/prompt.ts` | What an incoming message turns into. |
| `src/watch/cursor.ts` | The durable record of what has already been looked at. |
| `src/watch/pin.ts` | The chat an answer is owed to, enforced by the interceptor. |
| `src/utils/formatting.ts` | Domain types → display strings. Pure. |
| `src/utils/richtext.ts` | Markdown → the HTML subset Teams renders. Pure. |
| `src/utils/slots.ts` | Availability view → the free slots everybody shares. Pure, and the part of scheduling that can be wrong without looking wrong. |
| `src/utils/attachments.ts` | What a message carries besides text. Pure. |
| `src/utils/disclosure.ts` | The AI footer. Idempotent by construction. |
| `src/safety/` | The gates and the audit log. |
| `src/tools/` | One module per tool. |
| `src/extension/index.ts` | Registration, commands, the `tool_call` interceptor, and the listen-mode lifecycle. |

## Why the watch loop is split in two

`watch/index.ts` answers *should this wake pi?* from data alone — the chat list,
one message, and what has already been seen. `watch/loop.ts` does the polling and
the Graph calls. That line is what makes the decision rules testable without a
tenant, and they are the rules most likely to be wrong in a way nobody notices.

## Development

```bash
npm install
npm run typecheck
npm test
```

Tests cover the pure logic — scope matching, the config cascade, message body
construction, markdown rendering, HTML flattening, free-slot arithmetic, the
listen-mode decision rules, the safety gates, the download origin check — and run
without a tenant or a network connection.

`npm test` picks the right flags for the Node version: type stripping needs
`--experimental-strip-types` on Node 22 and nothing on 23.6+, which is why the
runner is a small script rather than a raw `node --test`.

## Release

`semantic-release` on `main` and `next`, driven by conventional commits. CI runs
typecheck and tests on Node 22 and 24.
