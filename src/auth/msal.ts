/**
 * MSAL application factory and the browser plumbing around it.
 *
 * MSAL handles the parts of OAuth that are tedious to get right — PKCE, the
 * loopback listener, refresh-token rotation, cross-tenant authorities — so
 * this module's job is narrow: build one application per connection, decide
 * whether a browser is actually reachable, and open it.
 */

import { spawn } from "node:child_process";
import {
	ConfidentialClientApplication,
	LogLevel,
	PublicClientApplication,
	type Configuration,
} from "@azure/msal-node";
import type { TeamsConnection } from "../config/index.ts";
import { browserUnavailableReason, canOpenBrowser } from "../utils/environment.ts";
import { createCachePlugin } from "./cache-plugin.ts";

export { browserUnavailableReason, canOpenBrowser };

// ---------------------------------------------------------------------------
// Reserved scopes
// ---------------------------------------------------------------------------

/**
 * MSAL adds these itself and rejects them in a request's scope list.
 * They stay in the documented defaults because the doctor reports on them.
 */
const RESERVED_SCOPES = new Set(["openid", "profile", "offline_access", "email"]);

export function requestScopes(conn: TeamsConnection): string[] {
	return conn.scopes.filter((scope) => !RESERVED_SCOPES.has(scope.toLowerCase()));
}

// ---------------------------------------------------------------------------
// Application factory
// ---------------------------------------------------------------------------

const publicApps = new Map<string, PublicClientApplication>();
const confidentialApps = new Map<string, ConfidentialClientApplication>();

function appKey(conn: TeamsConnection): string {
	return `${conn.account.toLowerCase()}::${conn.tenantId.toLowerCase()}::${conn.clientId}`;
}

function baseConfig(conn: TeamsConnection): Configuration {
	return {
		auth: {
			clientId: conn.clientId,
			authority: `${conn.authorityHost}/${conn.tenantId}`,
		},
		cache: {
			cachePlugin: createCachePlugin(conn.account, conn.tenantId),
		},
		system: {
			loggerOptions: {
				// MSAL logs to stdout by default, which would scribble over the TUI.
				loggerCallback: () => {},
				piiLoggingEnabled: false,
				logLevel: LogLevel.Error,
			},
		},
	};
}

/** The public client — used for every delegated (act-as-the-user) flow. */
export function getPublicApp(conn: TeamsConnection): PublicClientApplication {
	const key = appKey(conn);
	const existing = publicApps.get(key);
	if (existing) return existing;

	const app = new PublicClientApplication(baseConfig(conn));
	publicApps.set(key, app);
	return app;
}

/** The confidential client — app-only tokens. */
export function getConfidentialApp(conn: TeamsConnection): ConfidentialClientApplication {
	const key = appKey(conn);
	const existing = confidentialApps.get(key);
	if (existing) return existing;

	const config = baseConfig(conn);
	const app = new ConfidentialClientApplication({
		...config,
		auth: { ...config.auth, clientSecret: conn.clientSecret },
	});
	confidentialApps.set(key, app);
	return app;
}

/** Drop the cached applications (after a sign-out or a config change). */
export function resetApps(): void {
	publicApps.clear();
	confidentialApps.clear();
}

// ---------------------------------------------------------------------------
// Browser
// ---------------------------------------------------------------------------

/** Open a URL in the user's default browser, detached from this process. */
export function openBrowser(url: string): Promise<void> {
	return new Promise((resolve, reject) => {
		const [command, args] =
			process.platform === "darwin"
				? ["open", [url]]
				: process.platform === "win32"
					? // `start` is a shell builtin, and the empty string is the window
						// title cmd.exe would otherwise take the URL for.
						["cmd.exe", ["/c", "start", "", url]]
					: ["xdg-open", [url]];

		try {
			const child = spawn(command as string, args as string[], {
				detached: true,
				stdio: "ignore",
			});
			child.on("error", reject);
			child.unref();
			resolve();
		} catch (err) {
			reject(err instanceof Error ? err : new Error(String(err)));
		}
	});
}

// ---------------------------------------------------------------------------
// Browser response pages
// ---------------------------------------------------------------------------

function page(title: string, message: string, accent: string): string {
	return `<!doctype html>
<html lang="en"><head><meta charset="utf-8">
<title>${title}</title>
<style>
  :root { color-scheme: light dark; }
  body { margin: 0; min-height: 100vh; display: grid; place-items: center;
         font: 16px/1.5 system-ui, -apple-system, "Segoe UI", sans-serif;
         background: #fbfbfa; color: #1a1a18; }
  @media (prefers-color-scheme: dark) { body { background: #1a1a18; color: #f5f4ef; } }
  .card { max-width: 28rem; padding: 2.5rem; text-align: center; }
  .mark { width: 3rem; height: 3rem; margin: 0 auto 1.25rem; border-radius: 50%;
          background: ${accent}; display: grid; place-items: center;
          color: #fff; font-size: 1.5rem; }
  h1 { font-size: 1.25rem; margin: 0 0 .5rem; }
  p { margin: 0; opacity: .7; }
</style></head>
<body><div class="card">
  <div class="mark">${accent === "#2f7d4f" ? "&check;" : "!"}</div>
  <h1>${title}</h1>
  <p>${message}</p>
</div></body></html>`;
}

export const SUCCESS_TEMPLATE = page(
	"Signed in to Teams",
	"pi can now act as you in Microsoft Teams. You can close this tab and go back to your terminal.",
	"#2f7d4f",
);

export const ERROR_TEMPLATE = page(
	"Sign-in failed",
	"Microsoft rejected the sign-in. Close this tab and check the error in your terminal.",
	"#b03a2e",
);
