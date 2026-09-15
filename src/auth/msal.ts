/**
 * MSAL application factory and the browser plumbing around it.
 *
 * MSAL handles the parts of OAuth that are tedious to get right — PKCE, the
 * loopback listener, refresh-token rotation, cross-tenant authorities — so
 * this module's job is narrow: build one application per connection, decide
 * whether a browser is actually reachable, and open it.
 */

import { spawn, type ChildProcess } from "node:child_process";
import {
	ConfidentialClientApplication,
	LogLevel,
	PublicClientApplication,
	type Configuration,
} from "@azure/msal-node";
import type { TeamsConnection } from "../config/index.ts";
import { browserUnavailableReason, canOpenBrowser, isWsl } from "../utils/environment.ts";
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

/**
 * The Windows-side shell a WSL session hands its URL to.
 *
 * Windows PowerShell rather than `cmd.exe`: the authorize URL carries several
 * `&`, and `cmd /c start` reads the first of them as a command separator and
 * opens a truncated URL. Single quotes survive them.
 */
const WSL_POWERSHELL = "/mnt/c/Windows/System32/WindowsPowerShell/v1.0/powershell.exe";

/** How long to wait for a launcher that neither exits nor errors. */
const LAUNCH_EXIT_GRACE_MS = 2000;

/**
 * The command that puts `url` in front of the user in a browser.
 *
 * Split out of `openBrowser` so the shapes stay testable: the Windows and WSL
 * forms are easy to get subtly wrong, and all of them fail quietly when they
 * are.
 */
export function browserLaunch(
	url: string,
	target: { platform?: NodeJS.Platform; wsl?: boolean; browser?: string } = {},
): { command: string; args: string[] } {
	const override = (target.browser ?? process.env.PI_TEAMS_BROWSER)?.trim();
	if (override) return overrideLaunch(url, override);

	const platform = target.platform ?? process.platform;

	if (target.wsl ?? isWsl()) {
		// Escaping a quote inside a PowerShell single-quoted string means doubling
		// it. Nothing else in the URL needs attention.
		const quoted = url.replace(/'/g, "''");
		return {
			command: WSL_POWERSHELL,
			args: [
				"-NoProfile",
				"-NonInteractive",
				// Without this a console window flashes on the Windows side for the
				// fraction of a second the launcher is alive. Cosmetic, but the kind
				// of cosmetic that makes a tool feel broken on every single sign-in.
				"-WindowStyle",
				"Hidden",
				"-Command",
				`Start-Process '${quoted}'`,
			],
		};
	}

	if (platform === "darwin") return { command: "open", args: [url] };

	if (platform === "win32") {
		// `start` is a shell builtin, and the empty string is the window title
		// cmd.exe would otherwise take the URL for.
		return { command: "cmd.exe", args: ["/c", "start", "", url] };
	}

	return { command: "xdg-open", args: [url] };
}

/**
 * A launcher the user supplied through `PI_TEAMS_BROWSER`.
 *
 * Takes a complete command line. `{}` in any argument is replaced by the URL;
 * without it the URL is appended. This is the escape hatch for a WSL distro
 * with no `xdg-open`: point it at a script, or at a Windows binary, instead of
 * installing a shim for every program on the machine.
 */
function overrideLaunch(url: string, override: string): { command: string; args: string[] } {
	const [command, ...rest] = splitCommandLine(override);
	if (!command) throw new Error("PI_TEAMS_BROWSER is set but names no command.");

	if (rest.some((part) => part.includes("{}"))) {
		return { command, args: rest.map((part) => part.replaceAll("{}", url)) };
	}
	return { command, args: [...rest, url] };
}

/** Split on whitespace, honouring simple quoting. No expansion, no escaping. */
function splitCommandLine(input: string): string[] {
	const parts: string[] = [];
	let current = "";
	let quote: string | undefined;

	for (const char of input) {
		if (quote) {
			if (char === quote) quote = undefined;
			else current += char;
			continue;
		}
		if (char === '"' || char === "'") {
			quote = char;
			continue;
		}
		if (/\s/.test(char)) {
			if (current) parts.push(current);
			current = "";
			continue;
		}
		current += char;
	}

	if (current) parts.push(current);
	return parts;
}

/**
 * Run a launcher and wait until it is clear whether it worked.
 *
 * Two things have to be caught, and both were lost when this resolved on spawn.
 * A missing launcher — `xdg-open` on a fresh WSL distro — reports ENOENT on
 * `error` and never exits. A launcher that is present but finds no browser
 * exits non-zero, which is how `xdg-open` says so.
 *
 * Resolving on spawn made both look like success, so the sign-in sat waiting
 * for a redirect from a browser that had never opened, until the five minute
 * timeout replaced the real cause with "Browser sign-in timed out".
 */
export function launchBrowser(
	command: string,
	args: string[],
	graceMs = LAUNCH_EXIT_GRACE_MS,
): Promise<void> {
	return new Promise((resolve, reject) => {
		let child: ChildProcess;
		try {
			child = spawn(command, args, { detached: true, stdio: ["ignore", "ignore", "pipe"] });
		} catch (err) {
			reject(err instanceof Error ? err : new Error(String(err)));
			return;
		}

		let settled = false;
		let timer: ReturnType<typeof setTimeout> | undefined;
		let stderr = "";

		child.stderr?.on("data", (chunk: Buffer) => {
			stderr += String(chunk);
		});

		const settle = (err?: Error) => {
			if (settled) return;
			settled = true;
			if (timer) clearTimeout(timer);
			child.unref();
			if (err) reject(err);
			else resolve();
		};

		// A launcher that stays alive is waiting on the browser, not failing.
		timer = setTimeout(() => settle(), graceMs);

		child.once("error", (err: NodeJS.ErrnoException) => {
			const detail = err.code === "ENOENT" ? `${command} was not found` : err.message;
			settle(new Error(`could not open a browser: ${detail}`));
		});

		child.once("close", (code, signal) => {
			if (code === 0) settle();
			else {
				// `close`, not `exit`: the exit event can fire before the stderr pipe has
				// drained, and the reason xdg-open gives is on stderr.
				const why = stderr.trim() ? `: ${stderr.trim()}` : "";
				settle(
					new Error(
						`could not open a browser: ${command} exited with ` +
							`${code === null ? `signal ${signal ?? "unknown"}` : `code ${code}`}${why}`,
					),
				);
			}
		});
	});
}

/** Open a URL in the user's default browser, detached from this process. */
export async function openBrowser(url: string): Promise<void> {
	const { command, args } = browserLaunch(url);

	try {
		await launchBrowser(command, args);
	} catch (err) {
		const cause = err instanceof Error ? err.message : String(err);
		// The URL goes in every failure message. An exact exit code is only ever a
		// hint — a shim and some real xdg-open implementations exit 0 without
		// showing anything — so the one thing that always works is the address
		// itself, pasted into a browser by hand.
		const fallback = `Open it by hand instead: ${url}`;
		if (isWsl()) {
			throw new Error(
				`${cause}. ${fallback} — from WSL the handoff expects Windows PowerShell at ` +
					`${WSL_POWERSHELL}; set PI_TEAMS_BROWSER to use your own command instead.`,
			);
		}
		throw new Error(`${cause}. ${fallback}`);
	}
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
