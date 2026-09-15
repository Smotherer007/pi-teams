/**
 * Auth resolver — hands the Graph layer a valid bearer token.
 *
 * Built on MSAL, so the flows that matter are one call each:
 *
 *   - **interactive**: `acquireTokenInteractive` opens the system browser at
 *     the Microsoft sign-in page, listens on a loopback port for the redirect,
 *     and does PKCE. The default whenever a browser is actually reachable.
 *   - **device code**: `acquireTokenByDeviceCode` for SSH, containers and
 *     anything headless — the user types a short code on another device.
 *   - **client credentials**: app-only, no user. Cannot post as a person; the
 *     safety layer refuses those tools rather than letting Graph reject them.
 *
 * Silent renewal is MSAL's `acquireTokenSilent`, which uses the cached refresh
 * token and rotates it. The cache lives on disk (see ./cache-plugin.ts), so a
 * sign-in survives across pi sessions.
 */

import {
	InteractionRequiredAuthError,
	type AccountInfo,
	type AuthenticationResult,
} from "@azure/msal-node";
import type { TeamsConnection } from "../config/index.ts";
import { removeAllCaches, removeCache, readCacheSummary } from "./cache-plugin.ts";
import {
	ERROR_TEMPLATE,
	SUCCESS_TEMPLATE,
	browserUnavailableReason,
	canOpenBrowser,
	getConfidentialApp,
	getPublicApp,
	openBrowser,
	requestScopes,
	resetApps,
} from "./msal.ts";

// ---------------------------------------------------------------------------
// Types
// ---------------------------------------------------------------------------

/** A usable bearer token plus who it belongs to. */
export interface TokenInfo {
	accessToken: string;
	expiresOn?: Date;
	scopes: string[];
	account?: SignedInAccount;
}

export interface SignedInAccount {
	homeAccountId: string;
	username: string;
	name?: string;
	tenantId: string;
}

/** How a sign-in was performed, or should be. */
export type SignInMode = "interactive" | "device-code";

/** What the user must do to complete a device code sign-in. */
export interface DeviceCodeInfo {
	userCode: string;
	verificationUri: string;
	message: string;
	expiresIn: number;
}

/** A failure from the identity provider. */
export class AuthError extends Error {
	readonly code: string;
	constructor(code: string, message: string) {
		super(message);
		this.name = "AuthError";
		this.code = code;
	}
}

/** Raised when there is no usable token and only the user can fix it. */
export class NotSignedInError extends Error {
	readonly account: string;
	readonly tenant: string;
	constructor(account: string, tenant: string, detail?: string) {
		super(
			`Not signed in to Teams account "${account}"${tenant !== account ? ` (tenant ${tenant})` : ""}. ` +
				`Run the teams_login tool or the /teams-login command.${detail ? ` (${detail})` : ""}`,
		);
		this.name = "NotSignedInError";
		this.account = account;
		this.tenant = tenant;
	}
}

// ---------------------------------------------------------------------------
// Shaping
// ---------------------------------------------------------------------------

function toAccount(account: AccountInfo | null): SignedInAccount | undefined {
	if (!account) return undefined;
	return {
		homeAccountId: account.homeAccountId,
		username: account.username,
		name: account.name,
		tenantId: account.tenantId,
	};
}

function toTokenInfo(result: AuthenticationResult): TokenInfo {
	return {
		accessToken: result.accessToken,
		expiresOn: result.expiresOn ?? undefined,
		scopes: result.scopes ?? [],
		account: toAccount(result.account),
	};
}

function wrapError(err: unknown): never {
	if (err instanceof AuthError || err instanceof NotSignedInError) throw err;
	const error = err as { errorCode?: string; errorMessage?: string; message?: string };
	throw new AuthError(
		error.errorCode ?? "auth_failed",
		error.errorMessage ?? error.message ?? String(err),
	);
}

// ---------------------------------------------------------------------------
// Account lookup
// ---------------------------------------------------------------------------

/** The cached account for a connection, if the user has signed in before. */
export async function cachedAccount(conn: TeamsConnection): Promise<SignedInAccount | undefined> {
	const accounts = await getPublicApp(conn).getTokenCache().getAllAccounts();
	if (accounts.length === 0) return undefined;

	// One cache file per account+tenant, so there is normally exactly one. When
	// a user has signed in as more than one identity, prefer this tenant's.
	const match =
		accounts.find((account) => account.tenantId?.toLowerCase() === conn.tenantId.toLowerCase()) ??
		accounts[0];
	return toAccount(match ?? null);
}

/** Cheap, synchronous "has this connection ever signed in" — for status output. */
export function hasCachedSession(conn: TeamsConnection): boolean {
	return readCacheSummary(conn.account, conn.tenantId).present;
}

// ---------------------------------------------------------------------------
// Silent acquisition
// ---------------------------------------------------------------------------

/**
 * Get a token for this connection, renewing it silently when needed.
 *
 * @throws {NotSignedInError} when only an interactive sign-in can help
 * @throws {AuthError} when the identity provider refuses
 */
export async function getAccessToken(
	conn: TeamsConnection,
	_signal?: AbortSignal,
): Promise<TokenInfo> {
	if (conn.authMode === "client-credentials") {
		if (!conn.clientSecret) {
			throw new AuthError(
				"missing_client_secret",
				`Account "${conn.account}" uses client-credentials but has no clientSecret.`,
			);
		}
		try {
			const result = await getConfidentialApp(conn).acquireTokenByClientCredential({
				scopes: ["https://graph.microsoft.com/.default"],
			});
			if (!result) throw new AuthError("no_token", "Microsoft returned no app-only token.");
			return toTokenInfo(result);
		} catch (err) {
			wrapError(err);
		}
	}

	const app = getPublicApp(conn);
	const accounts = await app.getTokenCache().getAllAccounts();
	const account =
		accounts.find((entry) => entry.tenantId?.toLowerCase() === conn.tenantId.toLowerCase()) ??
		accounts[0];

	if (!account) throw new NotSignedInError(conn.account, conn.tenant);

	try {
		const result = await app.acquireTokenSilent({
			account,
			scopes: requestScopes(conn),
		});
		if (!result) throw new NotSignedInError(conn.account, conn.tenant);
		return toTokenInfo(result);
	} catch (err) {
		// Password change, revoked session, conditional access, a new scope the
		// cached refresh token was never granted — all need the user back.
		if (err instanceof InteractionRequiredAuthError) {
			throw new NotSignedInError(
				conn.account,
				conn.tenant,
				"the saved sign-in expired or needs to be renewed",
			);
		}
		wrapError(err);
	}
}

// ---------------------------------------------------------------------------
// Sign-in
// ---------------------------------------------------------------------------

export interface SignInOptions {
	/** Force a flow instead of letting the environment decide */
	mode?: SignInMode;
	/** Called once with the code and URL when the device code flow is used */
	onDeviceCode?: (info: DeviceCodeInfo) => void;
	/** Called just before the browser opens */
	onBrowser?: (url: string) => void;
	/** Fixed loopback port, for tenants that require an exact redirect URI */
	loopbackPort?: number;
	/** How long to wait for the user, in milliseconds (default 5 minutes) */
	timeoutMs?: number;
	signal?: AbortSignal;
}

export interface SignInResult extends TokenInfo {
	/** Which flow actually ran */
	mode: SignInMode;
}

/** Pick a flow for this environment. */
export function defaultSignInMode(): SignInMode {
	return canOpenBrowser() ? "interactive" : "device-code";
}

export { browserUnavailableReason, canOpenBrowser };

function withTimeout<T>(promise: Promise<T>, ms: number, label: string, hint?: string): Promise<T> {
	return new Promise<T>((resolve, reject) => {
		const timer = setTimeout(() => {
			reject(
				new AuthError(
					"timeout",
					`${label} timed out after ${Math.round(ms / 1000)}s.${hint ? ` ${hint}` : ""}`,
				),
			);
		}, ms);
		promise.then(
			(value) => {
				clearTimeout(timer);
				resolve(value);
			},
			(err) => {
				clearTimeout(timer);
				reject(err);
			},
		);
	});
}

/**
 * What it means when a browser sign-in never comes back.
 *
 * Three unrelated causes look identical from here: the browser opens and no
 * callback ever arrives. Naming all of them in the timeout text is the
 * difference between a five-second fix and hunting the launcher, which is
 * usually innocent — so the text asks the one question that separates them.
 */
const BROWSER_TIMEOUT_HINT =
	"The browser opened but nothing came back. Check what the browser showed: " +
	"a consent that is still pending (AADSTS65004 — an admin request waiting under " +
	"Enterprise applications), a Conditional Access policy that blocks the sign-in " +
	"(AADSTS53003), or a redirect URI on the app registration that does not match " +
	"this port (AADSTS50011). If no browser opened at all, the launcher is the " +
	"problem: set PI_TEAMS_BROWSER to your own command, or open the URL that came " +
	"with the error by hand.";

/**
 * Sign in as the user.
 *
 * Interactive by default: the browser opens, the user signs in with their own
 * credentials and MFA, and MSAL catches the redirect on a loopback port. On a
 * machine where that cannot work, the device code flow takes over.
 */
export async function signIn(
	conn: TeamsConnection,
	options: SignInOptions = {},
): Promise<SignInResult> {
	if (conn.authMode === "client-credentials") {
		const token = await getAccessToken(conn, options.signal);
		return { ...token, mode: "device-code" };
	}

	const mode = options.mode ?? (conn.authMode === "device-code" ? "device-code" : defaultSignInMode());
	const app = getPublicApp(conn);
	const scopes = requestScopes(conn);
	const timeoutMs = options.timeoutMs ?? 5 * 60 * 1000;

	if (mode === "interactive") {
		try {
			const result = await withTimeout(
				app.acquireTokenInteractive({
					scopes,
					openBrowser: async (url: string) => {
						options.onBrowser?.(url);
						await openBrowser(url);
					},
					successTemplate: SUCCESS_TEMPLATE,
					errorTemplate: ERROR_TEMPLATE,
					...(options.loopbackPort ? { preferredPort: options.loopbackPort } : {}),
				}),
				timeoutMs,
				"Browser sign-in",
				BROWSER_TIMEOUT_HINT,
			);
			if (!result) throw new AuthError("no_token", "Microsoft returned no token.");
			return { ...toTokenInfo(result), mode: "interactive" };
		} catch (err) {
			wrapError(err);
		}
	}

	try {
		const result = await withTimeout(
			app.acquireTokenByDeviceCode({
				scopes,
				deviceCodeCallback: (response) => {
					options.onDeviceCode?.({
						userCode: response.userCode,
						verificationUri: response.verificationUri,
						message: response.message,
						expiresIn: response.expiresIn,
					});
				},
			}),
			timeoutMs,
			"Device code sign-in",
		);
		if (!result) throw new AuthError("no_token", "Microsoft returned no token.");
		return { ...toTokenInfo(result), mode: "device-code" };
	} catch (err) {
		wrapError(err);
	}
}

// ---------------------------------------------------------------------------
// Sign-out
// ---------------------------------------------------------------------------

/** Remove one cached session. Returns true when something was removed. */
export async function signOut(conn: TeamsConnection): Promise<boolean> {
	// Ask MSAL to forget the account first so its in-memory copy goes too;
	// removing only the file would leave this process still holding tokens.
	try {
		const cache = getPublicApp(conn).getTokenCache();
		for (const account of await cache.getAllAccounts()) {
			await cache.removeAccount(account);
		}
	} catch {
		/* the file removal below is what actually matters */
	}

	const removed = removeCache(conn.account, conn.tenantId);
	resetApps();
	return removed;
}

/** Remove every cached session. Returns how many were removed. */
export function signOutAll(): number {
	const removed = removeAllCaches();
	resetApps();
	return removed;
}

export { resetApps };
