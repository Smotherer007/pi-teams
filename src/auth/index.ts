/**
 * Auth resolver — hands the Graph layer a valid bearer token.
 *
 * Order of preference for a delegated connection:
 *   1. a cached access token that is still fresh
 *   2. a refresh of the cached refresh token
 *   3. nothing — the caller must run `teams_login`
 *
 * App-only connections skip the cache-refresh dance; a new token is cheap and
 * there is no user to interrupt.
 */

import type { TeamsConnection } from "../config/index.ts";
import { AuthError, refreshAccessToken } from "./device-code.ts";
import { acquireAppToken } from "./client-credentials.ts";
import { getToken, saveToken, isFresh, type CachedToken } from "./token-store.ts";

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

/** In-process cache so a burst of Graph calls does not re-read the token file. */
const memoryCache = new Map<string, CachedToken>();

function memKey(conn: TeamsConnection): string {
	return `${conn.account.toLowerCase()}::${conn.tenantId.toLowerCase()}`;
}

/**
 * Get a bearer token for this connection, refreshing it when needed.
 *
 * @throws {NotSignedInError} when a delegated sign-in is required
 * @throws {AuthError} when the identity provider refuses
 */
export async function getAccessToken(
	conn: TeamsConnection,
	signal?: AbortSignal,
): Promise<CachedToken> {
	const key = memKey(conn);

	const inMemory = memoryCache.get(key);
	if (isFresh(inMemory)) return inMemory!;

	if (conn.authMode === "client-credentials") {
		const token = await acquireAppToken(conn, signal);
		memoryCache.set(key, token);
		saveToken(conn.account, conn.tenantId, token);
		return token;
	}

	const cached = getToken(conn.account, conn.tenantId);

	if (isFresh(cached)) {
		memoryCache.set(key, cached!);
		return cached!;
	}

	if (cached?.refreshToken) {
		try {
			const refreshed = await refreshAccessToken(conn, cached.refreshToken, signal);
			// Keep the identity we already know; a refresh response carries no id_token
			// unless openid was requested again.
			const merged: CachedToken = { ...refreshed, user: refreshed.user ?? cached.user };
			memoryCache.set(key, merged);
			saveToken(conn.account, conn.tenantId, merged);
			return merged;
		} catch (err) {
			if (err instanceof AuthError && err.code === "invalid_grant") {
				throw new NotSignedInError(conn.account, conn.tenant, "the saved sign-in expired or was revoked");
			}
			throw err;
		}
	}

	throw new NotSignedInError(conn.account, conn.tenant);
}

/** Store a freshly acquired token (used by `teams_login`). */
export function storeToken(conn: TeamsConnection, token: CachedToken): void {
	memoryCache.set(memKey(conn), token);
	saveToken(conn.account, conn.tenantId, token);
}

/** Forget the in-process token — the file cache is handled by token-store. */
export function forgetToken(conn: TeamsConnection): void {
	memoryCache.delete(memKey(conn));
}

/** Drop every in-process token. */
export function clearMemoryCache(): void {
	memoryCache.clear();
}

export { AuthError };
