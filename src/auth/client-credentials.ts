/**
 * Client credentials grant — app-only access, no user involved.
 *
 * Useful for unattended jobs (a nightly digest, a scheduled report) against
 * directory and presence data. It is deliberately *not* a way to post as
 * somebody: Microsoft Graph only accepts app-only writes to chats and channels
 * for migration scenarios (`Teamwork.Migrate.All`), so this extension refuses
 * message sends on an app-only token instead of letting them fail deep inside
 * a Graph call. See ../safety/index.ts.
 */

import type { TeamsConnection } from "../config/index.ts";
import { AuthError, tokenUrl } from "./device-code.ts";
import type { CachedToken } from "./token-store.ts";

interface TokenResponse {
	access_token: string;
	expires_in: number;
	scope?: string;
	token_type: string;
	error?: string;
	error_description?: string;
}

/**
 * Acquire an app-only token.
 *
 * @throws {AuthError} when the secret is missing, expired, or the app has no
 * admin-consented application permissions.
 */
export async function acquireAppToken(
	conn: TeamsConnection,
	signal?: AbortSignal,
): Promise<CachedToken> {
	if (!conn.clientSecret) {
		throw new AuthError(
			"missing_client_secret",
			`Account "${conn.account}" uses client-credentials but no clientSecret is configured.`,
		);
	}

	const body = new URLSearchParams({
		grant_type: "client_credentials",
		client_id: conn.clientId,
		client_secret: conn.clientSecret,
		scope: conn.scopes.length === 1 ? conn.scopes[0]! : "https://graph.microsoft.com/.default",
	});

	const response = await fetch(tokenUrl(conn.authorityHost, conn.tenantId), {
		method: "POST",
		headers: { "Content-Type": "application/x-www-form-urlencoded" },
		body: body.toString(),
		signal,
	});

	const json = (await response.json()) as TokenResponse;

	if (!response.ok || !json.access_token) {
		throw new AuthError(
			json.error ?? `http_${response.status}`,
			json.error_description ?? `Client credentials request failed with HTTP ${response.status}`,
		);
	}

	return {
		accessToken: json.access_token,
		expiresAt: Date.now() + Math.max(0, json.expires_in - 60) * 1000,
		scopes: json.scope ? json.scope.split(" ") : [],
		authMode: "client-credentials",
		updatedAt: new Date().toISOString(),
	};
}
