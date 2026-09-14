/**
 * OAuth 2.0 device authorization grant — the flow that makes pi *you*.
 *
 * The user opens a URL, types a short code, and signs in with their own
 * credentials and MFA. Everything pi does afterwards is attributed to that
 * person in Teams: messages come from them, and pi sees exactly what they see —
 * no more.
 *
 * Implemented against the token endpoint directly rather than through MSAL:
 * the extension ships with no build step, and one fetch-based module is easier
 * to audit than a dependency tree.
 */

import type { TeamsConnection } from "../config/index.ts";
import { decodeIdToken } from "./jwt.ts";
import type { CachedToken } from "./token-store.ts";

// ---------------------------------------------------------------------------
// Types
// ---------------------------------------------------------------------------

/** What the user needs in order to complete the sign-in. */
export interface DeviceCodeChallenge {
	deviceCode: string;
	userCode: string;
	verificationUri: string;
	/** Message Microsoft suggests showing to the user */
	message: string;
	/** Epoch ms after which the code is dead */
	expiresAt: number;
	/** Seconds between polls, as requested by the server */
	intervalSeconds: number;
}

interface DeviceCodeResponse {
	device_code: string;
	user_code: string;
	verification_uri: string;
	expires_in: number;
	interval: number;
	message: string;
}

interface TokenResponse {
	access_token: string;
	refresh_token?: string;
	id_token?: string;
	expires_in: number;
	scope?: string;
	token_type: string;
}

interface ErrorResponse {
	error: string;
	error_description?: string;
}

/** A token endpoint error that callers may want to branch on. */
export class AuthError extends Error {
	readonly code: string;
	constructor(code: string, message: string) {
		super(message);
		this.name = "AuthError";
		this.code = code;
	}
}

// ---------------------------------------------------------------------------
// Endpoints
// ---------------------------------------------------------------------------

function deviceCodeUrl(conn: TeamsConnection): string {
	return `${conn.authorityHost}/${encodeURIComponent(conn.tenantId)}/oauth2/v2.0/devicecode`;
}

export function tokenUrl(authorityHost: string, tenantId: string): string {
	return `${authorityHost}/${encodeURIComponent(tenantId)}/oauth2/v2.0/token`;
}

// ---------------------------------------------------------------------------
// Step 1 — request a code
// ---------------------------------------------------------------------------

/**
 * Ask Microsoft for a device code.
 *
 * @throws {AuthError} when the app registration rejects the request — most
 * often because "Allow public client flows" is still off.
 */
export async function requestDeviceCode(
	conn: TeamsConnection,
	signal?: AbortSignal,
): Promise<DeviceCodeChallenge> {
	const body = new URLSearchParams({
		client_id: conn.clientId,
		scope: conn.scopes.join(" "),
	});

	const response = await fetch(deviceCodeUrl(conn), {
		method: "POST",
		headers: { "Content-Type": "application/x-www-form-urlencoded" },
		body: body.toString(),
		signal,
	});

	const json = (await response.json()) as DeviceCodeResponse & Partial<ErrorResponse>;

	if (!response.ok || json.error) {
		throw new AuthError(
			json.error ?? `http_${response.status}`,
			json.error_description ?? `Device code request failed with HTTP ${response.status}`,
		);
	}

	return {
		deviceCode: json.device_code,
		userCode: json.user_code,
		verificationUri: json.verification_uri,
		message: json.message,
		expiresAt: Date.now() + json.expires_in * 1000,
		intervalSeconds: json.interval || 5,
	};
}

// ---------------------------------------------------------------------------
// Step 2 — poll until the user is done
// ---------------------------------------------------------------------------

function sleep(ms: number, signal?: AbortSignal): Promise<void> {
	return new Promise((resolve, reject) => {
		const timer = setTimeout(resolve, ms);
		signal?.addEventListener(
			"abort",
			() => {
				clearTimeout(timer);
				reject(new AuthError("aborted", "Sign-in cancelled."));
			},
			{ once: true },
		);
	});
}

/**
 * Poll the token endpoint until the user finishes signing in.
 *
 * `onPending` is called on every poll so a tool can keep the UI alive during
 * what may be a minute or two of waiting.
 */
export async function pollForToken(
	conn: TeamsConnection,
	challenge: DeviceCodeChallenge,
	options: { signal?: AbortSignal; onPending?: (secondsLeft: number) => void } = {},
): Promise<CachedToken> {
	const { signal, onPending } = options;
	let intervalMs = challenge.intervalSeconds * 1000;

	for (;;) {
		if (Date.now() > challenge.expiresAt) {
			throw new AuthError("expired_token", "The device code expired before sign-in completed.");
		}

		await sleep(intervalMs, signal);
		onPending?.(Math.max(0, Math.round((challenge.expiresAt - Date.now()) / 1000)));

		const body = new URLSearchParams({
			grant_type: "urn:ietf:params:oauth:grant-type:device_code",
			client_id: conn.clientId,
			device_code: challenge.deviceCode,
		});

		const response = await fetch(tokenUrl(conn.authorityHost, conn.tenantId), {
			method: "POST",
			headers: { "Content-Type": "application/x-www-form-urlencoded" },
			body: body.toString(),
			signal,
		});

		const json = (await response.json()) as TokenResponse & Partial<ErrorResponse>;

		if (response.ok && json.access_token) {
			return toCachedToken(json, "device-code");
		}

		switch (json.error) {
			case "authorization_pending":
				continue;
			case "slow_down":
				// The server asks for more breathing room; honour it permanently.
				intervalMs += 5000;
				continue;
			case "authorization_declined":
				throw new AuthError("authorization_declined", "Sign-in was declined in the browser.");
			case "expired_token":
				throw new AuthError("expired_token", "The device code expired before sign-in completed.");
			case "bad_verification_code":
				throw new AuthError("bad_verification_code", "The device code was rejected. Start over.");
			default:
				throw new AuthError(
					json.error ?? `http_${response.status}`,
					json.error_description ?? `Token request failed with HTTP ${response.status}`,
				);
		}
	}
}

// ---------------------------------------------------------------------------
// Step 3 — refresh
// ---------------------------------------------------------------------------

/**
 * Exchange a refresh token for a fresh access token.
 *
 * @throws {AuthError} with code `invalid_grant` when the refresh token is dead
 * (password change, revoked session, conditional access) — the caller should
 * then ask the user to run `teams_login` again.
 */
export async function refreshAccessToken(
	conn: TeamsConnection,
	refreshToken: string,
	signal?: AbortSignal,
): Promise<CachedToken> {
	const body = new URLSearchParams({
		grant_type: "refresh_token",
		client_id: conn.clientId,
		refresh_token: refreshToken,
		scope: conn.scopes.join(" "),
	});

	const response = await fetch(tokenUrl(conn.authorityHost, conn.tenantId), {
		method: "POST",
		headers: { "Content-Type": "application/x-www-form-urlencoded" },
		body: body.toString(),
		signal,
	});

	const json = (await response.json()) as TokenResponse & Partial<ErrorResponse>;

	if (!response.ok || !json.access_token) {
		throw new AuthError(
			json.error ?? `http_${response.status}`,
			json.error_description ?? `Token refresh failed with HTTP ${response.status}`,
		);
	}

	// Entra ID rotates refresh tokens; keep the old one only if none came back.
	return { ...toCachedToken(json, "device-code"), refreshToken: json.refresh_token ?? refreshToken };
}

// ---------------------------------------------------------------------------
// Shaping
// ---------------------------------------------------------------------------

function toCachedToken(
	json: TokenResponse,
	authMode: "device-code" | "client-credentials",
): CachedToken {
	const claims = json.id_token ? decodeIdToken(json.id_token) : undefined;
	return {
		accessToken: json.access_token,
		refreshToken: json.refresh_token,
		// Shave 60s off the stated lifetime so clock skew never hands Graph a
		// token that expired a second ago.
		expiresAt: Date.now() + Math.max(0, json.expires_in - 60) * 1000,
		scopes: json.scope ? json.scope.split(" ") : [],
		authMode,
		user: claims
			? {
					id: claims.oid,
					displayName: claims.name,
					upn: claims.preferred_username ?? claims.upn,
					tenantId: claims.tid,
				}
			: undefined,
		updatedAt: new Date().toISOString(),
	};
}
