/**
 * Token cache — `~/.pi/agent/pi-teams-tokens.json`, mode 0600.
 *
 * One entry per account+tenant. Holds the access token, its expiry, the
 * refresh token, the granted scopes, and who the token belongs to, so
 * `teams_status` can say *which user* pi is currently acting as without a
 * network round-trip.
 *
 * Writes go through a private temp file and an atomic rename: the file holds
 * long-lived refresh tokens, and a truncated cache would silently sign the
 * user out.
 */

import {
	existsSync,
	readFileSync,
	writeFileSync,
	mkdirSync,
	chmodSync,
	statSync,
	renameSync,
	unlinkSync,
} from "node:fs";
import { dirname, join } from "node:path";
import { getAgentDir } from "../config/index.ts";

// ---------------------------------------------------------------------------
// Types
// ---------------------------------------------------------------------------

export interface CachedToken {
	/** Bearer token */
	accessToken: string;
	/** Epoch milliseconds when the access token expires */
	expiresAt: number;
	/** Refresh token — absent for client-credentials */
	refreshToken?: string;
	/** Scopes the token was actually granted */
	scopes: string[];
	/** Which flow produced this token */
	authMode: "device-code" | "client-credentials";
	/** Signed-in identity (delegated tokens only) */
	user?: {
		id?: string;
		displayName?: string;
		upn?: string;
		tenantId?: string;
	};
	/** When the entry was last written */
	updatedAt: string;
}

interface TokenFile {
	version: 1;
	tokens: Record<string, CachedToken>;
}

// ---------------------------------------------------------------------------
// Paths & keys
// ---------------------------------------------------------------------------

export function getTokenPath(): string {
	return join(getAgentDir(), "pi-teams-tokens.json");
}

/** Cache key for an account+tenant pair. */
export function tokenKey(account: string, tenantId: string): string {
	return `${account.toLowerCase()}::${tenantId.toLowerCase()}`;
}

// ---------------------------------------------------------------------------
// I/O
// ---------------------------------------------------------------------------

function readFile(): TokenFile {
	const path = getTokenPath();
	if (!existsSync(path)) return { version: 1, tokens: {} };
	try {
		try {
			const mode = statSync(path).mode & 0o777;
			if (mode !== 0o600) chmodSync(path, 0o600);
		} catch {
			/* ignore */
		}
		const parsed = JSON.parse(readFileSync(path, "utf-8")) as TokenFile;
		if (!parsed || typeof parsed !== "object" || !parsed.tokens) return { version: 1, tokens: {} };
		return parsed;
	} catch {
		return { version: 1, tokens: {} };
	}
}

function writeFile(data: TokenFile): void {
	const path = getTokenPath();
	const dir = dirname(path);
	if (!existsSync(dir)) mkdirSync(dir, { recursive: true, mode: 0o700 });

	const tmpPath = `${path}.${process.pid}.tmp`;
	try {
		writeFileSync(tmpPath, `${JSON.stringify(data, null, 2)}\n`, { encoding: "utf-8", mode: 0o600 });
		chmodSync(tmpPath, 0o600);
		renameSync(tmpPath, path);
	} catch (err) {
		try {
			unlinkSync(tmpPath);
		} catch {
			/* ignore */
		}
		throw err;
	}
}

// ---------------------------------------------------------------------------
// Access
// ---------------------------------------------------------------------------

export function getToken(account: string, tenantId: string): CachedToken | undefined {
	return readFile().tokens[tokenKey(account, tenantId)];
}

export function saveToken(account: string, tenantId: string, token: CachedToken): void {
	const file = readFile();
	file.tokens[tokenKey(account, tenantId)] = { ...token, updatedAt: new Date().toISOString() };
	writeFile(file);
}

export function deleteToken(account: string, tenantId: string): boolean {
	const file = readFile();
	const key = tokenKey(account, tenantId);
	if (!file.tokens[key]) return false;
	delete file.tokens[key];
	writeFile(file);
	return true;
}

/** Drop every cached token — used by `teams_logout` with `all: true`. */
export function clearAllTokens(): number {
	const file = readFile();
	const count = Object.keys(file.tokens).length;
	writeFile({ version: 1, tokens: {} });
	return count;
}

export function listTokens(): Array<{ key: string; token: CachedToken }> {
	const file = readFile();
	return Object.entries(file.tokens).map(([key, token]) => ({ key, token }));
}

/**
 * Is the token still usable?
 * A 5-minute skew keeps a long Graph call from failing mid-flight.
 */
export function isFresh(token: CachedToken | undefined, skewMs = 5 * 60 * 1000): boolean {
	if (!token?.accessToken) return false;
	return token.expiresAt - skewMs > Date.now();
}
