/**
 * MSAL token cache, persisted to disk.
 *
 * MSAL keeps its cache in memory and hands us a serialized blob through the
 * `ICachePlugin` hooks; without persisting it, every pi session would start
 * with a fresh sign-in. One file per account+tenant under
 * `~/.pi/agent/pi-teams-tokens/`, mode 0600 — the blob contains refresh
 * tokens.
 *
 * Writes go through a private temp file and an atomic rename: a truncated
 * cache would silently sign the user out.
 */

import {
	existsSync,
	mkdirSync,
	readFileSync,
	writeFileSync,
	chmodSync,
	statSync,
	renameSync,
	unlinkSync,
	readdirSync,
} from "node:fs";
import { join } from "node:path";
import type { ICachePlugin, TokenCacheContext } from "@azure/msal-node";
import { getAgentDir } from "../config/index.ts";

// ---------------------------------------------------------------------------
// Paths
// ---------------------------------------------------------------------------

export function getCacheDir(): string {
	return join(getAgentDir(), "pi-teams-tokens");
}

/** Filesystem-safe cache key for an account+tenant pair. */
export function cacheKey(account: string, tenantId: string): string {
	const clean = (value: string) => value.toLowerCase().replace(/[^a-z0-9._-]/g, "_");
	return `${clean(account)}__${clean(tenantId)}`;
}

export function getCachePath(account: string, tenantId: string): string {
	return join(getCacheDir(), `${cacheKey(account, tenantId)}.json`);
}

// ---------------------------------------------------------------------------
// I/O
// ---------------------------------------------------------------------------

function readBlob(path: string): string | undefined {
	if (!existsSync(path)) return undefined;
	try {
		// Older files, or a cache copied between machines, may be too permissive
		// for something holding refresh tokens.
		try {
			const mode = statSync(path).mode & 0o777;
			if (mode !== 0o600) chmodSync(path, 0o600);
		} catch {
			/* ignore */
		}
		return readFileSync(path, "utf-8");
	} catch {
		return undefined;
	}
}

function writeBlob(path: string, blob: string): void {
	const dir = getCacheDir();
	if (!existsSync(dir)) mkdirSync(dir, { recursive: true, mode: 0o700 });

	const tmpPath = `${path}.${process.pid}.tmp`;
	try {
		writeFileSync(tmpPath, blob, { encoding: "utf-8", mode: 0o600 });
		chmodSync(tmpPath, 0o600);
		renameSync(tmpPath, path);
	} catch {
		try {
			unlinkSync(tmpPath);
		} catch {
			/* ignore */
		}
		// A cache that cannot be written is a degraded session, not a failed one:
		// the tokens still work until the process exits.
	}
}

// ---------------------------------------------------------------------------
// Plugin
// ---------------------------------------------------------------------------

/** Build the cache plugin for one account+tenant. */
export function createCachePlugin(account: string, tenantId: string): ICachePlugin {
	const path = getCachePath(account, tenantId);

	return {
		async beforeCacheAccess(context: TokenCacheContext): Promise<void> {
			const blob = readBlob(path);
			if (blob) context.tokenCache.deserialize(blob);
		},
		async afterCacheAccess(context: TokenCacheContext): Promise<void> {
			if (context.cacheHasChanged) writeBlob(path, context.tokenCache.serialize());
		},
	};
}

// ---------------------------------------------------------------------------
// Inspection
// ---------------------------------------------------------------------------

/** What the status surfaces need to know, readable without MSAL or a network call. */
export interface CacheSummary {
	/** A cache file exists and holds an account */
	present: boolean;
	/** Signed-in identity, when the cache holds one */
	username?: string;
	name?: string;
	/** Epoch milliseconds at which the newest access token expires */
	expiresAt?: number;
	/** Whether that access token is still valid */
	fresh: boolean;
}

/* eslint-disable @typescript-eslint/no-explicit-any */

/**
 * Read the identity and expiry out of a cache file.
 *
 * The blob is MSAL's own schema, so this reads it defensively: a shape change
 * upstream must degrade to "not signed in" in the status line, never throw in
 * the middle of rendering the TUI.
 */
export function readCacheSummary(account: string, tenantId: string): CacheSummary {
	const blob = readBlob(getCachePath(account, tenantId));
	if (!blob) return { present: false, fresh: false };

	try {
		const parsed = JSON.parse(blob) as Record<string, any>;

		const accounts = Object.values(parsed.Account ?? {}) as any[];
		const first = accounts[0];

		const tokens = Object.values(parsed.AccessToken ?? {}) as any[];
		// expires_on is seconds-since-epoch, as a string, in MSAL's schema.
		const expiries = tokens
			.map((token) => Number(token?.expires_on))
			.filter((value) => Number.isFinite(value) && value > 0);
		const expiresAt = expiries.length > 0 ? Math.max(...expiries) * 1000 : undefined;

		return {
			present: accounts.length > 0 || tokens.length > 0,
			username: first?.username,
			name: first?.name,
			expiresAt,
			fresh: expiresAt !== undefined && expiresAt - 5 * 60 * 1000 > Date.now(),
		};
	} catch {
		return { present: false, fresh: false };
	}
}

/** Delete one cached session. Returns true when a file was removed. */
export function removeCache(account: string, tenantId: string): boolean {
	const path = getCachePath(account, tenantId);
	if (!existsSync(path)) return false;
	try {
		unlinkSync(path);
		return true;
	} catch {
		return false;
	}
}

/** Delete every cached session. Returns how many were removed. */
export function removeAllCaches(): number {
	const dir = getCacheDir();
	if (!existsSync(dir)) return 0;

	let removed = 0;
	for (const entry of readdirSync(dir)) {
		if (!entry.endsWith(".json")) continue;
		try {
			unlinkSync(join(dir, entry));
			removed += 1;
		} catch {
			/* skip what we cannot remove */
		}
	}
	return removed;
}

/** Every cached session on disk, for `teams_status` and the doctor. */
export function listCachedSessions(): Array<{ key: string; summary: CacheSummary }> {
	const dir = getCacheDir();
	if (!existsSync(dir)) return [];

	const sessions: Array<{ key: string; summary: CacheSummary }> = [];
	for (const entry of readdirSync(dir)) {
		if (!entry.endsWith(".json")) continue;
		const key = entry.replace(/\.json$/, "");
		const blob = readBlob(join(dir, entry));
		if (!blob) continue;
		try {
			const parsed = JSON.parse(blob) as Record<string, any>;
			const first = (Object.values(parsed.Account ?? {}) as any[])[0];
			sessions.push({
				key,
				summary: { present: true, username: first?.username, name: first?.name, fresh: false },
			});
		} catch {
			/* ignore an unreadable cache file */
		}
	}
	return sessions;
}
