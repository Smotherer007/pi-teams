/**
 * Listen mode — the durable cursor.
 *
 * The watcher's memory of what it has already answered used to die with the
 * session, which made "switch listen mode on" mean "from now on": whatever
 * arrived while pi was not running was silently dropped. The Telegram bridge
 * solves the same problem with the Bot API's `offset` cursor — Telegram
 * re-delivers every update that has not been acknowledged yet, and the bridge
 * persists that offset, so a restart answers exactly the messages that came in
 * while it was down, and never the same one twice.
 *
 * Graph has no such cursor for chats, so this module keeps one: `chatId → the
 * activity marker of the newest message pi has already looked at`. A chat has
 * moved (and is therefore worth answering) when its current marker differs from
 * the stored one. Written after every tick, so a crash costs nothing but the
 * tick it died in.
 *
 * One file per account **and** tenant, next to the token cache: two accounts in
 * one process must not share a cursor, and neither must a guest tenant beneath
 * the same account.
 */

import { chmodSync, existsSync, mkdirSync, readFileSync, renameSync, unlinkSync, writeFileSync } from "node:fs";
import { join } from "node:path";
import { getAgentDir } from "../config/index.ts";

/** Shape of the file on disk. Kept small — it is read on every session start. */
interface WatchCursorFile {
	version: 1;
	/** ISO timestamp of the last write, for humans reading the file */
	updatedAt: string;
	/** chatId → activity marker of the newest message already examined */
	chats: Record<string, string>;
}

export function getWatchCursorDir(): string {
	return join(getAgentDir(), "pi-teams-watch");
}

/**
 * Path of one account+tenant cursor.
 *
 * Sanitised the same way the token cache is, so a hand-written account name
 * cannot escape the directory.
 */
export function getWatchCursorPath(account: string, tenantId: string): string {
	const clean = (value: string) => value.toLowerCase().replace(/[^a-z0-9._-]/g, "_");
	return join(getWatchCursorDir(), `${clean(account)}__${clean(tenantId)}.json`);
}

/**
 * The cursor for an account+tenant, or `undefined` when there is none yet.
 *
 * `undefined` is meaningful, not just "missing": it means listen mode has never
 * run for this account, which is the one case where history is deliberately
 * skipped instead of answered. A corrupt file is reported the same way — the
 * alternative is a watcher that refuses to start — and the next tick writes a
 * fresh one.
 */
export function readWatchCursor(account: string, tenantId: string): Map<string, string> | undefined {
	const path = getWatchCursorPath(account, tenantId);
	if (!existsSync(path)) return undefined;

	try {
		const parsed = JSON.parse(readFileSync(path, "utf-8")) as Partial<WatchCursorFile>;
		const chats = parsed?.chats;
		if (parsed?.version !== 1 || !chats || typeof chats !== "object") return undefined;

		return new Map(
			Object.entries(chats).filter((entry): entry is [string, string] => typeof entry[1] === "string"),
		);
	} catch {
		return undefined;
	}
}

/**
 * Persist the cursor. Never throws.
 *
 * Written to a temporary file and renamed, so a crash mid-write leaves the
 * previous cursor intact rather than a half-file that would be read as
 * "no cursor" and replay a whole history.
 */
export function writeWatchCursor(
	account: string,
	tenantId: string,
	chats: ReadonlyMap<string, string>,
): void {
	const path = getWatchCursorPath(account, tenantId);
	const file: WatchCursorFile = {
		version: 1,
		updatedAt: new Date().toISOString(),
		chats: Object.fromEntries(chats),
	};

	try {
		const dir = getWatchCursorDir();
		if (!existsSync(dir)) mkdirSync(dir, { recursive: true, mode: 0o700 });

		const tmpPath = `${path}.${process.pid}.tmp`;
		writeFileSync(tmpPath, `${JSON.stringify(file)}\n`, { encoding: "utf-8", mode: 0o600 });
		chmodSync(tmpPath, 0o600);
		renameSync(tmpPath, path);
	} catch {
		/* A cursor that cannot be written costs a repeated examination, not an operation. */
		try {
			unlinkSync(`${path}.${process.pid}.tmp`);
		} catch {
			/* ignore */
		}
	}
}
