/**
 * Append-only audit log of everything pi wrote in the user's name.
 *
 * One JSON object per line in `~/.pi/agent/pi-teams-audit.jsonl`, mode 0600.
 * The point is answerability: when a colleague asks "did you really send that
 * at 23:40?", the user can look it up instead of guessing. Logging never
 * blocks or fails an operation — a broken log is a smaller problem than a
 * refused message.
 */

import { appendFileSync, existsSync, mkdirSync, chmodSync } from "node:fs";
import { dirname, join } from "node:path";
import { getAgentDir } from "../config/index.ts";

export interface AuditEntry {
	/** ISO timestamp */
	at: string;
	/** Tool that performed the write */
	tool: string;
	account: string;
	tenant: string;
	/** Who pi acted as, when known */
	actor?: string;
	/** Human-readable target, e.g. "Engineering/General" */
	target: string;
	/** One-line description of what happened */
	summary: string;
	/** Present when the operation failed */
	error?: string;
}

export function getAuditPath(): string {
	return join(getAgentDir(), "pi-teams-audit.jsonl");
}

/** Append an entry. Never throws. */
export function recordAudit(entry: AuditEntry): void {
	try {
		const path = getAuditPath();
		const dir = dirname(path);
		if (!existsSync(dir)) mkdirSync(dir, { recursive: true, mode: 0o700 });
		const isNew = !existsSync(path);
		appendFileSync(path, `${JSON.stringify(entry)}\n`, { encoding: "utf-8", mode: 0o600 });
		if (isNew) chmodSync(path, 0o600);
	} catch {
		/* auditing must never break the operation it describes */
	}
}

/** Convenience wrapper used by the tools. */
export function auditWrite(
	enabled: boolean,
	fields: Omit<AuditEntry, "at">,
): void {
	if (!enabled) return;
	recordAudit({ at: new Date().toISOString(), ...fields });
}
