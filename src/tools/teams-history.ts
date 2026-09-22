/**
 * teams_history — what pi wrote in Teams, across chats.
 *
 * When every chat is answered by a pi process of its own (a lane, e.g. with
 * pi-lanes), no single context knows what pi said elsewhere. The audit log
 * does (every message pi sent, see ../safety/audit.ts), and this tool reads it
 * to answer "what did you tell Anna today?".
 *
 * Reading across chats is exactly what separate processes prevent by default,
 * so in a process that answers one chat the tool works only in a one-to-one
 * chat with somebody listed in `watch.historyReaders`. The session a person
 * types in is never restricted.
 */

import { existsSync, openSync, readSync, closeSync, statSync } from "node:fs";
import { Type } from "typebox";
import { matchesPattern } from "../config/scope.ts";
import { getAuditPath } from "../safety/audit.ts";
import { workerIdentity, type WorkerIdentity } from "../watch/worker.ts";
import { errorResult, run, textResult, type ToolContext, type ToolResult } from "./shared.ts";

/** How much of the end of a log is read. Old entries are the least asked for. */
const TAIL_BYTES = 2 * 1024 * 1024;

/** The last `TAIL_BYTES` of a JSONL file, parsed. Broken lines are skipped. */
export function readJsonlTail(path: string, maxBytes = TAIL_BYTES): Record<string, any>[] {
	if (!existsSync(path)) return [];
	const size = statSync(path).size;
	const start = Math.max(0, size - maxBytes);
	const buffer = Buffer.alloc(size - start);
	const fd = openSync(path, "r");
	try {
		readSync(fd, buffer, 0, buffer.length, start);
	} finally {
		closeSync(fd);
	}
	const lines = buffer.toString("utf-8").split("\n");
	if (start > 0) lines.shift(); // first line is cut
	const entries: Record<string, any>[] = [];
	for (const line of lines) {
		if (!line.trim()) continue;
		try {
			entries.push(JSON.parse(line));
		} catch {
			/* skip */
		}
	}
	return entries;
}

/**
 * Why this process may not read across chats, or undefined when it may.
 * Exported for tests.
 */
export function historyRefusal(identity: WorkerIdentity | undefined, readers: string[]): string | undefined {
	if (!identity) return undefined;
	const allowed =
		identity.chatType === "oneOnOne" &&
		!!identity.peer &&
		readers.some((pattern) => matchesPattern(pattern, identity.peer!));
	if (allowed) return undefined;
	return (
		"What pi wrote in other chats is not available in this chat. Each chat keeps its own context, and only " +
		"the people listed in watch.historyReaders may look across chats, from their one-to-one chat with pi."
	);
}

interface HistoryParams {
	chat?: string;
	sinceHours?: number;
	limit?: number;
}

/** Filter and render log entries. Exported for tests. */
export function renderHistory(entries: Record<string, any>[], params: HistoryParams, now = Date.now()): string {
	const limit = Math.min(Math.max(params.limit ?? 20, 1), 200);
	const since = now - Math.min(Math.max(params.sinceHours ?? 24, 1), 24 * 90) * 3_600_000;
	const needle = params.chat?.trim().toLowerCase();

	const hits = entries.filter((entry) => {
		const at = Date.parse(entry.at ?? "");
		if (!Number.isFinite(at) || at < since) return false;
		if (!needle) return true;
		return `${entry.target ?? ""} ${entry.summary ?? ""}`.toLowerCase().includes(needle);
	});

	const shown = hits.slice(-limit);
	if (shown.length === 0) return "No messages sent in that window.";

	const lines = shown.map((entry) => {
		const at = new Date(entry.at).toLocaleString();
		const failed = entry.error ? ` (failed: ${entry.error})` : "";
		return `- ${at} · ${entry.tool} → ${entry.target}: ${entry.summary}${failed}`;
	});
	const more = hits.length > shown.length ? `\n\n(${hits.length - shown.length} older entries not shown)` : "";
	return lines.join("\n") + more;
}

export const teamsHistoryTool = {
	name: "teams_history",
	description:
		"Look up what pi sent in Teams across chats (from the audit log). Use it when someone asks what pi told " +
		"another person or chat. Read-only.",
	parameters: Type.Object({
		chat: Type.Optional(Type.String({ description: "Only entries whose chat or person contains this text" })),
		sinceHours: Type.Optional(Type.Number({ description: "How far back, in hours (default 24, max 2160)" })),
		limit: Type.Optional(Type.Number({ description: "Most recent entries to show (default 20, max 200)" })),
	}),
	promptSnippet: "Look up what pi wrote in other Teams chats",

	async execute(
		_toolCallId: string,
		params: HistoryParams,
		_signal: AbortSignal | undefined,
		_onUpdate: undefined,
		ctx: ToolContext,
	): Promise<ToolResult> {
		return run(async () => {
			const readers = ctx.connection?.watch.historyReaders ?? [];
			const refusal = historyRefusal(workerIdentity(), readers);
			if (refusal) return errorResult(refusal);

			const path = getAuditPath();
			const text = renderHistory(readJsonlTail(path), params);
			return textResult(text, { path });
		});
	},
};
