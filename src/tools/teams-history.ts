/**
 * teams_history — what pi wrote in Teams, and what the chat workers were asked.
 *
 * In dispatch mode every chat has a session of its own, so no single context
 * knows what pi said elsewhere. This tool answers "what did you tell Anna
 * today?" from the two logs that do know: the audit log (every message pi
 * sent, see ../safety/audit.ts) and the dispatch journal (every request a chat
 * worker picked up, and its answer, see ../watch/dispatch.ts).
 *
 * Reading across chats is exactly what the per-chat sessions exist to prevent
 * by default, so inside a chat worker the tool answers only in a one-to-one
 * chat with somebody listed in `watch.dispatch.historyReaders`. The watching
 * session — the user at the keyboard — is never restricted.
 */

import { existsSync, openSync, readSync, closeSync, statSync } from "node:fs";
import { Type } from "typebox";
import { getAgentDir } from "../config/index.ts";
import { matchesPattern } from "../config/scope.ts";
import { getAuditPath } from "../safety/audit.ts";
import { getDispatchJournalPath } from "../watch/dispatch.ts";
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
		"the people listed in watch.dispatch.historyReaders may look across chats, from their one-to-one chat with pi."
	);
}

interface HistoryParams {
	kind?: "sent" | "dispatch";
	chat?: string;
	sinceHours?: number;
	limit?: number;
}

/** Filter and render log entries. Exported for tests. */
export function renderHistory(entries: Record<string, any>[], params: HistoryParams, now = Date.now()): string {
	const kind = params.kind ?? "sent";
	const limit = Math.min(Math.max(params.limit ?? 20, 1), 200);
	const since = now - Math.min(Math.max(params.sinceHours ?? 24, 1), 24 * 90) * 3_600_000;
	const needle = params.chat?.trim().toLowerCase();

	const hits = entries.filter((entry) => {
		const at = Date.parse(entry.at ?? "");
		if (!Number.isFinite(at) || at < since) return false;
		if (!needle) return true;
		const haystack =
			kind === "sent"
				? `${entry.target ?? ""} ${entry.summary ?? ""}`
				: `${entry.chat ?? ""} ${entry.chatId ?? ""} ${entry.from ?? ""}`;
		return haystack.toLowerCase().includes(needle);
	});

	const shown = hits.slice(-limit);
	if (shown.length === 0) return kind === "sent" ? "No messages sent in that window." : "No dispatch entries in that window.";

	const lines = shown.map((entry) => {
		const at = new Date(entry.at).toLocaleString();
		if (kind === "sent") {
			const failed = entry.error ? ` (failed: ${entry.error})` : "";
			return `- ${at} · ${entry.tool} → ${entry.target}: ${entry.summary}${failed}`;
		}
		const parts = [`- ${at} · ${entry.chat} · ${entry.event}`];
		if (entry.from) parts.push(`from ${entry.from}`);
		if (entry.request) parts.push(`asked: ${entry.request}`);
		if (entry.result) parts.push(`answered: ${entry.result}`);
		if (entry.durationMs) parts.push(`${Math.round(entry.durationMs / 1000)} s`);
		if (entry.detail) parts.push(entry.detail);
		return parts.join(" · ");
	});
	const more = hits.length > shown.length ? `\n\n(${hits.length - shown.length} older entries not shown)` : "";
	return lines.join("\n") + more;
}

export const teamsHistoryTool = {
	name: "teams_history",
	description:
		"Look up what pi did in Teams across chats: kind 'sent' lists messages pi sent (from the audit log), kind " +
		"'dispatch' lists requests the per-chat workers picked up and what they answered. Use it when someone asks " +
		"what pi told another person or chat. Read-only.",
	parameters: Type.Object({
		kind: Type.Optional(
			Type.Union([Type.Literal("sent"), Type.Literal("dispatch")], {
				description: "'sent' (default): messages pi sent. 'dispatch': requests and answers of the chat workers.",
			}),
		),
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
			const readers = ctx.connection?.watch.dispatch.historyReaders ?? [];
			const refusal = historyRefusal(workerIdentity(), readers);
			if (refusal) return errorResult(refusal);

			const path = (params.kind ?? "sent") === "sent" ? getAuditPath() : getDispatchJournalPath(getAgentDir());
			const text = renderHistory(readJsonlTail(path), params);
			return textResult(text, { path });
		});
	},
};
