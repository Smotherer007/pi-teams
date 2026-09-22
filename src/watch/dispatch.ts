/**
 * Listen mode — one pi process per chat.
 *
 * With `watch.dispatch.mode: "process"` a wake does not become a turn in the
 * session that runs the watcher. It goes to a worker: a `pi --mode rpc` process
 * that belongs to that one chat and keeps its own session from message to
 * message. Different chats are worked on at the same time; the watching
 * session only routes and stays small.
 *
 * The rules, in the order they are applied:
 *
 * - A chat whose worker is busy gets the new wake **steered** into the running
 *   turn. Nobody waits for a long task to finish before "stop, take repo X
 *   instead" is read.
 * - A chat without a busy worker starts one, if fewer than `maxConcurrent`
 *   chats are being worked on. Otherwise it waits in line; a newer wake for a
 *   chat already in line replaces the older one (its prompt carries every open
 *   message anyway).
 * - A stop word aborts the running turn; a reset word ends the worker and the
 *   chat starts over with a fresh session. Both prompts still reach the model,
 *   so the person gets a confirmation.
 * - An idle worker exits after `idleMinutes`. Its session stays on disk, and
 *   the next message in that chat continues it — unless the chat has been
 *   quiet for longer than `freshAfterHours`.
 *
 * Everything a worker is asked and answers is written to a journal
 * (`pi-teams-dispatch.jsonl`), so "what did pi tell whom" can be answered
 * across chats without mixing their contexts.
 *
 * The process handling is injectable (`spawn`), so the scheduling rules are
 * tested without starting pi.
 */

import { spawn as nodeSpawn } from "node:child_process";
import { createHash } from "node:crypto";
import { appendFileSync, chmodSync, existsSync, mkdirSync, readdirSync, statSync } from "node:fs";
import { join } from "node:path";
import type { ResolvedDispatchConfig } from "../config/index.ts";
import type { ControlAction } from "./control.ts";
import { workerEnv, type WorkerIdentity } from "./worker.ts";

// ---------------------------------------------------------------------------
// Types
// ---------------------------------------------------------------------------

/** One wake, ready to be handed to a worker. */
export interface DispatchJob {
	chatId: string;
	label: string;
	/** "oneOnOne" | "group" | "meeting" */
	chatType?: string;
	/** Display name of the sender of the newest message */
	from?: string;
	/** Address of the other person in a one-to-one chat */
	peer?: string;
	/** The wake prompt, exactly as a session-mode wake would receive it */
	prompt: string;
	/** Short form of the newest message, for the journal */
	request?: string;
	control?: ControlAction;
}

export type DispatchOutcome = "started" | "steered" | "queued" | "stopped" | "reset";

export type DispatchEvent =
	| "start"
	| "steer"
	| "queued"
	| "stop"
	| "reset"
	| "done"
	| "error"
	| "idle-exit";

export interface DispatchJournalEntry {
	at: string;
	chatId: string;
	chat: string;
	event: DispatchEvent;
	from?: string;
	request?: string;
	/** The worker's last answer text, shortened */
	result?: string;
	durationMs?: number;
	detail?: string;
}

/** The part of a child process the dispatcher uses. */
export interface WorkerProcess {
	pid?: number;
	stdin: { write(chunk: string): unknown; end(): unknown; on(event: "error", listener: (err: Error) => void): unknown } | null;
	stdout: { on(event: "data", listener: (chunk: Buffer | string) => void): unknown } | null;
	stderr: { on(event: "data", listener: (chunk: Buffer | string) => void): unknown } | null;
	on(event: "exit", listener: (code: number | null, signal: NodeJS.Signals | null) => void): unknown;
	on(event: "error", listener: (err: Error) => void): unknown;
	kill(signal?: NodeJS.Signals): unknown;
}

export type SpawnWorker = (
	command: string,
	args: string[],
	options: { cwd: string; env: NodeJS.ProcessEnv },
) => WorkerProcess;

export interface DispatcherOptions {
	config: ResolvedDispatchConfig;
	/** pi's agent directory: chat sessions and the journal live below it */
	agentDir: string;
	/** Working directory of the workers */
	cwd: string;
	spawn?: SpawnWorker;
	now?: () => number;
	/** Environment the workers inherit (default: process.env) */
	env?: NodeJS.ProcessEnv;
	/** Journal writer (default: append to pi-teams-dispatch.jsonl) */
	journal?: (entry: DispatchJournalEntry) => void;
	/** A worker failed in a way the user should hear about */
	onError?: (message: string) => void;
	/** Interval of the idle sweep in ms (default: 60 000, 0 = no timer) */
	sweepIntervalMs?: number;
}

export interface DispatchStatus {
	mode: "process";
	busy: number;
	alive: number;
	queued: number;
	maxConcurrent: number;
	chats: Array<{ chatId: string; label: string; busy: boolean; sinceMs: number }>;
	waiting: Array<{ chatId: string; label: string }>;
}

interface Worker {
	identity: WorkerIdentity;
	proc: WorkerProcess;
	busy: boolean;
	job?: DispatchJob;
	startedAt: number;
	lastActivity: number;
	buffer: string;
	lastText?: string;
	stderrTail: string;
	/** Set when the dispatcher itself ended the process */
	retiring: boolean;
}

// ---------------------------------------------------------------------------
// Paths and journal
// ---------------------------------------------------------------------------

export function getDispatchJournalPath(agentDir: string): string {
	return join(agentDir, "pi-teams-dispatch.jsonl");
}

/** The session directory of one chat. Hashed: chat IDs contain ':' and '@'. */
export function chatSessionDir(agentDir: string, chatId: string): string {
	const key = createHash("sha256").update(chatId).digest("hex").slice(0, 16);
	return join(agentDir, "pi-teams-chats", key);
}

/** Append a journal entry. Never throws. */
export function appendDispatchJournal(agentDir: string, entry: DispatchJournalEntry): void {
	try {
		if (!existsSync(agentDir)) mkdirSync(agentDir, { recursive: true, mode: 0o700 });
		const path = getDispatchJournalPath(agentDir);
		const isNew = !existsSync(path);
		appendFileSync(path, `${JSON.stringify(entry)}\n`, { encoding: "utf-8", mode: 0o600 });
		if (isNew) chmodSync(path, 0o600);
	} catch {
		/* the journal must never stop an answer */
	}
}

/** One line, at most `max` characters. */
export function excerpt(text: string | undefined, max = 300): string | undefined {
	if (!text) return undefined;
	const flat = text.replace(/\s+/g, " ").trim();
	if (!flat) return undefined;
	return flat.length > max ? `${flat.slice(0, max - 1)}…` : flat;
}

/**
 * Whether the chat's last session is recent enough to continue.
 *
 * `freshAfterHours: 0` means always continue. No session at all means there is
 * nothing to continue.
 */
export function shouldContinue(dir: string, freshAfterHours: number, now: number): boolean {
	let newest = 0;
	try {
		for (const name of readdirSync(dir)) {
			if (!name.endsWith(".jsonl")) continue;
			const mtime = statSync(join(dir, name)).mtimeMs;
			if (mtime > newest) newest = mtime;
		}
	} catch {
		return false;
	}
	if (newest === 0) return false;
	if (freshAfterHours === 0) return true;
	return now - newest <= freshAfterHours * 3_600_000;
}

/** The text of an assistant message, or undefined. */
function assistantText(message: any): string | undefined {
	if (!message || message.role !== "assistant") return undefined;
	const content = message.content;
	if (typeof content === "string") return content;
	if (!Array.isArray(content)) return undefined;
	const text = content
		.filter((part: any) => part?.type === "text" && typeof part.text === "string")
		.map((part: any) => part.text)
		.join("\n")
		.trim();
	return text || undefined;
}

const DIALOG_METHODS = new Set(["select", "confirm", "input", "editor"]);

// ---------------------------------------------------------------------------
// Dispatcher
// ---------------------------------------------------------------------------

export class ChatDispatcher {
	private config: ResolvedDispatchConfig;
	private readonly workers = new Map<string, Worker>();
	private queue: DispatchJob[] = [];
	private readonly spawnWorker: SpawnWorker;
	private readonly now: () => number;
	private readonly journalWriter: (entry: DispatchJournalEntry) => void;
	private sweepTimer?: ReturnType<typeof setInterval>;
	private stopped = false;
	private readonly options: DispatcherOptions;

	constructor(options: DispatcherOptions) {
		this.options = options;
		this.config = options.config;
		this.spawnWorker =
			options.spawn ??
			((command, args, opts) =>
				nodeSpawn(command, args, { cwd: opts.cwd, env: opts.env, stdio: ["pipe", "pipe", "pipe"] }) as WorkerProcess);
		this.now = options.now ?? Date.now;
		this.journalWriter = options.journal ?? ((entry) => appendDispatchJournal(options.agentDir, entry));
		const every = options.sweepIntervalMs ?? 60_000;
		if (every > 0) {
			this.sweepTimer = setInterval(() => this.sweep(), every);
			this.sweepTimer.unref?.();
		}
	}

	/** Apply changed settings. Running workers keep going; limits apply from now on. */
	configure(config: ResolvedDispatchConfig): void {
		this.config = config;
		this.drain();
	}

	/** Hand a wake to its chat's worker. */
	deliver(job: DispatchJob): DispatchOutcome {
		if (this.stopped) return "queued";
		const worker = this.workers.get(job.chatId);

		if (job.control === "reset") {
			this.dropQueued(job.chatId);
			if (worker) this.retire(worker, "reset");
			this.record(job, "reset");
			this.start(job, { fresh: true, force: true });
			return "reset";
		}

		if (job.control === "stop") {
			this.dropQueued(job.chatId);
			if (worker?.busy) {
				this.send(worker, { type: "clear_queue" });
				this.send(worker, { type: "abort" });
				this.record(job, "stop");
				// Queued behind the abort: the model confirms in the chat once it is idle.
				this.send(worker, { type: "prompt", message: job.prompt, streamingBehavior: "followUp" });
				worker.job = job;
				return "stopped";
			}
			this.record(job, "stop", { detail: "nothing was running" });
		}

		if (worker?.busy) {
			this.send(worker, { type: "prompt", message: job.prompt, streamingBehavior: "steer" });
			worker.lastActivity = this.now();
			this.record(job, "steer");
			return "steered";
		}

		if (this.busyCount() >= this.config.maxConcurrent) {
			this.dropQueued(job.chatId);
			this.queue.push(job);
			this.record(job, "queued", { detail: `${this.busyCount()} chat(s) in progress` });
			return "queued";
		}

		this.start(job, { fresh: false });
		return "started";
	}

	/** Snapshot for status commands. */
	status(): DispatchStatus {
		const now = this.now();
		return {
			mode: "process",
			busy: this.busyCount(),
			alive: this.workers.size,
			queued: this.queue.length,
			maxConcurrent: this.config.maxConcurrent,
			chats: [...this.workers.values()].map((w) => ({
				chatId: w.identity.chatId,
				label: w.identity.label,
				busy: w.busy,
				sinceMs: now - (w.busy ? w.startedAt : w.lastActivity),
			})),
			waiting: this.queue.map((job) => ({ chatId: job.chatId, label: job.label })),
		};
	}

	/** End idle workers that have been quiet for longer than `idleMinutes`. */
	sweep(): void {
		const now = this.now();
		const limit = this.config.idleMinutes * 60_000;
		for (const worker of [...this.workers.values()]) {
			if (!worker.busy && now - worker.lastActivity > limit) {
				this.journalWriter({
					at: new Date(now).toISOString(),
					chatId: worker.identity.chatId,
					chat: worker.identity.label,
					event: "idle-exit",
				});
				this.retire(worker, "idle");
			}
		}
	}

	/** End every worker. Their sessions stay on disk. */
	stopAll(): void {
		this.stopped = true;
		if (this.sweepTimer) clearInterval(this.sweepTimer);
		this.queue = [];
		for (const worker of [...this.workers.values()]) this.retire(worker, "shutdown");
	}

	// -----------------------------------------------------------------------

	private busyCount(): number {
		let busy = 0;
		for (const worker of this.workers.values()) if (worker.busy) busy += 1;
		return busy;
	}

	private dropQueued(chatId: string): void {
		this.queue = this.queue.filter((job) => job.chatId !== chatId);
	}

	private record(job: DispatchJob, event: DispatchEvent, extra: Partial<DispatchJournalEntry> = {}): void {
		this.journalWriter({
			at: new Date(this.now()).toISOString(),
			chatId: job.chatId,
			chat: job.label,
			event,
			from: job.from,
			request: job.request,
			...extra,
		});
	}

	private send(worker: Worker, command: Record<string, unknown>): void {
		try {
			worker.proc.stdin?.write(`${JSON.stringify(command)}\n`);
		} catch {
			/* a dead pipe shows up as an exit event */
		}
	}

	private start(job: DispatchJob, { fresh, force = false }: { fresh: boolean; force?: boolean }): void {
		let worker = this.workers.get(job.chatId);
		if (!worker) worker = this.launch(job, fresh);
		if (!worker) return;

		const now = this.now();
		worker.busy = true;
		worker.job = job;
		worker.startedAt = now;
		worker.lastActivity = now;
		worker.lastText = undefined;
		// followUp: harmless when idle, and correct if the worker is still
		// finishing something the dispatcher did not see yet.
		this.send(worker, { type: "prompt", message: job.prompt, streamingBehavior: "followUp" });
		this.record(job, "start", force ? { detail: "fresh session" } : {});
	}

	private launch(job: DispatchJob, fresh: boolean): Worker | undefined {
		this.makeRoom();

		const identity: WorkerIdentity = {
			chatId: job.chatId,
			label: job.label,
			chatType: job.chatType,
			peer: job.chatType === "oneOnOne" ? job.peer : undefined,
		};
		const dir = chatSessionDir(this.options.agentDir, job.chatId);
		try {
			mkdirSync(dir, { recursive: true, mode: 0o700 });
		} catch {
			/* spawn reports the real problem */
		}

		const resume = !fresh && shouldContinue(dir, this.config.freshAfterHours, this.now());
		const args = [
			...this.config.args,
			"--mode",
			"rpc",
			"--session-dir",
			dir,
			...(resume ? ["--continue"] : ["--name", `teams: ${job.label}`]),
		];

		let proc: WorkerProcess;
		try {
			proc = this.spawnWorker(this.config.command, args, {
				cwd: this.options.cwd,
				env: { ...(this.options.env ?? process.env), ...workerEnv(identity) },
			});
		} catch (err) {
			const message = err instanceof Error ? err.message : String(err);
			this.record(job, "error", { detail: `could not start ${this.config.command}: ${message}` });
			this.options.onError?.(`Teams dispatch: could not start a worker for "${job.label}": ${message}`);
			return undefined;
		}

		const now = this.now();
		const worker: Worker = {
			identity,
			proc,
			busy: false,
			startedAt: now,
			lastActivity: now,
			buffer: "",
			stderrTail: "",
			retiring: false,
		};
		this.workers.set(job.chatId, worker);

		proc.stdin?.on("error", () => undefined);
		proc.stdout?.on("data", (chunk) => this.onData(worker, chunk));
		proc.stderr?.on("data", (chunk) => {
			worker.stderrTail = (worker.stderrTail + chunk.toString()).slice(-2000);
		});
		proc.on("error", (err) => this.onExit(worker, `error: ${err.message}`));
		proc.on("exit", (code, signal) => this.onExit(worker, signal ? `signal ${signal}` : `exit code ${code}`));
		return worker;
	}

	/** Keep the number of live workers under `maxWorkers` by ending the longest-idle one. */
	private makeRoom(): void {
		while (this.workers.size >= this.config.maxWorkers) {
			const idle = [...this.workers.values()].filter((w) => !w.busy).sort((a, b) => a.lastActivity - b.lastActivity);
			if (idle.length === 0) return;
			this.retire(idle[0]!, "room");
		}
	}

	/** End a worker on purpose. Closing stdin lets pi shut down cleanly. */
	private retire(worker: Worker, _why: string): void {
		worker.retiring = true;
		if (this.workers.get(worker.identity.chatId) === worker) this.workers.delete(worker.identity.chatId);
		try {
			worker.proc.stdin?.end();
		} catch {
			/* ignore */
		}
		const timer = setTimeout(() => {
			try {
				worker.proc.kill("SIGTERM");
			} catch {
				/* already gone */
			}
		}, 10_000);
		timer.unref?.();
	}

	private onData(worker: Worker, chunk: Buffer | string): void {
		worker.buffer += chunk.toString();
		let index = worker.buffer.indexOf("\n");
		while (index >= 0) {
			const line = worker.buffer.slice(0, index).replace(/\r$/, "");
			worker.buffer = worker.buffer.slice(index + 1);
			if (line.trim()) this.onLine(worker, line);
			index = worker.buffer.indexOf("\n");
		}
	}

	private onLine(worker: Worker, line: string): void {
		let event: any;
		try {
			event = JSON.parse(line);
		} catch {
			return;
		}
		if (!event || typeof event !== "object") return;

		switch (event.type) {
			case "agent_start":
				// A steer that arrived just after the worker settled starts a new run.
				if (!worker.busy) {
					worker.busy = true;
					worker.startedAt = this.now();
				}
				return;
			case "message_end": {
				const text = assistantText(event.message);
				if (text) worker.lastText = text;
				return;
			}
			case "agent_settled":
				this.finish(worker);
				return;
			case "extension_ui_request":
				// Nobody sits in front of a worker. A dialog would wait forever.
				if (DIALOG_METHODS.has(event.method)) {
					this.send(worker, { type: "extension_ui_response", id: event.id, cancelled: true });
				}
				return;
			case "response":
				if (event.success === false) {
					const detail = `${event.command ?? "command"} failed: ${event.error ?? "unknown error"}`;
					if (worker.job) this.record(worker.job, "error", { detail });
					this.options.onError?.(`Teams dispatch (${worker.identity.label}): ${detail}`);
					if (event.command === "prompt" && worker.busy) this.finish(worker, true);
				}
				return;
			default:
				return;
		}
	}

	private finish(worker: Worker, failed = false): void {
		if (!worker.busy) return;
		const now = this.now();
		worker.busy = false;
		worker.lastActivity = now;
		if (worker.job && !failed) {
			this.record(worker.job, "done", { durationMs: now - worker.startedAt, result: excerpt(worker.lastText) });
		}
		this.drain();
	}

	private onExit(worker: Worker, how: string): void {
		const current = this.workers.get(worker.identity.chatId) === worker;
		if (current) this.workers.delete(worker.identity.chatId);
		if (!worker.retiring && worker.busy && worker.job) {
			const detail = `worker ended (${how})${worker.stderrTail ? `: ${excerpt(worker.stderrTail, 200)}` : ""}`;
			this.record(worker.job, "error", { detail });
			this.options.onError?.(`Teams dispatch (${worker.identity.label}): ${detail}`);
		}
		worker.busy = false;
		worker.retiring = true;
		if (!this.stopped) this.drain();
	}

	/** Start waiting chats while there is room. */
	private drain(): void {
		while (this.queue.length > 0 && this.busyCount() < this.config.maxConcurrent) {
			const job = this.queue.shift()!;
			const worker = this.workers.get(job.chatId);
			if (worker?.busy) {
				this.send(worker, { type: "prompt", message: job.prompt, streamingBehavior: "steer" });
				this.record(job, "steer");
				continue;
			}
			this.start(job, { fresh: false });
		}
	}
}
