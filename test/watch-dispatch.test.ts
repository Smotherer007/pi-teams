/**
 * Dispatch mode: one pi process per chat.
 *
 * The scheduling rules are the feature — parallel across chats, in order
 * within a chat, a steer for a busy worker, a line for a full house — so they
 * are tested against fake processes that record what they were told.
 */

import { test, describe } from "node:test";
import assert from "node:assert/strict";
import { EventEmitter } from "node:events";
import { mkdtempSync, mkdirSync, writeFileSync, utimesSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import {
	ChatDispatcher,
	chatSessionDir,
	excerpt,
	shouldContinue,
	type DispatchJob,
	type DispatchJournalEntry,
	type WorkerProcess,
} from "../src/watch/dispatch.ts";
import { resolveDispatchConfig, type DispatchConfig } from "../src/config/index.ts";

class FakeWorker extends EventEmitter {
	commands: any[] = [];
	ended = false;
	killed?: string;
	stdout = new EventEmitter();
	stderr = new EventEmitter();
	stdin = {
		write: (chunk: string) => {
			for (const line of chunk.split("\n").filter(Boolean)) this.commands.push(JSON.parse(line));
			return true;
		},
		end: () => {
			this.ended = true;
		},
		on: () => undefined,
	};
	command: string;
	args: string[];
	env: NodeJS.ProcessEnv;
	constructor(command: string, args: string[], env: NodeJS.ProcessEnv) {
		super();
		this.command = command;
		this.args = args;
		this.env = env;
	}
	kill(signal?: string) {
		this.killed = signal ?? "SIGTERM";
	}
	emitEvent(event: Record<string, unknown>) {
		this.stdout.emit("data", `${JSON.stringify(event)}\n`);
	}
	prompts() {
		return this.commands.filter((c) => c.type === "prompt");
	}
}

function setup(config: DispatchConfig = {}) {
	const agentDir = mkdtempSync(join(tmpdir(), "pi-teams-dispatch-"));
	const workers: FakeWorker[] = [];
	const journal: DispatchJournalEntry[] = [];
	const errors: string[] = [];
	let now = 1_000_000;
	const dispatcher = new ChatDispatcher({
		config: resolveDispatchConfig({ mode: "process", ...config }),
		agentDir,
		cwd: agentDir,
		env: {},
		now: () => now,
		journal: (entry) => journal.push(entry),
		onError: (message) => errors.push(message),
		sweepIntervalMs: 0,
		spawn: (command, args, options) => {
			const worker = new FakeWorker(command, args, options.env);
			workers.push(worker);
			return worker as unknown as WorkerProcess;
		},
	});
	return {
		agentDir,
		dispatcher,
		workers,
		journal,
		errors,
		advance: (ms: number) => {
			now += ms;
		},
	};
}

const job = (chatId: string, extra: Partial<DispatchJob> = {}): DispatchJob => ({
	chatId,
	label: `Chat ${chatId}`,
	chatType: "oneOnOne",
	from: "Anna",
	peer: "anna@example.com",
	prompt: `prompt for ${chatId}`,
	request: "hi",
	...extra,
});

describe("starting workers", () => {
	test("a wake starts a pi process in rpc mode for its chat", () => {
		const { dispatcher, workers, agentDir } = setup({ args: ["--model", "x/y"] });
		assert.equal(dispatcher.deliver(job("A")), "started");
		assert.equal(workers.length, 1);
		const [w] = workers;
		assert.equal(w!.command, "pi");
		assert.deepEqual(w!.args.slice(0, 4), ["--model", "x/y", "--mode", "rpc"]);
		assert.ok(w!.args.includes(chatSessionDir(agentDir, "A")));
		assert.equal(w!.env.PI_TEAMS_WORKER_CHAT, "A");
		assert.equal(w!.env.PI_TEAMS_WORKER_PEER, "anna@example.com");
		assert.deepEqual(w!.prompts()[0], { type: "prompt", message: "prompt for A", streamingBehavior: "followUp" });
	});

	test("a group chat carries no peer, so nobody reads across chats from it", () => {
		const { dispatcher, workers } = setup();
		dispatcher.deliver(job("G", { chatType: "group" }));
		assert.equal(workers[0]!.env.PI_TEAMS_WORKER_PEER, undefined);
	});

	test("different chats run in parallel", () => {
		const { dispatcher, workers } = setup({ maxConcurrent: 3 });
		assert.equal(dispatcher.deliver(job("A")), "started");
		assert.equal(dispatcher.deliver(job("B")), "started");
		assert.equal(dispatcher.deliver(job("C")), "started");
		assert.equal(workers.length, 3);
		assert.equal(dispatcher.status().busy, 3);
	});
});

describe("a chat whose worker is busy", () => {
	test("gets the new wake steered into the running turn", () => {
		const { dispatcher, workers } = setup();
		dispatcher.deliver(job("A"));
		assert.equal(dispatcher.deliver(job("A", { prompt: "second" })), "steered");
		assert.equal(workers.length, 1);
		assert.deepEqual(workers[0]!.prompts()[1], { type: "prompt", message: "second", streamingBehavior: "steer" });
	});

	test("continues in the same process once it has settled", () => {
		const { dispatcher, workers } = setup();
		dispatcher.deliver(job("A"));
		workers[0]!.emitEvent({ type: "agent_settled" });
		assert.equal(dispatcher.deliver(job("A", { prompt: "later" })), "started");
		assert.equal(workers.length, 1);
		assert.equal(workers[0]!.prompts()[1].streamingBehavior, "followUp");
	});
});

describe("a full house", () => {
	test("queues the next chat and starts it when a slot frees up", () => {
		const { dispatcher, workers, journal } = setup({ maxConcurrent: 1 });
		dispatcher.deliver(job("A"));
		assert.equal(dispatcher.deliver(job("B")), "queued");
		assert.equal(workers.length, 1);
		assert.equal(dispatcher.status().queued, 1);

		workers[0]!.emitEvent({ type: "message_end", message: { role: "assistant", content: [{ type: "text", text: "Done for A" }] } });
		workers[0]!.emitEvent({ type: "agent_settled" });

		assert.equal(workers.length, 2);
		assert.equal(workers[1]!.env.PI_TEAMS_WORKER_CHAT, "B");
		const done = journal.find((e) => e.event === "done");
		assert.equal(done?.chatId, "A");
		assert.equal(done?.result, "Done for A");
	});

	test("a newer wake for a waiting chat replaces the older one", () => {
		const { dispatcher, workers } = setup({ maxConcurrent: 1 });
		dispatcher.deliver(job("A"));
		dispatcher.deliver(job("B", { prompt: "old" }));
		dispatcher.deliver(job("B", { prompt: "new" }));
		assert.equal(dispatcher.status().queued, 1);
		workers[0]!.emitEvent({ type: "agent_settled" });
		assert.equal(workers[1]!.prompts()[0].message, "new");
	});

	test("a worker that crashes frees its slot and is reported", () => {
		const { dispatcher, workers, errors } = setup({ maxConcurrent: 1 });
		dispatcher.deliver(job("A"));
		dispatcher.deliver(job("B"));
		workers[0]!.emit("exit", 1, null);
		assert.equal(workers.length, 2);
		assert.match(errors[0] ?? "", /Chat A/);
	});
});

describe("control words", () => {
	test("stop aborts the running turn and still lets the model confirm", () => {
		const { dispatcher, workers } = setup();
		dispatcher.deliver(job("A"));
		assert.equal(dispatcher.deliver(job("A", { prompt: "stopp", control: "stop" })), "stopped");
		const types = workers[0]!.commands.map((c) => c.type);
		assert.deepEqual(types, ["prompt", "clear_queue", "abort", "prompt"]);
		assert.equal(workers[0]!.commands[3].streamingBehavior, "followUp");
	});

	test("stop with nothing running is an ordinary wake", () => {
		const { dispatcher } = setup();
		assert.equal(dispatcher.deliver(job("A", { control: "stop" })), "started");
	});

	test("reset ends the worker and starts over without --continue", () => {
		const { dispatcher, workers, agentDir } = setup();
		dispatcher.deliver(job("A"));
		workers[0]!.emitEvent({ type: "agent_settled" });
		// pretend the chat has a recent session
		const dir = chatSessionDir(agentDir, "A");
		mkdirSync(dir, { recursive: true });
		writeFileSync(join(dir, "s.jsonl"), "{}\n");

		assert.equal(dispatcher.deliver(job("A", { control: "reset" })), "reset");
		assert.equal(workers[0]!.ended, true);
		assert.equal(workers.length, 2);
		assert.ok(!workers[1]!.args.includes("--continue"));
	});
});

describe("idle workers", () => {
	test("exit after idleMinutes, and the next wake continues the session", () => {
		const { dispatcher, workers, advance, agentDir } = setup({ idleMinutes: 30 });
		dispatcher.deliver(job("A"));
		workers[0]!.emitEvent({ type: "agent_settled" });
		const dir = chatSessionDir(agentDir, "A");
		writeFileSync(join(dir, "s.jsonl"), "{}\n");

		advance(31 * 60_000);
		dispatcher.sweep();
		assert.equal(workers[0]!.ended, true);
		assert.equal(dispatcher.status().alive, 0);

		dispatcher.deliver(job("A"));
		assert.ok(workers[1]!.args.includes("--continue"));
	});

	test("the longest-idle worker makes room when maxWorkers is reached", () => {
		const { dispatcher, workers } = setup({ maxConcurrent: 1, maxWorkers: 1 });
		dispatcher.deliver(job("A"));
		workers[0]!.emitEvent({ type: "agent_settled" });
		dispatcher.deliver(job("B"));
		assert.equal(workers[0]!.ended, true);
		assert.equal(dispatcher.status().alive, 1);
	});
});

describe("dialogs in a worker", () => {
	test("are cancelled, because nobody is there to answer", () => {
		const { dispatcher, workers } = setup();
		dispatcher.deliver(job("A"));
		workers[0]!.emitEvent({ type: "extension_ui_request", id: "u1", method: "confirm", title: "?" });
		workers[0]!.emitEvent({ type: "extension_ui_request", id: "u2", method: "notify", message: "fyi" });
		const answers = workers[0]!.commands.filter((c) => c.type === "extension_ui_response");
		assert.deepEqual(answers, [{ type: "extension_ui_response", id: "u1", cancelled: true }]);
	});
});

describe("helpers", () => {
	test("shouldContinue respects freshAfterHours", () => {
		const dir = mkdtempSync(join(tmpdir(), "pi-teams-fresh-"));
		const now = Date.now();
		assert.equal(shouldContinue(dir, 72, now), false, "no session yet");
		const file = join(dir, "a.jsonl");
		writeFileSync(file, "{}\n");
		const old = (now - 100 * 3_600_000) / 1000;
		utimesSync(file, old, old);
		assert.equal(shouldContinue(dir, 72, now), false);
		assert.equal(shouldContinue(dir, 0, now), true);
		assert.equal(shouldContinue(dir, 200, now), true);
	});

	test("excerpt flattens and shortens", () => {
		assert.equal(excerpt("a\n  b"), "a b");
		assert.equal(excerpt("x".repeat(10), 5), "xxxx…");
		assert.equal(excerpt("  "), undefined);
	});

	test("config: defaults, clamps and maxWorkers >= maxConcurrent", () => {
		const d = resolveDispatchConfig(undefined);
		assert.equal(d.mode, "session");
		assert.equal(d.maxConcurrent, 3);
		const c = resolveDispatchConfig({ mode: "process", maxConcurrent: 99, maxWorkers: 2, stopWords: ["Halt "] });
		assert.equal(c.maxConcurrent, 16);
		assert.equal(c.maxWorkers, 16);
		assert.deepEqual(c.stopWords, ["halt"]);
		assert.deepEqual(resolveDispatchConfig({ mode: "bogus" as any }).mode, "session");
	});
});
