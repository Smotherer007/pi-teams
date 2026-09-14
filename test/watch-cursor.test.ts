/**
 * Listen mode — the durable cursor and the tick that advances it.
 *
 * These tests cover the two rules the cursor exists for: pi answers whatever
 * moved since its last look, and it never answers the same message twice. The
 * first run on an account is the deliberate exception — there is no cursor yet,
 * so the history that is already there is recorded and left alone.
 */

import { test, describe, beforeEach, afterEach } from "node:test";
import assert from "node:assert/strict";
import { mkdtempSync, readFileSync, rmSync, statSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import {
	MAX_EXAMINATIONS_PER_TICK,
	runWatchTick,
	type WatchTickDeps,
} from "../src/watch/loop.ts";import {
	getWatchCursorPath,
	readWatchCursor,
	writeWatchCursor,
} from "../src/watch/cursor.ts";
import { createWatchState, MAX_CHAT_ATTEMPTS } from "../src/watch/index.ts";
import { WATCH_DEFAULTS, type ResolvedWatchConfig } from "../src/config/index.ts";
import type { ChatSummary, MessageSummary, SignedInUser } from "../src/types.ts";

const ACCOUNT = "work";
const TENANT = "3c8a65a5-573a-4e62-b56c-748f9c048488";
const NOW = new Date("2026-09-14T16:00:00Z").getTime();

const ME: SignedInUser = {
	id: "me-1",
	displayName: "Patrick Weppelmann",
	upn: "patrick@contoso.com",
	mail: "patrick@contoso.com",
	tenantId: TENANT,
};

const watch: ResolvedWatchConfig = { ...WATCH_DEFAULTS, enabled: true, cooldownSeconds: 0 };

let agentDir = "";
let savedAgentDir: string | undefined;

beforeEach(() => {
	agentDir = mkdtempSync(join(tmpdir(), "pi-teams-watch-"));
	savedAgentDir = process.env.PI_CODING_AGENT_DIR;
	process.env.PI_CODING_AGENT_DIR = agentDir;
});

afterEach(() => {
	if (savedAgentDir === undefined) delete process.env.PI_CODING_AGENT_DIR;
	else process.env.PI_CODING_AGENT_DIR = savedAgentDir;
	rmSync(agentDir, { recursive: true, force: true });
});

function chat(id: string, minutesAgo: number): ChatSummary {
	return {
		id,
		chatType: "oneOnOne",
		label: `${id}, Patrick Weppelmann`,
		members: [{ displayName: id, mail: `${id}@contoso.com` }],
		lastUpdated: new Date(NOW - minutesAgo * 60_000).toISOString(),
		lastMessageFrom: id,
	};
}

function message(from: string): MessageSummary {
	return {
		id: `msg-${from}`,
		text: "Kannst du kurz schauen?",
		contentType: "text",
		from: { id: `id-${from}`, displayName: from },
		createdDateTime: new Date(NOW - 60_000).toISOString(),
		mentions: [],
		reactions: [],
		attachments: [],
	};
}

/** Deps over a fixed chat list, recording what the watcher did. */
function deps(chats: ChatSummary[], options: { failFor?: string; readAt?: string } = {}) {
	const examined: string[] = [];
	const readChecked: string[] = [];
	const persisted: Map<string, string>[] = [];

	const deps: WatchTickDeps = {
		listChats: async () => chats,
		// Every chat is unread unless the test says otherwise: read state is what
		// the walk through the chat list is there for.
		readState: async (chatId) => {
			readChecked.push(chatId);
			return options.readAt ?? "1970-01-01T00:00:00.000Z";
		},
		listMessages: async (chatId) => {
			if (options.failFor === chatId) throw new Error("chat unavailable");
			examined.push(chatId);
			return [message(chatId)];
		},
		now: () => NOW,
		persistCursor: (seen) => persisted.push(new Map(seen)),
	};

	return { deps, examined, readChecked, persisted };
}

describe("the cursor file", () => {
	test("is absent until listen mode has run for that account", () => {
		assert.equal(readWatchCursor(ACCOUNT, TENANT), undefined);
	});

	test("round-trips the chat markers", () => {
		writeWatchCursor(ACCOUNT, TENANT, new Map([["chat-1", "2026-09-14T15:00:00.000Z"]]));
		const restored = readWatchCursor(ACCOUNT, TENANT);

		assert.deepEqual([...(restored ?? [])], [["chat-1", "2026-09-14T15:00:00.000Z"]]);
	});

	test("is written owner-readable only, and is valid JSON", () => {
		writeWatchCursor(ACCOUNT, TENANT, new Map([["chat-1", "marker"]]));
		const path = getWatchCursorPath(ACCOUNT, TENANT);

		assert.equal(statSync(path).mode & 0o777, 0o600);
		assert.equal(JSON.parse(readFileSync(path, "utf-8")).version, 1);
	});

	test("keeps accounts and tenants apart", () => {
		writeWatchCursor("work", TENANT, new Map([["chat-1", "marker"]]));
		assert.equal(readWatchCursor("private", TENANT), undefined);
		assert.equal(readWatchCursor("work", "other-tenant"), undefined);
	});

	test("a corrupt file reads as no cursor rather than crashing the watcher", () => {
		const path = getWatchCursorPath(ACCOUNT, TENANT);
		writeWatchCursor(ACCOUNT, TENANT, new Map([["chat-1", "marker"]]));
		writeFileSync(path, "{ not json", "utf-8");

		assert.equal(readWatchCursor(ACCOUNT, TENANT), undefined);
	});
});

describe("runWatchTick", () => {
	test("the first run answers the unread backlog", async () => {
		// No cursor at all is not "answer nothing": the user's own read state is
		// the filter, so a chat nobody has opened is answered whenever pi starts.
		const { deps: tickDeps, readChecked, persisted } = deps([chat("Anna", 30), chat("Tolga", 500)]);
		const result = await runWatchTick(tickDeps, createWatchState(), watch, ME);

		assert.deepEqual(result.wakes.map((wake) => wake.chat.id), ["Anna", "Tolga"]);
		assert.deepEqual(readChecked, ["Anna", "Tolga"]);
		assert.equal(result.pending, 2);
		assert.equal(persisted.length, 1, "the cursor has to exist after the first tick");
	});

	test("answers nothing for a chat the user has already read", async () => {
		const anna = chat("Anna", 30);
		const { deps: tickDeps, examined } = deps([anna], { readAt: anna.lastUpdated });

		const result = await runWatchTick(tickDeps, createWatchState(), watch, ME);

		assert.deepEqual(result.wakes, []);
		assert.equal(result.alreadyRead, 1);
		assert.deepEqual(examined, [], "the message is not even fetched");
	});

	test("with a cursor, a chat that moved while pi was off wakes it", async () => {
		const stale = chat("Anna", 500);
		const fresh = chat("Tolga", 30);
		// Tolga was answered last time; Anna moved since.
		const state = createWatchState([[fresh.id, fresh.lastUpdated!]]);
		const { deps: tickDeps, persisted } = deps([stale, fresh]);

		const result = await runWatchTick(tickDeps, state, watch, ME);

		assert.deepEqual(result.wakes.map((wake) => wake.chat.id), ["Anna"]);
		assert.equal(result.pending, 1);
		assert.deepEqual(persisted.at(-1)?.get("Anna"), stale.lastUpdated);
	});

	test("answers a message once, not on every tick", async () => {
		const anna = chat("Anna", 30);
		const state = createWatchState();

		const first = deps([anna]);
		assert.equal((await runWatchTick(first.deps, state, watch, ME)).wakes.length, 1);

		const second = deps([anna]);
		assert.equal((await runWatchTick(second.deps, state, watch, ME)).wakes.length, 0);
		assert.deepEqual(second.readChecked, [], "an unmoved chat is not even asked about");
	});

	test("a failed read call is retried instead of dropping the message", async () => {
		const anna = chat("Anna", 30);
		const state = createWatchState();
		const { deps: base } = deps([anna]);
		const failing: WatchTickDeps = {
			...base,
			readState: async () => {
				throw new Error("viewpoint unavailable");
			},
		};

		const first = await runWatchTick(failing, state, watch, ME);
		assert.deepEqual(first.wakes, []);
		assert.equal(state.seen.has("Anna"), false, "the chat stays open for the next tick");

		const next = deps([anna]);
		assert.equal((await runWatchTick(next.deps, state, watch, ME)).wakes.length, 1);
	});

	test("a chat that keeps failing is given up on, not retried forever", async () => {
		const anna = chat("Anna", 30);
		const state = createWatchState();
		const { deps: base } = deps([anna], { failFor: "Anna" });

		for (let attempt = 1; attempt <= MAX_CHAT_ATTEMPTS; attempt++) {
			const result = await runWatchTick(base, state, watch, ME);
			assert.equal(result.examined, 1, `attempt ${attempt} examines the chat`);
		}

		assert.equal(state.seen.has("Anna"), true, "after the last attempt it is left alone");
		const after = deps([anna]);
		assert.equal((await runWatchTick(after.deps, state, watch, ME)).examined, 0);
	});

	test("a read state Graph will not give is answered, not ignored", async () => {
		const anna = chat("Anna", 30);
		const { deps: tickDeps } = deps([anna], { readAt: undefined });

		const result = await runWatchTick(tickDeps, createWatchState(), watch, ME);
		assert.equal(result.wakes.length, 1);
	});

	test("a backlog is drained over several ticks, not all at once", async () => {
		const chats = Array.from({ length: MAX_EXAMINATIONS_PER_TICK + 3 }, (_, index) =>
			chat(`chat-${index}`, index + 1),
		);
		const state = createWatchState();

		const first = deps(chats);
		const result = await runWatchTick(first.deps, state, watch, ME);

		assert.equal(result.examined, MAX_EXAMINATIONS_PER_TICK);
		assert.equal(result.pending, chats.length, "the rest is reported, not dropped");

		const second = deps(chats);
		const rest = await runWatchTick(second.deps, state, watch, ME);
		assert.equal(rest.examined, chats.length - MAX_EXAMINATIONS_PER_TICK);
		assert.equal(rest.pending, chats.length - MAX_EXAMINATIONS_PER_TICK);
	});

	test("lists no chat the read rules exclude", async () => {
		const anna = chat("Anna", 30);
		const { deps: tickDeps, readChecked } = deps([anna]);
		const result = await runWatchTick(
			{ ...tickDeps, allowed: () => false },
			createWatchState(),
			watch,
			ME,
		);
		assert.deepEqual(result.wakes, []);
		assert.deepEqual(readChecked, []);
	});
});
