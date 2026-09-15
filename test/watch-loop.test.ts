/**
 * Listen mode — one tick, with Graph stubbed out.
 *
 * The failure this file exists for: a message that could not wake pi *yet* —
 * because its chat was inside the cooldown, or because the hourly wake limit had
 * been reached — used to be written into the cursor anyway. The cursor is what
 * stops the watcher from ever looking at that chat again, so the message was
 * gone; two real ones disappeared that way before anyone noticed.
 *
 * Every test here asks the same question from a different angle: **may the
 * cursor learn about this chat?** It may only when the decision is final.
 */

import { describe, test } from "node:test";
import assert from "node:assert/strict";
import { runWatchTick, type WatchTickDeps } from "../src/watch/loop.ts";
import {
	activityMarker,
	createWatchState,
	noteDeferral,
	noteWake,
	type WatchState,
} from "../src/watch/index.ts";
import { WATCH_DEFAULTS, type ResolvedWatchConfig } from "../src/config/index.ts";
import type { ChatSummary, MessageSummary, SignedInUser } from "../src/types.ts";

const ME: SignedInUser = {
	id: "me-1",
	displayName: "Patrick Weppelmann",
	upn: "patrick@contoso.com",
	tenantId: "tenant-1",
};

const NOW = new Date("2026-09-14T17:00:00Z").getTime();

function watch(overrides: Partial<ResolvedWatchConfig> = {}): ResolvedWatchConfig {
	return { ...WATCH_DEFAULTS, enabled: true, from: ["Tolga Barlak"], ...overrides };
}

function chat(overrides: Partial<ChatSummary> = {}): ChatSummary {
	return {
		id: "chat-1",
		chatType: "oneOnOne",
		label: "Tolga Barlak, Patrick Weppelmann",
		members: [{ displayName: "Tolga Barlak" }],
		lastUpdated: new Date(NOW - 60_000).toISOString(),
		lastMessageFrom: "Tolga Barlak",
		...overrides,
	};
}

function message(overrides: Partial<MessageSummary> = {}): MessageSummary {
	return {
		id: "msg-1",
		text: "und entscheiden ob und wenn ja was ich wann zu meinem 40. gb mache",
		contentType: "text",
		from: { id: "tolga-1", displayName: "Tolga Barlak" },
		createdDateTime: new Date(NOW - 60_000).toISOString(),
		mentions: [],
		reactions: [],
		attachments: [],
		imageUrls: [],
		...overrides,
	};
}

interface Harness {
	deps: WatchTickDeps;
	state: WatchState;
	watch: ResolvedWatchConfig;
	/** Movable clock, so a test can wait out a cooldown without waiting */
	clock: { now: number };
	counters: { messagesRead: number; cursorWrites: number };
}

function harness(
	options: {
		chats?: ChatSummary[];
		messages?: MessageSummary[];
		readAt?: string | (() => string | undefined);
		watch?: ResolvedWatchConfig;
		state?: WatchState;
		failing?: boolean;
	} = {},
): Harness {
	const state = options.state ?? createWatchState();
	const clock = { now: NOW };
	const counters = { messagesRead: 0, cursorWrites: 0 };

	const deps: WatchTickDeps = {
		listChats: async () => options.chats ?? [chat()],
		readState: async () => {
			if (options.failing) throw new Error("graph is unhappy");
			return typeof options.readAt === "function" ? options.readAt() : options.readAt;
		},
		listMessages: async () => {
			counters.messagesRead += 1;
			if (options.failing) throw new Error("graph is unhappy");
			return options.messages ?? [message()];
		},
		now: () => clock.now,
		persistCursor: () => {
			counters.cursorWrites += 1;
		},
	};

	return { deps, state, watch: options.watch ?? watch(), clock, counters };
}

const tick = (s: Harness) => runWatchTick(s.deps, s.state, s.watch, ME);

describe("runWatchTick", () => {
	test("wakes for an unread message from a watched sender and writes it down", async () => {
		const s = harness();
		const result = await tick(s);

		assert.equal(result.wakes.length, 1);
		assert.equal(result.examined, 1);
		assert.equal(result.deferred, 0);
		assert.equal(s.state.seen.get("chat-1"), activityMarker(chat()));
		assert.equal(s.counters.cursorWrites, 1);
	});

	test("answers with the whole open thread, not only the newest line", async () => {
		const earlier = message({
			id: "msg-1",
			text: "ich muss noch in den Odyseus Film",
			createdDateTime: new Date(NOW - 300_000).toISOString(),
		});
		const newest = message({ id: "msg-2", createdDateTime: new Date(NOW - 60_000).toISOString() });
		const s = harness({
			messages: [newest, earlier],
			readAt: new Date(NOW - 600_000).toISOString(),
		});

		const result = await tick(s);
		assert.deepEqual(
			result.wakes[0]?.backlog?.map((entry) => entry.id),
			["msg-1", "msg-2"],
		);
	});

	test("a chat inside its cooldown keeps its message instead of losing it", async () => {
		const state = createWatchState();
		noteWake(state, "chat-1", NOW);
		const s = harness({ state });

		const first = await tick(s);
		assert.equal(first.wakes.length, 0);
		assert.equal(first.deferred, 1);
		// The whole point: nothing about a postponed chat may reach the cursor.
		assert.equal(s.state.seen.has("chat-1"), false);

		// Nor may it cost a Graph call per tick while it waits.
		s.clock.now = NOW + 60_000;
		assert.equal((await tick(s)).examined, 0);

		// Once the cooldown is over, the message is still there and gets answered.
		s.clock.now = NOW + 300_001;
		assert.equal((await tick(s)).wakes.length, 1);
	});

	test("the hourly limit postpones a wake, it does not consume the message", async () => {
		const state = createWatchState();
		noteWake(state, "other-0", NOW - 120_000);
		noteWake(state, "other-1", NOW - 60_000);
		const s = harness({
			state,
			watch: watch({ maxTriggersPerHour: 2, cooldownSeconds: 0 }),
		});

		const blocked = await tick(s);
		assert.equal(blocked.wakes.length, 0);
		assert.equal(blocked.deferred, 1);
		assert.equal(s.state.seen.has("chat-1"), false);

		// Held back until the oldest wake of the window falls out of it.
		s.clock.now = NOW + 3_000_000;
		assert.equal((await tick(s)).examined, 0);

		s.clock.now = NOW + 3_480_001;
		assert.equal((await tick(s)).wakes.length, 1);
	});

	test("a waiting chat does not take a batch slot from another", async () => {
		const waiting = chat({ id: "chat-waiting" });
		const others = Array.from({ length: 6 }, (_, index) =>
			chat({
				id: `chat-${index}`,
				lastUpdated: new Date(NOW - 30_000 - index * 1000).toISOString(),
			}),
		);
		const state = createWatchState();
		noteDeferral(state, waiting.id, NOW + 60_000);
		const s = harness({ state, chats: [waiting, ...others] });

		const result = await tick(s);
		assert.equal(result.examined, 5);
		assert.equal(s.counters.messagesRead, 5);
	});

	test("a sender outside the listen list is a final decision", async () => {
		const s = harness({ watch: watch({ from: ["Tobias Kupfer"] }) });

		const result = await tick(s);
		assert.equal(result.wakes.length, 0);
		assert.equal(result.deferred, 0);
		assert.equal(s.state.seen.get("chat-1"), activityMarker(chat()));

		// Written down, so the next tick does not even read the chat again.
		s.clock.now = NOW + 60_000;
		assert.equal((await tick(s)).examined, 0);
	});

	test("a message the user already read is settled before it is fetched", async () => {
		const s = harness({ readAt: new Date(NOW).toISOString() });

		const result = await tick(s);
		assert.equal(result.alreadyRead, 1);
		assert.equal(result.wakes.length, 0);
		assert.equal(s.state.seen.has("chat-1"), true);
		assert.equal(s.counters.messagesRead, 0);
	});

	test("a chat that cannot be read is retried, then given up on", async () => {
		const s = harness({ failing: true });

		await tick(s);
		assert.equal(s.state.seen.has("chat-1"), false, "the first failure keeps it open");
		await tick(s);
		assert.equal(s.state.seen.has("chat-1"), false);
		await tick(s);
		assert.equal(s.state.seen.has("chat-1"), true, "the third stops it taking a slot");
	});
});
