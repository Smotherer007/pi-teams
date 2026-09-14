/**
 * Listen mode — the decision rules.
 *
 * Every case here answers the same question: would this message cost a model
 * turn? A wrong "wake" is money and noise, a wrong "skip" is a missed message,
 * so both directions are tested explicitly.
 */

import { test, describe } from "node:test";
import assert from "node:assert/strict";
import {
	activityMarker,
	chatsToExamine,
	createWatchState,
	isFromMe,
	isUnread,
	isWatchedChat,
	isWatchedSender,
	mentionsMe,
	noteDeferral,
	noteWake,
	openMessages,
	pruneState,
	settleChat,
	shouldWake,
	waitingCount,
	wakesThisHour,
} from "../src/watch/index.ts";
import { composeWatchPrompt } from "../src/watch/prompt.ts";
import { WATCH_DEFAULTS, type ResolvedWatchConfig } from "../src/config/index.ts";
import type { ChatSummary, MessageSummary, SignedInUser } from "../src/types.ts";

const ME: SignedInUser = {
	id: "me-1",
	displayName: "Patrick Weppelmann",
	upn: "patrick@contoso.com",
	mail: "patrick@contoso.com",
	tenantId: "tenant-1",
};

const NOW = new Date("2026-09-14T16:00:00Z").getTime();

function watch(overrides: Partial<ResolvedWatchConfig> = {}): ResolvedWatchConfig {
	return { ...WATCH_DEFAULTS, enabled: true, ...overrides };
}

function chat(overrides: Partial<ChatSummary> = {}): ChatSummary {
	return {
		id: "chat-1",
		chatType: "oneOnOne",
		label: "Anna Schmidt, Patrick Weppelmann",
		members: [{ displayName: "Anna Schmidt", mail: "anna.schmidt@contoso.com" }],
		lastUpdated: new Date(NOW - 60_000).toISOString(),
		lastMessageFrom: "Anna Schmidt",
		...overrides,
	};
}

function message(overrides: Partial<MessageSummary> = {}): MessageSummary {
	return {
		id: "msg-1",
		text: "Kannst du kurz schauen?",
		contentType: "text",
		from: { id: "other-1", displayName: "Anna Schmidt", mail: "anna.schmidt@contoso.com" },
		createdDateTime: new Date(NOW - 60_000).toISOString(),
		mentions: [],
		reactions: [],
		attachments: [],
		...overrides,
	};
}

describe("isFromMe", () => {
	test("matches the signed-in user by ID and by name", () => {
		assert.equal(isFromMe({ id: "me-1", displayName: "Someone Else" }, ME), true);
		assert.equal(isFromMe({ displayName: "Patrick Weppelmann" }, ME), true);
		assert.equal(isFromMe({ id: "other-1", displayName: "Anna Schmidt" }, ME), false);
	});

	test("says nothing rather than guessing when either side is unknown", () => {
		assert.equal(isFromMe(undefined, ME), false);
		assert.equal(isFromMe({ displayName: "Anna Schmidt" }, undefined), false);
	});
});

describe("mentionsMe", () => {
	test("uses the mention list when Graph provides one", () => {
		const msg = message({ mentions: [{ id: "me-1", displayName: "Patrick Weppelmann" }] });
		assert.equal(mentionsMe(msg, ME), true);
	});

	test("falls back to the text, including first name only", () => {
		assert.equal(mentionsMe(message({ text: "@Patrick Weppelmann schaust du?" }), ME), true);
		assert.equal(mentionsMe(message({ text: "Patrick, schaust du?" }), ME), true);
		assert.equal(mentionsMe(message({ text: "Wer ist Patrick?" }), ME), true);
		assert.equal(mentionsMe(message({ text: "Alles gut hier" }), ME), false);
	});

	test("does not match inside a longer word", () => {
		assert.equal(mentionsMe(message({ text: "Patrickchen" }), ME), false);
	});
});

describe("isWatchedChat", () => {
	test("no patterns means every chat", () => {
		assert.equal(isWatchedChat(chat(), watch()), true);
	});

	test("matches topic, label, id and participants", () => {
		assert.equal(isWatchedChat(chat(), watch({ chats: ["Anna*"] })), true);
		assert.equal(isWatchedChat(chat(), watch({ chats: ["chat-1"] })), true);
		assert.equal(isWatchedChat(chat(), watch({ chats: ["*@contoso.com"] })), true);
		assert.equal(isWatchedChat(chat(), watch({ chats: ["Vertrieb*"] })), false);
	});
});

describe("isWatchedSender", () => {
	test("no patterns means any sender", () => {
		assert.equal(isWatchedSender(message(), watch()), true);
	});

	test("matches display name, UPN and e-mail", () => {
		assert.equal(isWatchedSender(message(), watch({ from: ["Anna*"] })), true);
		assert.equal(isWatchedSender(message(), watch({ from: ["anna.schmidt@contoso.com"] })), true);
		assert.equal(isWatchedSender(message(), watch({ from: ["Bernd*"] })), false);
	});

	test("a message without a sender is never a match", () => {
		assert.equal(isWatchedSender(message({ from: undefined }), watch({ from: ["*"] })), false);
	});
});

describe("shouldWake", () => {
	test("wakes for a new message from someone else", () => {
		assert.deepEqual(shouldWake(chat(), message(), ME, watch(), createWatchState(), NOW), { wake: true });
	});

	test("never wakes for pi's own message — this is the loop guard", () => {
		const own = message({ from: { id: "me-1", displayName: "Patrick Weppelmann" } });
		const decision = shouldWake(chat(), own, ME, watch(), createWatchState(), NOW);
		assert.equal(decision.wake, false);
	});

	test("honours the sender filter and the mention filter", () => {
		const from = watch({ from: ["Bernd*"] });
		const skipped = shouldWake(chat(), message(), ME, from, createWatchState(), NOW);
		assert.equal(skipped.wake, false);
		assert.match(skipped.wake === false ? skipped.reason : "", /sender/);

		const mentions = watch({ mentionOnly: true });
		assert.equal(shouldWake(chat(), message(), ME, mentions, createWatchState(), NOW).wake, false);
		const mentioned = message({ mentions: [{ id: "me-1", displayName: "Patrick Weppelmann" }] });
		assert.equal(shouldWake(chat(), mentioned, ME, mentions, createWatchState(), NOW).wake, true);
	});

	test("stays quiet inside the cooldown", () => {
		const state = createWatchState();
		noteWake(state, "chat-1", NOW);
		assert.equal(shouldWake(chat(), message(), ME, watch(), state, NOW + 1000).wake, false);
		assert.equal(shouldWake(chat(), message(), ME, watch({ cooldownSeconds: 0 }), state, NOW + 1).wake, true);
	});

	test("stops at the hourly limit", () => {
		const state = createWatchState();
		const limit = watch({ maxTriggersPerHour: 2, cooldownSeconds: 0 });
		assert.equal(shouldWake(chat(), message(), ME, limit, state, NOW).wake, true);
		noteWake(state, "chat-a", NOW);
		assert.equal(shouldWake(chat(), message(), ME, limit, state, NOW).wake, true);
		noteWake(state, "chat-b", NOW);
		assert.equal(shouldWake(chat(), message(), ME, limit, state, NOW).wake, false);
		// An hour later the budget is back.
		assert.equal(shouldWake(chat(), message(), ME, limit, state, NOW + 3600_001).wake, true);
	});

	test("a cooldown says how long to wait instead of writing the chat off", () => {
		const state = createWatchState();
		noteWake(state, "chat-1", NOW);
		const decision = shouldWake(chat(), message(), ME, watch(), state, NOW + 10_000);
		assert.equal(decision.wake, false);
		assert.equal(decision.wake === false ? decision.retryAfterMs : undefined, 290_000);
	});

	test("the hourly limit says when the next slot frees up", () => {
		const state = createWatchState();
		const limit = watch({ maxTriggersPerHour: 2, cooldownSeconds: 0 });
		noteWake(state, "chat-a", NOW - 120_000);
		noteWake(state, "chat-b", NOW - 60_000);

		const decision = shouldWake(chat(), message(), ME, limit, state, NOW);
		assert.equal(decision.wake, false);
		// The oldest wake has to leave the sliding window first.
		assert.equal(decision.wake === false ? decision.retryAfterMs : undefined, 3_480_000);
	});

	test("a final skip carries no retry, so the chat may be written off", () => {
		const from = watch({ from: ["Bernd*"] });
		const decision = shouldWake(chat(), message(), ME, from, createWatchState(), NOW);
		assert.equal(decision.wake, false);
		assert.equal(decision.wake === false ? decision.retryAfterMs : undefined, undefined);
	});

	test("ignores a deleted message", () => {
		const deleted = message({ deletedDateTime: new Date(NOW).toISOString() });
		assert.equal(shouldWake(chat(), deleted, ME, watch(), createWatchState(), NOW).wake, false);
	});
});

describe("isUnread", () => {
	test("compares the newest message against the user's read cursor", () => {
		const moved = chat({ lastUpdated: "2026-09-14T15:00:35.103Z" });
		assert.equal(isUnread(moved, "2026-09-14T14:55:57.835Z"), true);
		assert.equal(isUnread(moved, "2026-09-14T15:00:35.103Z"), false);
		assert.equal(isUnread(moved, "2026-09-14T16:00:00.000Z"), false);
	});

	test("answers rather than guesses when the read state is missing", () => {
		// A read state we cannot read must not switch listen mode off silently.
		assert.equal(isUnread(chat(), undefined), true);
		assert.equal(isUnread(chat(), "not a date"), true);
	});

	test("a chat the chat list cannot date is never unread", () => {
		assert.equal(isUnread(chat({ lastUpdated: undefined }), undefined), false);
	});
});

describe("chatsToExamine", () => {
	test("skips chats whose activity marker has not moved", () => {
		const state = createWatchState();
		const first = chat();
		state.seen.set(first.id, activityMarker(first));
		assert.deepEqual(chatsToExamine(state, [first], watch(), NOW), []);
	});

	test("examines a chat that moved while pi was not running, however long ago", () => {
		// The cursor is what makes listen mode pick up a backlog: age is not part
		// of the test, or exactly the messages the user switched it on for would
		// be dropped.
		const old = chat({ lastUpdated: new Date(NOW - 30 * 24 * 3600_000).toISOString() });
		assert.equal(chatsToExamine(createWatchState(), [old], watch(), NOW).length, 1);
	});

	test("examines a chat that just moved", () => {
		const fresh = chat({ lastUpdated: new Date(NOW - 30_000).toISOString() });
		assert.equal(chatsToExamine(createWatchState(), [fresh], watch(), NOW).length, 1);
	});

	test("a restored cursor is what an empty state would have examined", () => {
		const restored = chat({ lastUpdated: new Date(NOW - 30_000).toISOString() });
		const state = createWatchState([[restored.id, activityMarker(restored)]]);
		assert.deepEqual(chatsToExamine(state, [restored], watch(), NOW), []);
	});

	test("a chat the chat list cannot date is never a candidate", () => {
		const undated = chat({ lastUpdated: undefined });
		assert.deepEqual(chatsToExamine(createWatchState(), [undated], watch(), NOW), []);
	});

	test("respects the chat filter", () => {
		const fresh = chat({ lastUpdated: new Date(NOW - 30_000).toISOString() });
		const narrowed = watch({ chats: ["Vertrieb*"] });
		assert.deepEqual(chatsToExamine(createWatchState(), [fresh], narrowed, NOW), []);
	});
});

describe("state bookkeeping", () => {
	test("wakesThisHour only counts the last hour", () => {
		const state = createWatchState();
		noteWake(state, "a", NOW - 3600_001);
		noteWake(state, "b", NOW - 1000);
		assert.equal(wakesThisHour(state, NOW), 1);
	});

	test("pruning bounds the maps a long session accumulates", () => {
		const state = createWatchState();
		noteWake(state, "old", NOW - 25 * 3600_000);
		noteWake(state, "recent", NOW - 1000);
		pruneState(state, NOW);
		assert.deepEqual([...state.lastTriggered.keys()], ["recent"]);
		assert.equal(state.triggers.length, 1);
	});

	test("pruning never drops the cursor of a quiet chat", () => {
		// By age this would be a candidate again on the next tick, and pi would
		// answer a message from months ago.
		const state = createWatchState([["quiet", "2000-01-01T00:00:00.000Z"]]);
		pruneState(state, NOW);
		assert.equal(state.seen.get("quiet"), "2000-01-01T00:00:00.000Z");
	});

	test("the cursor is bounded by size, keeping the newest markers", () => {
		const entries: [string, string][] = [];
		for (let index = 0; index < 1005; index++) {
			entries.push([`chat-${index}`, new Date(NOW - index * 1000).toISOString()]);
		}
		const state = createWatchState(entries);
		pruneState(state, NOW);

		assert.equal(state.seen.size, 1000);
		assert.equal(state.seen.has("chat-0"), true);
		assert.equal(state.seen.has("chat-1004"), false);
	});
});

describe("deferral bookkeeping", () => {
	test("settling a chat clears its deferral and writes the marker", () => {
		const state = createWatchState();
		noteDeferral(state, "chat-1", NOW + 60_000);
		assert.equal(waitingCount(state, NOW), 1);

		settleChat(state, "chat-1", "marker-1");
		assert.equal(waitingCount(state, NOW), 0);
		assert.equal(state.seen.get("chat-1"), "marker-1");
	});

	test("a chat is not examined while it waits, and is again afterwards", () => {
		const state = createWatchState();
		const waiting = chat();
		noteDeferral(state, waiting.id, NOW + 60_000);

		assert.deepEqual(chatsToExamine(state, [waiting], watch(), NOW + 30_000), []);
		assert.equal(chatsToExamine(state, [waiting], watch(), NOW + 60_001).length, 1);
	});

	test("pruning forgets a wait that is over", () => {
		const state = createWatchState();
		noteDeferral(state, "chat-1", NOW - 1);
		pruneState(state, NOW);
		assert.equal(waitingCount(state, NOW), 0);
	});
});

describe("openMessages", () => {
	const earlier = message({ id: "msg-1", text: "ich brauche Urlaub", createdDateTime: new Date(NOW - 300_000).toISOString() });
	const middle = message({ id: "msg-2", text: "bzw ich mach drei kreuze", createdDateTime: new Date(NOW - 240_000).toISOString() });
	const newest = message({ id: "msg-3", text: "hilft dir das?", createdDateTime: new Date(NOW - 60_000).toISOString() });

	test("returns everything the read cursor has not caught up with, oldest first", () => {
		const readAt = new Date(NOW - 600_000).toISOString();
		assert.deepEqual(
			openMessages([newest, middle, earlier], readAt, ME).map((msg) => msg.id),
			["msg-1", "msg-2", "msg-3"],
		);
	});

	test("drops what the user already read", () => {
		const readAt = new Date(NOW - 270_000).toISOString();
		assert.deepEqual(
			openMessages([newest, middle, earlier], readAt, ME).map((msg) => msg.id),
			["msg-2", "msg-3"],
		);
	});

	test("never treats my own messages as open questions", () => {
		const mine = message({ id: "msg-4", from: { id: "me-1", displayName: "Patrick Weppelmann" } });
		const open = openMessages([mine, newest], new Date(NOW - 600_000).toISOString(), ME);
		assert.deepEqual(open.map((msg) => msg.id), ["msg-3"]);
	});

	test("skips deleted messages", () => {
		const gone = message({ id: "msg-5", deletedDateTime: new Date(NOW - 30_000).toISOString() });
		const open = openMessages([gone, newest], undefined, ME);
		assert.deepEqual(open.map((msg) => msg.id), ["msg-3"]);
	});

	test("treats a read cursor Graph will not give as everything being open", () => {
		assert.equal(openMessages([newest, middle, earlier], undefined, ME).length, 3);
	});
});

describe("composeWatchPrompt", () => {
	test("carries the chat ID, the sender and the instruction to use it", () => {
		const event = { chat: chat(), message: message(), wokeAt: NOW, me: ME };
		const prompt = composeWatchPrompt(event, ME);

		assert.match(prompt, /chat-1/);
		assert.match(prompt, /Anna Schmidt/);
		assert.match(prompt, /teams_send_chat_message/);
		assert.match(prompt, /> Kannst du kurz schauen\?/);
		assert.match(prompt, /If it does not: say so in one line and stop\./);
	});

	test("quotes a multi-line message line by line", () => {
		const event = { chat: chat(), message: message({ text: "Zeile 1\nZeile 2" }), wokeAt: NOW, me: ME };
		const prompt = composeWatchPrompt(event, ME);
		assert.match(prompt, /> Zeile 1\n> Zeile 2/);
	});

	test("says something useful when the message has no text", () => {
		const event = { chat: chat(), message: message({ text: "" }), wokeAt: NOW, me: ME };
		assert.match(composeWatchPrompt(event, ME), /no text content/);
	});

	test("shows the whole open thread and asks for all of it, not only the newest", () => {
		const earlier = message({ id: "msg-1", text: "ich brauche Urlaub" });
		const middle = message({ id: "msg-2", text: "bzw ich mach drei kreuze" });
		const newest = message({ id: "msg-3", text: "hilft dir das?" });
		const event = {
			chat: chat(),
			message: newest,
			backlog: [earlier, middle, newest],
			wokeAt: NOW,
			me: ME,
		};
		const prompt = composeWatchPrompt(event, ME);

		assert.match(prompt, /> 1\. .*\n> ich brauche Urlaub/);
		assert.match(prompt, /> 2\. /);
		assert.match(prompt, /> bzw ich mach drei kreuze/);
		assert.match(prompt, /newest/);
		assert.match(prompt, /Answer every one of them, not just the newest/);
		assert.match(prompt, /Unanswered messages in this chat \(3\), oldest first:/);
	});

	test("keeps the single-message shape when nothing older is open", () => {
		const event = { chat: chat(), message: message(), backlog: [message()], wokeAt: NOW, me: ME };
		const prompt = composeWatchPrompt(event, ME);

		assert.match(prompt, /^Message:$/m);
		assert.match(prompt, /Decide whether this needs an answer\./);
		assert.doesNotMatch(prompt, /Answer every one of them/);
	});
});
