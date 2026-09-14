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
	isWatchedChat,
	isWatchedSender,
	mentionsMe,
	noteWake,
	pruneState,
	shouldWake,
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

	test("ignores a deleted message", () => {
		const deleted = message({ deletedDateTime: new Date(NOW).toISOString() });
		assert.equal(shouldWake(chat(), deleted, ME, watch(), createWatchState(), NOW).wake, false);
	});
});

describe("chatsToExamine", () => {
	test("skips chats whose activity marker has not moved", () => {
		const state = createWatchState();
		const first = chat();
		state.seen.set(first.id, activityMarker(first));
		assert.deepEqual(chatsToExamine(state, [first], watch(), NOW, 5 * 60_000), []);
	});

	test("skips chats older than the freshness window", () => {
		const stale = chat({ lastUpdated: new Date(NOW - 3600_000).toISOString() });
		assert.deepEqual(chatsToExamine(createWatchState(), [stale], watch(), NOW, 5 * 60_000), []);
	});

	test("returns a chat that just moved", () => {
		const fresh = chat({ lastUpdated: new Date(NOW - 30_000).toISOString() });
		assert.equal(chatsToExamine(createWatchState(), [fresh], watch(), NOW, 5 * 60_000).length, 1);
	});

	test("respects the chat filter", () => {
		const fresh = chat({ lastUpdated: new Date(NOW - 30_000).toISOString() });
		const narrowed = watch({ chats: ["Vertrieb*"] });
		assert.deepEqual(chatsToExamine(createWatchState(), [fresh], narrowed, NOW, 5 * 60_000), []);
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
});
