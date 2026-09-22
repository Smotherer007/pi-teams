/**
 * Chat mapping — the activity timestamp.
 *
 * Listen mode dates every chat by this value: the freshness window in
 * `chatsToExamine` compares it against the poll interval, and the chat list
 * orders by it. Graph's `lastUpdatedDateTime` reports chat *metadata* and is
 * stale by more than a year on 1:1 chats in real tenants, so the message
 * timestamp has to win — otherwise a chat that just received a message looks
 * untouched and listen mode silently skips it.
 */

import { test, describe } from "node:test";
import assert from "node:assert/strict";
import { mapChat } from "../src/graph/mappers.ts";

const ME = "me-1";

function rawChat(overrides: Record<string, unknown> = {}) {
	return {
		id: "19:chat",
		chatType: "oneOnOne",
		members: [
			{ id: ME, displayName: "Patrick Weppelmann" },
			{ id: "t-1", displayName: "Tobias Kupfer" },
		],
		...overrides,
	} as never;
}

describe("mapChat activity timestamp", () => {
	test("prefers the last message over a stale lastUpdatedDateTime", () => {
		const chat = mapChat(
			rawChat({
				lastUpdatedDateTime: "2025-06-04T12:57:24.399Z",
				lastMessagePreview: {
					createdDateTime: "2026-09-14T14:57:57.508Z",
					body: { content: "<p>Hier nochmal ein Test</p>", contentType: "html" },
				},
			}),
			ME,
		);

		assert.equal(chat.lastUpdated, "2026-09-14T14:57:57.508Z");
	});

	test("falls back to lastUpdatedDateTime when there is no preview", () => {
		const chat = mapChat(rawChat({ lastUpdatedDateTime: "2026-09-14T13:25:48.079Z" }), ME);

		assert.equal(chat.lastUpdated, "2026-09-14T13:25:48.079Z");
	});

	test("reports no timestamp when Graph sends neither", () => {
		assert.equal(mapChat(rawChat(), ME).lastUpdated, undefined);
	});
});

describe("mapPerson — chat members", () => {
	test("a member's person id is its userId, not the membership id", async () => {
		const { mapPerson } = await import("../src/graph/mappers.ts");
		const person = mapPerson({
			id: "MCMjMSMjdGVuYW50",
			userId: "user-42",
			displayName: "Anna Schmidt",
			email: "anna@contoso.com",
		} as never);
		assert.equal(person?.id, "user-42");
		assert.equal(person?.mail, "anna@contoso.com");
	});
});
