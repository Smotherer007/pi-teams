/**
 * The routing hint pi-teams emits for a router such as pi-lanes.
 */

import { test } from "node:test";
import assert from "node:assert/strict";
import { laneHint } from "../src/extension/index.ts";

const conn = { account: "neo", watch: { historyReaders: ["patrick@example.com"] } } as any;
const event = (chatType: string, from: { displayName: string; mail?: string }, text = "Neo stopp!") =>
	({
		chat: { id: "19:abc@thread.v2", label: "Patrick", chatType, members: [] },
		message: { id: "1", text, contentType: "text", from, mentions: [{ displayName: "Neo" }], reactions: [], attachments: [] },
		me: { displayName: "Neo (neoimpulse)" },
		wokeAt: 0,
	}) as any;

test("one lane per chat and account, pinned by environment", () => {
	const hint = laneHint(conn, event("oneOnOne", { displayName: "Patrick", mail: "patrick@example.com" }), "PROMPT");
	assert.equal(hint.text, "PROMPT");
	assert.equal(hint.lane, "teams:neo:19:abc@thread.v2");
	assert.equal(hint.env.PI_TEAMS_WORKER_CHAT, "19:abc@thread.v2");
	assert.equal(hint.env.PI_TEAMS_WORKER_PEER, "patrick@example.com");
	assert.equal(hint.command, "stopp");
	assert.equal(hint.trusted, true);
});

test("trust only in a one-to-one chat with a history reader", () => {
	assert.equal(laneHint(conn, event("oneOnOne", { displayName: "Anna", mail: "anna@example.com" }), "P").trusted, false);
	const group = laneHint(conn, event("group", { displayName: "Patrick", mail: "patrick@example.com" }), "P");
	assert.equal(group.trusted, false);
	assert.equal(group.env.PI_TEAMS_WORKER_PEER, undefined);
});
