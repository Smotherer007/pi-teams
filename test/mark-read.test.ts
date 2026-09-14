/**
 * The read cursor: which endpoint it hits, and what the safety gates do with it.
 *
 * The Graph call itself is not mocked here — the repo has no fetch harness — so
 * the test targets the two things that actually go wrong in the wild: a wrong
 * path/body shape, and a mutation that slips past the safety gates.
 */

import { test, describe } from "node:test";
import assert from "node:assert/strict";
import { readStateRequest } from "../src/graph/chats.ts";
import { blockReason, formatMutationSummary, isMutationTool } from "../src/safety/index.ts";
import { teamsMarkReadTool } from "../src/tools/teams-mark-read.ts";

describe("readStateRequest", () => {
	test("read and unread are two endpoints, not one with a flag", () => {
		assert.equal(
			readStateRequest("19:abc", { id: "u1" }, true).path,
			"/chats/19%3Aabc/markChatReadForUser",
		);
		assert.equal(
			readStateRequest("19:abc", { id: "u1" }, false).path,
			"/chats/19%3Aabc/markChatUnreadForUser",
		);
	});

	test("the user goes in the body, which is where Graph reads it from", () => {
		const { body } = readStateRequest("19:abc", { id: "u1", tenantId: "t1" }, true);
		assert.deepEqual(body, { user: { id: "u1", tenantId: "t1" } });
	});

	test("a missing tenant id is dropped rather than sent as undefined", () => {
		const { body } = readStateRequest("19:abc", { id: "u1", tenantId: undefined }, true);
		assert.equal(JSON.stringify(body), '{"user":{"id":"u1"}}');
	});
});

describe("teams_mark_read safety", () => {
	test("it is a mutation, so readonly blocks it", () => {
		assert.equal(isMutationTool("teams_mark_read"), true);
		assert.match(blockReason("readonly", "device-code", "teams_mark_read") ?? "", /readonly/);
	});

	test("confirm lets it through to the dialog", () => {
		assert.equal(blockReason("confirm", "interactive", "teams_mark_read"), undefined);
	});

	test("an app-only token is refused with an explanation, not a Graph 403", () => {
		const reason = blockReason("open", "client-credentials", "teams_mark_read");
		assert.match(reason ?? "", /app-only/);
	});

	test("the confirmation dialog names the chat and the direction", () => {
		assert.match(
			formatMutationSummary("teams_mark_read", { chat: "SCODA_IPQT: Daily" }),
			/Mark as read.*SCODA_IPQT: Daily/,
		);
		assert.match(
			formatMutationSummary("teams_mark_read", { chat: "Anna", read: false }),
			/Mark as unread/,
		);
	});

	test("the tool is named and described for the model, with chat required", () => {
		assert.equal(teamsMarkReadTool.name, "teams_mark_read");
		assert.match(teamsMarkReadTool.description, /chat/i);
		assert.deepEqual(Object.keys(teamsMarkReadTool.parameters.properties), [
			"chat",
			"read",
			"account",
			"tenant",
		]);
	});
});
