/**
 * teams_history: reading across chats is allowed only where it is meant to be.
 */

import { test } from "node:test";
import assert from "node:assert/strict";
import { historyRefusal, renderHistory } from "../src/tools/teams-history.ts";
import { workerIdentity, workerViolation } from "../src/watch/worker.ts";

test("the watching session is never restricted", () => {
	assert.equal(historyRefusal(undefined, []), undefined);
});

test("a worker may read only in a one-to-one chat with a listed reader", () => {
	const readers = ["patrick@example.com"];
	assert.equal(historyRefusal({ chatId: "c", label: "P", chatType: "oneOnOne", peer: "patrick@example.com" }, readers), undefined);
	assert.ok(historyRefusal({ chatId: "c", label: "A", chatType: "oneOnOne", peer: "anna@example.com" }, readers));
	assert.ok(historyRefusal({ chatId: "c", label: "G", chatType: "group" }, readers));
});

test("renderHistory filters by time and chat", () => {
	const now = Date.parse("2026-09-22T12:00:00Z");
	const entries = [
		{ at: "2026-09-22T11:00:00Z", tool: "teams_send_chat_message", target: "chat:Anna", summary: "Hallo Anna" },
		{ at: "2026-09-22T11:30:00Z", tool: "teams_send_chat_message", target: "chat:Bob", summary: "Hallo Bob" },
		{ at: "2026-09-10T11:30:00Z", tool: "teams_send_chat_message", target: "chat:Anna", summary: "alt" },
	];
	const text = renderHistory(entries, { chat: "anna" }, now);
	assert.match(text, /Hallo Anna/);
	assert.doesNotMatch(text, /Bob|alt/);
});

test("worker identity comes from the environment, and config tools are closed to it", () => {
	const id = workerIdentity({ PI_TEAMS_WORKER_CHAT: "19:a", PI_TEAMS_WORKER_LABEL: "Anna" });
	assert.equal(id?.label, "Anna");
	assert.equal(workerIdentity({}), undefined);
	assert.ok(workerViolation("teams_watch", id));
	assert.equal(workerViolation("teams_send_chat_message", id), undefined);
	assert.equal(workerViolation("teams_watch", undefined), undefined);
});
