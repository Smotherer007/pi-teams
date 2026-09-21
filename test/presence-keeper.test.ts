/**
 * Listen mode keeps the user online: without a presence session Teams shows
 * Offline, whatever preferred presence was set.
 */

import { describe, test } from "node:test";
import assert from "node:assert/strict";
import { createPresenceKeeper, PRESENCE_SESSION_MINUTES } from "../src/watch/presence-keeper.ts";

function recorder(failSet = false) {
	const calls: string[] = [];
	const errors: string[] = [];
	const keeper = createPresenceKeeper({
		set: async (id, duration) => {
			calls.push(`set ${id} ${duration}`);
			if (failSet) throw new Error("403");
		},
		clear: async (id) => {
			calls.push(`clear ${id}`);
		},
		onError: (message) => errors.push(message),
	});
	return { keeper, calls, errors };
}

const tick = () => new Promise((resolve) => setImmediate(resolve));

describe("presence keeper", () => {
	test("start opens a session right away, stop clears it", async () => {
		const { keeper, calls } = recorder();
		keeper.start("me-1");
		await tick();
		assert.deepEqual(calls, [`set me-1 PT${PRESENCE_SESSION_MINUTES}M`]);
		assert.equal(keeper.active, true);
		await keeper.stop();
		assert.deepEqual(calls.at(-1), "clear me-1");
		assert.equal(keeper.active, false);
	});

	test("start is idempotent for the same user", async () => {
		const { keeper, calls } = recorder();
		keeper.start("me-1");
		keeper.start("me-1");
		await tick();
		assert.equal(calls.length, 1);
		await keeper.stop();
	});

	test("stop without start does nothing, and twice is safe", async () => {
		const { keeper, calls } = recorder();
		await keeper.stop();
		keeper.start("me-1");
		await tick();
		await keeper.stop();
		await keeper.stop();
		assert.deepEqual(calls, [`set me-1 PT${PRESENCE_SESSION_MINUTES}M`, "clear me-1"]);
	});

	test("a failing renewal is reported, not thrown", async () => {
		const { keeper, errors } = recorder(true);
		keeper.start("me-1");
		await tick();
		assert.equal(errors.length, 1);
		assert.match(errors[0]!, /presence/);
		await keeper.stop();
	});
});
