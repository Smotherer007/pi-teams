/**
 * The session-start notice.
 *
 * It exists for one specific failure: a listener that starts from the saved
 * setting, polls, and answers in the user's name while the transcript says
 * nothing about it. "There was a watch running although I never started it" is
 * not a bug in the watcher — it is a missing sentence.
 */

import { test, describe } from "node:test";
import assert from "node:assert/strict";
import { buildStartupNotice, type ConnectionCard } from "../src/status.ts";

const card = (overrides: Partial<ConnectionCard> = {}): ConnectionCard => ({
	account: "work",
	accountDisplayName: "neoimpulse",
	tenant: "work",
	tenantId: "3c8a65a5-573a-4e62-b56c-748f9c048488",
	authMode: "device-code",
	safetyLevel: "open",
	user: "Patrick Weppelmann",
	signedIn: true,
	otherAccounts: [],
	permissionSummary: "",
	...overrides,
});

describe("buildStartupNotice", () => {
	test("names listen mode when it is on, and how to stop it", () => {
		const notice = buildStartupNotice(card(), { listening: true, intervalSeconds: 60 });
		assert.match(notice.message, /listen mode ON/);
		assert.match(notice.message, /\/teams-listen off/);
	});

	test("says the state plainly when it is off, and how to start it", () => {
		const notice = buildStartupNotice(card(), { listening: false, intervalSeconds: 60 });
		assert.match(notice.message, /listen mode off/);
		assert.match(notice.message, /\/teams-listen on to start/);
	});

	test("the surprising state is the louder one", () => {
		// Being listened to is worth a warning; being off is the ordinary case.
		assert.equal(buildStartupNotice(card(), { listening: true, intervalSeconds: 60 }).level, "warning");
		assert.equal(buildStartupNotice(card(), { listening: false, intervalSeconds: 60 }).level, "info");
	});

	test("keeps saying who pi acts as", () => {
		const notice = buildStartupNotice(card(), { listening: true, intervalSeconds: 60 });
		assert.match(notice.message, /work as Patrick Weppelmann/);
		assert.match(notice.message, /safety: open/);
	});

	test("copes with a signed-in account that has no user name yet", () => {
		const notice = buildStartupNotice(card({ user: undefined }), {
			listening: false,
			intervalSeconds: 60,
		});
		assert.match(notice.message, /loaded \(work, safety: open\)/);
	});

	test("reports the interval it will actually poll at", () => {
		const notice = buildStartupNotice(card(), { listening: true, intervalSeconds: 15 });
		assert.match(notice.message, /every 15 s/);
	});
});
