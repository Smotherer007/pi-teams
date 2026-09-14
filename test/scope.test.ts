/**
 * Scope rules — the last line of defence before pi speaks as the user, so
 * these cases are deliberately paranoid.
 */

import { test, describe } from "node:test";
import assert from "node:assert/strict";
import {
	checkScope,
	defaultPermissions,
	findMatch,
	formatPermissions,
	matchesPattern,
	resolvePermissions,
	type PermissionBlock,
} from "../src/config/scope.ts";

describe("matchesPattern", () => {
	test("matches exactly, case-insensitively", () => {
		assert.equal(matchesPattern("Engineering", "engineering"), true);
		assert.equal(matchesPattern("Engineering", "Engineering Ops"), false);
	});

	test("* matches any run of characters, including slashes", () => {
		assert.equal(matchesPattern("*", "anything"), true);
		assert.equal(matchesPattern("Engineering/*", "Engineering/General"), true);
		assert.equal(matchesPattern("*/Announcements", "Sales/Announcements"), true);
		assert.equal(matchesPattern("Project *", "Project Alpha"), true);
		assert.equal(matchesPattern("*", "Team/Sub/Channel"), true);
	});

	test("? matches exactly one character", () => {
		assert.equal(matchesPattern("Team-?", "Team-1"), true);
		assert.equal(matchesPattern("Team-?", "Team-12"), false);
	});

	test("an id: prefix is ignored on both sides", () => {
		assert.equal(matchesPattern("id:19:abc@thread.tacv2", "19:abc@thread.tacv2"), true);
	});

	test("regex metacharacters in a name are not treated as regex", () => {
		assert.equal(matchesPattern("R&D (EU)", "R&D (EU)"), true);
		assert.equal(matchesPattern("a.b", "axb"), false);
	});

	test("blank values never match", () => {
		assert.equal(matchesPattern("", "x"), false);
		assert.equal(matchesPattern("*", ""), false);
	});
});

describe("findMatch", () => {
	test("returns the first pattern matching any candidate", () => {
		assert.equal(findMatch(["a", "b*"], ["zzz", "bcd"]), "b*");
		assert.equal(findMatch(["a"], ["zzz"]), undefined);
	});
});

describe("resolvePermissions", () => {
	test("no configuration allows everything", () => {
		const resolved = resolvePermissions([undefined, undefined]);
		assert.deepEqual(resolved.read.channels.allow, ["*"]);
		assert.deepEqual(resolved.write.people.allow, ["*"]);
		assert.deepEqual(resolved.write.chats.deny, []);
	});

	test("deny accumulates across every level", () => {
		const global: PermissionBlock = { channels: { deny: ["*/Announcements"] } };
		const account: PermissionBlock = { channels: { deny: ["HR/*"] } };
		const tenant: PermissionBlock = { write: { channels: { deny: ["Board/*"] } } };

		const resolved = resolvePermissions([global, account, tenant]);
		assert.deepEqual(resolved.write.channels.deny, ["*/Announcements", "HR/*", "Board/*"]);
		// The write-only deny does not leak into read.
		assert.deepEqual(resolved.read.channels.deny, ["*/Announcements", "HR/*"]);
	});

	test("the most specific allow list wins", () => {
		const global: PermissionBlock = { channels: { allow: ["*"] } };
		const account: PermissionBlock = { channels: { allow: ["Engineering/*"] } };

		const resolved = resolvePermissions([global, account]);
		assert.deepEqual(resolved.read.channels.allow, ["Engineering/*"]);
	});

	test("a mode-specific allow overrides the shorthand at the same level", () => {
		const block: PermissionBlock = {
			channels: { allow: ["Engineering/*", "Sales/*"] },
			write: { channels: { allow: ["Engineering/General"] } },
		};
		const resolved = resolvePermissions([block]);
		assert.deepEqual(resolved.read.channels.allow, ["Engineering/*", "Sales/*"]);
		assert.deepEqual(resolved.write.channels.allow, ["Engineering/General"]);
	});

	test("an empty allow list does not silently lock everything out", () => {
		const resolved = resolvePermissions([{ chats: { allow: [] } }]);
		assert.deepEqual(resolved.write.chats.allow, ["*"]);
	});

	test("duplicate deny patterns are not repeated", () => {
		const resolved = resolvePermissions([
			{ chats: { deny: ["ceo@x.com"] } },
			{ chats: { deny: ["ceo@x.com"] } },
		]);
		assert.deepEqual(resolved.write.chats.deny, ["ceo@x.com"]);
	});
});

describe("checkScope", () => {
	const permissions = resolvePermissions([
		{
			channels: { allow: ["Engineering/*", "Sales/General"], deny: ["*/Announcements"] },
			people: { deny: ["ceo@contoso.com"] },
		},
	]);

	test("allows a channel covered by an allow pattern", () => {
		const decision = checkScope(permissions, "write", "channels", ["Engineering/General"]);
		assert.equal(decision.allowed, true);
		assert.equal(decision.pattern, "Engineering/*");
	});

	test("deny beats allow", () => {
		const decision = checkScope(permissions, "write", "channels", ["Engineering/Announcements"]);
		assert.equal(decision.allowed, false);
		assert.match(decision.reason, /deny rule/);
	});

	test("a channel outside every allow pattern is refused", () => {
		const decision = checkScope(permissions, "write", "channels", ["HR/General"]);
		assert.equal(decision.allowed, false);
		assert.match(decision.reason, /not covered/);
	});

	test("any matching candidate is enough", () => {
		const decision = checkScope(permissions, "read", "channels", [
			undefined,
			"Sales/General",
			"19:abc@thread.tacv2",
		]);
		assert.equal(decision.allowed, true);
	});

	test("a denied person is blocked", () => {
		const decision = checkScope(permissions, "write", "people", ["Chris Example", "ceo@contoso.com"]);
		assert.equal(decision.allowed, false);
	});

	test("no identifier at all fails closed", () => {
		const decision = checkScope(permissions, "write", "chats", [undefined, "", "   "]);
		assert.equal(decision.allowed, false);
	});

	test("the permissive default allows anything", () => {
		const decision = checkScope(defaultPermissions(), "write", "chats", ["whatever"]);
		assert.equal(decision.allowed, true);
	});
});

describe("formatPermissions", () => {
	test("renders both modes and all categories", () => {
		const text = formatPermissions(resolvePermissions([{ channels: { deny: ["HR/*"] } }]));
		assert.match(text, /\*\*Read\*\*/);
		assert.match(text, /\*\*Write\*\*/);
		assert.match(text, /deny: HR\/\*/);
	});
});
