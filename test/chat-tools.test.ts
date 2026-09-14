/**
 * The chat write tools: editing a message, and changing who is in a chat.
 *
 * No fetch harness exists, so the tests cover what can be wrong without a live
 * tenant: the member matching that decides *who* is removed, and the safety
 * gates that decide whether the call may happen at all.
 */

import { test, describe } from "node:test";
import assert from "node:assert/strict";
import { matchChatMember } from "../src/graph/chats.ts";
import { blockReason, formatMutationSummary, isMutationTool } from "../src/safety/index.ts";
import { hasScope, missingScopeError } from "../src/tools/shared.ts";
import { teamsUpdateMessageTool } from "../src/tools/teams-update-message.ts";
import { teamsChatMembersTool } from "../src/tools/teams-chat-members.ts";
import type { TeamsConnection } from "../src/config/index.ts";
import type { ChatMemberSummary } from "../src/types.ts";

const members: ChatMemberSummary[] = [
	{
		membershipId: "MCMjMiMx",
		userId: "u-anna",
		displayName: "Anna Schmidt",
		upn: "anna.schmidt@contoso.com",
		mail: "anna@contoso.com",
		roles: ["owner"],
	},
	{
		membershipId: "MCMjMiMy",
		userId: "u-tom",
		displayName: "Tom Weber",
		upn: "tom.weber@contoso.com",
		mail: undefined,
		roles: [],
	},
];

describe("matchChatMember", () => {
	test("an object id finds the membership to remove", () => {
		assert.equal(matchChatMember(members, "u-tom")?.membershipId, "MCMjMiMy");
	});

	test("e-mail, UPN and display name all work, case-insensitively", () => {
		assert.equal(matchChatMember(members, "ANNA@contoso.com")?.membershipId, "MCMjMiMx");
		assert.equal(matchChatMember(members, "tom.weber@contoso.com")?.membershipId, "MCMjMiMy");
		assert.equal(matchChatMember(members, "tom weber")?.membershipId, "MCMjMiMy");
	});

	test("a partial name is not a match — removal must not guess", () => {
		assert.equal(matchChatMember(members, "Tom"), undefined);
		assert.equal(matchChatMember(members, "anna.schmidt"), undefined);
	});

	test("an empty reference matches nothing", () => {
		assert.equal(matchChatMember(members, "   "), undefined);
	});
});

describe("consent gates", () => {
	const conn = { scopes: ["Chat.ReadWrite", "Calendars.ReadWrite"] } as TeamsConnection;

	test("a requested scope is recognised regardless of case", () => {
		assert.equal(hasScope(conn, "Chat.ReadWrite"), true);
		assert.equal(hasScope(conn, "chat.readwrite"), true);
	});

	test("a scope the account never asked for is not", () => {
		assert.equal(hasScope(conn, "ChatMember.ReadWrite"), false);
	});

	test("the error names the scope, the feature and the fix", () => {
		const message = missingScopeError("ChatMember.ReadWrite", "Removing a member from a chat");
		assert.match(message, /ChatMember\.ReadWrite/);
		assert.match(message, /Removing a member from a chat/);
		assert.match(message, /pi-teams\.json/);
		assert.match(message, /teams_login/);
	});
});

describe("chat write tools are mutations", () => {
	test("readonly blocks editing and membership changes", () => {
		for (const tool of ["teams_update_message", "teams_chat_members"]) {
			assert.equal(isMutationTool(tool), true);
			assert.match(blockReason("readonly", "device-code", tool) ?? "", /readonly/);
			assert.equal(blockReason("confirm", "interactive", tool), undefined);
		}
	});

	test("app-only tokens cannot rewrite a message or change a chat's members", () => {
		assert.match(
			blockReason("open", "client-credentials", "teams_update_message") ?? "",
			/app-only/,
		);
		assert.match(blockReason("open", "client-credentials", "teams_chat_members") ?? "", /app-only/);
	});

	test("the confirmation dialog shows the new text, not just the tool name", () => {
		const summary = formatMutationSummary("teams_update_message", {
			messageId: "1700",
			chat: "Anna",
			body: "Termin ist 14:00 statt 13:00.",
		});
		assert.match(summary, /1700/);
		assert.match(summary, /14:00 statt 13:00/);
	});

	test("adding and removing read differently in the dialog", () => {
		assert.match(
			formatMutationSummary("teams_chat_members", {
				action: "add",
				chat: "Projekt Alpha",
				person: "Tom Weber",
			}),
			/Add Tom Weber to the chat "Projekt Alpha"/,
		);
		assert.match(
			formatMutationSummary("teams_chat_members", {
				action: "remove",
				chat: "Projekt Alpha",
				person: "Tom Weber",
			}),
			/Remove Tom Weber from the chat/,
		);
	});
});

describe("tool shapes", () => {
	test("teams_update_message needs a body and says which chat", () => {
		assert.equal(teamsUpdateMessageTool.name, "teams_update_message");
		assert.deepEqual(Object.keys(teamsUpdateMessageTool.parameters.properties), [
			"messageId",
			"body",
			"chat",
			"html",
			"account",
			"tenant",
		]);
		// No channel path: editing a channel post needs a scope pi does not request.
		assert.equal("channel" in teamsUpdateMessageTool.parameters.properties, false);
	});

	test("teams_chat_members offers add and remove, and says a 1:1 cannot grow", () => {
		assert.equal(teamsChatMembersTool.name, "teams_chat_members");
		assert.deepEqual(Object.keys(teamsChatMembersTool.parameters.properties), [
			"chat",
			"action",
			"person",
			"account",
			"tenant",
		]);
		assert.match(teamsChatMembersTool.description, /1:1/);
		assert.match(teamsChatMembersTool.description, /ChatMember\.ReadWrite/);
	});
});
