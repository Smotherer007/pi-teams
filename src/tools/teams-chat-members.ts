/**
 * teams_chat_members — bring someone into a group chat, or take them out.
 *
 * Teams has no "invite" for a chat: either the person is a member of the
 * conversation or they cannot see it. That makes adding a real change to who
 * can read the history, which is why it goes through the safety gates and the
 * person is checked against the member rules before Graph is called.
 */

import { Type } from "typebox";
import { addChatMember, listChatMembers, matchChatMember, removeChatMember } from "../graph/chats.ts";
import { resolveUserId } from "../graph/me.ts";
import { auditWrite } from "../safety/audit.ts";
import { assertAccess } from "../safety/index.ts";
import { requireChat } from "./resolve.ts";
import {
	AccountParam,
	TenantParam,
	connectionFor,
	currentUser,
	errorResult,
	hasScope,
	missingScopeError,
	run,
	textResult,
	type ToolContext,
	type ToolResult,
} from "./shared.ts";

export const teamsChatMembersTool = {
	name: "teams_chat_members",
	description:
		"Add a person to an existing Microsoft Teams group chat, or remove one. " +
		"Takes a chat ID, the group chat topic, or a participant's name/e-mail, plus who to add or remove. " +
		"A 1:1 chat cannot be extended — use teams_create_chat with everyone in it to make it a group chat. " +
		"Removing someone additionally needs the ChatMember.ReadWrite scope.",
	parameters: Type.Object({
		chat: Type.String({ description: "Chat ID, group chat topic, or a participant's name/e-mail" }),
		action: Type.Union([Type.Literal("add"), Type.Literal("remove")], {
			description: "'add' to bring someone in, 'remove' to take them out",
		}),
		person: Type.String({
			description: "Who to add or remove, by name, e-mail address or user principal name",
		}),
		account: AccountParam,
		tenant: TenantParam,
	}),
	promptSnippet: "Add or remove a member of a Teams chat",
	promptGuidelines: [
		"Adding a member gives them the chat's history from that point on, so confirm the person with the user when the name is ambiguous.",
	],

	async execute(
		_toolCallId: string,
		params: {
			chat: string;
			action: "add" | "remove";
			person: string;
			account?: string;
			tenant?: string;
		},
		signal: AbortSignal | undefined,
		_onUpdate: undefined,
		ctx: ToolContext,
	): Promise<ToolResult> {
		return run(async () => {
			const conn = connectionFor(ctx, params.account, params.tenant);
			const me = await currentUser(conn, signal);
			const chat = await requireChat(conn, params.chat, "write", signal);

			if (params.action === "add") {
				if (chat.chatType === "oneOnOne") {
					return errorResult(
						`"${chat.label}" is a 1:1 chat and cannot be extended. Use teams_create_chat with ` +
							`everyone who should be in it to start a group chat.`,
					);
				}

				const person = await resolveUserId(conn, params.person, signal);
				if (!person) {
					return errorResult(
						`No person matching "${params.person}". Use teams_find_user to find the exact name or e-mail.`,
					);
				}

				// The people already in the chat were checked by requireChat; the one
				// about to join has to be checked on its own.
				assertAccess(conn, "write", "people", person.displayName, [
					person.displayName,
					person.upn,
					person.mail,
					person.id,
				]);

				await addChatMember(conn, chat.id, person, { signal });

				auditWrite(conn.audit, {
					tool: "teams_chat_members",
					account: conn.account,
					tenant: conn.tenant,
					actor: me.upn,
					target: `chat:${chat.label}`,
					summary: `added ${person.displayName} to the chat`,
				});

				return textResult(`✅ ${person.displayName} added to "${chat.label}".`, {
					chatId: chat.id,
					action: "add",
				});
			}

			// Removal takes the membership id, not the user id — a scope the default
			// configuration leaves out, so the answer is the fix rather than a 403.
			if (!hasScope(conn, "ChatMember.ReadWrite")) {
				return errorResult(
					missingScopeError("ChatMember.ReadWrite", "Removing a member from a chat"),
				);
			}

			const person = await resolveUserId(conn, params.person, signal).catch(() => undefined);
			const members = await listChatMembers(conn, chat.id, { signal });
			const reference = person?.id ?? person?.upn ?? person?.mail ?? params.person;
			const member = matchChatMember(members, reference) ?? matchChatMember(members, params.person);

			if (!member) {
				const known = members.map((entry) => entry.displayName).join(", ") || "(none visible)";
				return errorResult(`"${params.person}" is not a member of "${chat.label}". Members: ${known}.`);
			}

			await removeChatMember(conn, chat.id, member.membershipId, { signal });

			auditWrite(conn.audit, {
				tool: "teams_chat_members",
				account: conn.account,
				tenant: conn.tenant,
				actor: me.upn,
				target: `chat:${chat.label}`,
				summary: `removed ${member.displayName} from the chat`,
			});

			return textResult(`✅ ${member.displayName} removed from "${chat.label}".`, {
				chatId: chat.id,
				action: "remove",
				membershipId: member.membershipId,
			});
		});
	},
};
