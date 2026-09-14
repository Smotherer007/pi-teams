/**
 * teams_create_chat — start a new 1:1 or group chat.
 *
 * Teams reuses an existing 1:1 conversation, so calling this for "message
 * Anna" does not litter the user's chat list with duplicates.
 */

import { Type } from "typebox";
import { createChat, resolveParticipants } from "../graph/chats.ts";
import { sendChatMessage } from "../graph/messages.ts";
import { auditWrite } from "../safety/audit.ts";
import { assertAccess } from "../safety/index.ts";
import { truncate } from "../utils/formatting.ts";
import {
	AccountParam,
	TenantParam,
	connectionFor,
	currentUser,
	errorResult,
	run,
	textResult,
	type ToolContext,
	type ToolResult,
} from "./shared.ts";

interface CreateChatParams {
	participants: string[];
	topic?: string;
	message?: string;
	account?: string;
	tenant?: string;
}

export const teamsCreateChatTool = {
	name: "teams_create_chat",
	description:
		"Start a new Microsoft Teams chat with one or more people, optionally sending the first message right " +
		"away. Participants are given as names, e-mail addresses or user principal names. " +
		"For one participant, Teams reuses the existing 1:1 chat if there is one.",
	parameters: Type.Object({
		participants: Type.Array(Type.String(), {
			description: "People to include, by name, e-mail or UPN",
		}),
		topic: Type.Optional(Type.String({ description: "Group chat topic (group chats only)" })),
		message: Type.Optional(Type.String({ description: "First message to send into the new chat" })),
		account: AccountParam,
		tenant: TenantParam,
	}),
	promptSnippet: "Start a new Teams chat",

	async execute(
		_toolCallId: string,
		params: CreateChatParams,
		signal: AbortSignal | undefined,
		_onUpdate: undefined,
		ctx: ToolContext,
	): Promise<ToolResult> {
		return run(async () => {
			const conn = connectionFor(ctx, params.account, params.tenant);

			if (!params.participants?.length) {
				return errorResult("At least one participant is required.");
			}

			const { resolved, unresolved } = await resolveParticipants(conn, params.participants, signal);
			if (unresolved.length > 0) {
				return errorResult(
					`Could not resolve: ${unresolved.join(", ")}. Use teams_find_user to look them up.`,
				);
			}

			for (const person of resolved) {
				assertAccess(conn, "write", "people", person.displayName, [
					person.displayName,
					person.upn,
					person.mail,
					person.id,
				]);
			}

			const chat = await createChat(conn, resolved, { topic: params.topic, signal });
			const me = await currentUser(conn, signal).catch(() => undefined);
			const label = params.topic ?? resolved.map((p) => p.displayName).join(", ");

			auditWrite(conn.audit, {
				tool: "teams_create_chat",
				account: conn.account,
				tenant: conn.tenant,
				actor: me?.upn,
				target: `chat:${label}`,
				summary: `created chat with ${resolved.map((p) => p.upn ?? p.displayName).join(", ")}`,
			});

			const lines = [`✅ Chat with **${label}** ready.`, "", `chatId: ${chat.id}`];

			if (params.message?.trim()) {
				const sent = await sendChatMessage(conn, chat.id, { body: params.message }, signal);
				auditWrite(conn.audit, {
					tool: "teams_create_chat",
					account: conn.account,
					tenant: conn.tenant,
					actor: me?.upn,
					target: `chat:${label}`,
					summary: truncate(params.message, 200),
				});
				lines.push("", `Sent as ${me?.displayName ?? "you"}:`, "", `> ${truncate(params.message, 300)}`, "", `messageId: ${sent.id}`);
			}

			return textResult(lines.join("\n"), { chatId: chat.id, participants: resolved.length });
		});
	},
};
