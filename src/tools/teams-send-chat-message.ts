/**
 * teams_send_chat_message — send a message in a chat, as the user.
 *
 * This is the sharpest tool in the box: what it writes is indistinguishable
 * from the user typing it. Hence the order of operations — resolve the chat,
 * check the chat rules, check every participant against the people rules,
 * resolve the mentions, send, then record the audit entry.
 */

import { Type } from "typebox";
import { sendChatMessage } from "../graph/messages.ts";
import { resolveUserId } from "../graph/me.ts";
import { auditWrite } from "../safety/audit.ts";
import { assertAccess } from "../safety/index.ts";
import { truncate } from "../utils/formatting.ts";
import { requireChat } from "./resolve.ts";
import {
	AccountParam,
	TenantParam,
	connectionFor,
	connectionLabel,
	currentUser,
	errorResult,
	run,
	textResult,
	type ToolContext,
	type ToolResult,
} from "./shared.ts";

interface SendChatParams {
	chat: string;
	body: string;
	account?: string;
	tenant?: string;
	mentions?: string[];
	importance?: string;
	html?: boolean;
}

export const teamsSendChatMessageTool = {
	name: "teams_send_chat_message",
	description:
		"Send a message in a Microsoft Teams chat as the signed-in user. The message appears exactly as if the " +
		"user typed it. Accepts a chat ID, group chat topic, or participant name/e-mail. " +
		"Use 'mentions' with names or e-mail addresses to @-mention people — write '@Display Name' in the body " +
		"where the mention should appear.",
	parameters: Type.Object({
		chat: Type.String({ description: "Chat ID, group chat topic, or a participant's name/e-mail" }),
		body: Type.String({ description: "Message text. Plain text unless 'html' is true." }),
		account: AccountParam,
		tenant: TenantParam,
		mentions: Type.Optional(
			Type.Array(Type.String(), {
				description: "People to @-mention, by name, UPN or e-mail",
			}),
		),
		importance: Type.Optional(
			Type.String({ description: "'normal' (default), 'high' or 'urgent'" }),
		),
		html: Type.Optional(
			Type.Boolean({ description: "Treat 'body' as raw HTML instead of plain text" }),
		),
	}),
	promptSnippet: "Send a message in a Teams chat as the user",
	promptGuidelines: [
		"Write the message in the user's voice — it is sent from their account, not from an assistant.",
		"Show the user the exact text before sending when the intent is even slightly ambiguous.",
		"Do not add signatures, disclaimers, or a note that the message was written by an AI unless the user asks.",
	],

	async execute(
		_toolCallId: string,
		params: SendChatParams,
		signal: AbortSignal | undefined,
		_onUpdate: undefined,
		ctx: ToolContext,
	): Promise<ToolResult> {
		return run(async () => {
			const conn = connectionFor(ctx, params.account, params.tenant);

			if (!params.body.trim()) return errorResult("Refusing to send an empty message.");

			const chat = await requireChat(conn, params.chat, "write", signal);

			// Mentions address people directly, so they are checked individually
			// even when the chat itself is allowed.
			const mentions = [];
			for (const reference of params.mentions ?? []) {
				const person = await resolveUserId(conn, reference, signal);
				if (!person) {
					return errorResult(
						`Could not resolve "${reference}" to a person. Use teams_find_user to look them up.`,
					);
				}
				assertAccess(conn, "write", "people", person.displayName, [
					person.displayName,
					person.upn,
					person.mail,
					person.id,
				]);
				mentions.push(person);
			}

			const me = await currentUser(conn, signal).catch(() => undefined);

			try {
				const sent = await sendChatMessage(
					conn,
					chat.id,
					{
						body: params.body,
						html: params.html,
						importance: params.importance,
						mentions,
					},
					signal,
				);

				auditWrite(conn.audit, {
					tool: "teams_send_chat_message",
					account: conn.account,
					tenant: conn.tenant,
					actor: me?.upn,
					target: `chat:${chat.label}`,
					summary: truncate(params.body, 200),
				});

				return textResult(
					[
						`✅ Message sent to **${chat.label}** as ${me?.displayName ?? "you"} (${connectionLabel(conn)}).`,
						"",
						`> ${truncate(params.body, 300)}`,
						"",
						`messageId: ${sent.id}`,
					].join("\n"),
					{ chatId: chat.id, messageId: sent.id },
				);
			} catch (err) {
				auditWrite(conn.audit, {
					tool: "teams_send_chat_message",
					account: conn.account,
					tenant: conn.tenant,
					actor: me?.upn,
					target: `chat:${chat.label}`,
					summary: truncate(params.body, 200),
					error: err instanceof Error ? err.message : String(err),
				});
				throw err;
			}
		});
	},
};
