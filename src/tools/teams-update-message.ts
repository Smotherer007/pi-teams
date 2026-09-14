/**
 * teams_update_message — rewrite a message the user already sent.
 *
 * Editing beats delete-and-resend: the message keeps its ID and its place in
 * the conversation, and Teams shows it as edited rather than as a gap followed
 * by a near-duplicate. Only the sender's own messages can be edited, which is
 * the only case that makes sense for something pi wrote in the user's name.
 *
 * Chats only. Editing a channel post needs the `ChannelMessage.ReadWrite`
 * scope, which this package does not request and does not assume will be
 * granted — so the tool does not offer a path that could not run.
 */

import { Type } from "typebox";
import { updateMessage } from "../graph/messages.ts";
import { auditWrite } from "../safety/audit.ts";
import { requireChat } from "./resolve.ts";
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

export const teamsUpdateMessageTool = {
	name: "teams_update_message",
	description:
		"Edit a message the signed-in user already sent in a Microsoft Teams chat: replaces the text and keeps " +
		"the message in place, so Teams shows it as edited instead of as a new message. " +
		"Use this instead of deleting a message and sending a corrected copy.",
	parameters: Type.Object({
		messageId: Type.String({ description: "ID of the message to edit" }),
		body: Type.String({ description: "The new message text (lightweight markdown by default)" }),
		chat: Type.String({ description: "Chat ID, group chat topic, or a participant's name/e-mail" }),
		html: Type.Optional(Type.Boolean({ description: "Treat 'body' as raw HTML instead of markdown" })),
		account: AccountParam,
		tenant: TenantParam,
	}),
	promptSnippet: "Edit a message the user sent in Teams",
	promptGuidelines: [
		"Use teams_update_message to correct a message the user already sent, rather than deleting it and posting a replacement.",
	],

	async execute(
		_toolCallId: string,
		params: {
			messageId: string;
			body: string;
			chat: string;
			html?: boolean;
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

			await updateMessage(conn, chat.id, params.messageId, me.id, {
				body: params.body,
				html: params.html,
			}, { signal });

			auditWrite(conn.audit, {
				tool: "teams_update_message",
				account: conn.account,
				tenant: conn.tenant,
				actor: me.upn,
				target: `chat:${chat.label}`,
				summary: `edited message ${params.messageId}: ${params.body.slice(0, 120)}`,
			});

			return textResult(`✅ Message ${params.messageId} edited in "${chat.label}".`, {
				messageId: params.messageId,
				chatId: chat.id,
			});
		});
	},
};
