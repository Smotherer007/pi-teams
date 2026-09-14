/**
 * teams_delete_message — take back a message the user sent in a chat.
 *
 * Graph only permits this for the sender, which is also the only case that
 * makes sense here: pi may take back something it said in the user's name, not
 * edit anyone else's history.
 *
 * Chats only. Deleting a channel post needs the `ChannelMessage.ReadWrite`
 * scope, which this package does not request — a channel post is deleted in
 * Teams, or edited into what it should have said.
 */

import { Type } from "typebox";
import { softDeleteMessage, undoSoftDeleteMessage } from "../graph/messages.ts";
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

export const teamsDeleteMessageTool = {
	name: "teams_delete_message",
	description:
		"Delete a message the signed-in user sent in a Microsoft Teams chat (soft delete, exactly like deleting " +
		"it in the app). Set undo: true to restore a message that was soft-deleted. " +
		"Only the sender's own messages can be deleted, and only in chats — a channel post cannot be deleted " +
		"through Graph with the permissions pi asks for.",
	parameters: Type.Object({
		messageId: Type.String({ description: "ID of the message to delete" }),
		chat: Type.String({ description: "Chat ID, group chat topic, or a participant's name/e-mail" }),
		undo: Type.Optional(Type.Boolean({ description: "Restore a previously deleted message" })),
		account: AccountParam,
		tenant: TenantParam,
	}),
	promptSnippet: "Delete a message the user sent in Teams",
	promptGuidelines: [
		"Prefer teams_update_message over deleting and resending: an edit keeps the message in place, a delete leaves a gap and the replacement arrives as a new notification.",
	],

	async execute(
		_toolCallId: string,
		params: {
			messageId: string;
			chat: string;
			undo?: boolean;
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

			if (params.undo) {
				await undoSoftDeleteMessage(conn, chat.id, params.messageId, me.id, { signal });
			} else {
				await softDeleteMessage(conn, chat.id, params.messageId, me.id, { signal });
			}

			auditWrite(conn.audit, {
				tool: "teams_delete_message",
				account: conn.account,
				tenant: conn.tenant,
				actor: me.upn,
				target: `chat:${chat.label}`,
				summary: `${params.undo ? "restored" : "deleted"} message ${params.messageId}`,
			});

			return textResult(
				`✅ Message ${params.messageId} ${params.undo ? "restored" : "deleted"} in "${chat.label}".`,
				{ messageId: params.messageId, chatId: chat.id, undone: !!params.undo },
			);
		});
	},
};
