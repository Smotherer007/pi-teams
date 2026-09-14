/**
 * teams_delete_message — soft-delete one of the user's own messages.
 *
 * Graph only permits this for the sender, which is also the only case that
 * makes sense here: pi may take back something it said in the user's name, not
 * edit anyone else's history.
 */

import { Type } from "typebox";
import { softDeleteMessage, undoSoftDeleteMessage } from "../graph/messages.ts";
import { auditWrite } from "../safety/audit.ts";
import { requireChannel, requireChat } from "./resolve.ts";
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
import type { MessageLocation } from "../types.ts";

export const teamsDeleteMessageTool = {
	name: "teams_delete_message",
	description:
		"Delete a message the signed-in user sent in Microsoft Teams (soft delete, exactly like deleting it in " +
		"the app). Specify 'chat' or 'channel' (plus 'team') to say where it lives. " +
		"Set undo: true to restore a message that was soft-deleted. Only the sender's own messages can be deleted.",
	parameters: Type.Object({
		messageId: Type.String({ description: "ID of the message to delete" }),
		chat: Type.Optional(Type.String({ description: "Chat ID or name, when the message is in a chat" })),
		channel: Type.Optional(
			Type.String({ description: "Channel name/ID or 'Team/Channel' path, when it is in a channel" }),
		),
		team: Type.Optional(Type.String({ description: "Team name or ID" })),
		replyId: Type.Optional(
			Type.String({ description: "Reply ID, when deleting a reply inside a channel thread" }),
		),
		undo: Type.Optional(Type.Boolean({ description: "Restore a previously deleted message" })),
		account: AccountParam,
		tenant: TenantParam,
	}),
	promptSnippet: "Delete a message the user sent in Teams",

	async execute(
		_toolCallId: string,
		params: {
			messageId: string;
			chat?: string;
			channel?: string;
			team?: string;
			replyId?: string;
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

			let location: MessageLocation;
			let target: string;

			if (params.chat) {
				const chat = await requireChat(conn, params.chat, "write", signal);
				location = { kind: "chat", chatId: chat.id };
				target = `chat:${chat.label}`;
			} else if (params.channel) {
				const { team, channel } = await requireChannel(conn, params.team, params.channel, "write", signal);
				location = {
					kind: "channel",
					teamId: team.id,
					teamName: team.displayName,
					channelId: channel.id,
					channelName: channel.displayName,
				};
				target = `channel:${team.displayName}/${channel.displayName}`;
			} else {
				return errorResult("Specify either 'chat' or 'channel' so pi knows where the message lives.");
			}

			// Graph needs an explicit user id for the chat form of this action, and
			// only the sender may delete — so failing here is better than a 403.
			const me = await currentUser(conn, signal);

			const options = { replyId: params.replyId, signal };
			if (params.undo) {
				await undoSoftDeleteMessage(conn, location, params.messageId, me.id, options);
			} else {
				await softDeleteMessage(conn, location, params.messageId, me.id, options);
			}

			auditWrite(conn.audit, {
				tool: "teams_delete_message",
				account: conn.account,
				tenant: conn.tenant,
				actor: me.upn,
				target,
				summary: `${params.undo ? "restored" : "deleted"} message ${params.messageId}`,
			});

			return textResult(
				`✅ Message ${params.messageId} ${params.undo ? "restored" : "deleted"} in ${target}.`,
				{ messageId: params.messageId, undone: !!params.undo },
			);
		});
	},
};
