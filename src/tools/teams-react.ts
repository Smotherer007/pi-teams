/**
 * teams_react — add or remove an emoji reaction on a message.
 */

import { Type } from "typebox";
import { setReaction, unsetReaction } from "../graph/messages.ts";
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

interface ReactParams {
	messageId: string;
	reaction: string;
	chat?: string;
	channel?: string;
	team?: string;
	replyId?: string;
	remove?: boolean;
	account?: string;
	tenant?: string;
}

export const teamsReactTool = {
	name: "teams_react",
	description:
		"Add or remove an emoji reaction on a Microsoft Teams message, as the signed-in user. " +
		"Specify either 'chat' or 'channel' (plus 'team') to say where the message lives. " +
		"For a reply inside a channel thread, pass the thread's messageId and the reply's ID as 'replyId'. " +
		"The reaction is a Unicode emoji such as 👍 or ✅.",
	parameters: Type.Object({
		messageId: Type.String({ description: "ID of the message to react to" }),
		reaction: Type.String({ description: "Unicode emoji, e.g. 👍, ❤️, ✅" }),
		chat: Type.Optional(Type.String({ description: "Chat ID or name, when the message is in a chat" })),
		channel: Type.Optional(
			Type.String({ description: "Channel name/ID or 'Team/Channel' path, when it is in a channel" }),
		),
		team: Type.Optional(Type.String({ description: "Team name or ID" })),
		replyId: Type.Optional(Type.String({ description: "Reply ID, when reacting to a reply in a thread" })),
		remove: Type.Optional(Type.Boolean({ description: "Remove the reaction instead of adding it" })),
		account: AccountParam,
		tenant: TenantParam,
	}),
	promptSnippet: "React to a Teams message",

	async execute(
		_toolCallId: string,
		params: ReactParams,
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

			const action = params.remove ? unsetReaction : setReaction;
			await action(conn, location, params.messageId, params.reaction, {
				replyId: params.replyId,
				signal,
			});

			const me = await currentUser(conn, signal).catch(() => undefined);
			auditWrite(conn.audit, {
				tool: "teams_react",
				account: conn.account,
				tenant: conn.tenant,
				actor: me?.upn,
				target,
				summary: `${params.remove ? "removed" : "added"} reaction ${params.reaction} on ${params.messageId}`,
			});

			return textResult(
				`✅ ${params.remove ? "Removed" : "Added"} ${params.reaction} on message ${params.messageId} (${target}).`,
				{ messageId: params.messageId, reaction: params.reaction, removed: !!params.remove },
			);
		});
	},
};
