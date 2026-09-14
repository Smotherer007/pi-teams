/**
 * teams_read_thread — the opening post and all replies of one channel thread.
 */

import { Type } from "typebox";
import { getChannelMessage, listChannelReplies } from "../graph/messages.ts";
import { formatMessage } from "../utils/formatting.ts";
import { requireChannel } from "./resolve.ts";
import {
	AccountParam,
	LimitParam,
	TenantParam,
	connectionFor,
	errorResult,
	run,
	textResult,
	type ToolContext,
	type ToolResult,
} from "./shared.ts";

export const teamsReadThreadTool = {
	name: "teams_read_thread",
	description:
		"Read one conversation thread in a Microsoft Teams channel: the opening post plus every reply, " +
		"oldest first. Get the messageId from teams_read_channel or teams_search_messages.",
	parameters: Type.Object({
		channel: Type.String({ description: "Channel name or ID, or a 'Team/Channel' path" }),
		messageId: Type.String({ description: "ID of the thread's opening message" }),
		team: Type.Optional(Type.String({ description: "Team name or ID (omit when using a path)" })),
		account: AccountParam,
		tenant: TenantParam,
		limit: LimitParam,
	}),
	promptSnippet: "Read a Teams channel thread with all replies",

	async execute(
		_toolCallId: string,
		params: {
			channel: string;
			messageId: string;
			team?: string;
			account?: string;
			tenant?: string;
			limit?: number;
		},
		signal: AbortSignal | undefined,
		_onUpdate: undefined,
		ctx: ToolContext,
	): Promise<ToolResult> {
		return run(async () => {
			const conn = connectionFor(ctx, params.account, params.tenant);
			const { team, channel } = await requireChannel(conn, params.team, params.channel, "read", signal);

			const root = await getChannelMessage(conn, channel, params.messageId, signal);
			if (!root) {
				return errorResult(
					`Message ${params.messageId} not found in ${team.displayName}/${channel.displayName}.`,
				);
			}

			const replies = await listChannelReplies(conn, channel, params.messageId, {
				max: params.limit ?? 50,
				signal,
			});

			const lines = [
				`## Thread in ${team.displayName}/${channel.displayName}`,
				"",
				formatMessage(root),
				"",
				`### ${replies.length} repl${replies.length === 1 ? "y" : "ies"}`,
				"",
			];

			// Graph returns replies newest first; a conversation reads the other way.
			for (const reply of [...replies].reverse()) {
				lines.push(formatMessage(reply), "");
			}

			return textResult(lines.join("\n").trimEnd(), {
				teamId: team.id,
				channelId: channel.id,
				messageId: params.messageId,
				replyCount: replies.length,
			});
		});
	},
};
