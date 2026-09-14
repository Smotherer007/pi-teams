/**
 * teams_read_channel — read the top-level posts of a channel.
 */

import { Type } from "typebox";
import { listChannelMessages } from "../graph/messages.ts";
import { formatMessageList } from "../utils/formatting.ts";
import { requireChannel } from "./resolve.ts";
import {
	AccountParam,
	LimitParam,
	TenantParam,
	connectionFor,
	run,
	textResult,
	type ToolContext,
	type ToolResult,
} from "./shared.ts";

export const teamsReadChannelTool = {
	name: "teams_read_channel",
	description:
		"Read the posts of a Microsoft Teams channel (top-level messages, not the replies inside each thread). " +
		"Pass team and channel by name or ID, or a single 'Team/Channel' path as 'channel'. " +
		"Set withReplyCounts to also show how many replies each thread has — use teams_read_thread for the replies themselves.",
	parameters: Type.Object({
		channel: Type.String({ description: "Channel name or ID, or a 'Team/Channel' path" }),
		team: Type.Optional(Type.String({ description: "Team name or ID (omit when using a path)" })),
		account: AccountParam,
		tenant: TenantParam,
		limit: LimitParam,
		withReplyCounts: Type.Optional(
			Type.Boolean({ description: "Resolve reply counts for each post (slower, one call per post)" }),
		),
	}),
	promptSnippet: "Read posts from a Teams channel",
	promptGuidelines: [
		"Channel conversations are threaded: teams_read_channel returns the openers, teams_read_thread the replies.",
	],

	async execute(
		_toolCallId: string,
		params: {
			channel: string;
			team?: string;
			account?: string;
			tenant?: string;
			limit?: number;
			withReplyCounts?: boolean;
		},
		signal: AbortSignal | undefined,
		_onUpdate: undefined,
		ctx: ToolContext,
	): Promise<ToolResult> {
		return run(async () => {
			const conn = connectionFor(ctx, params.account, params.tenant);
			const { team, channel } = await requireChannel(conn, params.team, params.channel, "read", signal);

			const messages = await listChannelMessages(conn, channel, {
				max: params.limit ?? conn.maxMessages,
				withReplies: params.withReplyCounts,
				signal,
			});

			return textResult(
				formatMessageList(messages, `${team.displayName}/${channel.displayName}`),
				{
					teamId: team.id,
					channelId: channel.id,
					count: messages.length,
				},
			);
		});
	},
};
