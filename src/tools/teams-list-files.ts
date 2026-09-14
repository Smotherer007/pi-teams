/**
 * teams_list_files — files in a channel's SharePoint folder.
 */

import { Type } from "typebox";
import { listChannelFiles } from "../graph/teams.ts";
import { formatFileList } from "../utils/formatting.ts";
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

export const teamsListFilesTool = {
	name: "teams_list_files",
	description:
		"List the files stored in a Microsoft Teams channel (its SharePoint folder). Returns names, sizes, " +
		"last-modified times and web links. Use the link to open or download a file — this tool does not " +
		"download file contents.",
	parameters: Type.Object({
		channel: Type.String({ description: "Channel name or ID, or a 'Team/Channel' path" }),
		team: Type.Optional(Type.String({ description: "Team name or ID (omit when using a path)" })),
		account: AccountParam,
		tenant: TenantParam,
		limit: LimitParam,
	}),
	promptSnippet: "List files in a Teams channel",

	async execute(
		_toolCallId: string,
		params: { channel: string; team?: string; account?: string; tenant?: string; limit?: number },
		signal: AbortSignal | undefined,
		_onUpdate: undefined,
		ctx: ToolContext,
	): Promise<ToolResult> {
		return run(async () => {
			const conn = connectionFor(ctx, params.account, params.tenant);
			const { team, channel } = await requireChannel(conn, params.team, params.channel, "read", signal);

			const files = await listChannelFiles(conn, team.id, channel.id, params.limit ?? 50, signal);

			return textResult(
				formatFileList(files, `Files in ${team.displayName}/${channel.displayName}`),
				{ teamId: team.id, channelId: channel.id, count: files.length },
			);
		});
	},
};
