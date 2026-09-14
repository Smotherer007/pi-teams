/**
 * teams_list_channels — channels of a team, filtered by the read rules.
 */

import { Type } from "typebox";
import { listChannels } from "../graph/teams.ts";
import { channelCandidates } from "../graph/mappers.ts";
import { hasAccess } from "../safety/index.ts";
import { formatChannelList } from "../utils/formatting.ts";
import { requireTeam } from "./resolve.ts";
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

export const teamsListChannelsTool = {
	name: "teams_list_channels",
	description:
		"List the channels of a Microsoft Teams team. Accepts the team's display name or ID. " +
		"Channels excluded by the configured read rules are not shown.",
	parameters: Type.Object({
		team: Type.String({ description: "Team display name or ID" }),
		account: AccountParam,
		tenant: TenantParam,
		limit: LimitParam,
	}),
	promptSnippet: "List the channels of a Teams team",

	async execute(
		_toolCallId: string,
		params: { team: string; account?: string; tenant?: string; limit?: number },
		signal: AbortSignal | undefined,
		_onUpdate: undefined,
		ctx: ToolContext,
	): Promise<ToolResult> {
		return run(async () => {
			const conn = connectionFor(ctx, params.account, params.tenant);
			const team = await requireTeam(conn, params.team, "read", signal);

			const all = await listChannels(conn, team, params.limit ?? 100, signal);
			const visible = all.filter((channel) =>
				hasAccess(conn, "read", "channels", channelCandidates(channel)),
			);

			const hidden = all.length - visible.length;
			const note = hidden > 0 ? `\n\n_${hidden} channel(s) hidden by your scope rules._` : "";

			return textResult(`${formatChannelList(visible)}${note}`, {
				team: team.displayName,
				teamId: team.id,
				count: visible.length,
				hidden,
			});
		});
	},
};
