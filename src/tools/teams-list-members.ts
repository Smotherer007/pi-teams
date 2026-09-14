/**
 * teams_list_members — members of a team or a channel.
 */

import { Type } from "typebox";
import { listChannelMembers, listTeamMembers } from "../graph/teams.ts";
import { hasAccess } from "../safety/index.ts";
import { formatMemberList } from "../utils/formatting.ts";
import { requireChannel, requireTeam } from "./resolve.ts";
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

export const teamsListMembersTool = {
	name: "teams_list_members",
	description:
		"List the members of a Microsoft Teams team, or of one channel when 'channel' is given. " +
		"Returns display names, e-mail addresses and roles — useful before @-mentioning someone or " +
		"starting a chat with the right person.",
	parameters: Type.Object({
		team: Type.String({ description: "Team display name or ID" }),
		channel: Type.Optional(
			Type.String({ description: "Channel name or ID; omit to list the whole team" }),
		),
		account: AccountParam,
		tenant: TenantParam,
		limit: LimitParam,
	}),
	promptSnippet: "List members of a Teams team or channel",

	async execute(
		_toolCallId: string,
		params: { team: string; channel?: string; account?: string; tenant?: string; limit?: number },
		signal: AbortSignal | undefined,
		_onUpdate: undefined,
		ctx: ToolContext,
	): Promise<ToolResult> {
		return run(async () => {
			const conn = connectionFor(ctx, params.account, params.tenant);

			if (params.channel) {
				const { team, channel } = await requireChannel(conn, params.team, params.channel, "read", signal);
				const members = await listChannelMembers(conn, team.id, channel.id, params.limit ?? 200, signal);
				const visible = members.filter((member) =>
					hasAccess(conn, "read", "people", [member.displayName, member.upn, member.mail, member.id]),
				);
				return textResult(
					formatMemberList(visible, `${team.displayName}/${channel.displayName}`),
					{ team: team.displayName, channel: channel.displayName, count: visible.length },
				);
			}

			const team = await requireTeam(conn, params.team, "read", signal);
			const members = await listTeamMembers(conn, team.id, params.limit ?? 200, signal);
			const visible = members.filter((member) =>
				hasAccess(conn, "read", "people", [member.displayName, member.upn, member.mail, member.id]),
			);

			return textResult(formatMemberList(visible, team.displayName), {
				team: team.displayName,
				count: visible.length,
			});
		});
	},
};
