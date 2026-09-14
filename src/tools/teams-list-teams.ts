/**
 * teams_list_teams — the teams the signed-in user belongs to, filtered by the
 * read rules.
 */

import { Type } from "typebox";
import { listJoinedTeams } from "../graph/teams.ts";
import { teamCandidates } from "../graph/mappers.ts";
import { hasAccess } from "../safety/index.ts";
import { formatTeamList } from "../utils/formatting.ts";
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

export const teamsListTeamsTool = {
	name: "teams_list_teams",
	description:
		"List the Microsoft Teams teams the signed-in user is a member of. Teams excluded by the configured " +
		"read rules are not shown. Use this to discover team names and IDs before reading or posting in a channel.",
	parameters: Type.Object({
		account: AccountParam,
		tenant: TenantParam,
		limit: LimitParam,
		filter: Type.Optional(
			Type.String({ description: "Only show teams whose name contains this text" }),
		),
	}),
	promptSnippet: "List the user's Microsoft Teams teams",

	async execute(
		_toolCallId: string,
		params: { account?: string; tenant?: string; limit?: number; filter?: string },
		signal: AbortSignal | undefined,
		_onUpdate: undefined,
		ctx: ToolContext,
	): Promise<ToolResult> {
		return run(async () => {
			const conn = connectionFor(ctx, params.account, params.tenant);
			const all = await listJoinedTeams(conn, params.limit ?? 100, signal);

			const visible = all.filter((team) => hasAccess(conn, "read", "teams", teamCandidates(team)));
			const filtered = params.filter
				? visible.filter((team) =>
						team.displayName.toLowerCase().includes(params.filter!.toLowerCase()),
					)
				: visible;

			const hidden = all.length - visible.length;
			const note = hidden > 0 ? `\n\n_${hidden} team(s) hidden by your scope rules._` : "";

			return textResult(`${formatTeamList(filtered)}${note}`, {
				count: filtered.length,
				hidden,
				account: conn.account,
			});
		});
	},
};
