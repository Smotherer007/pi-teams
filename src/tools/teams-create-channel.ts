/**
 * teams_create_channel — add a channel to a team.
 */

import { Type } from "typebox";
import { createChannel } from "../graph/teams.ts";
import { auditWrite } from "../safety/audit.ts";
import { requireTeam } from "./resolve.ts";
import {
	AccountParam,
	TenantParam,
	connectionFor,
	currentUser,
	run,
	textResult,
	type ToolContext,
	type ToolResult,
} from "./shared.ts";

export const teamsCreateChannelTool = {
	name: "teams_create_channel",
	description:
		"Create a new channel in a Microsoft Teams team. membershipType 'standard' (default) is visible to the " +
		"whole team, 'private' only to the members you add afterwards.",
	parameters: Type.Object({
		team: Type.String({ description: "Team name or ID" }),
		name: Type.String({ description: "Channel display name" }),
		description: Type.Optional(Type.String({ description: "Channel description" })),
		membershipType: Type.Optional(
			Type.String({ description: "'standard' (default) or 'private'" }),
		),
		account: AccountParam,
		tenant: TenantParam,
	}),
	promptSnippet: "Create a channel in a Teams team",

	async execute(
		_toolCallId: string,
		params: {
			team: string;
			name: string;
			description?: string;
			membershipType?: string;
			account?: string;
			tenant?: string;
		},
		signal: AbortSignal | undefined,
		_onUpdate: undefined,
		ctx: ToolContext,
	): Promise<ToolResult> {
		return run(async () => {
			const conn = connectionFor(ctx, params.account, params.tenant);
			const team = await requireTeam(conn, params.team, "write", signal);

			const channel = await createChannel(
				conn,
				team,
				{
					displayName: params.name,
					description: params.description,
					membershipType: params.membershipType,
				},
				signal,
			);

			const me = await currentUser(conn, signal).catch(() => undefined);
			auditWrite(conn.audit, {
				tool: "teams_create_channel",
				account: conn.account,
				tenant: conn.tenant,
				actor: me?.upn,
				target: `team:${team.displayName}`,
				summary: `created channel "${params.name}" (${params.membershipType ?? "standard"})`,
			});

			return textResult(
				[
					`✅ Channel **${team.displayName}/${channel.displayName}** created.`,
					"",
					`channelId: ${channel.id}`,
					channel.webUrl ? `Open in Teams: ${channel.webUrl}` : "",
				]
					.filter(Boolean)
					.join("\n"),
				{ teamId: team.id, channelId: channel.id },
			);
		});
	},
};
