/**
 * Target resolution with scope enforcement.
 *
 * Every tool that touches a specific channel or chat goes through here, so the
 * "resolve the name, then check the rules" pair can never be half-done: the
 * functions below either return an in-scope target or throw.
 */

import type { TeamsConnection } from "../config/index.ts";
import type { ChannelSummary, ChatSummary, TeamSummary } from "../types.ts";
import { chatCandidates, channelCandidates, teamCandidates } from "../graph/mappers.ts";
import { findChats } from "../graph/chats.ts";
import { resolveTeam, resolveTeamAndChannel } from "../graph/teams.ts";
import { assertAccess } from "../safety/index.ts";
import type { ScopeMode } from "../config/scope.ts";
import { currentUser } from "./shared.ts";

/** Raised when a name matches nothing, or too much to be safe. */
export class TargetError extends Error {
	constructor(message: string) {
		super(message);
		this.name = "TargetError";
	}
}

/**
 * Resolve a team by name or ID and check it against the rules.
 *
 * @throws {TargetError} when it does not exist
 * @throws {ScopeDeniedError} when the rules forbid it
 */
export async function requireTeam(
	conn: TeamsConnection,
	reference: string,
	mode: ScopeMode,
	signal?: AbortSignal,
): Promise<TeamSummary> {
	const team = await resolveTeam(conn, reference, signal);
	if (!team) {
		throw new TargetError(
			`Team "${reference}" not found among the teams you are a member of. ` +
				`Use teams_list_teams to see the available names.`,
		);
	}
	assertAccess(conn, mode, "teams", team.displayName, teamCandidates(team));
	return team;
}

/**
 * Resolve a channel (optionally via "Team/Channel") and check both the team
 * and the channel against the rules.
 */
export async function requireChannel(
	conn: TeamsConnection,
	teamRef: string | undefined,
	channelRef: string,
	mode: ScopeMode,
	signal?: AbortSignal,
): Promise<{ team: TeamSummary; channel: ChannelSummary }> {
	if (!teamRef && !channelRef.includes("/")) {
		throw new TargetError(
			`Specify the team as well: pass 'team' explicitly, or use a "Team/Channel" path for 'channel'.`,
		);
	}

	const { team, channel } = await resolveTeamAndChannel(conn, teamRef, channelRef, signal);

	if (!team) {
		throw new TargetError(
			`Team "${teamRef ?? channelRef.split("/")[0]}" not found. Use teams_list_teams to see the available names.`,
		);
	}
	assertAccess(conn, mode, "teams", team.displayName, teamCandidates(team));

	if (!channel) {
		throw new TargetError(
			`Channel "${channelRef}" not found in team "${team.displayName}". ` +
				`Use teams_list_channels to see the available names.`,
		);
	}
	assertAccess(
		conn,
		mode,
		"channels",
		`${team.displayName}/${channel.displayName}`,
		channelCandidates(channel),
	);

	return { team, channel };
}

/**
 * Resolve a chat by ID, topic or participant and check it against the rules.
 *
 * Refuses ambiguous matches: sending "the usual update" to the wrong Anna is
 * not something a confirmation dialog can undo.
 */
export async function requireChat(
	conn: TeamsConnection,
	reference: string,
	mode: ScopeMode,
	signal?: AbortSignal,
): Promise<ChatSummary> {
	const me = await currentUser(conn, signal).catch(() => undefined);
	const matches = await findChats(conn, reference, { meId: me?.id, signal });

	if (matches.length === 0) {
		throw new TargetError(
			`No chat matching "${reference}". Use teams_list_chats to see recent chats, ` +
				`or teams_create_chat to start a new one.`,
		);
	}

	if (matches.length > 1) {
		const options = matches.slice(0, 8).map((chat) => `"${chat.label}" (${chat.id})`);
		throw new TargetError(
			`"${reference}" matches ${matches.length} chats: ${options.join(", ")}. ` +
				`Pass the chat ID to be unambiguous.`,
		);
	}

	const chat = matches[0]!;
	assertAccess(conn, mode, "chats", chat.label, chatCandidates(chat));

	// A chat is also the people in it: a deny rule on a person must stop a
	// message to any chat they are part of, not just their 1:1.
	for (const member of chat.members) {
		if (me && member.id === me.id) continue;
		assertAccess(conn, mode, "people", member.displayName, [
			member.displayName,
			member.upn,
			member.mail,
			member.id,
		]);
	}

	return chat;
}
