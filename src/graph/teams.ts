/**
 * Teams, channels, members and channel files.
 *
 * Every lookup accepts a name *or* an ID, because that is how people refer to
 * their own teams ("post this in Engineering/General"). Name resolution is
 * case-insensitive and reports ambiguity rather than guessing — posting in the
 * wrong channel is not a recoverable mistake.
 */

import type { TeamsConnection } from "../config/index.ts";
import type { ChannelSummary, DriveItemSummary, MemberSummary, TeamSummary } from "../types.ts";
import { graphGetOptional, graphList, graphPost, graphRequest } from "./client.ts";
import { mapChannel, mapDriveItem, mapMember, mapTeam } from "./mappers.ts";

/** Raised when a name matches more than one team or channel. */
export class AmbiguousNameError extends Error {
	readonly candidates: string[];
	constructor(kind: string, name: string, candidates: string[]) {
		super(
			`${kind} "${name}" is ambiguous — ${candidates.length} matches: ${candidates.join(", ")}. ` +
				`Use the ID instead.`,
		);
		this.name = "AmbiguousNameError";
		this.candidates = candidates;
	}
}

const isGuid = (value: string): boolean => /^[0-9a-f-]{36}$/i.test(value);
const isChannelId = (value: string): boolean => value.startsWith("19:");

// ---------------------------------------------------------------------------
// Teams
// ---------------------------------------------------------------------------

/** Teams the signed-in user is a member of. */
export async function listJoinedTeams(
	conn: TeamsConnection,
	max = 100,
	signal?: AbortSignal,
): Promise<TeamSummary[]> {
	const raw = await graphList<Record<string, unknown>>(conn, "/me/joinedTeams", { max, signal });
	return raw.map((entry) => mapTeam(entry as Record<string, never>));
}

export async function getTeam(
	conn: TeamsConnection,
	teamId: string,
	signal?: AbortSignal,
): Promise<TeamSummary | undefined> {
	const raw = await graphGetOptional<Record<string, unknown>>(
		conn,
		`/teams/${encodeURIComponent(teamId)}`,
		{ signal },
	);
	return raw ? mapTeam(raw as Record<string, never>) : undefined;
}

/**
 * Resolve a team by ID or display name.
 *
 * @throws {AmbiguousNameError} when a name matches several teams
 */
export async function resolveTeam(
	conn: TeamsConnection,
	reference: string,
	signal?: AbortSignal,
): Promise<TeamSummary | undefined> {
	if (isGuid(reference)) {
		const direct = await getTeam(conn, reference, signal);
		if (direct) return direct;
	}

	const teams = await listJoinedTeams(conn, 200, signal);
	const lower = reference.toLowerCase();

	const exact = teams.filter((t) => t.displayName.toLowerCase() === lower || t.id === reference);
	if (exact.length === 1) return exact[0];
	if (exact.length > 1) {
		throw new AmbiguousNameError("Team", reference, exact.map((t) => `${t.displayName} (${t.id})`));
	}

	const partial = teams.filter((t) => t.displayName.toLowerCase().includes(lower));
	if (partial.length === 1) return partial[0];
	if (partial.length > 1) {
		throw new AmbiguousNameError("Team", reference, partial.map((t) => t.displayName));
	}

	return undefined;
}

// ---------------------------------------------------------------------------
// Channels
// ---------------------------------------------------------------------------

export async function listChannels(
	conn: TeamsConnection,
	team: TeamSummary,
	max = 100,
	signal?: AbortSignal,
): Promise<ChannelSummary[]> {
	const raw = await graphList<Record<string, unknown>>(
		conn,
		`/teams/${encodeURIComponent(team.id)}/channels`,
		{ max, signal },
	);
	return raw.map((entry) =>
		mapChannel(entry as Record<string, never>, { id: team.id, name: team.displayName }),
	);
}

export async function getChannel(
	conn: TeamsConnection,
	team: TeamSummary,
	channelId: string,
	signal?: AbortSignal,
): Promise<ChannelSummary | undefined> {
	const raw = await graphGetOptional<Record<string, unknown>>(
		conn,
		`/teams/${encodeURIComponent(team.id)}/channels/${encodeURIComponent(channelId)}`,
		{ signal },
	);
	return raw
		? mapChannel(raw as Record<string, never>, { id: team.id, name: team.displayName })
		: undefined;
}

/**
 * Resolve a channel by ID or display name inside a team.
 *
 * @throws {AmbiguousNameError} when a name matches several channels
 */
export async function resolveChannel(
	conn: TeamsConnection,
	team: TeamSummary,
	reference: string,
	signal?: AbortSignal,
): Promise<ChannelSummary | undefined> {
	if (isChannelId(reference)) {
		const direct = await getChannel(conn, team, reference, signal);
		if (direct) return direct;
	}

	const channels = await listChannels(conn, team, 200, signal);
	const lower = reference.toLowerCase();

	const exact = channels.filter((c) => c.displayName.toLowerCase() === lower || c.id === reference);
	if (exact.length === 1) return exact[0];
	if (exact.length > 1) {
		throw new AmbiguousNameError("Channel", reference, exact.map((c) => `${c.displayName} (${c.id})`));
	}

	const partial = channels.filter((c) => c.displayName.toLowerCase().includes(lower));
	if (partial.length === 1) return partial[0];
	if (partial.length > 1) {
		throw new AmbiguousNameError("Channel", reference, partial.map((c) => c.displayName));
	}

	return undefined;
}

/**
 * Resolve "Team/Channel", or a team plus channel reference, in one step.
 * Returns whatever it could resolve so callers can produce a precise error.
 */
export async function resolveTeamAndChannel(
	conn: TeamsConnection,
	teamRef: string | undefined,
	channelRef: string,
	signal?: AbortSignal,
): Promise<{ team?: TeamSummary; channel?: ChannelSummary }> {
	let teamReference = teamRef;
	let channelReference = channelRef;

	if (!teamReference && channelRef.includes("/")) {
		const [first, ...rest] = channelRef.split("/");
		teamReference = first;
		channelReference = rest.join("/");
	}

	if (!teamReference) return {};

	const team = await resolveTeam(conn, teamReference, signal);
	if (!team) return {};

	const channel = await resolveChannel(conn, team, channelReference, signal);
	return { team, channel };
}

export async function createChannel(
	conn: TeamsConnection,
	team: TeamSummary,
	input: { displayName: string; description?: string; membershipType?: string },
	signal?: AbortSignal,
): Promise<ChannelSummary> {
	const raw = await graphPost<Record<string, unknown>>(
		conn,
		`/teams/${encodeURIComponent(team.id)}/channels`,
		{
			displayName: input.displayName,
			description: input.description,
			membershipType: input.membershipType ?? "standard",
		},
		{ signal },
	);
	return mapChannel((raw ?? {}) as Record<string, never>, { id: team.id, name: team.displayName });
}

// ---------------------------------------------------------------------------
// Members
// ---------------------------------------------------------------------------

export async function listTeamMembers(
	conn: TeamsConnection,
	teamId: string,
	max = 200,
	signal?: AbortSignal,
): Promise<MemberSummary[]> {
	const raw = await graphList<Record<string, unknown>>(
		conn,
		`/teams/${encodeURIComponent(teamId)}/members`,
		{ max, signal },
	);
	return raw.map((entry) => mapMember(entry as Record<string, never>));
}

export async function listChannelMembers(
	conn: TeamsConnection,
	teamId: string,
	channelId: string,
	max = 200,
	signal?: AbortSignal,
): Promise<MemberSummary[]> {
	const raw = await graphList<Record<string, unknown>>(
		conn,
		`/teams/${encodeURIComponent(teamId)}/channels/${encodeURIComponent(channelId)}/members`,
		{ max, signal },
	);
	return raw.map((entry) => mapMember(entry as Record<string, never>));
}

// ---------------------------------------------------------------------------
// Files
// ---------------------------------------------------------------------------

/** Files in a channel's SharePoint folder. */
export async function listChannelFiles(
	conn: TeamsConnection,
	teamId: string,
	channelId: string,
	max = 50,
	signal?: AbortSignal,
): Promise<DriveItemSummary[]> {
	const folder = await graphRequest<Record<string, unknown>>(
		conn,
		"GET",
		`/teams/${encodeURIComponent(teamId)}/channels/${encodeURIComponent(channelId)}/filesFolder`,
		{ signal },
	);

	const parentReference = folder?.["parentReference"] as { driveId?: string } | undefined;
	const driveId = parentReference?.driveId;
	const itemId = folder?.["id"] as string | undefined;
	if (!driveId || !itemId) return [];

	const raw = await graphList<Record<string, unknown>>(
		conn,
		`/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(itemId)}/children`,
		{ max, signal },
	);
	return raw.map((entry) => mapDriveItem(entry as Record<string, never>));
}
