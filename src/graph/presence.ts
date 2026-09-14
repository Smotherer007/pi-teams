/**
 * Presence — reading who is available and setting the user's own status.
 *
 * Setting presence uses `setUserPreferredPresence`, not `setPresence`: the
 * latter is tied to an application session and gets overwritten by the Teams
 * client within seconds. Preferred presence is what the user themselves sets
 * in the app, so it behaves the way they expect — including the expiry.
 */

import type { TeamsConnection } from "../config/index.ts";
import type { PresenceInfo } from "../types.ts";
import { graphList, graphPost, graphRequest } from "./client.ts";
import { mapPresence } from "./mappers.ts";

/* eslint-disable @typescript-eslint/no-explicit-any */
type Raw = Record<string, any>;

/** Availability values Graph accepts for preferred presence. */
export const AVAILABILITY_VALUES = [
	"Available",
	"Busy",
	"DoNotDisturb",
	"BeRightBack",
	"Away",
	"Offline",
] as const;

export type Availability = (typeof AVAILABILITY_VALUES)[number];

/** Graph pairs each availability with a specific activity. */
const ACTIVITY_FOR: Record<Availability, string> = {
	Available: "Available",
	Busy: "Busy",
	DoNotDisturb: "DoNotDisturb",
	BeRightBack: "BeRightBack",
	Away: "Away",
	Offline: "OffWork",
};

export function activityFor(availability: Availability): string {
	return ACTIVITY_FOR[availability];
}

export async function getMyPresence(
	conn: TeamsConnection,
	signal?: AbortSignal,
): Promise<PresenceInfo> {
	const raw = await graphRequest<Raw>(conn, "GET", "/me/presence", { signal });
	return mapPresence(raw ?? {}, "you");
}

export async function getPresenceForUser(
	conn: TeamsConnection,
	userId: string,
	displayName?: string,
	signal?: AbortSignal,
): Promise<PresenceInfo> {
	const raw = await graphRequest<Raw>(
		conn,
		"GET",
		`/users/${encodeURIComponent(userId)}/presence`,
		{ signal },
	);
	return mapPresence(raw ?? {}, displayName);
}

/**
 * Presence for several people at once.
 *
 * Graph caps this action at 650 IDs per call; batching in 100s keeps request
 * bodies small and the throttling risk low.
 */
export async function getPresenceForUsers(
	conn: TeamsConnection,
	userIds: string[],
	names: Map<string, string>,
	signal?: AbortSignal,
): Promise<PresenceInfo[]> {
	const results: PresenceInfo[] = [];

	for (let i = 0; i < userIds.length; i += 100) {
		const batch = userIds.slice(i, i + 100);
		const response = await graphPost<Raw>(
			conn,
			"/communications/getPresencesByUserId",
			{ ids: batch },
			{ signal },
		);
		for (const entry of response?.value ?? []) {
			results.push(mapPresence(entry, names.get(entry.id)));
		}
	}

	return results;
}

/**
 * Set the signed-in user's preferred presence.
 *
 * @param expirationDuration - ISO 8601 duration, e.g. "PT2H". Omit to let
 * Microsoft apply its default (1 day for Busy/DoNotDisturb, 7 days otherwise).
 */
export async function setPreferredPresence(
	conn: TeamsConnection,
	availability: Availability,
	options: { expirationDuration?: string; signal?: AbortSignal } = {},
): Promise<void> {
	const body: Record<string, unknown> = {
		availability,
		activity: activityFor(availability),
	};
	if (options.expirationDuration) body.expirationDuration = options.expirationDuration;

	await graphPost(conn, "/me/presence/setUserPreferredPresence", body, { signal: options.signal });
}

/** Hand presence back to Teams' automatic calculation. */
export async function clearPreferredPresence(
	conn: TeamsConnection,
	signal?: AbortSignal,
): Promise<void> {
	await graphPost(conn, "/me/presence/clearUserPreferredPresence", undefined, { signal });
}

/** Set or clear the status message shown under the user's name. */
export async function setStatusMessage(
	conn: TeamsConnection,
	message: string,
	options: { expiresAt?: string; signal?: AbortSignal } = {},
): Promise<void> {
	const statusMessage: Record<string, unknown> = {
		message: { content: message, contentType: "text" },
	};
	if (options.expiresAt) {
		statusMessage.expiryDateTime = { dateTime: options.expiresAt, timeZone: "UTC" };
	}
	await graphPost(conn, "/me/presence/setStatusMessage", { statusMessage }, { signal: options.signal });
}

/** Presence of everyone in a chat — handy before pinging a group. */
export async function getChatPresence(
	conn: TeamsConnection,
	chatId: string,
	signal?: AbortSignal,
): Promise<PresenceInfo[]> {
	const members = await graphList<Raw>(conn, `/chats/${encodeURIComponent(chatId)}/members`, {
		max: 50,
		signal,
	});
	const names = new Map<string, string>();
	const ids: string[] = [];
	for (const member of members) {
		if (member.userId) {
			ids.push(member.userId);
			names.set(member.userId, member.displayName ?? "(unknown)");
		}
	}
	if (ids.length === 0) return [];
	return getPresenceForUsers(conn, ids, names, signal);
}
