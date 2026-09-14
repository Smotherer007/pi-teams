/**
 * 1:1, group and meeting chats.
 */

import type { TeamsConnection } from "../config/index.ts";
import type { ChatSummary, PersonRef } from "../types.ts";
import { graphGetOptional, graphList, graphPost } from "./client.ts";
import { GraphError } from "../utils/errors.ts";
import { mapChat } from "./mappers.ts";
import { resolveUserId } from "./me.ts";

/**
 * The signed-in user's chats, newest activity first.
 *
 * `$expand=members` is what turns an opaque `19:…@thread.v2` into "Anna Schmidt,
 * Tom Weber", which is both what a person recognizes and what the scope rules
 * match on — so it is always expanded, never optional.
 */
export async function listChats(
	conn: TeamsConnection,
	options: { max?: number; meId?: string; signal?: AbortSignal } = {},
): Promise<ChatSummary[]> {
	const max = options.max ?? 25;

	const fetchChats = (orderBy: boolean) =>
		graphList<Record<string, unknown>>(conn, "/me/chats", {
			query: {
				$expand: "members,lastMessagePreview",
				$orderby: orderBy ? "lastMessagePreview/createdDateTime desc" : undefined,
				$top: 50,
			},
			max,
			signal: options.signal,
		});

	let raw: Record<string, unknown>[];
	try {
		raw = await fetchChats(true);
	} catch (err) {
		// $orderby combined with $expand is not accepted in every tenant. Losing
		// the server-side sort is survivable; losing the chat list is not.
		if (err instanceof GraphError && err.status === 400) {
			raw = await fetchChats(false);
		} else {
			throw err;
		}
	}

	const chats = raw.map((entry) => mapChat(entry as Record<string, never>, options.meId));

	// Sort defensively: the fallback path is unsorted, and even the ordered one
	// puts chats without a preview in an arbitrary place.
	return chats.sort((a, b) => {
		const left = a.lastUpdated ? new Date(a.lastUpdated).getTime() : 0;
		const right = b.lastUpdated ? new Date(b.lastUpdated).getTime() : 0;
		return right - left;
	});
}

export async function getChat(
	conn: TeamsConnection,
	chatId: string,
	options: { meId?: string; signal?: AbortSignal } = {},
): Promise<ChatSummary | undefined> {
	const raw = await graphGetOptional<Record<string, unknown>>(
		conn,
		`/chats/${encodeURIComponent(chatId)}`,
		{ query: { $expand: "members" }, signal: options.signal },
	);
	return raw ? mapChat(raw as Record<string, never>, options.meId) : undefined;
}

/**
 * Find an existing chat by label, topic or participant.
 *
 * Returns every match so the caller can refuse to act on an ambiguous one —
 * "message Anna" must not land in the wrong Anna's chat.
 */
export async function findChats(
	conn: TeamsConnection,
	reference: string,
	options: { max?: number; meId?: string; signal?: AbortSignal } = {},
): Promise<ChatSummary[]> {
	if (reference.startsWith("19:")) {
		const direct = await getChat(conn, reference, options);
		return direct ? [direct] : [];
	}

	const chats = await listChats(conn, { max: options.max ?? 100, meId: options.meId, signal: options.signal });
	const lower = reference.toLowerCase();

	const exact = chats.filter(
		(chat) =>
			chat.label.toLowerCase() === lower ||
			chat.topic?.toLowerCase() === lower ||
			chat.members.some(
				(m) => m.upn?.toLowerCase() === lower || m.mail?.toLowerCase() === lower,
			),
	);
	if (exact.length > 0) return exact;

	return chats.filter(
		(chat) =>
			chat.label.toLowerCase().includes(lower) ||
			chat.members.some((m) => m.displayName.toLowerCase().includes(lower)),
	);
}

/**
 * Create a 1:1 or group chat with the given people.
 *
 * Teams reuses the existing 1:1 chat when one is already there, so this is
 * safe to call for "send Anna a message" without checking first.
 */
export async function createChat(
	conn: TeamsConnection,
	participants: PersonRef[],
	options: { topic?: string; signal?: AbortSignal } = {},
): Promise<ChatSummary> {
	const chatType = participants.length > 1 ? "group" : "oneOnOne";

	const members = [
		// The signed-in user must be listed explicitly.
		{
			"@odata.type": "#microsoft.graph.aadUserConversationMember",
			roles: ["owner"],
			"user@odata.bind": "https://graph.microsoft.com/v1.0/me",
		},
		...participants.map((person) => ({
			"@odata.type": "#microsoft.graph.aadUserConversationMember",
			roles: ["owner"],
			"user@odata.bind": `https://graph.microsoft.com/v1.0/users('${person.id ?? person.upn ?? person.mail}')`,
		})),
	];

	const body: Record<string, unknown> = { chatType, members };
	if (chatType === "group" && options.topic) body.topic = options.topic;

	const raw = await graphPost<Record<string, unknown>>(conn, "/chats", body, { signal: options.signal });
	return mapChat((raw ?? {}) as Record<string, never>);
}

/** Resolve people references (UPN, e-mail, name) to PersonRefs. */
export async function resolveParticipants(
	conn: TeamsConnection,
	references: string[],
	signal?: AbortSignal,
): Promise<{ resolved: PersonRef[]; unresolved: string[] }> {
	const resolved: PersonRef[] = [];
	const unresolved: string[] = [];

	for (const reference of references) {
		const person = await resolveUserId(conn, reference, signal);
		if (person) resolved.push(person);
		else unresolved.push(reference);
	}

	return { resolved, unresolved };
}
