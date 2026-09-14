/**
 * 1:1, group and meeting chats.
 */

import type { TeamsConnection } from "../config/index.ts";
import type { ChatMemberSummary, ChatSummary, PersonRef } from "../types.ts";
import { graphDelete, graphGetOptional, graphList, graphPost } from "./client.ts";
import { GraphError } from "../utils/errors.ts";
import { mapChat, mapChatMember } from "./mappers.ts";
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

// ---------------------------------------------------------------------------
// Read state
// ---------------------------------------------------------------------------

/**
 * Endpoint and body for the two read-cursor actions.
 *
 * Graph takes the user in the **body**, not in the URL: the delegated token
 * says who is allowed to act, the body says whose read state changes. Kept as
 * a pure function (like `deletePath`) because this is the part worth testing.
 */
export function readStateRequest(
	chatId: string,
	user: { id: string; tenantId?: string },
	read: boolean,
): { path: string; body: Record<string, unknown> } {
	return {
		path: `/chats/${encodeURIComponent(chatId)}/${
			read ? "markChatReadForUser" : "markChatUnreadForUser"
		}`,
		// `tenantId` may be undefined; JSON.stringify drops it, and Graph accepts
		// the id alone.
		body: { user: { id: user.id, tenantId: user.tenantId } },
	};
}

/**
 * Mark a chat as read — or unread again — for the signed-in user.
 *
 * This is the only read-state a delegated token may move, and it moves the
 * same cursor the app moves when you open a chat: it does not touch what
 * anyone else sees, but read receipts (where the tenant has them on) do make
 * the chat look seen. Channel messages have no equivalent action.
 */
export async function setChatReadState(
	conn: TeamsConnection,
	chatId: string,
	user: { id: string; tenantId?: string },
	read: boolean,
	options: { signal?: AbortSignal } = {},
): Promise<void> {
	const { path, body } = readStateRequest(chatId, user, read);
	await graphPost(conn, path, body, { signal: options.signal });
}

// ---------------------------------------------------------------------------
// Membership
// ---------------------------------------------------------------------------

/**
 * The members of a chat, with the membership id each one is removed by.
 *
 * Read raw rather than through `mapChat`, because a `ChatSummary` keeps only
 * the people (what a label and the scope rules need) and drops the membership
 * id that removing someone requires.
 */
export async function listChatMembers(
	conn: TeamsConnection,
	chatId: string,
	options: { signal?: AbortSignal } = {},
): Promise<ChatMemberSummary[]> {
	const raw = await graphList<Record<string, never>>(
		conn,
		`/chats/${encodeURIComponent(chatId)}/members`,
		{ max: 100, signal: options.signal },
	);
	return raw.map((entry) => mapChatMember(entry));
}

/**
 * Add a person to an existing chat.
 *
 * Graph binds the member by user id or UPN; a 1:1 chat cannot be extended (the
 * Teams client converts it to a group chat, Graph refuses), so callers should
 * check the chat type first.
 */
export async function addChatMember(
	conn: TeamsConnection,
	chatId: string,
	person: PersonRef,
	options: { signal?: AbortSignal } = {},
): Promise<void> {
	const bind = person.id ?? person.upn ?? person.mail;
	if (!bind) throw new Error(`Cannot add ${person.displayName}: no user id, UPN or e-mail address.`);

	await graphPost(
		conn,
		`/chats/${encodeURIComponent(chatId)}/members`,
		{
			"@odata.type": "#microsoft.graph.aadUserConversationMember",
			"user@odata.bind": `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(bind)}`,
			roles: [],
		},
		{ signal: options.signal },
	);
}

/** Remove one membership. Graph takes the membership id, not the user id. */
export async function removeChatMember(
	conn: TeamsConnection,
	chatId: string,
	membershipId: string,
	options: { signal?: AbortSignal } = {},
): Promise<void> {
	await graphDelete(
		conn,
		`/chats/${encodeURIComponent(chatId)}/members/${encodeURIComponent(membershipId)}`,
		{ signal: options.signal },
	);
}

/**
 * Find the membership of one person in a chat.
 *
 * Matching goes through the identifiers a caller can actually have: object id,
 * UPN, e-mail, display name.
 */
export function matchChatMember(
	members: readonly ChatMemberSummary[],
	reference: string,
): ChatMemberSummary | undefined {
	const needle = reference.trim().toLowerCase();
	if (!needle) return undefined;

	return members.find((member) =>
		[member.userId, member.upn, member.mail, member.displayName]
			.filter((value): value is string => !!value)
			.some((value) => value.toLowerCase() === needle),
	);
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
