/**
 * Reading and writing messages — chats, channels, replies, reactions.
 *
 * This is the module that speaks in the user's name, so the shapes it builds
 * are deliberately explicit: a mention is only sent when the caller resolved a
 * real person for it, and an edit or delete always targets one message ID that
 * the caller obtained from a read.
 */

import type { TeamsConnection } from "../config/index.ts";
import type { ChannelSummary, MessageLocation, MessageSummary, PersonRef } from "../types.ts";
import { graphGetOptional, graphList, graphPatch, graphPost } from "./client.ts";
import { mapMessage } from "./mappers.ts";
import { markdownToTeamsHtml } from "../utils/richtext.ts";

/* eslint-disable @typescript-eslint/no-explicit-any */
type Raw = Record<string, any>;

// ---------------------------------------------------------------------------
// Message bodies
// ---------------------------------------------------------------------------

export interface MessageInput {
	/**
	 * Message text.
	 *
	 * Treated as lightweight markdown by default — bold, italic, code, bullets,
	 * numbered lists, links and `# headings` are converted to the HTML Teams
	 * renders. Set `html` to true to send raw HTML instead.
	 */
	body: string;
	/** Treat `body` as raw HTML instead of converting markdown */
	html?: boolean;
	/** Channel messages only */
	subject?: string;
	/** "normal" | "high" | "urgent" */
	importance?: string;
	/** People to @-mention; each must appear in the body as @DisplayName */
	mentions?: PersonRef[];
}

/**
 * Build the Graph message body.
 *
 * Mentions are the fiddly part: Teams requires an `<at id="N">` span in the
 * HTML *and* a matching entry in the `mentions` array. The display name is
 * replaced in the text so the caller can just write "@Anna Schmidt".
 *
 * Markdown is converted unless the caller says the body is already HTML: what
 * comes back from a model is markdown, and markdown sent raw arrives as
 * asterisks and dashes.
 */
export function buildMessageBody(input: MessageInput): Record<string, unknown> {
	const mentions = input.mentions ?? [];
	let content = input.html ? input.body : markdownToTeamsHtml(input.body);

	const mentionEntries: Record<string, unknown>[] = [];
	mentions.forEach((person, index) => {
		const tag = `<at id="${index}">${escapeHtml(person.displayName)}</at>`;
		const needle = new RegExp(`@${escapeRegExp(person.displayName)}`, "g");
		if (needle.test(content)) {
			content = content.replace(needle, tag);
		} else {
			// Not referenced in the text — prepend it so the mention still fires.
			content = `${tag} ${content}`;
		}
		mentionEntries.push({
			id: index,
			mentionText: person.displayName,
			mentioned: {
				user: {
					displayName: person.displayName,
					id: person.id,
					userIdentityType: "aadUser",
				},
			},
		});
	});

	const body: Record<string, unknown> = {
		body: { contentType: "html", content },
	};
	if (input.subject) body.subject = input.subject;
	if (input.importance) body.importance = input.importance;
	if (mentionEntries.length > 0) body.mentions = mentionEntries;

	return body;
}

function escapeHtml(value: string): string {
	return value.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

function escapeRegExp(value: string): string {
	return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

// ---------------------------------------------------------------------------
// Paths
// ---------------------------------------------------------------------------

function chatPath(chatId: string): string {
	return `/chats/${encodeURIComponent(chatId)}/messages`;
}

function channelPath(teamId: string, channelId: string): string {
	return `/teams/${encodeURIComponent(teamId)}/channels/${encodeURIComponent(channelId)}/messages`;
}

// ---------------------------------------------------------------------------
// Reading — chats
// ---------------------------------------------------------------------------

export async function listChatMessages(
	conn: TeamsConnection,
	chatId: string,
	options: { max?: number; signal?: AbortSignal } = {},
): Promise<MessageSummary[]> {
	const raw = await graphList<Raw>(conn, chatPath(chatId), {
		query: { $top: Math.min(options.max ?? 25, 50) },
		max: options.max ?? 25,
		signal: options.signal,
	});
	const location: MessageLocation = { kind: "chat", chatId };
	return raw.map((entry) => mapMessage(entry, location));
}

export async function getChatMessage(
	conn: TeamsConnection,
	chatId: string,
	messageId: string,
	signal?: AbortSignal,
): Promise<MessageSummary | undefined> {
	const raw = await graphGetOptional<Raw>(
		conn,
		`${chatPath(chatId)}/${encodeURIComponent(messageId)}`,
		{ signal },
	);
	return raw ? mapMessage(raw, { kind: "chat", chatId }) : undefined;
}

// ---------------------------------------------------------------------------
// Reading — channels
// ---------------------------------------------------------------------------

export async function listChannelMessages(
	conn: TeamsConnection,
	channel: ChannelSummary,
	options: { max?: number; withReplies?: boolean; signal?: AbortSignal } = {},
): Promise<MessageSummary[]> {
	if (!channel.teamId) throw new Error("Channel is missing its teamId.");

	const raw = await graphList<Raw>(conn, channelPath(channel.teamId, channel.id), {
		query: { $top: Math.min(options.max ?? 25, 50) },
		max: options.max ?? 25,
		signal: options.signal,
	});

	const location: MessageLocation = {
		kind: "channel",
		teamId: channel.teamId,
		teamName: channel.teamName,
		channelId: channel.id,
		channelName: channel.displayName,
	};

	const messages = raw.map((entry) => mapMessage(entry, location));

	if (!options.withReplies) return messages;

	// Threads are what a channel conversation actually is, so resolve them —
	// sequentially, because a burst of parallel calls is the fastest way to get
	// throttled by Graph.
	for (const message of messages) {
		const replies = await listChannelReplies(conn, channel, message.id, {
			max: 20,
			signal: options.signal,
		});
		message.replyCount = replies.length;
	}
	return messages;
}

export async function listChannelReplies(
	conn: TeamsConnection,
	channel: ChannelSummary,
	messageId: string,
	options: { max?: number; signal?: AbortSignal } = {},
): Promise<MessageSummary[]> {
	if (!channel.teamId) throw new Error("Channel is missing its teamId.");

	const raw = await graphList<Raw>(
		conn,
		`${channelPath(channel.teamId, channel.id)}/${encodeURIComponent(messageId)}/replies`,
		{
			query: { $top: Math.min(options.max ?? 30, 50) },
			max: options.max ?? 30,
			signal: options.signal,
		},
	);

	const location: MessageLocation = {
		kind: "channel",
		teamId: channel.teamId,
		teamName: channel.teamName,
		channelId: channel.id,
		channelName: channel.displayName,
	};
	return raw.map((entry) => mapMessage(entry, location));
}

export async function getChannelMessage(
	conn: TeamsConnection,
	channel: ChannelSummary,
	messageId: string,
	signal?: AbortSignal,
): Promise<MessageSummary | undefined> {
	if (!channel.teamId) throw new Error("Channel is missing its teamId.");
	const raw = await graphGetOptional<Raw>(
		conn,
		`${channelPath(channel.teamId, channel.id)}/${encodeURIComponent(messageId)}`,
		{ signal },
	);
	return raw
		? mapMessage(raw, {
				kind: "channel",
				teamId: channel.teamId,
				teamName: channel.teamName,
				channelId: channel.id,
				channelName: channel.displayName,
			})
		: undefined;
}

// ---------------------------------------------------------------------------
// Writing
// ---------------------------------------------------------------------------

export async function sendChatMessage(
	conn: TeamsConnection,
	chatId: string,
	input: MessageInput,
	signal?: AbortSignal,
): Promise<MessageSummary> {
	const raw = await graphPost<Raw>(conn, chatPath(chatId), buildMessageBody(input), { signal });
	return mapMessage(raw ?? {}, { kind: "chat", chatId });
}

export async function sendChannelMessage(
	conn: TeamsConnection,
	channel: ChannelSummary,
	input: MessageInput,
	signal?: AbortSignal,
): Promise<MessageSummary> {
	if (!channel.teamId) throw new Error("Channel is missing its teamId.");
	const raw = await graphPost<Raw>(
		conn,
		channelPath(channel.teamId, channel.id),
		buildMessageBody(input),
		{ signal },
	);
	return mapMessage(raw ?? {}, {
		kind: "channel",
		teamId: channel.teamId,
		teamName: channel.teamName,
		channelId: channel.id,
		channelName: channel.displayName,
	});
}

export async function replyToChannelMessage(
	conn: TeamsConnection,
	channel: ChannelSummary,
	messageId: string,
	input: MessageInput,
	signal?: AbortSignal,
): Promise<MessageSummary> {
	if (!channel.teamId) throw new Error("Channel is missing its teamId.");
	const raw = await graphPost<Raw>(
		conn,
		`${channelPath(channel.teamId, channel.id)}/${encodeURIComponent(messageId)}/replies`,
		buildMessageBody(input),
		{ signal },
	);
	return mapMessage(raw ?? {}, {
		kind: "channel",
		teamId: channel.teamId,
		teamName: channel.teamName,
		channelId: channel.id,
		channelName: channel.displayName,
	});
}

/**
 * Path to one message in a chat — the base for editing and deleting.
 *
 * The user id is explicit: Graph does not accept `/me` in this position, which
 * is why every caller resolves the signed-in user first.
 *
 * There is no channel form. Editing or deleting a channel post needs the
 * `ChannelMessage.ReadWrite` scope, and this package deliberately does not ask
 * for a permission it cannot rely on being granted.
 */
export function messagePath(chatId: string, messageId: string, userId: string): string {
	return (
		`/users/${encodeURIComponent(userId)}` +
		`/chats/${encodeURIComponent(chatId)}` +
		`/messages/${encodeURIComponent(messageId)}`
	);
}

/**
 * Delete one of the signed-in user's own chat messages (soft delete, as in the app).
 *
 * Graph exposes this as an action, not an HTTP DELETE, and only for the sender.
 */
export async function softDeleteMessage(
	conn: TeamsConnection,
	chatId: string,
	messageId: string,
	userId: string,
	options: { signal?: AbortSignal } = {},
): Promise<void> {
	await graphPost(conn, `${messagePath(chatId, messageId, userId)}/softDelete`, undefined, {
		signal: options.signal,
	});
}

/** Restore a message that was soft-deleted. */
export async function undoSoftDeleteMessage(
	conn: TeamsConnection,
	chatId: string,
	messageId: string,
	userId: string,
	options: { signal?: AbortSignal } = {},
): Promise<void> {
	await graphPost(conn, `${messagePath(chatId, messageId, userId)}/undoSoftDelete`, undefined, {
		signal: options.signal,
	});
}

// ---------------------------------------------------------------------------
// Editing
// ---------------------------------------------------------------------------

/**
 * Rewrite one of the signed-in user's own chat messages.
 *
 * Only the body is patched: an edit replaces the text as it stands, and sending
 * a fresh set of mention entities would silently re-notify people who were
 * already notified by the original message.
 */
export async function updateMessage(
	conn: TeamsConnection,
	chatId: string,
	messageId: string,
	userId: string,
	input: { body: string; html?: boolean },
	options: { signal?: AbortSignal } = {},
): Promise<void> {
	const { body } = buildMessageBody({ body: input.body, html: input.html });
	await graphPatch(conn, messagePath(chatId, messageId, userId), { body }, {
		signal: options.signal,
	});
}

// ---------------------------------------------------------------------------
// Reactions
// ---------------------------------------------------------------------------

function reactionBase(location: MessageLocation, messageId: string, replyId?: string): string {
	if (location.kind === "chat") {
		return `/chats/${encodeURIComponent(location.chatId!)}/messages/${encodeURIComponent(messageId)}`;
	}
	const channelBase = `${channelPath(location.teamId!, location.channelId!)}/${encodeURIComponent(messageId)}`;
	return replyId ? `${channelBase}/replies/${encodeURIComponent(replyId)}` : channelBase;
}

export async function setReaction(
	conn: TeamsConnection,
	location: MessageLocation,
	messageId: string,
	reactionType: string,
	options: { replyId?: string; signal?: AbortSignal } = {},
): Promise<void> {
	await graphPost(
		conn,
		`${reactionBase(location, messageId, options.replyId)}/setReaction`,
		{ reactionType },
		{ signal: options.signal },
	);
}

export async function unsetReaction(
	conn: TeamsConnection,
	location: MessageLocation,
	messageId: string,
	reactionType: string,
	options: { replyId?: string; signal?: AbortSignal } = {},
): Promise<void> {
	await graphPost(
		conn,
		`${reactionBase(location, messageId, options.replyId)}/unsetReaction`,
		{ reactionType },
		{ signal: options.signal },
	);
}
