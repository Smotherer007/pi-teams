/**
 * Listen mode — the pure half.
 *
 * "Watching" Teams means answering one question over and over: *is there
 * anything here that should wake pi up?* That question is decided entirely
 * from data — the chat list, one message, and what we have already seen — so it
 * lives here as pure functions. The polling, the Graph calls and the prompt
 * injection are in ./loop.ts.
 *
 * Two things are deliberately not pure: the clock and the fetch. Both are
 * injected, which is what makes the decision rules testable without a tenant.
 *
 * `seen` is the durable cursor (see ./cursor.ts): a chat whose marker differs
 * from the stored one has moved since pi last looked, and `isUnread` then says
 * whether the user has already opened it in Teams. That pair is what makes
 * switching listen mode back on answer exactly the messages that are still open
 * — including the ones that arrived while pi was not running.
 */

import type { ChatSummary, MessageSummary, PersonRef, SignedInUser } from "../types.ts";
import type { ResolvedWatchConfig } from "../config/index.ts";
import { matchesPattern } from "../config/scope.ts";
import { chatCandidates } from "../graph/mappers.ts";

// ---------------------------------------------------------------------------
// State
// ---------------------------------------------------------------------------

/**
 * What the watcher remembers between ticks.
 *
 * `seen` is keyed by chat and holds the activity marker of the last message we
 * looked at, so a tick can skip every chat that has not moved. It survives the
 * session on disk, which is what turns "new since the last poll" into "new
 * since pi last looked". It is a high-water mark, not a read receipt: nothing
 * here is ever sent to Teams.
 */
export interface WatchState {
	/** chatId → activity marker of the last examined message, restored from disk */
	seen: Map<string, string>;
	/**
	 * chatId → how often examining it has failed in a row.
	 *
	 * The cursor is written before the chat is examined, which is what keeps a
	 * broken chat from being retried forever — but on its own it would also turn
	 * a token renewal or a throttle into a message that is never answered. So a
	 * failure un-marks the chat again, up to MAX_CHAT_ATTEMPTS.
	 */
	attempts: Map<string, number>;
	/** chatId → epoch ms of the last wake in that chat */
	lastTriggered: Map<string, number>;
	/** epoch ms of every wake, pruned to the last hour */
	triggers: number[];
}

/**
 * How often one chat may fail before pi stops trying and moves on.
 *
 * Three is enough to ride out a token renewal and a throttling window, and few
 * enough that a chat Graph will never serve does not occupy a batch slot on
 * every tick.
 */
export const MAX_CHAT_ATTEMPTS = 3;

/**
 * @param seen The persisted cursor, when this account has been watched before.
 *   Omitted means the very first run, which is the only run that skips history.
 */
export function createWatchState(seen?: Iterable<[string, string]>): WatchState {
	return { seen: new Map(seen), attempts: new Map(), lastTriggered: new Map(), triggers: [] };
}

/** Record a failed examination; true while the chat is still worth retrying. */
export function noteFailure(state: WatchState, chatId: string): boolean {
	const attempts = (state.attempts.get(chatId) ?? 0) + 1;
	state.attempts.set(chatId, attempts);
	return attempts < MAX_CHAT_ATTEMPTS;
}

/** Forget the failure count of a chat that was examined successfully. */
export function noteSuccess(state: WatchState, chatId: string): void {
	state.attempts.delete(chatId);
}

/** Marker for "this is the newest state of that chat we have looked at". */
export function activityMarker(chat: ChatSummary): string {
	return chat.lastUpdated ?? "";
}

/**
 * Upper bound on cursor entries.
 *
 * A user can accumulate chats for years, and the file is read on every session
 * start; a thousand is two orders of magnitude more than the window the watcher
 * polls, so the bound is only ever felt by chats that are long out of sight.
 */
const MAX_CURSOR_ENTRIES = 1000;

/** How many wakes happened within the last hour. */
export function wakesThisHour(state: WatchState, now: number): number {
	return state.triggers.filter((at) => now - at < 3600_000).length;
}

/**
 * Drop bookkeeping that can no longer influence a decision.
 *
 * A session can run for days; without this the maps grow with every chat the
 * user ever touches. The hour is the longest window any rule looks at.
 */
export function pruneState(state: WatchState, now: number): void {
	state.triggers = state.triggers.filter((at) => now - at < 3600_000);
	for (const [chatId, at] of state.lastTriggered) {
		if (now - at > 24 * 3600_000) state.lastTriggered.delete(chatId);
	}

	// The cursor is pruned by size, never by age. Dropping the entry of a chat
	// that has simply been quiet for a long time would make it a candidate again
	// and answer an old message; the count bound only ever touches chats far
	// outside the window the chat list is read in.
	if (state.seen.size > MAX_CURSOR_ENTRIES) {
		const byMarker = [...state.seen.entries()].sort((a, b) => (b[1] ?? "").localeCompare(a[1] ?? ""));
		state.seen = new Map(byMarker.slice(0, MAX_CURSOR_ENTRIES));
	}

	// Counters are bounded by size rather than by the cursor: a chat that keeps
	// failing is deliberately *not* in the cursor, and its counter is what stops
	// it from being retried forever.
	if (state.attempts.size > MAX_CURSOR_ENTRIES) {
		const overflow = state.attempts.size - MAX_CURSOR_ENTRIES;
		for (const chatId of [...state.attempts.keys()].slice(0, overflow)) state.attempts.delete(chatId);
	}
}

// ---------------------------------------------------------------------------
// Which chats are worth a look
// ---------------------------------------------------------------------------

/**
 * Chats that are both allowed by the watch filter and have moved since we last
 * looked.
 *
 * Two questions are asked in sequence and they are deliberately different:
 * "has this chat moved since pi last examined it?" is this function, and "does
 * the user still need to see it?" is `isUnread`, which needs one Graph call per
 * chat and therefore lives in the loop. Age is part of neither: a message that
 * is still unread is still open, whether it arrived a minute or a week ago.
 */
export function chatsToExamine(
	state: WatchState,
	chats: ChatSummary[],
	watch: ResolvedWatchConfig,
): ChatSummary[] {
	return chats.filter((chat) => {
		if (!isWatchedChat(chat, watch)) return false;

		const marker = activityMarker(chat);
		// A chat the chat list cannot date has nothing to compare, so it is never
		// "new" — and never marked, which costs no Graph call at all.
		if (!marker) return false;

		return state.seen.get(chat.id) !== marker;
	});
}

/** Does this chat fall inside the configured chat filter? */
export function isWatchedChat(chat: ChatSummary, watch: ResolvedWatchConfig): boolean {
	if (watch.chats.length === 0) return true;
	const candidates = chatCandidates(chat).filter((value): value is string => !!value);
	return watch.chats.some((pattern) => candidates.some((candidate) => matchesPattern(pattern, candidate)));
}

// ---------------------------------------------------------------------------
// Who pi listens to
// ---------------------------------------------------------------------------

/**
 * Does the sender fall inside the configured listen list?
 *
 * `chats` says where pi listens, this says to whom — the person filter is what
 * keeps listen mode from becoming "answer the whole company". Matched against
 * every identifier the message carries, like the scope rules do.
 */
export function isWatchedSender(message: MessageSummary, watch: ResolvedWatchConfig): boolean {
	if (watch.from.length === 0) return true;

	const sender = message.from;
	if (!sender) return false;

	const candidates = [sender.displayName, sender.upn, sender.mail, sender.id].filter(
		(value): value is string => !!value,
	);
	return watch.from.some((pattern) => candidates.some((candidate) => matchesPattern(pattern, candidate)));
}

// ---------------------------------------------------------------------------
// What the answer has to cover
// ---------------------------------------------------------------------------

/**
 * How many recent messages of a chat the watcher pulls in.
 *
 * One would be enough to decide *whether* to wake, and that is what this used to
 * be — but not enough to answer well: someone who sends three lines in a row
 * leaves three open messages, and a reply to only the newest reads as if the
 * other two were ignored. The window is small on purpose; it only exists to
 * carry the context of the newest message, not to replay the conversation.
 */
export const WATCH_MESSAGE_WINDOW = 10;

/**
 * The messages in this chat that are still open for the user, oldest first.
 *
 * "Open" is the Teams read cursor talking, exactly as `isUnread` means it: a
 * message from somebody else that the user has not caught up with yet. Messages
 * of my own are dropped — pi posts as the user, so they are the answers that
 * were already given, not questions waiting for one.
 *
 * `messages` is newest first, as Graph returns it. A `readAt` Teams will not
 * give is treated as "all of them are open" rather than as settled, for the same
 * reason `isUnread` does: a read state we cannot read must not silently switch
 * listen mode off.
 */
export function openMessages(
	messages: MessageSummary[],
	readAt: string | undefined,
	me: SignedInUser | undefined,
): MessageSummary[] {
	const read = readAt ? new Date(readAt).getTime() : Number.NaN;
	const known = Number.isFinite(read);

	return messages
		.filter((entry) => {
			if (entry.deletedDateTime) return false;
			if (isFromMe(entry.from, me)) return false;
			if (!known) return true;

			const at = entry.createdDateTime ? new Date(entry.createdDateTime).getTime() : Number.NaN;
			return Number.isFinite(at) && at > read;
		})
		.reverse();
}

// ---------------------------------------------------------------------------
// Whether a message should wake pi
// ---------------------------------------------------------------------------

export type WakeDecision = { wake: true } | { wake: false; reason: string };

const WAKE: WakeDecision = { wake: true };
const skip = (reason: string): WakeDecision => ({ wake: false, reason });

/**
 * Is the newest message in this chat still unread for the user?
 *
 * This is the Teams read cursor talking, not pi's own memory: a message the
 * user already opened in Teams does not need an answer from pi, and a message
 * that is still unread does — even if it arrived days ago while pi was not
 * running. Graph withholds the read time for some chat shapes, and a chat we
 * cannot date is not "new" at all; both fall back to answering, because a read
 * state we cannot read must not silently switch listen mode off.
 */
export function isUnread(chat: ChatSummary, lastMessageReadAt?: string): boolean {
	const marker = activityMarker(chat);
	if (!marker) return false;
	if (!lastMessageReadAt) return true;

	const read = new Date(lastMessageReadAt).getTime();
	const moved = new Date(marker).getTime();
	if (!Number.isFinite(read) || !Number.isFinite(moved)) return true;

	return read < moved;
}

/**
 * Should this message wake pi?
 *
 * Ordered so the cheapest and most obvious reason to stay quiet is also the
 * one that gets reported — a wrong wake costs a model turn, a wrong silence
 * costs nothing but the reason in the log.
 */
export function shouldWake(
	chat: ChatSummary,
	message: MessageSummary,
	me: SignedInUser | undefined,
	watch: ResolvedWatchConfig,
	state: WatchState,
	now: number,
): WakeDecision {
	if (message.deletedDateTime) return skip("the message was deleted");
	if (isFromMe(message.from, me)) return skip("the last message is my own");

	if (!isWatchedChat(chat, watch)) return skip("the chat is outside the configured watch list");
	if (!isWatchedSender(message, watch)) return skip("the sender is outside the configured listen list");
	if (watch.mentionOnly && !mentionsMe(message, me)) return skip("the chat is set to mentions only");
	if (watch.maxTriggersPerHour > 0 && wakesThisHour(state, now) >= watch.maxTriggersPerHour) {
		return skip("the hourly wake limit is reached");
	}

	const last = state.lastTriggered.get(chat.id);
	if (last !== undefined && now - last < watch.cooldownSeconds * 1000) {
		return skip("this chat is inside its cooldown");
	}

	return WAKE;
}

/**
 * Record that pi was woken for a chat.
 *
 * Split from the decision so `shouldWake` stays free of side effects and can be
 * called speculatively.
 */
export function noteWake(state: WatchState, chatId: string, now: number): void {
	state.lastTriggered.set(chatId, now);
	state.triggers.push(now);
}

/**
 * Was this message sent by the signed-in user?
 *
 * pi posts as the user, so its own replies come back as "my own message" on the
 * next tick — which is what stops the watcher from waking on itself. Matching
 * on the ID first and the display name second covers both a chat member entry
 * that carries an ID and one that only carries a name.
 */
export function isFromMe(from: PersonRef | undefined, me: SignedInUser | undefined): boolean {
	if (!from || !me) return false;
	if (from.id && me.id && from.id === me.id) return true;
	return !!from.displayName && from.displayName === me.displayName;
}

/**
 * Does this message mention the signed-in user?
 *
 * Graph fills `mentions` for real @-mentions; the text fallback covers the
 * common case of someone typing a first name, which is what "mentions only"
 * means to a person waiting for an answer.
 */
export function mentionsMe(message: MessageSummary, me: SignedInUser | undefined): boolean {
	if (!me) return false;

	const mentioned = message.mentions.some(
		(person) =>
			(person.id && person.id === me.id) ||
			(person.displayName && person.displayName === me.displayName),
	);
	if (mentioned) return true;

	const text = message.text ?? "";
	if (me.displayName && text.includes(`@${me.displayName}`)) return true;

	// First name only: "Patrick, kannst du…". Word-boundary-ish, case-insensitive.
	const firstName = me.displayName.split(/\s+/)[0];
	if (!firstName || firstName.length < 3) return false;
	return new RegExp(`(^|[^\\p{L}])${escapeRegExp(firstName)}($|[^\\p{L}])`, "iu").test(text);
}

function escapeRegExp(value: string): string {
	return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}
