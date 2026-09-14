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
 * looked at, so a tick can skip every chat that has not moved. It is a
 * high-water mark, not a read receipt: nothing here is ever sent to Teams.
 */
export interface WatchState {
	/** chatId → activity marker of the last examined message */
	seen: Map<string, string>;
	/** chatId → epoch ms of the last wake in that chat */
	lastTriggered: Map<string, number>;
	/** epoch ms of every wake, pruned to the last hour */
	triggers: number[];
}

export function createWatchState(): WatchState {
	return { seen: new Map(), lastTriggered: new Map(), triggers: [] };
}

/** Marker for "this is the newest state of that chat we have looked at". */
export function activityMarker(chat: ChatSummary): string {
	return chat.lastUpdated ?? "";
}

/** How many wakes happened within the last hour. */
export function wakesThisHour(state: WatchState, now: number): number {
	return state.triggers.filter((at) => now - at < 3600_000).length;
}

/**
 * Drop bookkeeping that can no longer influence a decision.
 *
 * A session can run for days; without this the two maps grow with every chat
 * the user ever touches. The hour is the longest window any rule looks at.
 */
export function pruneState(state: WatchState, now: number): void {
	state.triggers = state.triggers.filter((at) => now - at < 3600_000);
	for (const [chatId, at] of state.lastTriggered) {
		if (now - at > 24 * 3600_000) state.lastTriggered.delete(chatId);
	}
}

// ---------------------------------------------------------------------------
// Which chats are worth a look
// ---------------------------------------------------------------------------

/**
 * Chats that are both allowed by the watch filter and have moved since we last
 * looked.
 *
 * @param freshnessMs How far back "new" reaches. Enabling listen mode must not
 *   wake pi for every conversation that happened to be active this morning, so
 *   anything older than the window is recorded as seen and then ignored.
 */
export function chatsToExamine(
	state: WatchState,
	chats: ChatSummary[],
	watch: ResolvedWatchConfig,
	now: number,
	freshnessMs: number,
): ChatSummary[] {
	const cutoff = now - freshnessMs;

	return chats.filter((chat) => {
		if (!isWatchedChat(chat, watch)) return false;
		if (state.seen.get(chat.id) === activityMarker(chat)) return false;

		const updated = chat.lastUpdated ? new Date(chat.lastUpdated).getTime() : 0;
		// A chat with no activity timestamp cannot be dated, so it is never new.
		if (!updated || updated < cutoff) return false;

		return true;
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
// Whether a message should wake pi
// ---------------------------------------------------------------------------

export type WakeDecision = { wake: true } | { wake: false; reason: string };

const WAKE: WakeDecision = { wake: true };
const skip = (reason: string): WakeDecision => ({ wake: false, reason });

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
