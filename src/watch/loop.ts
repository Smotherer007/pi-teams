/**
 * Listen mode — the loop.
 *
 * Polls the signed-in user's chats and reports the ones that should wake pi.
 * Deliberately knows nothing about pi: it takes a callback, so the decision
 * rules (./index.ts) and the polling can each be tested on their own, and the
 * extension stays the only place that touches the session.
 *
 * The watcher's memory of what it has already looked at does not live in this
 * process — it is read at startup and written back after every tick, so a
 * restart picks up exactly the messages that arrived while pi was not running.
 * See ./cursor.ts for why, and ./index.ts for what counts as "worth a look".
 *
 * Why polling and not a webhook: Microsoft Graph change notifications for chats
 * need a publicly reachable HTTPS endpoint plus subscription renewal, which a
 * laptop on a home network cannot offer. Polling costs one `/me/chats` call per
 * interval and works everywhere.
 */

import type { ChatSummary, MessageSummary, SignedInUser } from "../types.ts";
import type { ResolvedWatchConfig, TeamsConnection } from "../config/index.ts";
import { getChatViewpoint, listChats } from "../graph/chats.ts";
import { listChatMessages } from "../graph/messages.ts";
import { getMe } from "../graph/me.ts";
import { chatCandidates } from "../graph/mappers.ts";
import { hasAccess } from "../safety/index.ts";
import { formatGraphError } from "../utils/errors.ts";
import {
	activityMarker,
	chatsToExamine,
	createWatchState,
	isUnread,
	noteFailure,
	noteSuccess,
	noteWake,
	pruneState,
	shouldWake,
	wakesThisHour,
	type WatchState,
} from "./index.ts";

// ---------------------------------------------------------------------------
// Events and status
// ---------------------------------------------------------------------------

/** One message that deserves a model turn. */
export interface WatchEvent {
	chat: ChatSummary;
	message: MessageSummary;
	/** Epoch ms the wake was decided */
	wokeAt: number;
	/** Who pi is acting as, carried along for the prompt */
	me?: SignedInUser;
}

/** What the footer and `teams_watch` report about a running watcher. */
export interface WatchRuntimeStatus {
	running: boolean;
	enabled: boolean;
	intervalSeconds: number;
	lastTickAt?: number;
	lastWakeAt?: number;
	/** Chat label of the last wake, for the status line */
	lastWakeChat?: string;
	lastError?: string;
	wakesThisHour: number;
	/** Chats seen so far — a quick sanity signal that the loop is alive */
	trackedChats: number;
	/** How many chats had moved when the watcher came up, before the per-tick cap */
	catchUp?: number;
	/** Of those, how many the user had already read in Teams */
	catchUpRead?: number;
}

// ---------------------------------------------------------------------------
// One tick
// ---------------------------------------------------------------------------

/**
 * Bound on how many chats one tick examines.
 *
 * A chat list can move a lot at once — a Monday morning, a release — and each
 * examination is a Graph call. The cap keeps a burst from turning into a
 * thundering herd; whatever is left over is simply picked up next tick.
 */
export const MAX_EXAMINATIONS_PER_TICK = 5;

export interface WatchTickDeps {
	/** Newest chats, most recent activity first */
	listChats: () => Promise<ChatSummary[]>;
	/** The most recent messages of one chat, newest first */
	listMessages: (chatId: string) => Promise<MessageSummary[]>;
	/**
	 * The user's read cursor for one chat (`viewpoint.lastMessageReadDateTime`).
	 * `undefined` means Graph did not say, which is treated as unread rather than
	 * as settled.
	 */
	readState: (chatId: string) => Promise<string | undefined>;
	now: () => number;
	/**
	 * Read guard, so the watcher respects the same allow/deny rules as every
	 * other read. Defaults to allowing everything.
	 */
	allowed?: (chat: ChatSummary) => boolean;
	/**
	 * Persist the cursor after the tick. The durable half of "pi has looked at
	 * this chat": without it a restart replays what was already answered.
	 */
	persistCursor?: (seen: ReadonlyMap<string, string>) => void;
}

export interface WatchTickResult {
	/** Chats looked at in this tick */
	examined: number;
	/** Chats that had moved, before the per-tick cap trimmed the batch */
	pending: number;
	/** Of the examined ones, how many the user had already read */
	alreadyRead: number;
	wakes: WatchEvent[];
}

/**
 * Advance the watcher by one tick.
 *
 * Nothing here is a baseline: with a cursor restored, "moved since pi last
 * looked" is the backlog the user switched listen mode on to catch up with, and
 * without one every chat is compared against the user's own read state on the
 * very first tick. What keeps that from becoming a flood is the per-tick cap
 * and the hourly wake limit, not a window that quietly drops messages.
 */
export async function runWatchTick(
	deps: WatchTickDeps,
	state: WatchState,
	watch: ResolvedWatchConfig,
	me: SignedInUser | undefined,
): Promise<WatchTickResult> {
	const now = deps.now();
	pruneState(state, now);

	const chats = (await deps.listChats()).filter((chat) => deps.allowed?.(chat) ?? true);

	const candidates = chatsToExamine(state, chats, watch);
	const batch = candidates.slice(0, MAX_EXAMINATIONS_PER_TICK);

	const wakes: WatchEvent[] = [];
	let alreadyRead = 0;

	for (const chat of batch) {
		// Mark first, then undo the mark if this chat cannot be examined: a chat
		// that fails must not be retried forever, but a token renewal must not turn
		// into a message nobody ever answers.
		state.seen.set(chat.id, activityMarker(chat));

		// The read cursor first: a chat the user has already opened in Teams ends
		// here, one Graph call in, before the message is even fetched.
		let readAt: string | undefined;
		try {
			readAt = await deps.readState(chat.id);
		} catch {
			if (noteFailure(state, chat.id)) state.seen.delete(chat.id);
			continue;
		}

		if (!isUnread(chat, readAt)) {
			noteSuccess(state, chat.id);
			alreadyRead += 1;
			continue;
		}

		let messages: MessageSummary[];
		try {
			messages = await deps.listMessages(chat.id);
		} catch {
			// A single unreadable chat must not end the tick — the rest of the
			// list is still worth examining.
			if (noteFailure(state, chat.id)) state.seen.delete(chat.id);
			continue;
		}

		noteSuccess(state, chat.id);

		const message = messages.find((entry) => !entry.deletedDateTime) ?? messages[0];
		if (!message) continue;

		const decision = shouldWake(chat, message, me, watch, state, now);
		if (!decision.wake) continue;

		noteWake(state, chat.id, now);
		wakes.push({ chat, message, wokeAt: now, me });
	}

	// Written even when nothing woke: marking a chat examined is the progress
	// that must survive a restart.
	if (batch.length > 0) deps.persistCursor?.(state.seen);

	return { examined: batch.length, pending: candidates.length, alreadyRead, wakes };
}

// ---------------------------------------------------------------------------
// The loop
// ---------------------------------------------------------------------------

export interface WatchLoopOptions {
	connection: TeamsConnection;
	/**
	 * The persisted cursor read at startup. `undefined` means listen mode has
	 * never run for this account, which is the one case where the read state of
	 * every chat is asked instead — the backlog the user switched it on for.
	 */
	cursor?: Map<string, string>;
	/** Called after a tick that moved the cursor, to make it durable */
	persistCursor?: (seen: ReadonlyMap<string, string>) => void;
	/** Called once per wake, in arrival order */
	onWake: (event: WatchEvent) => void;
	/** Called when a tick fails; the loop keeps running */
	onError?: (message: string) => void;
	/** Called after every tick, for the footer */
	onTick?: (status: WatchRuntimeStatus) => void;
}

export interface WatchLoop {
	/** Stop polling. Safe to call twice. */
	stop(): void;
	/** A snapshot for the status line and the tool */
	status(): WatchRuntimeStatus;
	/** Force a tick now, instead of waiting out the interval */
	poll(): Promise<void>;
}

/**
 * The loop this process is running, if any.
 *
 * One session runs one watcher, and the tools that report on listen mode need
 * to see it. Keeping the slot here — rather than in the extension — means
 * `teams_watch` can answer with live runtime facts without importing the
 * extension, which imports the tools.
 */
let activeLoop: WatchLoop | undefined;

export function setActiveWatchLoop(loop: WatchLoop | undefined): void {
	activeLoop = loop;
}

export function getActiveWatchLoop(): WatchLoop | undefined {
	return activeLoop;
}

/** Stop the running watcher, if there is one. */
export function stopActiveWatchLoop(): void {
	activeLoop?.stop();
	activeLoop = undefined;
}

/**
 * Start polling for the given connection.
 *
 * Each tick is scheduled after the previous one finished rather than on a fixed
 * timer, so a slow Graph call cannot stack ticks on top of each other.
 */
export function startWatchLoop(options: WatchLoopOptions): WatchLoop {
	const { connection } = options;
	const watch = connection.watch;
	const state = createWatchState(options.cursor);

	let stopped = false;
	let timer: ReturnType<typeof setTimeout> | undefined;
	let me: SignedInUser | undefined;
	let meResolved = false;
	let failureCount = 0;

	const status: WatchRuntimeStatus = {
		running: true,
		enabled: watch.enabled,
		intervalSeconds: watch.intervalSeconds,
		wakesThisHour: 0,
		trackedChats: 0,
	};

	const schedule = (delayMs: number) => {
		if (stopped) return;
		timer = setTimeout(() => void tick(), delayMs);
		// Do not hold the process open for a poll that can wait.
		timer.unref?.();
	};

	async function tick(): Promise<void> {
		if (stopped) return;

		try {
			if (!meResolved) {
				me = await getMe(connection).catch(() => undefined);
				meResolved = true;
			}

			const result = await runWatchTick(
				{
					listChats: () => listChats(connection, { max: 50, meId: me?.id }),
					listMessages: (chatId) => listChatMessages(connection, chatId, { max: 1 }),
					readState: async (chatId) => (await getChatViewpoint(connection, chatId))?.lastMessageReadAt,
					now: () => Date.now(),
					allowed: (chat) => hasAccess(connection, "read", "chats", chatCandidates(chat)),
					persistCursor: options.persistCursor,
				},
				state,
				watch,
				me,
			);

			// Reported once, from the first tick: the number the user wants after
			// switching listen mode on is "how much did I miss", and every later
			// tick would only show the steady-state traffic.
			if (status.catchUp === undefined) {
				status.catchUp = result.pending;
				status.catchUpRead = result.alreadyRead;
			}
			failureCount = 0;
			status.lastTickAt = Date.now();
			status.lastError = undefined;
			status.wakesThisHour = wakesThisHour(state, status.lastTickAt);
			status.trackedChats = state.seen.size;

			for (const event of result.wakes) {
				status.lastWakeAt = event.wokeAt;
				status.lastWakeChat = event.chat.label;
				status.wakesThisHour = wakesThisHour(state, event.wokeAt);
				try {
					options.onWake(event);
				} catch (err) {
					options.onError?.(formatGraphError(err));
				}
			}
		} catch (err) {
			// Transient failures are normal — a sleeping laptop, an expired token,
			// a throttled tenant. Back off instead of giving up, and report it so
			// the user is not left wondering why nothing happens.
			failureCount += 1;
			status.lastTickAt = Date.now();
			status.lastError = formatGraphError(err);
			options.onError?.(status.lastError);
		}

		options.onTick?.(status);
		schedule(watch.intervalSeconds * 1000 * Math.min(2 ** failureCount, 10));
	}

	// The first tick is immediate: it answers what is still open since the last
	// run, and a user who switched listen mode on should not wait a full interval
	// to see it alive.
	void tick();

	return {
		stop() {
			stopped = true;
			status.running = false;
			if (timer) clearTimeout(timer);
			timer = undefined;
		},
		status: () => ({ ...status, wakesThisHour: wakesThisHour(state, Date.now()) }),
		async poll() {
			if (timer) clearTimeout(timer);
			timer = undefined;
			await tick();
		},
	};
}
