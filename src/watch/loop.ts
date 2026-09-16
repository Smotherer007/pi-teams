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
	noteDeferral,
	noteFailure,
	noteSuccess,
	noteWake,
	openMessages,
	pruneState,
	settleChat,
	shouldWake,
	waitingCount,
	wakesThisHour,
	WATCH_MESSAGE_WINDOW,
	type WatchState,
} from "./index.ts";

// ---------------------------------------------------------------------------
// Events and status
// ---------------------------------------------------------------------------

/** One message that deserves a model turn. */
export interface WatchEvent {
	chat: ChatSummary;
	/** The newest message in the chat — the one the wake was decided on */
	message: MessageSummary;
	/**
	 * Everything still open in that chat, oldest first, the newest included.
	 *
	 * Waking up is about one message, answering is usually about more than one:
	 * three lines sent in a row are three open messages, and answering only the
	 * last of them is what the user notices. Empty means the trigger was the only
	 * thing the read cursor left open.
	 */
	backlog?: MessageSummary[];
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
	/**
	 * Chats still open but held back right now — inside their cooldown, or waiting
	 * for a slot under the hourly limit. The visible form of "not lost, not yet
	 * answered": without it, a burst that hits the limit looks like silence.
	 */
	waiting: number;
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
	/**
	 * A window of the chat's most recent messages, newest first — enough to find
	 * the newest one and whatever the user has not read yet (see
	 * `WATCH_MESSAGE_WINDOW`).
	 */
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
	/** Of the examined ones, how many were only postponed */
	deferred: number;
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
 *
 * The one rule the whole function hangs on: **the cursor is written only for a
 * decision that is final.** It is the record of what pi has looked at, so what
 * lands in it is never looked at again — which is right for a message from
 * somebody pi does not listen to, and wrong for a message that was merely too
 * early. Those go to `noteDeferral` instead and keep their place in line.
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

	const candidates = chatsToExamine(state, chats, watch, now);
	const batch = candidates.slice(0, MAX_EXAMINATIONS_PER_TICK);

	const wakes: WatchEvent[] = [];
	let alreadyRead = 0;
	let deferred = 0;

	for (const chat of batch) {
		const marker = activityMarker(chat);

		// A chat that cannot be read must not occupy a batch slot on every tick, and
		// a token renewal must not turn into a message nobody ever answers: it is
		// written off only once it has used up its attempts.
		const failed = () => {
			if (!noteFailure(state, chat.id)) settleChat(state, chat.id, marker);
		};

		// The read cursor first: a chat the user has already opened in Teams ends
		// here, one Graph call in, before the message is even fetched.
		let readAt: string | undefined;
		try {
			readAt = await deps.readState(chat.id);
		} catch {
			failed();
			continue;
		}

		if (!isUnread(chat, readAt)) {
			noteSuccess(state, chat.id);
			settleChat(state, chat.id, marker);
			alreadyRead += 1;
			continue;
		}

		let messages: MessageSummary[];
		try {
			messages = await deps.listMessages(chat.id);
		} catch {
			// A single unreadable chat must not end the tick — the rest of the
			// list is still worth examining.
			failed();
			continue;
		}

		noteSuccess(state, chat.id);

		const message = messages.find((entry) => !entry.deletedDateTime) ?? messages[0];
		if (!message) {
			settleChat(state, chat.id, marker);
			continue;
		}

		const decision = shouldWake(chat, message, me, watch, state, now);
		if (!decision.wake) {
			if (decision.retryAfterMs !== undefined) {
				// Postponed, not answered and not written off: the chat keeps its
				// message and waits out the cooldown or the wake limit.
				noteDeferral(state, chat.id, now + decision.retryAfterMs);
				deferred += 1;
			} else {
				settleChat(state, chat.id, marker);
			}
			continue;
		}

		noteWake(state, chat.id, now);
		settleChat(state, chat.id, marker);

		// The answer has to cover the whole open thread, not just its newest line:
		// the trigger message is the reason for the wake, the backlog is the reason
		// for a reply that sounds like the conversation was read.
		const backlog = openMessages(messages, readAt, me);
		wakes.push({
			chat,
			message,
			backlog: backlog.length > 0 ? backlog : [message],
			wokeAt: now,
			me,
		});
	}

	// Written even when nothing woke: marking a chat examined is the progress
	// that must survive a restart. Postponed chats are absent from it on purpose.
	if (batch.length > 0) deps.persistCursor?.(state.seen);

	return { examined: batch.length, pending: candidates.length, alreadyRead, deferred, wakes };
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
	/**
	 * The tick currently running, if any.
	 *
	 * Every tick schedules the next one when it finishes, so two ticks running at
	 * once would leave two chains behind and quietly double the polling rate for
	 * the rest of the session. `poll()` therefore joins the tick in flight
	 * instead of starting a second one.
	 */
	let inFlight: Promise<void> | undefined;
	let me: SignedInUser | undefined;
	let meResolved = false;
	let failureCount = 0;

	const status: WatchRuntimeStatus = {
		running: true,
		enabled: watch.enabled,
		intervalSeconds: watch.intervalSeconds,
		wakesThisHour: 0,
		trackedChats: 0,
		waiting: 0,
	};

	const schedule = (delayMs: number) => {
		if (stopped) return;
		timer = setTimeout(() => void runTick(), delayMs);
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
					listMessages: (chatId) =>
						listChatMessages(connection, chatId, { max: WATCH_MESSAGE_WINDOW }),
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
			status.waiting = waitingCount(state, status.lastTickAt);

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

		// Outside the try above on purpose, and guarded on its own: a throwing
		// status callback must not become an unhandled rejection that takes the
		// watcher down without a word.
		try {
			options.onTick?.(status);
		} catch {
			/* a footer that cannot render is not worth stopping the loop for */
		}

		schedule(watch.intervalSeconds * 1000 * Math.min(2 ** failureCount, 10));
	}

	/** Run a tick unless one is already running; either way, resolve with it. */
	function runTick(): Promise<void> {
		if (inFlight) return inFlight;
		inFlight = tick().finally(() => {
			inFlight = undefined;
		});
		return inFlight;
	}

	// The first tick is immediate: it answers what is still open since the last
	// run, and a user who switched listen mode on should not wait a full interval
	// to see it alive.
	void runTick();

	return {
		stop() {
			stopped = true;
			status.running = false;
			if (timer) clearTimeout(timer);
			timer = undefined;
		},
		status: () => ({ ...status, wakesThisHour: wakesThisHour(state, Date.now()) }),
		async poll() {
			// Join a tick already in flight rather than racing it; only start one
			// when nothing is running.
			if (inFlight) {
				await inFlight;
				return;
			}
			if (timer) clearTimeout(timer);
			timer = undefined;
			await runTick();
		},
	};
}
