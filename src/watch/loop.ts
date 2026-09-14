/**
 * Listen mode — the loop.
 *
 * Polls the signed-in user's chats and reports the ones that should wake pi.
 * Deliberately knows nothing about pi: it takes a callback, so the decision
 * rules (./index.ts) and the polling can each be tested on their own, and the
 * extension stays the only place that touches the session.
 *
 * Why polling and not a webhook: Microsoft Graph change notifications for chats
 * need a publicly reachable HTTPS endpoint plus subscription renewal, which a
 * laptop on a home network cannot offer. Polling costs one `/me/chats` call per
 * interval and works everywhere.
 */

import type { ChatSummary, MessageSummary, SignedInUser } from "../types.ts";
import type { ResolvedWatchConfig, TeamsConnection } from "../config/index.ts";
import { listChats } from "../graph/chats.ts";
import { listChatMessages } from "../graph/messages.ts";
import { getMe } from "../graph/me.ts";
import { chatCandidates } from "../graph/mappers.ts";
import { hasAccess } from "../safety/index.ts";
import { formatGraphError } from "../utils/errors.ts";
import {
	activityMarker,
	chatsToExamine,
	createWatchState,
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
	now: () => number;
	/**
	 * Read guard, so the watcher respects the same allow/deny rules as every
	 * other read. Defaults to allowing everything.
	 */
	allowed?: (chat: ChatSummary) => boolean;
}

export interface WatchTickResult {
	/** Chats looked at in this tick */
	examined: number;
	wakes: WatchEvent[];
}

/**
 * Advance the watcher by one tick.
 *
 * The first tick is a **baseline**: it records what is currently there and
 * wakes nobody. Switching listen mode on must not answer the backlog — it means
 * "from now on", and anything older than the freshness window would be a
 * surprise to wake for.
 */
export async function runWatchTick(
	deps: WatchTickDeps,
	state: WatchState,
	watch: ResolvedWatchConfig,
	me: SignedInUser | undefined,
	options: { baseline: boolean },
): Promise<WatchTickResult> {
	const now = deps.now();
	pruneState(state, now);

	const chats = (await deps.listChats()).filter((chat) => deps.allowed?.(chat) ?? true);

	if (options.baseline) {
		for (const chat of chats) state.seen.set(chat.id, activityMarker(chat));
		return { examined: 0, wakes: [] };
	}

	const freshnessMs = Math.max(5 * 60_000, watch.intervalSeconds * 2000);
	const candidates = chatsToExamine(state, chats, watch, now, freshnessMs).slice(0, MAX_EXAMINATIONS_PER_TICK);

	const wakes: WatchEvent[] = [];

	for (const chat of candidates) {
		// Mark first: a chat that fails to load must not be retried forever, and
		// the marker is the chat list's own, so the next tick sees it as moved
		// only if something new actually arrives.
		state.seen.set(chat.id, activityMarker(chat));

		let messages: MessageSummary[];
		try {
			messages = await deps.listMessages(chat.id);
		} catch {
			// A single unreadable chat must not end the tick — the rest of the
			// list is still worth examining.
			continue;
		}

		const message = messages.find((entry) => !entry.deletedDateTime) ?? messages[0];
		if (!message) continue;

		const decision = shouldWake(chat, message, me, watch, state, now);
		if (!decision.wake) continue;

		noteWake(state, chat.id, now);
		wakes.push({ chat, message, wokeAt: now, me });
	}

	return { examined: candidates.length, wakes };
}

// ---------------------------------------------------------------------------
// The loop
// ---------------------------------------------------------------------------

export interface WatchLoopOptions {
	connection: TeamsConnection;
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
	const state = createWatchState();

	let stopped = false;
	let baselineDone = false;
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
					now: () => Date.now(),
					allowed: (chat) => hasAccess(connection, "read", "chats", chatCandidates(chat)),
				},
				state,
				watch,
				me,
				{ baseline: !baselineDone },
			);

			baselineDone = true;
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

	// The first tick is immediate: it establishes the baseline, and a user who
	// just switched listen mode should not wait a full interval to see it alive.
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
