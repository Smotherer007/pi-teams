/**
 * Listen mode — keeping the user online while pi listens.
 *
 * Teams shows a user as Offline unless at least one *presence session* exists
 * for them. A running Teams client is such a session; pi talking to Graph is
 * not. Even a preferred presence set with `setUserPreferredPresence` only takes
 * effect "when at least one presence session exists for the user. Otherwise, the
 * user's presence shows as Offline" (Microsoft Graph docs). A user who only
 * works through pi therefore always looked Offline, including while pi was
 * answering their chats.
 *
 * While listen mode runs, pi now holds its own application presence session
 * (`setPresence`, sessionId = the app's client ID) and renews it before it
 * expires. When the watcher stops, the session is cleared so the user drops to
 * Offline right away instead of lingering as Available. A preferred presence
 * the user picked (`teams_set_presence`, e.g. DoNotDisturb) still wins over the
 * session — which is the point: it now actually shows.
 */

import type { TeamsConnection } from "../config/index.ts";
import { clearSessionPresence, setSessionPresence } from "../graph/presence.ts";
import { formatGraphError } from "../utils/errors.ts";

/** How long each session write lasts. Graph accepts PT5M … PT4H. */
export const PRESENCE_SESSION_MINUTES = 15;
/** Renew well before expiry so one slow or failed call does not drop the user. */
export const PRESENCE_RENEW_MINUTES = 5;

export interface PresenceKeeperDeps {
	set: (userId: string, expirationDuration: string) => Promise<void>;
	clear: (userId: string) => Promise<void>;
	onError?: (message: string) => void;
}

export interface PresenceKeeper {
	/** Start (or keep) holding the session for this user. Idempotent. */
	start(userId: string): void;
	/** Stop renewing and clear the session. Safe to call twice. */
	stop(): Promise<void>;
	/** Whether a renewal timer is running */
	readonly active: boolean;
}

export function createPresenceKeeper(deps: PresenceKeeperDeps): PresenceKeeper {
	let userId: string | undefined;
	let timer: ReturnType<typeof setInterval> | undefined;
	/** Report a failure once, not every renewal, until a renewal succeeds again. */
	let failing = false;
	const duration = `PT${PRESENCE_SESSION_MINUTES}M`;

	const renew = async () => {
		if (!userId) return;
		try {
			await deps.set(userId, duration);
			failing = false;
		} catch (err) {
			if (!failing) deps.onError?.(`Could not keep Teams presence online: ${formatGraphError(err)}`);
			failing = true;
		}
	};

	return {
		start(id: string) {
			if (timer && userId === id) return;
			if (timer) clearInterval(timer);
			userId = id;
			void renew();
			timer = setInterval(() => void renew(), PRESENCE_RENEW_MINUTES * 60_000);
			timer.unref?.();
		},
		async stop() {
			if (timer) clearInterval(timer);
			timer = undefined;
			const id = userId;
			userId = undefined;
			if (!id) return;
			try {
				await deps.clear(id);
			} catch {
				/* the session expires on its own within PRESENCE_SESSION_MINUTES */
			}
		},
		get active() {
			return timer !== undefined;
		},
	};
}

/** The keeper wired to Graph for one connection. */
export function presenceKeeperFor(
	conn: TeamsConnection,
	onError?: (message: string) => void,
): PresenceKeeper {
	return createPresenceKeeper({
		set: (userId, expirationDuration) =>
			setSessionPresence(conn, userId, "Available", { expirationDuration }),
		clear: (userId) => clearSessionPresence(conn, userId),
		onError,
	});
}
