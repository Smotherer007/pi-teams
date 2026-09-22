/**
 * Listen mode — who a chat worker is.
 *
 * Where a chat is answered by a pi process of its own (a lane, e.g. with
 * pi-lanes), that process loads this extension like any other, so it has to
 * know two things the moment it starts: that it must never run a watcher of
 * its own, and which single chat it may write to.
 *
 * Both travel in the environment, set by the router that starts the lane (see
 * `laneHint` in ../extension/index.ts). An environment variable is the one
 * channel that exists before the first line of the extension runs, and the
 * model cannot change it.
 */

/** Environment variables a chat worker is started with. */
export const WORKER_ENV = {
	/** The chat this worker answers, and the only one it may write to */
	chat: "PI_TEAMS_WORKER_CHAT",
	/** Chat label, for refusal messages and the session name */
	label: "PI_TEAMS_WORKER_LABEL",
	/** "oneOnOne" | "group" | "meeting" */
	chatType: "PI_TEAMS_WORKER_CHAT_TYPE",
	/**
	 * Address of the other person, set for one-to-one chats only.
	 *
	 * Used to decide whether this chat may look at what pi wrote elsewhere
	 * (`teams_history`). In a group chat there is no single person to ask, so
	 * the variable is absent and the answer is no.
	 */
	peer: "PI_TEAMS_WORKER_PEER",
} as const;

export interface WorkerIdentity {
	chatId: string;
	label: string;
	chatType?: string;
	peer?: string;
}

/** The chat this process works for, or undefined in an ordinary pi session. */
export function workerIdentity(env: NodeJS.ProcessEnv = process.env): WorkerIdentity | undefined {
	const chatId = env[WORKER_ENV.chat]?.trim();
	if (!chatId) return undefined;
	return {
		chatId,
		label: env[WORKER_ENV.label]?.trim() || chatId,
		chatType: env[WORKER_ENV.chatType]?.trim() || undefined,
		peer: env[WORKER_ENV.peer]?.trim() || undefined,
	};
}

/** The environment a worker for this chat is started with. */
export function workerEnv(identity: WorkerIdentity): Record<string, string> {
	const env: Record<string, string> = {
		[WORKER_ENV.chat]: identity.chatId,
		[WORKER_ENV.label]: identity.label,
	};
	if (identity.chatType) env[WORKER_ENV.chatType] = identity.chatType;
	if (identity.peer) env[WORKER_ENV.peer] = identity.peer;
	return env;
}

/**
 * Tools a chat worker may not call at all.
 *
 * A worker answers one conversation. Switching listen mode, rewriting the
 * config or signing the account out would change what every other worker and
 * the watcher itself do, which is not a decision one chat gets to make.
 */
export const WORKER_FORBIDDEN_TOOLS = new Set<string>(["teams_watch", "teams_setup", "teams_logout"]);

export function workerViolation(toolName: string, identity: WorkerIdentity | undefined): string | undefined {
	if (!identity) return undefined;
	if (!WORKER_FORBIDDEN_TOOLS.has(toolName)) return undefined;
	return (
		`This pi process answers the Teams chat "${identity.label}" only. ${toolName} changes listen mode or the ` +
		"account for every chat, so it is not available here. Ask the user to do it in the main session."
	);
}
