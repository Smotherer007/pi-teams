/**
 * Listen mode — pinning the answer to the chat that asked for it.
 *
 * A wake puts somebody else's words into pi's prompt, and pi holds tools that
 * write in the user's name. The prompt asks it to answer only the chat it was
 * woken for, but an instruction inside a message is exactly the thing that can
 * talk a model out of following an instruction in a prompt. So the rule is
 * enforced here instead: while an answer is pending, sending anywhere other
 * than the chat that woke pi is refused by the safety interceptor.
 *
 * This is deliberately narrow. It gates outgoing messages only — the tools that
 * speak to other people — and not reading, presence or the calendar, because a
 * pin that blocks unrelated work would be a bug the user meets far more often
 * than an attack.
 *
 * Module state rather than a parameter: the interceptor sees a tool call, not a
 * turn, and one session runs one watcher.
 */

import { workerIdentity, type WorkerIdentity } from "./worker.ts";

/**
 * How long a pin can survive without the turn ending.
 *
 * The pin is normally cleared when the agent settles. The timeout is the
 * backstop for the case where that never arrives — a crash, a cancelled run —
 * so a stuck pin cannot lock the user out of their own tools.
 */
export const PIN_TTL_MS = 15 * 60 * 1000;

interface Pin {
	chatId: string;
	/** Chat label, for the refusal message */
	label: string;
	at: number;
}

let pin: Pin | undefined;

/**
 * Wakes whose prompt is queued but whose turn has not started yet.
 *
 * A wake is delivered as a follow-up: when pi is busy it waits in the queue.
 * Pinning at wake time would therefore move the pin away from the answer that
 * is still being written — the reply to chat A would be refused because chat B
 * had just arrived. The pin is taken instead when the queued prompt actually
 * enters the transcript (see `claimWakePin`), keyed by its exact text.
 */
const pending = new Map<string, Pin>();

/** The worker identity of this process, read once. Overridable for tests. */
let worker: WorkerIdentity | undefined = workerIdentity();

/** Tests only: pretend this process is (or is not) a chat worker. */
export function setWorkerIdentityForTests(identity: WorkerIdentity | undefined): void {
	worker = identity;
}

/**
 * Remember a queued wake, so its pin can be taken when its turn starts.
 *
 * Entries older than the pin TTL are dropped on the way: a prompt that never
 * made it into a turn (cancelled queue, restart) must not pin a later one.
 */
export function queueWakePin(prompt: string, chatId: string, label: string, now = Date.now()): void {
	for (const [text, entry] of pending) {
		if (now - entry.at > PIN_TTL_MS) pending.delete(text);
	}
	pending.set(prompt, { chatId, label, at: now });
}

/**
 * A user message entered the transcript: if it is a queued wake, pin it now.
 *
 * Returns whether a pin was taken. Anything else — typed input, other
 * extensions' prompts — leaves the pin alone.
 */
export function claimWakePin(text: string, now = Date.now()): boolean {
	const entry = pending.get(text);
	if (!entry) return false;
	pending.delete(text);
	pin = { chatId: entry.chatId, label: entry.label, at: now };
	return true;
}

/** How many wakes are queued but not yet pinned. */
export function pendingWakePins(): number {
	return pending.size;
}

/** Pin the answer of the turn that is about to start to one chat. */
export function pinWakeTarget(chatId: string, label: string, now = Date.now()): void {
	pin = { chatId, label, at: now };
}

/** Release the pin. Safe to call when there is none. */
export function clearWakeTarget(): void {
	pin = undefined;
}

/** Tests only: forget every queued wake as well. */
export function resetWakePinsForTests(): void {
	pin = undefined;
	pending.clear();
}

/** The chat an answer is currently owed to, if any. */
export function getWakeTarget(now = Date.now()): { chatId: string; label: string } | undefined {
	// A chat worker is pinned for its whole life: it exists to answer one chat,
	// so there is no turn after which it may write anywhere else.
	if (worker) return { chatId: worker.chatId, label: worker.label };
	if (!pin) return undefined;
	if (now - pin.at > PIN_TTL_MS) {
		pin = undefined;
		return undefined;
	}
	return { chatId: pin.chatId, label: pin.label };
}

/**
 * Tools that speak to people, and must therefore stay inside the pinned chat.
 *
 * Reading, presence and calendar tools are absent on purpose: they do not put
 * words in anyone's inbox, and gating them would break ordinary work for no
 * security gain.
 */
const OUTGOING_TOOLS = new Set<string>([
	"teams_send_chat_message",
	"teams_send_channel_message",
	"teams_reply_channel_message",
	"teams_create_chat",
	"teams_chat_members",
	"teams_update_message",
	"teams_react",
]);

/**
 * Why this call must not go through while an answer is pinned, or undefined.
 *
 * The chat is matched on its ID, because that is what the wake prompt hands the
 * model and the only identifier that cannot be talked into meaning a different
 * conversation.
 */
export function pinViolation(
	toolName: string,
	params: Record<string, unknown>,
	now = Date.now(),
): string | undefined {
	const target = getWakeTarget(now);
	if (!target) return undefined;
	if (!OUTGOING_TOOLS.has(toolName)) return undefined;

	const explain = (what: string) =>
		`Listen mode is answering "${target.label}". ${what} Only that chat may be written to until the ` +
		`answer is done — a message can ask for anything, and this is the part that does not take its word for it.`;

	if (toolName === "teams_send_channel_message" || toolName === "teams_reply_channel_message") {
		return explain("This would post in a channel.");
	}
	if (toolName === "teams_create_chat") {
		return explain("This would start a different conversation.");
	}

	const chat = typeof params.chat === "string" ? params.chat : undefined;
	if (!chat) return explain("This does not name the chat it would write to.");
	if (chat !== target.chatId) {
		return explain(`This would write to "${chat}" instead.`);
	}

	return undefined;
}
