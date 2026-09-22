/**
 * Listen mode — control words in a chat.
 *
 * In dispatch mode a chat has a pi process that may be in the middle of
 * something. Two short messages steer that process instead of adding to its
 * work: a stop word aborts what it is doing, a reset word starts the chat over
 * with a fresh session. Both still reach the model afterwards, so the person
 * gets a one-line confirmation in the chat instead of silence.
 *
 * Deliberately strict: the whole message has to be the word (a mention of pi
 * and punctuation aside). "Stop the deployment on hera" is a request, not a
 * control word, and must be read by the model like any other.
 */

import type { ResolvedDispatchConfig } from "../config/index.ts";
import type { MessageSummary } from "../types.ts";

export type ControlAction = "stop" | "reset";

const escapeRegExp = (value: string) => value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");

/** Lower-case the text and drop mentions, punctuation and extra spaces. */
export function normalizeCommand(text: string, names: string[] = []): string {
	let result = ` ${text.toLowerCase()} `;
	const sorted = [...new Set(names.map((n) => n.trim().toLowerCase()).filter(Boolean))].sort((a, b) => b.length - a.length);
	for (const name of sorted) {
		result = result.replace(new RegExp(`(^|[\\s@])${escapeRegExp(name)}(?=[\\s,.:;!?]|$)`, "g"), " ");
	}
	return result
		.replace(/@/g, " ")
		.replace(/[.,;:!?"'„“”«»()[\]]/g, " ")
		.replace(/\s+/g, " ")
		.trim();
}

/**
 * The names pi can be addressed by, for stripping a leading mention.
 *
 * "Neo (neoimpulse)" is mentioned as "Neo" in Teams, so the part before the
 * first space or bracket counts as well.
 */
export function selfNames(displayName: string | undefined): string[] {
	if (!displayName) return [];
	const short = displayName.split(/[\s(]/)[0] ?? "";
	return [displayName, short].filter(Boolean);
}

/** Whether the newest message is a control word, and which. */
export function controlAction(
	message: Pick<MessageSummary, "text" | "mentions">,
	dispatch: Pick<ResolvedDispatchConfig, "stopWords" | "resetWords">,
	myDisplayName?: string,
): ControlAction | undefined {
	const names = [...(message.mentions ?? []).map((m) => m.displayName), ...selfNames(myDisplayName)];
	const text = normalizeCommand(message.text ?? "", names);
	if (!text) return undefined;
	if (dispatch.resetWords.includes(text)) return "reset";
	if (dispatch.stopWords.includes(text)) return "stop";
	return undefined;
}
