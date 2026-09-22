/**
 * Listen mode — the text of a trigger, as a command.
 *
 * A router such as pi-lanes can treat a short message as a control word
 * ("stopp", "neues Thema") instead of a request. It needs the message the way
 * the person meant it: without the mention of pi in front, without
 * punctuation. This module produces exactly that, and nothing more; which
 * words mean what is the router's business.
 */

import type { MessageSummary } from "../types.ts";

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

/** The newest message as a command: no mentions, no punctuation, lower case. */
export function commandText(message: Pick<MessageSummary, "text" | "mentions">, myDisplayName?: string): string {
	const names = [...(message.mentions ?? []).map((m) => m.displayName), ...selfNames(myDisplayName)];
	return normalizeCommand(message.text ?? "", names);
}
