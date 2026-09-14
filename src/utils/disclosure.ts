/**
 * The AI disclosure footer.
 *
 * When `aiFooter` is on, every message pi sends in the user's name ends with a
 * note saying so. It is applied here rather than asked for in a prompt because
 * a disclosure that depends on the model remembering is not a disclosure: this
 * is the one piece of an outgoing message that must not be negotiable.
 *
 * Two properties matter more than the wording:
 *
 *  - **Idempotent.** The same body run through this twice carries one footer,
 *    so an edit of an already-footered message does not stack them — and a
 *    model that copied the footer from the chat it is answering does not
 *    produce two.
 *  - **Body-shaped.** The body may be markdown or raw HTML depending on the
 *    tool call, and the footer has to arrive as text in both. HTML output
 *    escapes it and wraps it in its own paragraph.
 */

import type { ResolvedAiFooter } from "../config/index.ts";

/**
 * Append the disclosure, or return the body untouched.
 *
 * A disabled footer, an empty body, and a body that already ends with the note
 * all come back unchanged — the caller never has to ask whether it should.
 */
export function applyAiFooter(
	body: string,
	footer: ResolvedAiFooter | undefined,
	options: { html?: boolean } = {},
): string {
	if (!footer?.enabled) return body;

	const text = footer.text.trim();
	// An empty message is refused before it is sent, and a disclosure on its own
	// is not a message.
	if (!text || !body.trim()) return body;
	if (hasAiFooter(body, text)) return body;

	const trimmed = body.trimEnd();
	return options.html
		? `${trimmed}\n<p>${escapeHtml(text)}</p>`
		: `${trimmed}\n\n${text}`;
}

/** Does this body already carry the disclosure? */
export function hasAiFooter(body: string, text: string): boolean {
	const trimmed = body.trimEnd();
	return trimmed.endsWith(text) || trimmed.endsWith(`<p>${escapeHtml(text)}</p>`);
}

function escapeHtml(value: string): string {
	return value.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}
