/**
 * The shape of a Teams message.
 *
 * A chat bubble is read at a glance, on a phone, between two other things. The
 * failure mode is well known and it is not a lack of language skill: the model
 * writes a *document*. It opens with context, builds up to the point, and
 * closes with a summary — so the reader has to hunt for the answer, which is
 * exactly the "wall of text" people complain about.
 *
 * Telling the model to "be brief" does not fix it. Vague length advice is the
 * weakest kind of formatting instruction; an explicit shape and a worked
 * example are the strongest, and they have to sit where the message is being
 * composed rather than in a document the model may or may not have loaded.
 *
 * So one shape, stated the same way in every place that matters:
 *
 *   - the `body` parameter description — in context at the moment of writing,
 *     which is the highest-attention placement available,
 *   - the per-tool guideline bullet — in context for the whole session, and
 *     flat in the Guidelines list, so it has to name its tools,
 *   - the skill and the listen-mode prompt — which carry the worked example
 *     (`skills/teams-collaboration/SKILL.md`).
 *
 * Keeping it in one module is the point. Four tools used to carry their own
 * copy of one sentence, they had drifted, and the sentence they shared was
 * "three short paragraphs at most" — which permits prose, and got prose.
 */

/**
 * What a message body is, and what shape it has.
 *
 * Used as the `body` parameter description on every tool that sends one, so it
 * is read at the moment the text is written. Kept short on purpose: it is in
 * the tool schema, and every tool in this package is in context at once.
 */
export const CHAT_MESSAGE_BODY_DESCRIPTION =
	"Message text as lightweight markdown, in chat shape: the answer on the first line, then at most 5 short " +
	"lines, `- ` bullets for anything enumerable, one idea per line. No greeting, no closing, no summary. " +
	"Rendered: **bold**, *italic*, `code`, ~~struck~~, [label](url), `- ` bullets, `1. ` numbered, ``` fences. " +
	"A `#` heading arrives as a bold line; tables are not rendered at all. Longer posts the user asked for " +
	"(digest, summary, list of news): an intro line, a blank line, then one `- ` bullet per item, a blank line " +
	"before any closing note.";

/**
 * The same shape as one guideline bullet.
 *
 * `promptGuidelines` are appended flat to the system prompt's Guidelines
 * section with no tool-name prefix, so this names every tool it applies to —
 * otherwise the model cannot tell which "the message body" means. It stays on
 * one line for the same reason: it is a bullet in someone else's list.
 */
export const CHAT_SHAPE_GUIDELINE =
	"teams_send_chat_message, teams_send_channel_message, teams_reply_channel_message and teams_update_message " +
	"write chat shape, not a document: the answer on the first line, at most 5 short lines, `- ` bullets for " +
	"enumerable items, no greeting, no closing summary, no tables; a longer post (digest) is one `- ` bullet per " +
	"item with blank lines between intro, list and closing note.";

/**
 * The shape as prose, for the listen-mode prompt.
 *
 * Listen mode is the case where this matters most: an answer that arrives
 * unrequested has to be readable at a glance, because the recipient is not
 * waiting for it and did not ask for it.
 */
export const CHAT_SHAPE_FOR_LISTEN_MODE =
	"Format it as a chat bubble, not a document: the answer on the first line, at most 5 short lines, " +
	"`- ` bullets for anything enumerable, no greeting and no closing. The recipient is not expecting this " +
	"message, so it has to be readable at a glance.";
