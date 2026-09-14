/**
 * Listen mode — the prompt.
 *
 * An incoming Teams message is turned into a prompt, and this is the whole of
 * that translation. It is a separate module because the wording is the feature:
 * what pi does with a message it was woken for is decided here, and nothing in
 * the polling loop should be able to change it.
 *
 * Two things the prompt must always carry: the chat ID — so the reply cannot
 * land in the wrong conversation, since pi posts as a person and a name is not
 * an identifier — and the instruction that doing nothing is a valid answer.
 */

import type { MessageSummary, SignedInUser } from "../types.ts";
import type { WatchEvent } from "./loop.ts";

function senderName(message: MessageSummary): string {
	return message.from?.displayName ?? message.from?.upn ?? "someone";
}

/**
 * The prompt for one wake.
 *
 * Written to be read by the model, not by the user: it states the facts, names
 * the tool to use, and then gets out of the way.
 */
export function composeWatchPrompt(event: WatchEvent, me?: SignedInUser): string {
	const { chat, message } = event;
	const text = message.text?.trim() ? message.text.trim() : "(no text content — open the chat to see it)";
	const at = message.createdDateTime ? new Date(message.createdDateTime).toLocaleString() : "just now";

	return [
		"An incoming Microsoft Teams message arrived while you were idle.",
		"",
		`- Chat: ${chat.label} (${chat.chatType})`,
		`- From: ${senderName(message)}`,
		`- At: ${at}`,
		`- Chat ID: ${chat.id}`,
		"",
		"Message:",
		text
			.split("\n")
			.map((line) => `> ${line}`)
			.join("\n"),
		"",
		"Decide whether this needs an answer.",
		"",
		`- If it does: reply with teams_send_chat_message, passing chat: "${chat.id}". pi posts as ` +
			`${me?.displayName ?? "me"}, so write in my voice, in the language of the conversation, as briefly as ` +
			"the message allows. No greeting boilerplate, no signature.",
		"- If it does not: say so in one line and stop.",
		"",
		"Format it for a chat bubble, not a document: the answer first, three short paragraphs at most, " +
			"bullets for anything enumerable, no tables. The recipient is not expecting this message, so it has " +
			"to be readable at a glance.",
		"",
		"Do not react, do not touch other chats, and do not send anything to a channel.",
	].join("\n");
}
