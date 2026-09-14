/**
 * Listen mode — the prompt.
 *
 * An incoming Teams message is turned into a prompt, and this is the whole of
 * that translation. It is a separate module because the wording is the feature:
 * what pi does with a message it was woken for is decided here, and nothing in
 * the polling loop should be able to change it.
 *
 * Three things the prompt must always carry: the chat ID — so the reply cannot
 * land in the wrong conversation, since pi posts as a person and a name is not
 * an identifier — the instruction that doing nothing is a valid answer, and
 * every message still open in that chat. The last one is not decoration: a wake
 * is decided on the newest message, but people send three lines in a row, and a
 * reply that answers only the last of them is the failure mode this prompt
 * exists to prevent.
 */

import type { MessageSummary, SignedInUser } from "../types.ts";
import type { WatchEvent } from "./loop.ts";

function senderName(message: MessageSummary): string {
	return message.from?.displayName ?? message.from?.upn ?? "someone";
}

/** `17:00`, or nothing when Graph did not date the message. */
function timeOf(message: MessageSummary): string {
	if (!message.createdDateTime) return "";
	const at = new Date(message.createdDateTime);
	return Number.isFinite(at.getTime()) ? at.toLocaleTimeString() : "";
}

function quote(text: string | undefined): string {
	const body = text?.trim() ? text.trim() : "(no text content — open the chat to see it)";
	return body
		.split("\n")
		.map((line) => `> ${line}`)
		.join("\n");
}

/**
 * The open messages of the chat, oldest first.
 *
 * Falls back to the trigger message alone when the backlog is missing — a wake
 * without one still has to produce a usable prompt.
 */
function openThread(event: WatchEvent): MessageSummary[] {
	const backlog = event.backlog?.filter((entry) => !entry.deletedDateTime) ?? [];
	return backlog.length > 0 ? backlog : [event.message];
}

/** One numbered block per open message, oldest first. */
function renderThread(messages: MessageSummary[]): string {
	return messages
		.map((entry, index) => {
			const newest = index === messages.length - 1;
			const stamp = timeOf(entry);
			const heading = [stamp, senderName(entry)].filter(Boolean).join(" · ");
			const label = newest ? `${heading} — newest` : heading;
			return [`> ${index + 1}. ${label}`, quote(entry.text)].join("\n");
		})
		.join("\n>\n");
}

/**
 * The prompt for one wake.
 *
 * Written to be read by the model, not by the user: it states the facts, names
 * the tool to use, and then gets out of the way.
 */
export function composeWatchPrompt(event: WatchEvent, me?: SignedInUser): string {
	const { chat, message } = event;
	const thread = openThread(event);
	const at = message.createdDateTime ? new Date(message.createdDateTime).toLocaleString() : "just now";

	const header = [
		"The following Microsoft Teams message arrived while you were idle.",
		"",
		`- Chat: ${chat.label} (${chat.chatType})`,
		`- From: ${senderName(message)}`,
		`- At: ${at}`,
		`- Chat ID: ${chat.id}`,
	];

	const body =
		thread.length > 1
			? [
					"",
					`Unanswered messages in this chat (${thread.length}), oldest first:`,
					"",
					renderThread(thread),
				]
			: ["", "Message:", "", quote(message.text)];

	const decision =
		thread.length > 1
			? [
					"Decide whether these need an answer.",
					"",
					"All of them are still unread in Teams, and they arrived together as one conversation. Answer every " +
						"one of them, not just the newest — one short reply can cover all of it, but a reply that " +
						"ignores the earlier messages is wrong even when its wording is fine.",
				]
			: ["Decide whether this needs an answer."];

	return [
		...header,
		...body,
		"",
		...decision,
		"",
		`- If it does: reply with teams_send_chat_message, passing chat: "${chat.id}". pi posts as ` +
			`${me?.displayName ?? "me"}, so write in my voice, in the language of the conversation, as briefly as ` +
			"the message allows. No greeting boilerplate, no signature, and no note that a model wrote it — " +
			"the sending tool appends the configured AI disclosure itself.",
		"- If it does not: say so in one line and stop.",
		"",
		"Format it for a chat bubble, not a document: the answer first, three short paragraphs at most, " +
			"bullets for anything enumerable, no tables. The recipient is not expecting this message, so it has " +
			"to be readable at a glance.",
		"",
		"Do not react, do not touch other chats, and do not send anything to a channel.",
	].join("\n");
}
