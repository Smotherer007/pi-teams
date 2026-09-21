/**
 * The shape of a Teams message.
 *
 * The failure this guards against is not a crash. It is a message that is
 * correct and reads like a report: context first, the answer somewhere in the
 * middle, a summary at the end — the "wall of text" people complain about.
 *
 * The root cause was wording. Four tools each carried their own copy of one
 * sentence, and the sentence they shared was "three short paragraphs at most",
 * which permits exactly the prose it was meant to prevent. So these tests are
 * about a contract, not a behaviour: the shape is stated once, every tool that
 * sends a body states it the same way, and the sentence that caused the drift
 * cannot come back.
 *
 * The last case is the one that makes the rest worth anything: the shape the
 * prompt asks for has to survive the markdown → Teams HTML conversion. Asking
 * for something the renderer mangles would be the same bug in a new place.
 */

import { test, describe } from "node:test";
import assert from "node:assert/strict";
import {
	CHAT_MESSAGE_BODY_DESCRIPTION,
	CHAT_SHAPE_FOR_LISTEN_MODE,
	CHAT_SHAPE_GUIDELINE,
} from "../src/utils/chat-style.ts";
import { teamsSendChatMessageTool } from "../src/tools/teams-send-chat-message.ts";
import {
	teamsReplyChannelMessageTool,
	teamsSendChannelMessageTool,
} from "../src/tools/teams-send-channel-message.ts";
import { teamsUpdateMessageTool } from "../src/tools/teams-update-message.ts";
import { composeWatchPrompt } from "../src/watch/prompt.ts";
import { markdownToTeamsHtml } from "../src/utils/richtext.ts";
import type { ChatSummary, MessageSummary, SignedInUser } from "../src/types.ts";

const NOW = new Date("2026-09-14T16:00:00Z").getTime();

const ME: SignedInUser = {
	id: "me-1",
	displayName: "Patrick Weppelmann",
	upn: "patrick@contoso.com",
	mail: "patrick@contoso.com",
	tenantId: "tenant-1",
};

const CHAT: ChatSummary = {
	id: "chat-1",
	chatType: "oneOnOne",
	label: "Anna Schmidt, Patrick Weppelmann",
	members: [{ displayName: "Anna Schmidt", mail: "anna.schmidt@contoso.com" }],
	lastUpdated: new Date(NOW - 60_000).toISOString(),
	lastMessageFrom: "Anna Schmidt",
};

const MESSAGE: MessageSummary = {
	id: "msg-1",
	text: "Kannst du kurz schauen?",
	contentType: "text",
	from: { id: "other-1", displayName: "Anna Schmidt", mail: "anna.schmidt@contoso.com" },
	createdDateTime: new Date(NOW - 60_000).toISOString(),
	mentions: [],
	reactions: [],
	attachments: [],
	imageUrls: [],
};

/** Every tool that puts a body into a Teams message. */
const SENDERS = [
	teamsSendChatMessageTool,
	teamsSendChannelMessageTool,
	teamsReplyChannelMessageTool,
	teamsUpdateMessageTool,
] as const;

function bodyDescription(tool: (typeof SENDERS)[number]): string {
	return (tool.parameters.properties as Record<string, { description?: string }>).body?.description ?? "";
}

function guidelines(tool: (typeof SENDERS)[number]): string[] {
	return "promptGuidelines" in tool ? ((tool as { promptGuidelines?: string[] }).promptGuidelines ?? []) : [];
}

describe("chat shape — one statement, everywhere", () => {
	test("every sending tool describes its body with the same shape", () => {
		for (const tool of SENDERS) {
			assert.equal(
				bodyDescription(tool),
				CHAT_MESSAGE_BODY_DESCRIPTION,
				`${tool.name} does not use the shared body description`,
			);
		}
	});

	test("the shape asks for the answer first and a hard line budget", () => {
		assert.match(CHAT_MESSAGE_BODY_DESCRIPTION, /the answer on the first line/);
		assert.match(CHAT_MESSAGE_BODY_DESCRIPTION, /at most 5 short lines/);
		assert.match(CHAT_MESSAGE_BODY_DESCRIPTION, /No greeting, no closing, no summary/);
	});

	test("every sending tool puts the shape in its system-prompt guidelines", () => {
		for (const tool of SENDERS) {
			assert.ok(
				guidelines(tool).includes(CHAT_SHAPE_GUIDELINE),
				`${tool.name} is missing the shared shape guideline`,
			);
		}
	});

	test("the guideline names every tool it applies to", () => {
		// promptGuidelines land flat in the Guidelines section with no tool-name
		// prefix, so an unnamed "the message body" would be unattributable.
		for (const tool of SENDERS) {
			assert.match(CHAT_SHAPE_GUIDELINE, new RegExp(tool.name));
		}
	});

	test("the guideline stays on one line", () => {
		assert.equal(CHAT_SHAPE_GUIDELINE.includes("\n"), false);
	});

	test("no longer offers paragraphs as the budget", () => {
		// The exact sentence that produced walls of text.
		const modelFacing = [
			CHAT_MESSAGE_BODY_DESCRIPTION,
			CHAT_SHAPE_GUIDELINE,
			CHAT_SHAPE_FOR_LISTEN_MODE,
			...SENDERS.flatMap((tool) => [bodyDescription(tool), ...guidelines(tool)]),
		];
		for (const text of modelFacing) {
			assert.doesNotMatch(text, /three short paragraphs/);
		}
	});
});

describe("chat shape — listen mode", () => {
	test("carries the shape, because nobody asked for the message", () => {
		const prompt = composeWatchPrompt({ chat: CHAT, message: MESSAGE, wokeAt: NOW, me: ME }, ME);
		assert.ok(prompt.includes(CHAT_SHAPE_FOR_LISTEN_MODE));
		assert.match(prompt, /readable at a glance/);
	});

	test("no longer tells the model a paragraph budget is fine", () => {
		const prompt = composeWatchPrompt({ chat: CHAT, message: MESSAGE, wokeAt: NOW, me: ME }, ME);
		assert.doesNotMatch(prompt, /three short paragraphs/);
	});
});

describe("chat shape — the renderer has to keep it", () => {
	test("renders the documented before/after shape as Teams HTML", () => {
		assert.equal(
			markdownToTeamsHtml("QA läuft wieder.\n\n- Ursache: abgelaufenes Zertifikat\n- Fix: neu deployt"),
			"<p>QA läuft wieder.</p><p>&nbsp;</p><ul><li>Ursache: abgelaufenes Zertifikat</li><li>Fix: neu deployt</li></ul>",
		);
	});

	test("a bold line and a link survive, a heading arrives as a bold line", () => {
		assert.equal(
			markdownToTeamsHtml("# Deployment\n\n**Fertig.** Release Notes: [Wiki](https://example.com/w)"),
			'<p><b>Deployment</b></p><p>&nbsp;</p><p><b>Fertig.</b> Release Notes: <a href="https://example.com/w">Wiki</a></p>',
		);
	});

	test("a one-line answer is a single paragraph, not a document", () => {
		assert.equal(markdownToTeamsHtml("Ja, Freitag passt."), "<p>Ja, Freitag passt.</p>");
	});
});
