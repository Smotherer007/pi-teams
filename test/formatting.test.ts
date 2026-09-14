/**
 * Formatting — the agent reads Teams entirely through these functions, so a
 * bug here shows up as pi misquoting a colleague.
 */

import { test, describe } from "node:test";
import assert from "node:assert/strict";
import {
	formatBytes,
	formatMessage,
	formatMessageList,
	formatPresence,
	htmlToText,
	textToHtml,
	truncate,
} from "../src/utils/formatting.ts";
import type { MessageSummary } from "../src/types.ts";

const message = (overrides: Partial<MessageSummary> = {}): MessageSummary => ({
	id: "1700000000000",
	text: "Hello there",
	contentType: "html",
	from: { displayName: "Anna Schmidt", upn: "anna@contoso.com" },
	createdDateTime: "2026-09-14T08:31:00Z",
	mentions: [],
	reactions: [],
	attachments: [],
	...overrides,
});

describe("htmlToText", () => {
	test("keeps line structure from block elements", () => {
		assert.equal(htmlToText("<p>One</p><p>Two</p>"), "One\nTwo");
		assert.equal(htmlToText("A<br>B"), "A\nB");
	});

	test("keeps link targets", () => {
		assert.equal(
			htmlToText('<a href="https://example.com">docs</a>'),
			"docs (https://example.com)",
		);
	});

	test("does not duplicate a link whose text is its URL", () => {
		assert.equal(
			htmlToText('<a href="https://example.com">https://example.com</a>'),
			"https://example.com",
		);
	});

	test("names images instead of dropping them silently", () => {
		assert.equal(htmlToText('<img alt="chart">'), "[image: chart]");
		assert.equal(htmlToText("<img src='x.png'>"), "[image]");
	});

	test("decodes entities", () => {
		assert.equal(htmlToText("<p>A &amp; B &lt;c&gt;&nbsp;d</p>"), "A & B <c> d");
		assert.equal(htmlToText("<p>&#65;</p>"), "A");
	});

	test("turns list items into bullets", () => {
		assert.equal(htmlToText("<ul><li>one</li><li>two</li></ul>"), "- one\n- two");
	});

	test("leaves plain text untouched", () => {
		assert.equal(htmlToText("  already plain  ", "text"), "already plain");
	});

	test("collapses runs of blank lines", () => {
		assert.equal(htmlToText("<p>a</p><p></p><p></p><p>b</p>"), "a\n\nb");
	});
});

describe("textToHtml", () => {
	test("escapes markup and keeps paragraphs", () => {
		assert.equal(textToHtml("a<b>"), "<p>a&lt;b&gt;</p>");
		assert.equal(textToHtml("one\n\ntwo"), "<p>one</p><p>two</p>");
		assert.equal(textToHtml("one\ntwo"), "<p>one<br>two</p>");
	});

	test("round-trips through htmlToText", () => {
		const original = "Line one\nLine two";
		assert.equal(htmlToText(textToHtml(original)), original);
	});
});

describe("truncate", () => {
	test("collapses newlines and adds an ellipsis only when needed", () => {
		assert.equal(truncate("a\n b", 10), "a b");
		assert.equal(truncate("abcdefghij", 5), "abcd…");
	});
});

describe("formatBytes", () => {
	test("scales units", () => {
		assert.equal(formatBytes(512), "512 B");
		assert.equal(formatBytes(2048), "2.0 KB");
		assert.equal(formatBytes(5 * 1024 * 1024), "5.0 MB");
		assert.equal(formatBytes(undefined), "");
	});
});

describe("formatMessage", () => {
	test("shows sender, body and message id", () => {
		const text = formatMessage(message());
		assert.match(text, /Anna Schmidt/);
		assert.match(text, /Hello there/);
		assert.match(text, /messageId: 1700000000000/);
	});

	test("marks a deleted message instead of showing an empty body", () => {
		const text = formatMessage(message({ text: "", deletedDateTime: "2026-09-14T09:00:00Z" }));
		assert.match(text, /message deleted/);
	});

	test("groups repeated reactions", () => {
		const text = formatMessage(
			message({ reactions: [{ type: "👍" }, { type: "👍" }, { type: "🎉" }] }),
		);
		assert.match(text, /👍×2/);
		assert.match(text, /🎉/);
	});

	test("names attachments", () => {
		const text = formatMessage(message({ attachments: [{ name: "spec.pdf" }] }));
		assert.match(text, /spec\.pdf/);
	});
});

describe("formatMessageList", () => {
	test("reports an empty conversation plainly", () => {
		assert.match(formatMessageList([], "Team/General"), /no messages/);
	});

	test("orders oldest first — Graph returns newest first", () => {
		const older = message({ id: "1", text: "first", createdDateTime: "2026-09-14T08:00:00Z" });
		const newer = message({ id: "2", text: "second", createdDateTime: "2026-09-14T09:00:00Z" });
		const text = formatMessageList([newer, older], "Chat");
		assert.ok(text.indexOf("first") < text.indexOf("second"));
	});
});

describe("formatPresence", () => {
	test("includes an icon, the availability and the status message", () => {
		const text = formatPresence({
			id: "1",
			displayName: "Anna",
			availability: "Busy",
			activity: "InAMeeting",
			statusMessage: "Workshop until 5",
		});
		assert.match(text, /🔴/);
		assert.match(text, /Anna: Busy \(InAMeeting\)/);
		assert.match(text, /Workshop until 5/);
	});
});
