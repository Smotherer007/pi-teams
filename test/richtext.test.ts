/**
 * Teams message formatting.
 *
 * These rules decide what a colleague actually sees in the chat bubble, so the
 * cases below are written from the reader's side: would this look right in
 * Teams, and would plain text survive untouched?
 */

import { test, describe } from "node:test";
import assert from "node:assert/strict";
import { markdownToTeamsHtml } from "../src/utils/richtext.ts";
import { htmlToText } from "../src/utils/formatting.ts";

describe("markdownToTeamsHtml — paragraphs", () => {
	test("keeps a blank line as a paragraph break and a single newline as a <br>", () => {
		assert.equal(markdownToTeamsHtml("one\n\ntwo"), "<p>one</p><p>two</p>");
		assert.equal(markdownToTeamsHtml("one\ntwo"), "<p>one<br>two</p>");
	});

	test("escapes HTML so a message cannot inject markup", () => {
		assert.equal(markdownToTeamsHtml("a<b>c&d"), "<p>a&lt;b&gt;c&amp;d</p>");
		assert.equal(markdownToTeamsHtml("<script>alert(1)</script>"), "<p>&lt;script&gt;alert(1)&lt;/script&gt;</p>");
	});

	test("drops runs of blank lines instead of emitting empty paragraphs", () => {
		assert.equal(markdownToTeamsHtml("a\n\n\n\nb"), "<p>a</p><p>b</p>");
	});
});

describe("markdownToTeamsHtml — inline", () => {
	test("bold, italic, strike and code", () => {
		assert.equal(markdownToTeamsHtml("**fett**"), "<p><b>fett</b></p>");
		assert.equal(markdownToTeamsHtml("*kursiv*"), "<p><i>kursiv</i></p>");
		assert.equal(markdownToTeamsHtml("_kursiv_"), "<p><i>kursiv</i></p>");
		assert.equal(markdownToTeamsHtml("~~weg~~"), "<p><s>weg</s></p>");
		assert.equal(markdownToTeamsHtml("`code`"), "<p><code>code</code></p>");
	});

	test("leaves underscores inside a word alone", () => {
		assert.equal(markdownToTeamsHtml("snake_case_name"), "<p>snake_case_name</p>");
		assert.equal(markdownToTeamsHtml("die Datei ist order_id_42 wert"), "<p>die Datei ist order_id_42 wert</p>");
	});

	test("keeps emphasis inside a code span literal", () => {
		assert.equal(markdownToTeamsHtml("`a ** b`"), "<p><code>a ** b</code></p>");
	});

	test("link with label and href", () => {
		assert.equal(
			markdownToTeamsHtml("[Docs](https://example.com/a?b=1)"),
			'<p><a href="https://example.com/a?b=1">Docs</a></p>',
		);
	});

	test("refuses anything that is not http(s) or mailto", () => {
		assert.equal(
			markdownToTeamsHtml("[klick](javascript:alert(1))"),
			'<p>[klick](javascript:alert(1))</p>',
		);
	});

	test("bare asterisks survive", () => {
		assert.equal(markdownToTeamsHtml("3 * 4 = 12"), "<p>3 * 4 = 12</p>");
	});
});

describe("markdownToTeamsHtml — blocks", () => {
	test("bullets become a list", () => {
		assert.equal(
			markdownToTeamsHtml("- eins\n- zwei"),
			"<ul><li>eins</li><li>zwei</li></ul>",
		);
	});

	test("numbered lines become an ordered list, starting at any number", () => {
		assert.equal(
			markdownToTeamsHtml("1. erstens\n2. zweitens"),
			"<ol><li>erstens</li><li>zweitens</li></ol>",
		);
		assert.equal(markdownToTeamsHtml("3. drittens"), "<ol><li>drittens</li></ol>");
	});

	test("switching between list kinds starts a new list", () => {
		assert.equal(
			markdownToTeamsHtml("- eins\n1. zwei"),
			"<ul><li>eins</li></ul><ol><li>zwei</li></ol>",
		);
	});

	test("a heading becomes a bold line — Teams has no headings in a chat", () => {
		assert.equal(markdownToTeamsHtml("## Lage"), "<p><b>Lage</b></p>");
	});

	test("a fenced block keeps its text verbatim", () => {
		assert.equal(
			markdownToTeamsHtml("```\nif (a < b) {}\n```"),
			"<pre>if (a &lt; b) {}</pre>",
		);
	});

	test("an unterminated fence is still code", () => {
		assert.equal(markdownToTeamsHtml("```\nrest"), "<pre>rest</pre>");
	});

	test("a realistic answer keeps its structure", () => {
		const html = markdownToTeamsHtml(
			"Deployment ist durch.\n\n- DEV: grün\n- QA: grün\n\nNächster Schritt: **Freigabe** von @Anna Schmidt.",
		);
		assert.equal(
			html,
			"<p>Deployment ist durch.</p><ul><li>DEV: grün</li><li>QA: grün</li></ul>" +
				"<p>Nächster Schritt: <b>Freigabe</b> von @Anna Schmidt.</p>",
		);
	});

	test("round-trips through htmlToText back to readable text", () => {
		// The reader gets every line back in order. A blank line before a list is a
		// paragraph break we put in and htmlToText does not re-invent — the text is
		// unchanged, only that one empty line disappears.
		const original = "Lage\n\n- DEV grün\n- QA rot";
		const text = htmlToText(markdownToTeamsHtml(original));
		assert.equal(text, "Lage\n- DEV grün\n- QA rot");
	});
});

describe("markdownToTeamsHtml — edges", () => {
	test("empty and whitespace-only input", () => {
		assert.equal(markdownToTeamsHtml(""), "");
		assert.equal(markdownToTeamsHtml("   \n  "), "");
	});

	test("normalises CRLF", () => {
		assert.equal(markdownToTeamsHtml("a\r\n\r\nb"), "<p>a</p><p>b</p>");
	});

	test("a list directly after a paragraph is not swallowed by it", () => {
		assert.equal(
			markdownToTeamsHtml("Ergebnis:\n- eins"),
			"<p>Ergebnis:</p><ul><li>eins</li></ul>",
		);
	});
});
