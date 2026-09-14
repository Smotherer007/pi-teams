/**
 * The AI disclosure footer.
 *
 * A disclosure is only worth anything if it is always there and never twice, so
 * the tests are about the two things that can go wrong: forgetting it, and
 * stacking it.
 */

import { test, describe } from "node:test";
import assert from "node:assert/strict";
import { applyAiFooter, hasAiFooter } from "../src/utils/disclosure.ts";
import { AI_FOOTER_DEFAULT, resolveAiFooter, type ResolvedAiFooter } from "../src/config/index.ts";
import { formatMutationSummary } from "../src/safety/index.ts";

const ON: ResolvedAiFooter = { enabled: true, text: "🤖 Generated with pi (an AI agent)" };
const OFF: ResolvedAiFooter = { enabled: false, text: "🤖 Generated with pi (an AI agent)" };

describe("applyAiFooter", () => {
	test("leaves the body alone when the setting is off", () => {
		assert.equal(applyAiFooter("Hallo Anna", OFF), "Hallo Anna");
	});

	test("appends the disclosure after the body, separated by a blank line", () => {
		assert.equal(applyAiFooter("Hallo Anna", ON), `Hallo Anna\n\n${ON.text}`);
	});

	test("does not double a footer the body already carries", () => {
		const once = applyAiFooter("Hallo Anna", ON);
		assert.equal(applyAiFooter(once, ON), once);
		assert.equal(hasAiFooter(once, ON.text), true);
	});

	test("tolerates stray whitespace around the body", () => {
		assert.equal(applyAiFooter("Hallo Anna\n\n  ", ON), `Hallo Anna\n\n${ON.text}`);
	});

	test("never footers an empty body", () => {
		assert.equal(applyAiFooter("", ON), "");
		assert.equal(applyAiFooter("   ", ON), "   ");
	});

	test("has nothing to add when the footer text is blank", () => {
		assert.equal(applyAiFooter("Hallo Anna", { enabled: true, text: "  " }), "Hallo Anna");
	});

	test("escapes and wraps the text when the body is raw HTML", () => {
		assert.equal(
			applyAiFooter("<p>Hallo</p>", { enabled: true, text: "KI & <pi>" }, { html: true }),
			"<p>Hallo</p>\n<p>KI &amp; &lt;pi&gt;</p>",
		);
	});

	test("recognises an HTML footer as already present", () => {
		const once = applyAiFooter("<p>Hallo</p>", ON, { html: true });
		assert.equal(applyAiFooter(once, ON, { html: true }), once);
	});
});

describe("resolveAiFooter", () => {
	test("is on unless something switches it off", () => {
		assert.deepEqual(resolveAiFooter(undefined, undefined), {
			enabled: true,
			text: AI_FOOTER_DEFAULT,
		});
	});

	test("an account can switch the default off", () => {
		assert.equal(resolveAiFooter(undefined, { enabled: false }).enabled, false);
		assert.equal(resolveAiFooter({ enabled: true }, { enabled: false }).enabled, false);
	});

	test("the account overrides the global setting field by field", () => {
		assert.deepEqual(resolveAiFooter({ enabled: true }, { text: "Nur Text" }), {
			enabled: true,
			text: "Nur Text",
		});
		assert.deepEqual(resolveAiFooter({ enabled: true, text: "Global" }, { enabled: false }), {
			enabled: false,
			text: "Global",
		});
	});

	test("a whitespace-only text falls back to the default wording", () => {
		assert.equal(resolveAiFooter({ enabled: true, text: "   " }, undefined).text, AI_FOOTER_DEFAULT);
		assert.equal(resolveAiFooter(undefined, { enabled: true, text: "  " }).text, AI_FOOTER_DEFAULT);
	});
});

describe("the confirmation dialog", () => {
	test("shows the disclosure, so the user confirms what is actually sent", () => {
		const summary = formatMutationSummary(
			"teams_send_chat_message",
			{ chat: "Anna Schmidt", body: "Hallo Anna" },
			ON,
		);
		assert.match(summary, /Hallo Anna/);
		assert.match(summary, /Generated with pi/);
	});

	test("shows the first message of a new chat, which used to be invisible", () => {
		const summary = formatMutationSummary(
			"teams_create_chat",
			{ participants: ["anna@contoso.com"], message: "Hallo Anna" },
			ON,
		);
		assert.match(summary, /First message/);
		assert.match(summary, /Generated with pi/);
	});

	test("says nothing extra while the setting is off", () => {
		const summary = formatMutationSummary(
			"teams_send_chat_message",
			{ chat: "Anna Schmidt", body: "Hallo Anna" },
			OFF,
		);
		assert.doesNotMatch(summary, /Generated with pi/);
	});
});
