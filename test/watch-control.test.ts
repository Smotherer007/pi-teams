/**
 * Control words: only a message that is nothing but the word counts.
 */

import { test } from "node:test";
import assert from "node:assert/strict";
import { controlAction, normalizeCommand } from "../src/watch/control.ts";
import { resolveDispatchConfig } from "../src/config/index.ts";

const dispatch = resolveDispatchConfig({ mode: "process" });
const msg = (text: string, mentions: string[] = []) => ({ text, mentions: mentions.map((displayName) => ({ displayName })) });

test("the bare words are recognised, case and punctuation aside", () => {
	assert.equal(controlAction(msg("Stopp!"), dispatch), "stop");
	assert.equal(controlAction(msg("abbrechen"), dispatch), "stop");
	assert.equal(controlAction(msg("Neues Thema."), dispatch), "reset");
});

test("a mention of pi in front does not hide the word", () => {
	assert.equal(controlAction(msg("Neo stopp", ["Neo"]), dispatch), "stop");
	assert.equal(controlAction(msg("@Neo neues Thema"), dispatch, "Neo (neoimpulse)"), "reset");
});

test("a request that merely contains the word is not a control word", () => {
	assert.equal(controlAction(msg("stopp bitte das Deployment auf hera"), dispatch), undefined);
	assert.equal(controlAction(msg("lass uns ein neues Thema anfangen: EWM"), dispatch), undefined);
});

test("normalizeCommand keeps words that only contain a name", () => {
	assert.equal(normalizeCommand("Neonlicht", ["Neo"]), "neonlicht");
});
