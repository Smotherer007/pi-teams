/**
 * The newest message as a command: what a router compares control words to.
 */

import { test } from "node:test";
import assert from "node:assert/strict";
import { commandText, normalizeCommand } from "../src/watch/control.ts";

const msg = (text: string, mentions: string[] = []) => ({ text, mentions: mentions.map((displayName) => ({ displayName })) });

test("case and punctuation are dropped", () => {
	assert.equal(commandText(msg("Stopp!")), "stopp");
	assert.equal(commandText(msg("Neues Thema.")), "neues thema");
});

test("a mention of pi in front is dropped", () => {
	assert.equal(commandText(msg("Neo stopp", ["Neo"])), "stopp");
	assert.equal(commandText(msg("@Neo neues Thema"), "Neo (neoimpulse)"), "neues thema");
});

test("a longer request stays a request", () => {
	assert.equal(commandText(msg("stopp bitte das Deployment auf hera")), "stopp bitte das deployment auf hera");
});

test("words that only contain a name are kept", () => {
	assert.equal(normalizeCommand("Neonlicht", ["Neo"]), "neonlicht");
});
