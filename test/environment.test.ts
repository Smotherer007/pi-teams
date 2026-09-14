/**
 * Browser detection.
 *
 * This decides whether teams_login opens a browser or shows a code, so getting
 * it wrong either hangs an SSH session waiting for a browser nobody can see, or
 * makes a perfectly capable desktop ask the user to type a code.
 */

import { test, describe, beforeEach, afterEach } from "node:test";
import assert from "node:assert/strict";
import { browserUnavailableReason, canOpenBrowser } from "../src/utils/environment.ts";

const TOUCHED = [
	"PI_TEAMS_NO_BROWSER",
	"SSH_CONNECTION",
	"SSH_TTY",
	"SSH_CLIENT",
	"DISPLAY",
	"WAYLAND_DISPLAY",
];

let saved: Record<string, string | undefined> = {};
const realPlatform = process.platform;

function setPlatform(value: NodeJS.Platform): void {
	Object.defineProperty(process, "platform", { value, configurable: true });
}

beforeEach(() => {
	saved = Object.fromEntries(TOUCHED.map((key) => [key, process.env[key]]));
	for (const key of TOUCHED) delete process.env[key];
});

afterEach(() => {
	for (const [key, value] of Object.entries(saved)) {
		if (value === undefined) delete process.env[key];
		else process.env[key] = value;
	}
	setPlatform(realPlatform);
});

describe("canOpenBrowser", () => {
	test("a desktop macOS or Windows session can", () => {
		setPlatform("darwin");
		assert.equal(canOpenBrowser(), true);

		setPlatform("win32");
		assert.equal(canOpenBrowser(), true);
	});

	test("Linux needs a display server", () => {
		setPlatform("linux");
		assert.equal(canOpenBrowser(), false);

		process.env.DISPLAY = ":0";
		assert.equal(canOpenBrowser(), true);

		delete process.env.DISPLAY;
		process.env.WAYLAND_DISPLAY = "wayland-0";
		assert.equal(canOpenBrowser(), true);
	});

	test("an SSH session cannot, even on a Mac", () => {
		// `open` would launch a browser on the remote machine, where nobody is
		// sitting — the single most confusing possible failure.
		setPlatform("darwin");
		process.env.SSH_CONNECTION = "10.0.0.1 22 10.0.0.2 22";
		assert.equal(canOpenBrowser(), false);

		delete process.env.SSH_CONNECTION;
		process.env.SSH_TTY = "/dev/pts/0";
		assert.equal(canOpenBrowser(), false);
	});

	test("PI_TEAMS_NO_BROWSER forces the fallback anywhere", () => {
		setPlatform("darwin");
		process.env.PI_TEAMS_NO_BROWSER = "1";
		assert.equal(canOpenBrowser(), false);
	});
});

describe("browserUnavailableReason", () => {
	test("names the env var when it is the cause", () => {
		process.env.PI_TEAMS_NO_BROWSER = "1";
		assert.match(browserUnavailableReason(), /PI_TEAMS_NO_BROWSER/);
	});

	test("explains the SSH case in terms of the wrong machine", () => {
		process.env.SSH_CONNECTION = "x";
		assert.match(browserUnavailableReason(), /wrong machine/);
	});

	test("names the display variables on Linux", () => {
		setPlatform("linux");
		assert.match(browserUnavailableReason(), /DISPLAY/);
	});
});
