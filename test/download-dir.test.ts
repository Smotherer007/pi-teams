/**
 * Where a download may land.
 *
 * The directory is a tool parameter, so the model picks it — while reading
 * messages other people wrote. Confinement is what keeps "save the picture" from
 * becoming "write into ~/.ssh".
 */

import { test, describe } from "node:test";
import assert from "node:assert/strict";
import { join, resolve } from "node:path";
import { homedir, tmpdir } from "node:os";
import { resolveDownloadDir } from "../src/tools/teams-download-files.ts";
import { getAgentDir } from "../src/config/index.ts";

const FALLBACK = join(getAgentDir(), "pi-teams-files", "anna");

describe("resolveDownloadDir", () => {
	test("uses the fallback when nothing was asked for", () => {
		assert.equal(resolveDownloadDir(undefined, FALLBACK, undefined), FALLBACK);
	});

	test("a configured directory becomes the default", () => {
		const configured = join(tmpdir(), "teams-downloads");
		assert.equal(resolveDownloadDir(undefined, FALLBACK, configured), resolve(configured));
	});

	test("allows a path inside the working directory", () => {
		const target = join(process.cwd(), "downloads", "teams");
		assert.equal(resolveDownloadDir(target, FALLBACK, undefined), resolve(target));
	});

	test("allows a path inside the agent directory", () => {
		const target = join(getAgentDir(), "pi-teams-files", "x");
		assert.equal(resolveDownloadDir(target, FALLBACK, undefined), resolve(target));
	});

	test("refuses somewhere else entirely", () => {
		assert.throws(() => resolveDownloadDir(join(homedir(), ".ssh"), FALLBACK, undefined), /Refusing/);
		assert.throws(() => resolveDownloadDir("/etc", FALLBACK, undefined), /Refusing/);
	});

	test("refuses an escape through ..", () => {
		assert.throws(
			() => resolveDownloadDir(join(process.cwd(), "..", "..", "elsewhere"), FALLBACK, undefined),
			/Refusing/,
		);
	});

	test("refuses a sibling whose name merely starts with an allowed root", () => {
		// `/home/me/project-evil` must not pass because `/home/me/project` does.
		assert.throws(() => resolveDownloadDir(`${process.cwd()}-evil`, FALLBACK, undefined), /Refusing/);
	});

	test("a configured directory widens the allowance", () => {
		const configured = join(tmpdir(), "teams-downloads");
		const target = join(configured, "anna");
		assert.throws(() => resolveDownloadDir(target, FALLBACK, undefined), /Refusing/);
		assert.equal(resolveDownloadDir(target, FALLBACK, configured), resolve(target));
	});

	test("the refusal names the allowed roots, so the user can act on it", () => {
		try {
			resolveDownloadDir("/etc", FALLBACK, undefined);
			assert.fail("expected a refusal");
		} catch (err) {
			assert.match(String(err), /downloadDir/);
			assert.match(String(err), new RegExp(resolve(process.cwd()).replace(/[.*+?^${}()|[\]\\]/g, "\\$&")));
		}
	});
});
