/**
 * The browser handoff.
 *
 * Everything here fails quietly in production, which is why it is tested at
 * all. A missing launcher reports ENOENT on an event nobody was waiting for, a
 * launcher that finds no browser says so only in its exit code, and a URL given
 * to `cmd.exe` is cut at its first `&`. In each case the sign-in did not fail —
 * it sat waiting for a redirect until the timeout, and reported the timeout.
 *
 * No test here opens a real browser: `browserLaunch` is asked for its arguments
 * and `launchBrowser` is driven with node itself.
 */

import { test, describe, beforeEach, afterEach } from "node:test";
import assert from "node:assert/strict";
import { browserLaunch, launchBrowser } from "../src/auth/msal.ts";

const realPlatform = process.platform;
const SCRUBBED = ["WSL_DISTRO_NAME", "WSL_INTEROP", "PI_TEAMS_BROWSER"];
let savedEnv: Record<string, string | undefined> = {};

function setPlatform(value: NodeJS.Platform): void {
	Object.defineProperty(process, "platform", { value, configurable: true });
}

/** Shaped like the authorize URL MSAL hands to `openBrowser`. */
const URL_WITH_QUERY =
	"https://login.microsoftonline.com/common/oauth2/v2.0/authorize" +
	"?client_id=abc&response_type=code&redirect_uri=http%3A%2F%2Flocalhost%3A3000&scope=User.Read";

beforeEach(() => {
	savedEnv = Object.fromEntries(SCRUBBED.map((key) => [key, process.env[key]]));
	for (const key of SCRUBBED) delete process.env[key];
});

afterEach(() => {
	for (const [key, value] of Object.entries(savedEnv)) {
		if (value === undefined) delete process.env[key];
		else process.env[key] = value;
	}
	setPlatform(realPlatform);
});

describe("browserLaunch", () => {
	test("macOS uses open", () => {
		assert.deepEqual(browserLaunch("https://example.com", { platform: "darwin", wsl: false }), {
			command: "open",
			args: ["https://example.com"],
		});
	});

	test("Windows uses cmd start, with the empty title it needs", () => {
		assert.deepEqual(browserLaunch("https://example.com", { platform: "win32", wsl: false }), {
			command: "cmd.exe",
			args: ["/c", "start", "", "https://example.com"],
		});
	});

	test("plain Linux uses xdg-open", () => {
		assert.deepEqual(browserLaunch("https://example.com", { platform: "linux", wsl: false }), {
			command: "xdg-open",
			args: ["https://example.com"],
		});
	});

	test("WSL hands the URL to Windows PowerShell instead", () => {
		const launch = browserLaunch("https://example.com", { platform: "linux", wsl: true });
		assert.match(launch.command, /WindowsPowerShell\/v1\.0\/powershell\.exe$/);
		assert.deepEqual(launch.args.slice(0, -1), [
			"-NoProfile",
			"-NonInteractive",
			// Without Hidden, a console window flashes on every sign-in.
			"-WindowStyle",
			"Hidden",
			"-Command",
		]);
		assert.equal(launch.args.at(-1), "Start-Process 'https://example.com'");
	});

	test("an authorize URL keeps every & through the handoff", () => {
		// The regression: cmd.exe treats the first & as a command separator, so
		// the browser got a truncated URL and the sign-in waited for a redirect
		// that could never arrive. The URL has to arrive whole, in one argument.
		const launch = browserLaunch(URL_WITH_QUERY, { platform: "linux", wsl: true });
		const payload = launch.args.at(-1) as string;
		assert.equal(launch.args.length, 6, "the URL must stay a single argument");
		assert.equal(payload.split("&").length, URL_WITH_QUERY.split("&").length);
		assert.ok(payload.includes(URL_WITH_QUERY));
	});

	test("a single quote in the URL is doubled, not dropped", () => {
		// PowerShell ends a single-quoted string at the first quote, so an
		// unescaped one would truncate the URL just as fatally as & does.
		const launch = browserLaunch("https://example.com/a'b", { platform: "linux", wsl: true });
		assert.equal(launch.args.at(-1), "Start-Process 'https://example.com/a''b'");
	});

	test("WSL is recognised from the environment when no target is given", () => {
		setPlatform("linux");
		process.env.WSL_DISTRO_NAME = "Ubuntu";
		assert.match(browserLaunch("https://example.com").command, /powershell\.exe$/);
	});

	test("WSL without the marker falls back to xdg-open", () => {
		setPlatform("linux");
		assert.equal(browserLaunch("https://example.com").command, "xdg-open");
	});

	test("PI_TEAMS_BROWSER replaces the platform default", () => {
		setPlatform("linux");
		process.env.PI_TEAMS_BROWSER = "my-opener --quiet";
		assert.deepEqual(browserLaunch("https://example.com"), {
			command: "my-opener",
			args: ["--quiet", "https://example.com"],
		});
	});

	test("PI_TEAMS_BROWSER puts the URL where the {} is", () => {
		// So a launcher that wants the URL in the middle of its own flags works
		// without a wrapper script around it.
		process.env.PI_TEAMS_BROWSER = "/usr/local/bin/open --url={} --new-window";
		assert.deepEqual(browserLaunch("https://example.com"), {
			command: "/usr/local/bin/open",
			args: ["--url=https://example.com", "--new-window"],
		});
	});

	test("a quoted PI_TEAMS_BROWSER keeps its arguments together", () => {
		process.env.PI_TEAMS_BROWSER = '"/mnt/c/Program Files/App/open.exe" --quiet';
		assert.deepEqual(browserLaunch("https://example.com"), {
			command: "/mnt/c/Program Files/App/open.exe",
			args: ["--quiet", "https://example.com"],
		});
	});

	test("the override beats the WSL handoff, and keeps the & intact", () => {
		setPlatform("linux");
		process.env.WSL_DISTRO_NAME = "Ubuntu";
		process.env.PI_TEAMS_BROWSER = "my-opener";
		const launch = browserLaunch(URL_WITH_QUERY);
		assert.equal(launch.command, "my-opener");
		assert.equal(launch.args.at(-1), URL_WITH_QUERY);
	});
});

describe("launchBrowser", () => {
	test("rejects when the launcher is not installed", async () => {
		// This is the WSL case: no xdg-open in the distro. Reporting it turns a
		// five minute timeout into an error that names the cause.
		await assert.rejects(
			launchBrowser("/nonexistent/definitely-not-a-browser-launcher", []),
			/was not found/,
		);
	});

	test("rejects when the launcher exits non-zero, keeping its stderr", async () => {
		// And this is the case xdg-open uses to say it found no browser.
		await assert.rejects(
			launchBrowser(process.execPath, ["-e", "console.error('no browser found'); process.exit(3)"]),
			/code 3: no browser found/,
		);
	});

	test("resolves when the launcher succeeds", async () => {
		await launchBrowser(process.execPath, ["-e", "process.exit(0)"]);
	});

	test("resolves when the launcher stays alive, waiting on the browser", async () => {
		// A launcher that neither exits nor errors is what a successful handoff
		// looks like, so it must not be reported as a failure.
		await launchBrowser(process.execPath, ["-e", "setTimeout(() => process.exit(0), 300)"], 50);
	});
});
