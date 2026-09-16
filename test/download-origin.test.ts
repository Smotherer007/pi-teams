/**
 * The access token must never leave the Graph origin.
 *
 * Message content is written by whoever sent the message: an attachment's
 * `contentUrl` and an `<img src>` in the body are both attacker-controlled. If
 * a download attached the bearer token to one of those, anyone able to message
 * the user could collect a delegated Teams token — Chat.ReadWrite,
 * ChannelMessage.Send, Calendars.ReadWrite — and act as them for its lifetime.
 *
 * These are regression tests for exactly that.
 */

import { test, describe } from "node:test";
import assert from "node:assert/strict";
import { graphOrigin, isGraphUrl, sharesContentPath } from "../src/graph/client.ts";
import { imageUrlsInHtml } from "../src/utils/attachments.ts";
import type { TeamsConnection } from "../src/config/index.ts";

const conn = { graphBaseUrl: "https://graph.microsoft.com/v1.0" } as TeamsConnection;
const sovereign = { graphBaseUrl: "https://microsoftgraph.chinacloudapi.cn/v1.0" } as TeamsConnection;

describe("graphOrigin", () => {
	test("is derived from the configured base URL, not hardcoded", () => {
		assert.equal(graphOrigin(conn), "https://graph.microsoft.com");
		assert.equal(graphOrigin(sovereign), "https://microsoftgraph.chinacloudapi.cn");
	});
});

describe("isGraphUrl", () => {
	test("accepts the tenant's own Graph endpoint", () => {
		assert.equal(
			isGraphUrl(conn, "https://graph.microsoft.com/v1.0/chats/x/messages/y/hostedContents/1/$value"),
			true,
		);
	});

	test("accepts a relative path — ours by construction", () => {
		assert.equal(isGraphUrl(conn, "/me/chats"), true);
	});

	test("refuses a foreign host that imitates a Graph path", () => {
		// The exact shape an attacker would choose: the path looks like hosted
		// content, only the host differs.
		assert.equal(
			isGraphUrl(conn, "https://attacker.example.net/hostedContents/1/$value"),
			false,
		);
	});

	test("refuses a host that merely contains the Graph host", () => {
		assert.equal(isGraphUrl(conn, "https://graph.microsoft.com.evil.example/x"), false);
		assert.equal(isGraphUrl(conn, "https://evil.example/graph.microsoft.com/x"), false);
	});

	test("refuses a downgrade to http", () => {
		assert.equal(isGraphUrl(conn, "http://graph.microsoft.com/v1.0/me"), false);
	});

	test("refuses userinfo trickery", () => {
		assert.equal(isGraphUrl(conn, "https://graph.microsoft.com@evil.example/x"), false);
	});

	test("refuses a non-http scheme", () => {
		assert.equal(isGraphUrl(conn, "file:///etc/passwd"), false);
		assert.equal(isGraphUrl(conn, "data:text/plain;base64,AAAA"), false);
	});

	test("does not treat one cloud's endpoint as another's", () => {
		assert.equal(isGraphUrl(sovereign, "https://graph.microsoft.com/v1.0/me"), false);
	});
});

describe("sharesContentPath", () => {
	test("encodes a SharePoint URL into Graph's own addressing", () => {
		const path = sharesContentPath("https://contoso.sharepoint.com/sites/x/Shared Documents/a.pdf");
		assert.ok(path.startsWith("/shares/u!"));
		assert.ok(path.endsWith("/driveItem/content"));
	});

	test("produces base64url — no padding, no + or /", () => {
		const path = sharesContentPath("https://contoso.sharepoint.com/sites/x/y/z/????>>>>");
		const token = path.slice("/shares/u!".length, -"/driveItem/content".length);
		assert.ok(!token.includes("="), "no padding");
		assert.ok(!token.includes("+"), "no plus");
		assert.ok(!token.includes("/"), "no slash — it would split the Graph path");
	});

	test("the result stays on the Graph origin", () => {
		// The whole point: a foreign URL becomes a Graph request, never a request
		// to the foreign host.
		assert.equal(isGraphUrl(conn, sharesContentPath("https://attacker.example.net/x")), true);
	});
});

describe("inline image extraction", () => {
	test("still finds a genuine hosted image", () => {
		const html = '<img src="https://graph.microsoft.com/v1.0/chats/1/messages/2/hostedContents/3/$value">';
		assert.deepEqual(imageUrlsInHtml(html), [
			"https://graph.microsoft.com/v1.0/chats/1/messages/2/hostedContents/3/$value",
		]);
	});

	test("a foreign hosted-content URL survives extraction but fails the origin check", () => {
		// Extraction is deliberately permissive; the origin check is the gate, and
		// it is the one this test pins down.
		const [url] = imageUrlsInHtml('<img src="https://attacker.example.net/hostedContents/1/$value">');
		assert.ok(url);
		assert.equal(isGraphUrl(conn, url), false);
	});
});
