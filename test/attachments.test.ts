/**
 * What a message carries besides text.
 *
 * The fixture is the shape Teams actually sends: an `<img>` whose `src` is a
 * Graph URL into the message's own `hostedContents`. That URL is the whole
 * point — it was in the body all along, and treating the body as text threw it
 * away, which is why a screenshot used to reach a reader as `[image: Bild]`.
 */

import { test, describe } from "node:test";
import assert from "node:assert/strict";
import {
	extensionForContentType,
	imageUrlsInHtml,
	safeFileName,
} from "../src/utils/attachments.ts";
import { mapMessage } from "../src/graph/mappers.ts";

const CHAT_ID = "19:abc_def@unq.gbl.spaces";
const MESSAGE_ID = "1789471823027";
const HOSTED = (id: string) =>
	`https://graph.microsoft.com/v1.0/chats/${CHAT_ID}/messages/${MESSAGE_ID}/hostedContents/${id}/$value`;

describe("imageUrlsInHtml", () => {
	test("finds the hosted content of an inline image", () => {
		const html = `<p><img src="${HOSTED("aWQ9eF8w")}" width="546" height="250" alt="Bild"></p>`;
		assert.deepEqual(imageUrlsInHtml(html), [HOSTED("aWQ9eF8w")]);
	});

	test("keeps the order and drops duplicates", () => {
		const html =
			`<p><img src="${HOSTED("one")}" alt="Bild"><img src="${HOSTED("two")}" alt="Bild"></p>` +
			`<p>&nbsp;</p><p><img src="${HOSTED("one")}" alt="Bild"></p>`;
		assert.deepEqual(imageUrlsInHtml(html), [HOSTED("one"), HOSTED("two")]);
	});

	test("ignores images that are not part of the message", () => {
		// An <img> can point at an outside host, and there is nothing of ours
		// there to fetch.
		const html = `<p><img src="https://example.com/logo.png" alt="logo"></p>`;
		assert.deepEqual(imageUrlsInHtml(html), []);
	});

	test("decodes entities in the URL", () => {
		const html = `<img src="${HOSTED("a")}?x=1&amp;y=2">`;
		assert.deepEqual(imageUrlsInHtml(html), [`${HOSTED("a")}?x=1&y=2`]);
	});

	test("finds nothing in a text-only body", () => {
		assert.deepEqual(imageUrlsInHtml("<p>hello world</p>"), []);
		assert.deepEqual(imageUrlsInHtml(""), []);
	});
});

describe("mapMessage", () => {
	test("carries the image URLs through from the raw body", () => {
		// The regression: htmlToText turns this into "[image: Bild]" and the URL
		// never reached anyone, so nothing could fetch the picture.
		const summary = mapMessage({
			id: MESSAGE_ID,
			body: { contentType: "html", content: `<p><img src="${HOSTED("x")}" alt="Bild"></p>` },
		});
		assert.deepEqual(summary.imageUrls, [HOSTED("x")]);
		assert.match(summary.text, /\[image: Bild\]/);
	});

	test("leaves imageUrls empty for a text-only message", () => {
		const summary = mapMessage({ id: "1", body: { contentType: "text", content: "hi" } });
		assert.deepEqual(summary.imageUrls, []);
	});
});

describe("extensionForContentType", () => {
	test("maps the types a message actually carries", () => {
		assert.equal(extensionForContentType("image/png"), ".png");
		assert.equal(extensionForContentType("image/jpeg"), ".jpg");
		assert.equal(extensionForContentType("application/pdf"), ".pdf");
	});

	test("ignores a charset parameter and case", () => {
		assert.equal(extensionForContentType("IMAGE/PNG; charset=utf-8"), ".png");
	});

	test("returns an empty string rather than guessing", () => {
		assert.equal(extensionForContentType("application/octet-stream"), "");
		assert.equal(extensionForContentType(undefined), "");
	});
});

describe("safeFileName", () => {
	test("keeps a plain name unchanged", () => {
		assert.equal(safeFileName("quarterly-report.pdf", "fallback"), "quarterly-report.pdf");
	});

	test("cannot escape the directory it is given", () => {
		// The name comes from the sender, so this is the one that matters.
		assert.equal(safeFileName("../../etc/passwd", "fallback"), "passwd");
		assert.equal(safeFileName("..\\..\\windows\\system32\\cmd.exe", "fallback"), "cmd.exe");
	});

	test("replaces characters Windows refuses, and hides no file", () => {
		assert.equal(safeFileName('a:b|c?d*e"f<g>h.txt', "fallback"), "a_b_c_d_e_f_g_h.txt");
		assert.equal(safeFileName(".hidden", "fallback"), "hidden");
	});

	test("falls back rather than writing an unnamed file", () => {
		assert.equal(safeFileName("", "fallback"), "fallback");
		assert.equal(safeFileName(undefined, "fallback"), "fallback");
		assert.equal(safeFileName("   ", "fallback"), "fallback");
	});

	test("shortens from the front so the extension survives", () => {
		const long = `${"x".repeat(200)}.pdf`;
		const result = safeFileName(long, "fallback");
		assert.equal(result.length, 120);
		assert.ok(result.endsWith(".pdf"));
	});
});
