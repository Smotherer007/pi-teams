/**
 * Reading a local image for an inline send.
 *
 * The limits are the point: Graph caps hosted content at 4 MB, and a file that
 * is not an image Teams renders would reach the reader as a broken picture. Both
 * are refused here, before a message is built around them.
 */

import { test, describe, after } from "node:test";
import assert from "node:assert/strict";
import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import {
	MAX_HOSTED_CONTENT_BYTES,
	hostedContentType,
	imageNote,
	readHostedImage,
} from "../src/utils/hosted-content.ts";

const dir = mkdtempSync(join(tmpdir(), "pi-teams-hosting-"));

after(() => rmSync(dir, { recursive: true, force: true }));

/** A real 1×1 PNG, so what round-trips through base64 is image data. */
const PNG = Buffer.from(
	"iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8DwHwAFAAH/q842iQAAAABJRU5ErkJggg==",
	"base64",
);

function file(name: string, bytes: Buffer = PNG): string {
	const path = join(dir, name);
	writeFileSync(path, bytes);
	return path;
}

describe("hostedContentType", () => {
	test("maps the image types Teams renders, ignoring case", () => {
		assert.equal(hostedContentType("Bild.PNG"), "image/png");
		assert.equal(hostedContentType("foto.jpeg"), "image/jpeg");
		assert.equal(hostedContentType("anim.gif"), "image/gif");
	});

	test("is undefined for everything else", () => {
		assert.equal(hostedContentType("bericht.pdf"), undefined);
		assert.equal(hostedContentType("logo.svg"), undefined);
		assert.equal(hostedContentType("noextension"), undefined);
	});
});

describe("readHostedImage", () => {
	test("reads a file into base64 with its content type", async () => {
		const path = file("diagramm.png");
		const image = await readHostedImage(path);

		assert.equal(image.name, "diagramm.png");
		assert.equal(image.contentType, "image/png");
		assert.deepEqual(Buffer.from(image.contentBytes, "base64"), PNG);
	});

	test("refuses a file Teams would not show inline", async () => {
		const path = file("bericht.pdf");
		await assert.rejects(() => readHostedImage(path), /not an image Teams shows inline/);
	});

	test("refuses a file over the Graph limit", async () => {
		const path = file("riesig.png", Buffer.alloc(MAX_HOSTED_CONTENT_BYTES + 1));
		await assert.rejects(() => readHostedImage(path), /at most 4\.0 MB per image/);
	});

	test("refuses an empty file", async () => {
		const path = file("leer.png", Buffer.alloc(0));
		await assert.rejects(() => readHostedImage(path), /is empty/);
	});

	test("refuses a directory and a path that does not exist", async () => {
		const path = join(dir, "ordner.png");
		mkdirSync(path, { recursive: true });
		await assert.rejects(() => readHostedImage(path), /is a directory/);
		await assert.rejects(() => readHostedImage(join(dir, "fehlt.png")), /could not be read/);
	});
});

describe("imageNote", () => {
	test("says nothing without images", () => {
		assert.equal(imageNote([]), "");
	});

	test("names one and several images", () => {
		const one = [{ name: "a.png", contentType: "image/png", contentBytes: "AA==" }];
		assert.equal(imageNote(one), " [image: a.png]");

		const two = [
			...one,
			{ name: "b.jpg", contentType: "image/jpeg", contentBytes: "BB==" },
		];
		assert.equal(imageNote(two), " [images: a.png, b.jpg]");
	});
});
