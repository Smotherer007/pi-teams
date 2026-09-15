/**
 * What a message carries besides text.
 *
 * Two different mechanisms hide behind "there was a picture in that message":
 * an inline image is a `hostedContent` referenced from the message HTML, and a
 * file is an `attachment` on the message. For a long time only the second one
 * reached the mapper, so a reader saw `[image: Bild]` and nothing else — the
 * URL was in the body all along, pointing straight at Graph.
 *
 * The helpers here are pure, because the shapes are the part that quietly goes
 * wrong and they can be tested without a Graph call.
 */

/**
 * Every string inside a value, however it is nested.
 *
 * Exists for the JSON that Graph hands back for a forwarded message: the
 * original message is in there as data of an undocumented depth, and the only
 * thing needed from it is whatever strings it contains.
 */
export function collectStrings(value: unknown, out: string[] = [], depth = 0): string[] {
	// A guard against a pathological document, not a real limit.
	if (depth > 8) return out;

	if (typeof value === "string") {
		out.push(value);
		// A forwarded message inside a forwarded message is JSON stored as a string
		// inside JSON, and its quotes stay escaped until it is parsed. Without this
		// the images two levels down are simply not there to be found.
		const trimmed = value.trim();
		if (trimmed.startsWith("{") || trimmed.startsWith("[")) {
			try {
				collectStrings(JSON.parse(trimmed), out, depth + 1);
			} catch {
				/* a string that merely starts with a brace */
			}
		}
		return out;
	}
	if (Array.isArray(value)) {
		for (const entry of value) collectStrings(entry, out, depth + 1);
		return out;
	}
	if (value && typeof value === "object") {
		for (const entry of Object.values(value as Record<string, unknown>)) {
			collectStrings(entry, out, depth + 1);
		}
	}
	return out;
}

/** Absolute URLs of the inline images in a message body, in the order shown. */
export function imageUrlsInHtml(html: string): string[] {
	const urls: string[] = [];
	for (const match of html.matchAll(/<img\b[^>]*\bsrc=["']([^"']+)["']/gi)) {
		const src = decodeEntities(match[1] ?? "");
		// An <img> can also point at an outside host. Only content that lives in
		// the message itself is ours to fetch.
		if (/\/hostedContents\//i.test(src) && !urls.includes(src)) urls.push(src);
	}
	return urls;
}

function decodeEntities(input: string): string {
	return input
		.replace(/&amp;/g, "&")
		.replace(/&lt;/g, "<")
		.replace(/&gt;/g, ">")
		.replace(/&quot;/g, '"')
		.replace(/&#(\d+);/g, (_match, code: string) => String.fromCharCode(Number(code)));
}

const CONTENT_TYPE_EXTENSIONS: Record<string, string> = {
	"image/png": ".png",
	"image/jpeg": ".jpg",
	"image/jpg": ".jpg",
	"image/gif": ".gif",
	"image/webp": ".webp",
	"image/bmp": ".bmp",
	"image/svg+xml": ".svg",
	"application/pdf": ".pdf",
	"text/plain": ".txt",
	"text/csv": ".csv",
	"application/json": ".json",
	"application/zip": ".zip",
	"application/vnd.openxmlformats-officedocument.wordprocessingml.document": ".docx",
	"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": ".xlsx",
	"application/vnd.openxmlformats-officedocument.presentationml.presentation": ".pptx",
};

/** Extension for a response content type, including the dot. Empty when unknown. */
export function extensionForContentType(contentType: string | undefined): string {
	const type = (contentType ?? "").split(";")[0]?.trim().toLowerCase() ?? "";
	return CONTENT_TYPE_EXTENSIONS[type] ?? "";
}

/**
 * A name that can be written on any platform, and is never empty.
 *
 * Keeps the extension when shortening: `slice` from the front would turn
 * `quarterly-report.pdf` into a name ending in `-report.pdf` at best and into
 * an unnamed blob at worst.
 */
export function safeFileName(name: string | undefined, fallback: string): string {
	const base = (name ?? "").split(/[\\/]/).pop() ?? "";
	const cleaned = base
		// Reserved on Windows, plus control characters.
		.replace(/[\u0000-\u001f<>:"|?*]/g, "_")
		// A leading dot would make the file invisible in most file managers.
		.replace(/^[.\s]+/, "")
		.trim();
	if (!cleaned) return fallback;
	return cleaned.length > 120 ? cleaned.slice(-120) : cleaned;
}
