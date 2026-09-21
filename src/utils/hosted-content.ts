/**
 * Attaching a local image to a message.
 *
 * Teams keeps an inline image inside the message itself as a `hostedContent`:
 * the body points at `../hostedContents/{id}/$value` and the bytes travel in the
 * same POST. That is what makes this cheap — nothing has to be uploaded to
 * SharePoint first, which is what a real file attachment needs, and which would
 * mean asking for a write scope on the user's drive.
 *
 * Graph caps a hosted content at 4 MB, and an image of a type Teams does not
 * render would reach the reader as a broken picture, so both are checked here,
 * before the message that carries it is built.
 */

import { readFile, stat } from "node:fs/promises";
import { basename, extname } from "node:path";

/** Graph rejects a hosted content larger than this. */
export const MAX_HOSTED_CONTENT_BYTES = 4 * 1024 * 1024;

/**
 * The image types Teams shows inline, by file extension.
 *
 * A list rather than "anything starting with image/": a type Teams does not
 * render arrives as a broken image, and refusing it up front is friendlier than
 * sending it and having the reader find out.
 */
export const HOSTED_CONTENT_TYPES: Record<string, string> = {
	".png": "image/png",
	".jpg": "image/jpeg",
	".jpeg": "image/jpeg",
	".gif": "image/gif",
	".webp": "image/webp",
	".bmp": "image/bmp",
};

/** The content type for a path, or undefined when Teams does not render it inline. */
export function hostedContentType(path: string): string | undefined {
	return HOSTED_CONTENT_TYPES[extname(path).toLowerCase()];
}

/**
 * One image, ready to send.
 *
 * `contentBytes` is base64 because that is what Graph asks for: the bytes are
 * JSON string data, not a multipart upload.
 */
export interface HostedImage {
	/** The file name, as the attachment is described in audit entries */
	name: string;
	contentType: string;
	contentBytes: string;
}

/** A byte count as a human reads it — the limits here are in megabytes. */
export function formatBytes(bytes: number): string {
	return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

/**
 * Read one local image for sending.
 *
 * Throws with the reason instead of sending something partial: a message that
 * quietly goes out without the picture the user asked for is worse than a tool
 * call that says why it did not.
 */
export async function readHostedImage(path: string): Promise<HostedImage> {
	const name = basename(path);

	const contentType = hostedContentType(path);
	if (!contentType) {
		const supported = [...new Set(Object.values(HOSTED_CONTENT_TYPES))].join(", ");
		throw new Error(`"${name}" is not an image Teams shows inline. Supported types: ${supported}.`);
	}

	let info;
	try {
		info = await stat(path);
	} catch {
		throw new Error(`"${path}" could not be read. Give a path to a file that exists locally.`);
	}
	if (info.isDirectory()) throw new Error(`"${path}" is a directory, not an image.`);
	if (info.size === 0) throw new Error(`"${name}" is empty.`);
	if (info.size > MAX_HOSTED_CONTENT_BYTES) {
		throw new Error(
			`"${name}" is ${formatBytes(info.size)}; Teams accepts at most ` +
				`${formatBytes(MAX_HOSTED_CONTENT_BYTES)} per image.`,
		);
	}

	const bytes = await readFile(path);
	return { name, contentType, contentBytes: bytes.toString("base64") };
}

/**
 * The images named in one phrase, for the audit entry.
 *
 * What went out is the thing an audit entry has to answer, and a message with a
 * picture in it is not described by its text alone.
 */
export function imageNote(images: readonly HostedImage[]): string {
	if (images.length === 0) return "";
	const label = images.length === 1 ? "image" : "images";
	return ` [${label}: ${images.map((image) => image.name).join(", ")}]`;
}
