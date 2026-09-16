/**
 * Downloading what a message carries.
 *
 * A tool result is text, so bytes have to reach the disk before an agent can
 * look at them — that is the whole job here. Both sources are Graph: inline
 * images live in `hostedContents`, files are `attachments`, and a small
 * attachment arrives inline as base64 instead of a URL.
 */

import { mkdir, writeFile } from "node:fs/promises";
import { join } from "node:path";
import type { TeamsConnection } from "../config/index.ts";
import type { MessageAttachment, MessageSummary } from "../types.ts";
import { extensionForContentType, safeFileName } from "../utils/attachments.ts";
import { formatGraphError } from "../utils/errors.ts";
import { graphDownload, isGraphUrl, sharesContentPath } from "./client.ts";

/**
 * Refuse anything larger.
 *
 * A message can carry a two gigabyte video, and a tool call that quietly fills
 * the disk is worse than one that says no.
 */
export const MAX_FILE_BYTES = 25 * 1024 * 1024;

export interface SavedFile {
	path: string;
	name: string;
	bytes: number;
	kind: "image" | "attachment";
	contentType?: string;
}

export interface SkippedFile {
	name: string;
	reason: string;
	/** The URL to fetch by hand, when there is one. */
	url?: string;
}

export interface SaveResult {
	saved: SavedFile[];
	skipped: SkippedFile[];
}

function tooBig(bytes: number): string {
	const mb = Math.round((bytes / (1024 * 1024)) * 10) / 10;
	return `larger than ${Math.round(MAX_FILE_BYTES / (1024 * 1024))} MB (${mb} MB)`;
}

/** Append the extension the content type implies, when the name has none. */
function nameWithExtension(name: string | undefined, contentType: string | undefined, fallback: string): string {
	const safe = safeFileName(name, fallback);
	if (/\.[a-z0-9]{1,5}$/i.test(safe)) return safe;
	return safe + extensionForContentType(contentType);
}

/** Keep two attachments with the same name from overwriting each other. */
function uniqueName(used: Set<string>, name: string): string {
	if (!used.has(name)) {
		used.add(name);
		return name;
	}
	const dot = name.lastIndexOf(".");
	const stem = dot > 0 ? name.slice(0, dot) : name;
	const suffix = dot > 0 ? name.slice(dot) : "";
	for (let n = 2; ; n += 1) {
		const candidate = `${stem}-${n}${suffix}`;
		if (!used.has(candidate)) {
			used.add(candidate);
			return candidate;
		}
	}
}

/**
 * Save every image and attachment of one message into `dir`.
 *
 * Never throws for a single failed file: a message with three images and one
 * broken link should still leave the three on disk, with the failure reported
 * next to them.
 */
export async function saveMessageFiles(
	conn: TeamsConnection,
	message: MessageSummary,
	dir: string,
	options: { signal?: AbortSignal } = {},
): Promise<SaveResult> {
	const saved: SavedFile[] = [];
	const skipped: SkippedFile[] = [];
	const used = new Set<string>();

	await mkdir(dir, { recursive: true });

	const write = async (
		baseName: string,
		contentType: string | undefined,
		data: Buffer,
		kind: SavedFile["kind"],
	): Promise<void> => {
		const name = uniqueName(used, baseName);
		const path = join(dir, name);
		await writeFile(path, data);
		saved.push({ path, name, bytes: data.byteLength, kind, contentType });
	};

	// Inline images carry no name, so they are numbered the way they appear.
	let imageIndex = 0;
	for (const url of message.imageUrls) {
		imageIndex += 1;
		const label = `image ${imageIndex}`;

		// An inline image is only ours to fetch when Graph hosts it. The body is
		// written by whoever sent the message, so an <img> pointing anywhere else
		// is a request to hand out the access token, not a picture.
		if (!isGraphUrl(conn, url)) {
			skipped.push({ name: label, reason: foreignHost(url), url });
			continue;
		}

		try {
			const { data, contentType } = await graphDownload(conn, url, { signal: options.signal });
			if (data.byteLength > MAX_FILE_BYTES) {
				skipped.push({ name: label, reason: tooBig(data.byteLength), url });
				continue;
			}
			const extension = extensionForContentType(contentType) || ".bin";
			await write(`${message.id}-${imageIndex}${extension}`, contentType, data, "image");
		} catch (err) {
			skipped.push({ name: label, reason: formatGraphError(err), url });
		}
	}

	for (const attachment of message.attachments) {
		const label = attachment.name ?? "attachment";

		if (attachment.contentBytes) {
			// Small attachments arrive with the message itself.
			const data = Buffer.from(attachment.contentBytes, "base64");
			if (data.byteLength > MAX_FILE_BYTES) {
				skipped.push({ name: label, reason: tooBig(data.byteLength), url: attachment.contentUrl });
				continue;
			}
			await write(
				nameWithExtension(attachment.name, attachment.contentType, `attachment-${message.id}`),
				attachment.contentType,
				data,
				"attachment",
			);
			continue;
		}

		if (!attachment.contentUrl) {
			// A forwarded message is not a file being withheld: its images and files
			// are hosted content of the message that carries it, and they are fetched
			// above. Listing it as unsaved would put a line of noise under every
			// forwarded message.
			if (attachment.contentType === "forwardedMessageReference") continue;
			skipped.push({ name: label, reason: describeMissing(attachment) });
			continue;
		}

		// A file attachment usually points at the SharePoint site that stores it,
		// which is not an origin the token may go to. Graph fetches it for us when
		// the URL is encoded as a sharing token, and the user's own permissions
		// still decide what comes back.
		const source = resolveAttachmentSource(conn, attachment.contentUrl);
		if (!source) {
			skipped.push({ name: label, reason: foreignHost(attachment.contentUrl), url: attachment.contentUrl });
			continue;
		}

		try {
			const { data, contentType } = await graphDownload(conn, source, {
				signal: options.signal,
			});
			if (data.byteLength > MAX_FILE_BYTES) {
				skipped.push({ name: label, reason: tooBig(data.byteLength), url: attachment.contentUrl });
				continue;
			}
			await write(
				nameWithExtension(attachment.name, contentType ?? attachment.contentType, `attachment-${message.id}`),
				contentType ?? attachment.contentType,
				data,
				"attachment",
			);
		} catch (err) {
			skipped.push({ name: label, reason: formatGraphError(err), url: attachment.contentUrl });
		}
	}

	return { saved, skipped };
}

function describeMissing(attachment: MessageAttachment): string {
	const type = attachment.contentType ? ` (${attachment.contentType})` : "";
	return `no content URL and no inline bytes${type} — likely a card or a code snippet, not a file`;
}

/**
 * Where to fetch an attachment from, or `undefined` when it must not be fetched.
 *
 * Graph's own URLs are used as they are; an `http(s)` URL elsewhere is rewritten
 * to Graph's `/shares` addressing so the token never leaves the Graph origin.
 * Anything that is not http(s) — a `file:` path, a custom scheme — has no
 * business being fetched at all.
 */
function resolveAttachmentSource(conn: TeamsConnection, contentUrl: string): string | undefined {
	if (isGraphUrl(conn, contentUrl)) return contentUrl;
	if (!/^https?:\/\//i.test(contentUrl)) return undefined;
	return sharesContentPath(contentUrl);
}

/** Why a URL was left alone, phrased so the user knows it was a choice. */
function foreignHost(url: string): string {
	let host = url;
	try {
		host = new URL(url).host;
	} catch {
		/* keep the raw string */
	}
	return `hosted by ${host}, not by Microsoft Graph — not fetched, so your access token stays with Graph`;
}
