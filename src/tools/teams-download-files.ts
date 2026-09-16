/**
 * teams_download_files — save the files and images a message carries.
 *
 * The read tools flatten a message to text, which is right for prose and wrong
 * for anything with bytes in it. This is the counterpart: it fetches the images
 * embedded in messages and the files attached to them into a directory, and
 * returns the paths, so a reader can open them.
 */

import { isAbsolute, join, resolve, sep } from "node:path";
import { homedir } from "node:os";
import { Type } from "typebox";
import { getAgentDir } from "../config/index.ts";
import { saveMessageFiles, type SavedFile, type SkippedFile } from "../graph/media.ts";
import {
	getChannelMessage,
	getChatMessage,
	listChannelMessages,
	listChatMessages,
} from "../graph/messages.ts";
import type { MessageSummary } from "../types.ts";
import { safeFileName } from "../utils/attachments.ts";
import { formatDownloadResult } from "../utils/formatting.ts";
import { requireChannel, requireChat, TargetError } from "./resolve.ts";
import {
	AccountParam,
	TenantParam,
	connectionFor,
	run,
	textResult,
	type ToolContext,
	type ToolResult,
} from "./shared.ts";

export const teamsDownloadFilesTool = {
	name: "teams_download_files",
	description:
		"Download the images embedded in Teams messages and the files attached to them to a local " +
		"directory, and return the paths so they can be opened. Covers files attached to a message, not " +
		"a channel's SharePoint folder (see teams_list_files).",
	parameters: Type.Object({
		chat: Type.Optional(Type.String({ description: "Chat ID, topic, or a participant's name/e-mail" })),
		channel: Type.Optional(Type.String({ description: "Channel name or ID, or a 'Team/Channel' path" })),
		team: Type.Optional(Type.String({ description: "Team name or ID (omit when using a path)" })),
		messageId: Type.Optional(
			Type.String({ description: "Download only this message instead of scanning recent ones" }),
		),
		limit: Type.Optional(
			Type.Number({ description: "How many recent messages to scan (default 10, max 50)" }),
		),
		dir: Type.Optional(
			Type.String({
				description:
					"Directory to save into. Must be inside the working directory, the pi agent directory, or the " +
					"configured downloadDir. Defaults to the agent directory.",
			}),
		),
		account: AccountParam,
		tenant: TenantParam,
	}),
	promptSnippet: "Download files and images from Teams messages",

	async execute(
		_toolCallId: string,
		params: {
			chat?: string;
			channel?: string;
			team?: string;
			messageId?: string;
			limit?: number;
			dir?: string;
			account?: string;
			tenant?: string;
		},
		signal: AbortSignal | undefined,
		_onUpdate: undefined,
		ctx: ToolContext,
	): Promise<ToolResult> {
		return run(async () => {
			if (params.chat && params.channel) {
				throw new TargetError("Pass either 'chat' or 'channel', not both.");
			}
			if (!params.chat && !params.channel) {
				throw new TargetError("Pass a 'chat' or a 'channel' to download from.");
			}

			const conn = connectionFor(ctx, params.account, params.tenant);
			const scanLimit = Math.min(Math.max(params.limit ?? 10, 1), 50);

			let label: string;
			let slug: string;
			let messages: MessageSummary[];

			if (params.chat) {
				const chat = await requireChat(conn, params.chat, "read", signal);
				label = chat.label;
				slug = safeFileName(chat.label, chat.id);
				messages = params.messageId
					? await one(await getChatMessage(conn, chat.id, params.messageId, signal), params.messageId)
					: await listChatMessages(conn, chat.id, { max: scanLimit, signal });
			} else {
				const { team, channel } = await requireChannel(
					conn,
					params.team,
					params.channel as string,
					"read",
					signal,
				);
				label = `${team.displayName}/${channel.displayName}`;
				slug = safeFileName(`${team.displayName}-${channel.displayName}`, channel.id);
				messages = params.messageId
					? await one(await getChannelMessage(conn, channel, params.messageId, signal), params.messageId)
					: await listChannelMessages(conn, channel, { max: scanLimit, signal });
			}

			const dir = resolveDownloadDir(params.dir, join(getAgentDir(), "pi-teams-files", slug), conn.downloadDir);

			const saved: SavedFile[] = [];
			const skipped: SkippedFile[] = [];
			for (const message of messages) {
				const result = await saveMessageFiles(conn, message, dir, { signal });
				saved.push(...result.saved);
				skipped.push(...result.skipped);
			}

			const withFiles = messages.filter(
				(m) => m.imageUrls.length > 0 || m.attachments.length > 0,
			).length;

			return textResult(
				[
					formatDownloadResult(label, dir, saved, skipped),
					"",
					`Scanned ${messages.length} message${messages.length === 1 ? "" : "s"}` +
						`${withFiles === 0 ? ", none with files" : ""}.`,
				].join("\n"),
				{
					dir,
					saved: saved.map((file) => file.path),
					skipped: skipped.length,
					scanned: messages.length,
				},
			);
		});
	},
};

/** A single message, or a clear error — rather than silently downloading nothing. */
async function one(message: MessageSummary | undefined, messageId: string): Promise<MessageSummary[]> {
	if (!message) throw new TargetError(`No message ${messageId} in this conversation.`);
	return [message];
}

/**
 * Where the download may land.
 *
 * The directory is a tool parameter, so it is chosen by the model — and the
 * model is reading messages other people wrote. Confining it keeps a download
 * from landing in `~/.ssh` or a shell profile: the working directory and the
 * agent directory are always allowed, and a user who wants somewhere else says
 * so once in the configuration rather than per call.
 *
 * `~` is expanded first, because a path that silently means a literal `./~`
 * would be its own small surprise.
 */
export function resolveDownloadDir(
	requested: string | undefined,
	fallback: string,
	configured: string | undefined,
): string {
	if (!requested) return configured ? resolve(expandHome(configured)) : fallback;

	const target = resolve(expandHome(requested));
	const roots = [resolve(process.cwd()), resolve(getAgentDir())];
	if (configured) roots.push(resolve(expandHome(configured)));

	if (roots.some((root) => target === root || target.startsWith(root + sep))) return target;

	throw new TargetError(
		`Refusing to write downloads to "${target}". Allowed: ${roots.join(", ")}. ` +
			`Set "downloadDir" in pi-teams.json to allow another location.`,
	);
}

function expandHome(value: string): string {
	if (value === "~") return homedir();
	if (value.startsWith("~/") || value.startsWith("~\\")) return join(homedir(), value.slice(2));
	return isAbsolute(value) ? value : value;
}
