/**
 * Message search across chats and channels.
 *
 * Uses the Microsoft Search API (`POST /search/query`, entityType
 * `chatMessage`), which is the only endpoint that searches a user's whole
 * Teams history in one call. Results come back as search hits with a summary
 * and a partial resource, so they are normalized here into MessageSummary.
 */

import type { TeamsConnection } from "../config/index.ts";
import type { MessageLocation, MessageSummary } from "../types.ts";
import { graphPost } from "./client.ts";
import { mapMessage } from "./mappers.ts";

/* eslint-disable @typescript-eslint/no-explicit-any */
type Raw = Record<string, any>;

export interface SearchHit {
	message: MessageSummary;
	/** Search summary with <c0> highlight markers stripped */
	summary?: string;
	/** Where the hit lives, when the resource carried enough information */
	channelIdentity?: { teamId?: string; channelId?: string };
	chatId?: string;
}

/**
 * Search Teams messages.
 *
 * @param query - KQL-ish query string; plain words work, as do `from:` filters
 * @param max - hits to return (Graph pages in blocks of 25)
 */
export async function searchMessages(
	conn: TeamsConnection,
	query: string,
	options: { max?: number; from?: number; signal?: AbortSignal } = {},
): Promise<SearchHit[]> {
	const size = Math.min(options.max ?? 25, 25);

	const response = await graphPost<Raw>(
		conn,
		"/search/query",
		{
			requests: [
				{
					entityTypes: ["chatMessage"],
					query: { queryString: query },
					from: options.from ?? 0,
					size,
				},
			],
		},
		{ signal: options.signal },
	);

	const containers = response?.value?.[0]?.hitsContainers ?? [];
	const hits: SearchHit[] = [];

	for (const container of containers) {
		for (const hit of container.hits ?? []) {
			const resource = hit.resource ?? {};
			const channelIdentity = resource.channelIdentity as
				| { teamId?: string; channelId?: string }
				| undefined;

			const location: MessageLocation | undefined = channelIdentity?.channelId
				? {
						kind: "channel",
						teamId: channelIdentity.teamId,
						channelId: channelIdentity.channelId,
					}
				: resource.chatId
					? { kind: "chat", chatId: resource.chatId }
					: undefined;

			hits.push({
				message: mapMessage(resource, location),
				summary: typeof hit.summary === "string"
					? hit.summary.replace(/<\/?c\d+>/g, "").trim()
					: undefined,
				channelIdentity,
				chatId: resource.chatId,
			});
		}
	}

	return hits.slice(0, options.max ?? 25);
}

/** Render search hits for the agent, keeping the IDs needed for follow-ups. */
export function formatSearchHits(hits: readonly SearchHit[], query: string): string {
	if (hits.length === 0) return `No Teams messages found for "${query}".`;

	const lines = [`## Search: "${query}"`, "", `${hits.length} hit(s):`, ""];
	for (const hit of hits) {
		const who = hit.message.from?.displayName ?? "(unknown)";
		const when = hit.message.createdDateTime
			? new Date(hit.message.createdDateTime).toISOString().slice(0, 16).replace("T", " ")
			: "";
		const where = hit.channelIdentity?.channelId
			? `channel ${hit.channelIdentity.channelId}`
			: hit.chatId
				? `chat ${hit.chatId}`
				: "unknown location";

		lines.push(`- **${who}**${when ? ` · ${when}` : ""} · ${where}`);
		lines.push(`  ${hit.summary || hit.message.text.slice(0, 200)}`);
		lines.push(`  messageId: ${hit.message.id}`);
	}
	return lines.join("\n");
}
