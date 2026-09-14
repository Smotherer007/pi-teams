/**
 * teams_search_messages — full-text search across the user's Teams history.
 */

import { Type } from "typebox";
import { formatSearchHits, searchMessages } from "../graph/search.ts";
import { hasAccess } from "../safety/index.ts";
import {
	AccountParam,
	LimitParam,
	TenantParam,
	connectionFor,
	run,
	textResult,
	type ToolContext,
	type ToolResult,
} from "./shared.ts";

export const teamsSearchMessagesTool = {
	name: "teams_search_messages",
	description:
		"Search the signed-in user's Microsoft Teams messages — chats and channels — with the Microsoft Search " +
		"API. Plain keywords work; so do qualifiers such as from:anna@contoso.com. " +
		"Use this to find a conversation before reading or replying to it.",
	parameters: Type.Object({
		query: Type.String({ description: "Search terms, e.g. 'deployment friday' or 'from:anna@contoso.com budget'" }),
		account: AccountParam,
		tenant: TenantParam,
		limit: LimitParam,
	}),
	promptSnippet: "Search the user's Teams messages",
	promptGuidelines: [
		"Use teams_search_messages when the user refers to something said in Teams without naming the chat or channel.",
		"Search hits carry messageId plus the chat or channel ID — pass those to teams_read_thread or teams_read_chat for full context.",
	],

	async execute(
		_toolCallId: string,
		params: { query: string; account?: string; tenant?: string; limit?: number },
		signal: AbortSignal | undefined,
		_onUpdate: undefined,
		ctx: ToolContext,
	): Promise<ToolResult> {
		return run(async () => {
			const conn = connectionFor(ctx, params.account, params.tenant);
			const hits = await searchMessages(conn, params.query, { max: params.limit ?? 25, signal });

			// Search bypasses the listing endpoints, so the rules are applied to the
			// results: a denied channel must not leak through a search box.
			const visible = hits.filter((hit) => {
				if (hit.channelIdentity?.channelId) {
					return hasAccess(conn, "read", "channels", [
						hit.channelIdentity.channelId,
						hit.channelIdentity.teamId,
					]);
				}
				if (hit.chatId) return hasAccess(conn, "read", "chats", [hit.chatId]);
				return true;
			});

			const hidden = hits.length - visible.length;
			const note = hidden > 0 ? `\n\n_${hidden} hit(s) hidden by your scope rules._` : "";

			return textResult(`${formatSearchHits(visible, params.query)}${note}`, {
				query: params.query,
				count: visible.length,
				hidden,
			});
		});
	},
};
