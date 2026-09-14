/**
 * teams_inbox — what needs the user's attention right now.
 *
 * Teams has no single "unread" endpoint a delegated token can rely on, so this
 * composes the two signals that are reliable: recent chat activity (with the
 * last message and who sent it) and a search for mentions of the user. It
 * deliberately answers "what changed since you last looked", not "list
 * everything".
 */

import { Type } from "typebox";
import { listChats } from "../graph/chats.ts";
import { searchMessages } from "../graph/search.ts";
import { chatCandidates } from "../graph/mappers.ts";
import { hasAccess } from "../safety/index.ts";
import { formatRelative, truncate } from "../utils/formatting.ts";
import {
	AccountParam,
	LimitParam,
	TenantParam,
	connectionFor,
	currentUser,
	run,
	textResult,
	type ToolContext,
	type ToolResult,
} from "./shared.ts";

export const teamsInboxTool = {
	name: "teams_inbox",
	description:
		"Show what needs the user's attention in Microsoft Teams: chats with recent activity where the last " +
		"message came from someone else, plus recent @-mentions of the user across chats and channels. " +
		"Use this for 'what did I miss', 'anything new in Teams', or a morning catch-up.",
	parameters: Type.Object({
		account: AccountParam,
		tenant: TenantParam,
		limit: LimitParam,
		hours: Type.Optional(
			Type.Number({ description: "Only consider activity from the last N hours (default 24)" }),
		),
		includeMentions: Type.Optional(
			Type.Boolean({ description: "Also search for @-mentions of the user (default true)" }),
		),
	}),
	promptSnippet: "Show new Teams activity and mentions",
	promptGuidelines: [
		"Use teams_inbox for open-ended catch-up questions about Teams.",
		"Summarize what is waiting and who is waiting on it; do not reply to anything unless the user asks.",
		"After a catch-up, offer to clear the list with teams_mark_read — never mark a chat read unasked.",
	],

	async execute(
		_toolCallId: string,
		params: {
			account?: string;
			tenant?: string;
			limit?: number;
			hours?: number;
			includeMentions?: boolean;
		},
		signal: AbortSignal | undefined,
		_onUpdate: undefined,
		ctx: ToolContext,
	): Promise<ToolResult> {
		return run(async () => {
			const conn = connectionFor(ctx, params.account, params.tenant);
			const me = await currentUser(conn, signal).catch(() => undefined);
			const cutoff = Date.now() - (params.hours ?? 24) * 3600 * 1000;
			const limit = params.limit ?? 15;

			const chats = await listChats(conn, { max: 50, meId: me?.id, signal });

			const active = chats
				.filter((chat) => hasAccess(conn, "read", "chats", chatCandidates(chat)))
				.filter((chat) => {
					if (!chat.lastUpdated) return false;
					if (new Date(chat.lastUpdated).getTime() < cutoff) return false;
					// Something the user themselves said last is not waiting on them.
					if (me && chat.lastMessageFrom && chat.lastMessageFrom === me.displayName) return false;
					return true;
				})
				.slice(0, limit);

			const lines = [`## Teams inbox — last ${params.hours ?? 24} h`, ""];

			if (active.length === 0) {
				lines.push("No chat activity waiting for you.", "");
			} else {
				lines.push(`### ${active.length} chat(s) with new messages`, "");
				for (const chat of active) {
					lines.push(`- **${chat.label}** · ${formatRelative(chat.lastUpdated)}`);
					if (chat.lastMessagePreview) {
						const who = chat.lastMessageFrom ? `${chat.lastMessageFrom}: ` : "";
						lines.push(`  ${who}${truncate(chat.lastMessagePreview, 140)}`);
					}
					lines.push(`  chatId: ${chat.id}`);
				}
				lines.push("");
			}

			if (params.includeMentions !== false && me) {
				try {
					const hits = await searchMessages(conn, `"${me.displayName}"`, { max: 15, signal });
					const recent = hits.filter((hit) => {
						const created = hit.message.createdDateTime;
						if (!created || new Date(created).getTime() < cutoff) return false;
						// Only real mentions, not every message containing the name.
						return hit.message.mentions.some(
							(mention) => mention.id === me.id || mention.displayName === me.displayName,
						);
					});

					if (recent.length > 0) {
						lines.push(`### ${recent.length} mention(s) of you`, "");
						for (const hit of recent) {
							const where = hit.channelIdentity?.channelId
								? `channel ${hit.channelIdentity.channelId}`
								: hit.chatId
									? `chat ${hit.chatId}`
									: "unknown";
							lines.push(
								`- **${hit.message.from?.displayName ?? "(unknown)"}** · ${formatRelative(hit.message.createdDateTime)} · ${where}`,
							);
							lines.push(`  ${truncate(hit.message.text, 160)}`);
							lines.push(`  messageId: ${hit.message.id}`);
						}
						lines.push("");
					}
				} catch {
					// Search needs its own consent and is not granted everywhere; the
					// chat half of the answer is still worth returning.
					lines.push("_Mention search unavailable (Microsoft Search may not be consented)._", "");
				}
			}

			return textResult(lines.join("\n").trimEnd(), {
				chats: active.length,
				account: conn.account,
			});
		});
	},
};
