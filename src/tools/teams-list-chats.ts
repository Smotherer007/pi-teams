/**
 * teams_list_chats — recent chats with a preview of the last message.
 */

import { Type } from "typebox";
import { listChats } from "../graph/chats.ts";
import { chatCandidates } from "../graph/mappers.ts";
import { hasAccess } from "../safety/index.ts";
import { formatChatList } from "../utils/formatting.ts";
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

export const teamsListChatsTool = {
	name: "teams_list_chats",
	description:
		"List the signed-in user's Microsoft Teams chats, most recent activity first, each with a preview of " +
		"the last message. Chats excluded by the configured read rules are not shown. " +
		"Use this to find a chat ID before reading or replying.",
	parameters: Type.Object({
		account: AccountParam,
		tenant: TenantParam,
		limit: LimitParam,
		filter: Type.Optional(
			Type.String({ description: "Only show chats whose label or participants contain this text" }),
		),
	}),
	promptSnippet: "List the user's recent Teams chats",

	async execute(
		_toolCallId: string,
		params: { account?: string; tenant?: string; limit?: number; filter?: string },
		signal: AbortSignal | undefined,
		_onUpdate: undefined,
		ctx: ToolContext,
	): Promise<ToolResult> {
		return run(async () => {
			const conn = connectionFor(ctx, params.account, params.tenant);
			const me = await currentUser(conn, signal).catch(() => undefined);

			const all = await listChats(conn, {
				max: params.limit ?? conn.maxMessages,
				meId: me?.id,
				signal,
			});

			const visible = all.filter((chat) => hasAccess(conn, "read", "chats", chatCandidates(chat)));
			const filtered = params.filter
				? visible.filter((chat) =>
						`${chat.label} ${chat.members.map((m) => m.displayName).join(" ")}`
							.toLowerCase()
							.includes(params.filter!.toLowerCase()),
					)
				: visible;

			const hidden = all.length - visible.length;
			const note = hidden > 0 ? `\n\n_${hidden} chat(s) hidden by your scope rules._` : "";

			return textResult(`${formatChatList(filtered)}${note}`, {
				count: filtered.length,
				hidden,
				account: conn.account,
			});
		});
	},
};
