/**
 * teams_read_chat — read the messages of one chat.
 */

import { Type } from "typebox";
import { listChatMessages } from "../graph/messages.ts";
import { formatChatDetail, formatMessageList } from "../utils/formatting.ts";
import { requireChat } from "./resolve.ts";
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

export const teamsReadChatTool = {
	name: "teams_read_chat",
	description:
		"Read messages from a Microsoft Teams chat. Accepts a chat ID, a group chat topic, or a participant's " +
		"name or e-mail — ambiguous matches are reported rather than guessed. Returns messages oldest first " +
		"with sender, timestamp, reactions and message IDs.",
	parameters: Type.Object({
		chat: Type.String({ description: "Chat ID, group chat topic, or a participant's name/e-mail" }),
		account: AccountParam,
		tenant: TenantParam,
		limit: LimitParam,
		includeDetails: Type.Optional(
			Type.Boolean({ description: "Also show chat metadata (members, type, link)" }),
		),
	}),
	promptSnippet: "Read messages from a Teams chat",
	promptGuidelines: [
		"Use teams_read_chat to catch up on a conversation before replying.",
		"Quote the messageId when the user wants to react to or delete a specific message.",
	],

	async execute(
		_toolCallId: string,
		params: {
			chat: string;
			account?: string;
			tenant?: string;
			limit?: number;
			includeDetails?: boolean;
		},
		signal: AbortSignal | undefined,
		_onUpdate: undefined,
		ctx: ToolContext,
	): Promise<ToolResult> {
		return run(async () => {
			const conn = connectionFor(ctx, params.account, params.tenant);
			const chat = await requireChat(conn, params.chat, "read", signal);

			const messages = await listChatMessages(conn, chat.id, {
				max: params.limit ?? conn.maxMessages,
				signal,
			});

			const body = formatMessageList(messages, chat.label);
			const detail = params.includeDetails ? `${formatChatDetail(chat)}\n\n` : "";

			return textResult(`${detail}${body}`, {
				chatId: chat.id,
				label: chat.label,
				count: messages.length,
			});
		});
	},
};
