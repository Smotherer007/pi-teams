/**
 * teams_mark_read — move the read cursor of a chat.
 *
 * "Mark as read" is the one piece of chat state a delegated token may change
 * on the user's behalf, and it is the natural end of a catch-up: pi summarizes
 * the inbox, then clears it. It is a mutation all the same — the chat stops
 * being bold, and where the tenant shows read receipts, the sender sees that
 * it was opened.
 */

import { Type } from "typebox";
import { setChatReadState } from "../graph/chats.ts";
import { auditWrite } from "../safety/audit.ts";
import { requireChat } from "./resolve.ts";
import {
	AccountParam,
	TenantParam,
	connectionFor,
	currentUser,
	run,
	textResult,
	type ToolContext,
	type ToolResult,
} from "./shared.ts";

export const teamsMarkReadTool = {
	name: "teams_mark_read",
	description:
		"Mark a Microsoft Teams chat as read — or unread again — as the signed-in user, exactly like opening " +
		"or re-marking it in the app. Accepts a chat ID, a group chat topic, or a participant's name/e-mail. " +
		"Set read: false to make the chat unread again. Channels have no equivalent: channel posts cannot be " +
		"marked read or unread.",
	parameters: Type.Object({
		chat: Type.String({
			description: "Chat ID, group chat topic, or a participant's name/e-mail",
		}),
		read: Type.Optional(
			Type.Boolean({
				description: "true (default) marks the chat read, false marks it unread again",
			}),
		),
		account: AccountParam,
		tenant: TenantParam,
	}),
	promptSnippet: "Mark a Teams chat read or unread",
	promptGuidelines: [
		"Use teams_mark_read for 'mark that as read', 'clear my inbox', 'mark it unread again'.",
		"Marking a chat read is visible to the sender where the tenant shows read receipts, so do it when asked, not on your own initiative.",
	],

	async execute(
		_toolCallId: string,
		params: { chat: string; read?: boolean; account?: string; tenant?: string },
		signal: AbortSignal | undefined,
		_onUpdate: undefined,
		ctx: ToolContext,
	): Promise<ToolResult> {
		return run(async () => {
			const conn = connectionFor(ctx, params.account, params.tenant);
			const me = await currentUser(conn, signal);

			// "write": this changes state in Teams, so the write rules apply —
			// a chat pi may read but not write is one whose unread marker stays.
			const chat = await requireChat(conn, params.chat, "write", signal);
			const read = params.read !== false;

			await setChatReadState(conn, chat.id, me, read, { signal });

			auditWrite(conn.audit, {
				tool: "teams_mark_read",
				account: conn.account,
				tenant: conn.tenant,
				actor: me.upn,
				target: `chat:${chat.label}`,
				summary: `marked chat as ${read ? "read" : "unread"}`,
			});

			return textResult(
				`✅ "${chat.label}" marked as ${read ? "read" : "unread"}.`,
				{ chatId: chat.id, label: chat.label, read },
			);
		});
	},
};
