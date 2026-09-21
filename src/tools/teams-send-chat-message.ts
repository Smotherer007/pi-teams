/**
 * teams_send_chat_message — send a message in a chat, as the user.
 *
 * This is the sharpest tool in the box: what it writes is indistinguishable
 * from the user typing it. Hence the order of operations — resolve the chat,
 * check the chat rules, check every participant against the people rules,
 * resolve the mentions, send, then record the audit entry.
 *
 * The last step is the read cursor. Sending from Graph does not move it — only
 * the user opening the chat, or teams_mark_read, does — so without this a chat
 * pi has just answered keeps its unread marker in Teams, which is the one thing
 * that looks like nobody dealt with it. Replying is reading, so the message
 * would arrive with the chat still bold; the mark is best effort and comes last,
 * because a send that went out must not be reported as failed.
 */

import { Type } from "typebox";
import { setChatReadState } from "../graph/chats.ts";
import { sendChatMessage } from "../graph/messages.ts";
import { formatGraphError } from "../utils/errors.ts";
import { resolveUserId } from "../graph/me.ts";
import { auditWrite } from "../safety/audit.ts";
import { assertAccess } from "../safety/index.ts";
import { CHAT_MESSAGE_BODY_DESCRIPTION, CHAT_SHAPE_GUIDELINE } from "../utils/chat-style.ts";
import { applyAiFooter } from "../utils/disclosure.ts";
import { imageNote, readHostedImage } from "../utils/hosted-content.ts";
import { truncate } from "../utils/formatting.ts";
import { requireChat } from "./resolve.ts";
import {
	AccountParam,
	TenantParam,
	connectionFor,
	connectionLabel,
	currentUser,
	errorResult,
	run,
	textResult,
	type ToolContext,
	type ToolResult,
} from "./shared.ts";

interface SendChatParams {
	chat: string;
	body: string;
	account?: string;
	tenant?: string;
	mentions?: string[];
	importance?: string;
	html?: boolean;
	images?: string[];
}

export const teamsSendChatMessageTool = {
	name: "teams_send_chat_message",
	description:
		"Send a message in a Microsoft Teams chat as the signed-in user. The message appears exactly as if the " +
		"user typed it. Accepts a chat ID, group chat topic, or participant name/e-mail. " +
		"Use 'mentions' with names or e-mail addresses to @-mention people — write '@Display Name' in the body " +
		"where the mention should appear. " +
		"'images' attaches local image files to the message itself; pi cannot attach other kinds of files.",
	parameters: Type.Object({
		chat: Type.String({ description: "Chat ID, group chat topic, or a participant's name/e-mail" }),
		body: Type.String({ description: CHAT_MESSAGE_BODY_DESCRIPTION }),
		account: AccountParam,
		tenant: TenantParam,
		mentions: Type.Optional(
			Type.Array(Type.String(), {
				description: "People to @-mention, by name, UPN or e-mail",
			}),
		),
		importance: Type.Optional(
			Type.String({ description: "'normal' (default), 'high' or 'urgent'" }),
		),
		html: Type.Optional(
			Type.Boolean({ description: "Send 'body' as raw HTML instead of converting markdown" }),
		),
		images: Type.Optional(
			Type.Array(Type.String(), {
				description:
					"Local paths of images to attach to the message (png, jpg, gif, webp, bmp; 4 MB each at most). " +
					"They are shown inline, under the text.",
			}),
		),
	}),
	promptSnippet: "Send a message in a Teams chat as the user",
	promptGuidelines: [
		"Write the message in the user's voice — it is sent from their account, not from an assistant.",
		"Show the user the exact text before sending when the intent is even slightly ambiguous.",
		"Do not add signatures, or a note that the message was written by an AI: the sending tool appends the configured AI disclosure itself (on by default), so writing one into the body duplicates it.",
		CHAT_SHAPE_GUIDELINE,
		"Answer in the language of the conversation you are writing into, and match its register.",
	],

	async execute(
		_toolCallId: string,
		params: SendChatParams,
		signal: AbortSignal | undefined,
		_onUpdate: undefined,
		ctx: ToolContext,
	): Promise<ToolResult> {
		return run(async () => {
			const conn = connectionFor(ctx, params.account, params.tenant);

			// Read before anything is sent: a picture that cannot be read has to stop
			// the message, not arrive as text promising one.
			const images = [];
			for (const path of params.images ?? []) images.push(await readHostedImage(path));

			if (!params.body.trim() && images.length === 0) {
				return errorResult("Refusing to send an empty message.");
			}

			const chat = await requireChat(conn, params.chat, "write", signal);

			// Mentions address people directly, so they are checked individually
			// even when the chat itself is allowed.
			const mentions = [];
			for (const reference of params.mentions ?? []) {
				const person = await resolveUserId(conn, reference, signal);
				if (!person) {
					return errorResult(
						`Could not resolve "${reference}" to a person. Use teams_find_user to look them up.`,
					);
				}
				assertAccess(conn, "write", "people", person.displayName, [
					person.displayName,
					person.upn,
					person.mail,
					person.id,
				]);
				mentions.push(person);
			}

			const me = await currentUser(conn, signal).catch(() => undefined);

			// The disclosure is added here, not asked for: what leaves the account
			// has to carry it whether or not the model remembered.
			const body = applyAiFooter(params.body, conn.aiFooter, { html: params.html });

			try {
				const sent = await sendChatMessage(
					conn,
					chat.id,
					{
						body,
						html: params.html,
						importance: params.importance,
						mentions,
						images,
					},
					signal,
				);

				auditWrite(conn.audit, {
					tool: "teams_send_chat_message",
					account: conn.account,
					tenant: conn.tenant,
					actor: me?.upn,
					target: `chat:${chat.label}`,
					summary: truncate(body, 200) + imageNote(images),
				});

				// Replying is reading: clear the unread marker the answer was written for.
				// Deliberately after the audit entry, so the message that went out is
				// recorded even when the read-mark cannot be moved.
				let readNote = "";
				if (me) {
					try {
						await setChatReadState(conn, chat.id, me, true, { signal });
					} catch (err) {
						readNote = `\n\n⚠️ The chat could not be marked as read: ${formatGraphError(err)}`;
					}
				}

				return textResult(
					[
						`✅ Message sent to **${chat.label}** as ${me?.displayName ?? "you"} (${connectionLabel(conn)}).`,
						"",
						`> ${truncate(body, 300)}`,
						"",
						`messageId: ${sent.id}`,
					].join("\n") + readNote,
					{ chatId: chat.id, messageId: sent.id },
				);
			} catch (err) {
				auditWrite(conn.audit, {
					tool: "teams_send_chat_message",
					account: conn.account,
					tenant: conn.tenant,
					actor: me?.upn,
					target: `chat:${chat.label}`,
					summary: truncate(body, 200),
					error: err instanceof Error ? err.message : String(err),
				});
				throw err;
			}
		});
	},
};
