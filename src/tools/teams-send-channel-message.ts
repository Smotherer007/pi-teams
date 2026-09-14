/**
 * teams_send_channel_message — start a new post in a channel, as the user.
 * teams_reply_channel_message — reply inside an existing thread.
 *
 * Both live here because they share the mention resolution and the audit
 * shape; splitting them would mean two copies of the part that matters.
 */

import { Type } from "typebox";
import { replyToChannelMessage, sendChannelMessage } from "../graph/messages.ts";
import { resolveUserId } from "../graph/me.ts";
import { auditWrite } from "../safety/audit.ts";
import { applyAiFooter } from "../utils/disclosure.ts";
import { assertAccess } from "../safety/index.ts";
import { truncate } from "../utils/formatting.ts";
import { requireChannel } from "./resolve.ts";
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
import type { PersonRef } from "../types.ts";
import type { TeamsConnection } from "../config/index.ts";

/** Resolve mention references and check each against the write rules. */
async function resolveMentions(
	conn: TeamsConnection,
	references: string[] | undefined,
	signal?: AbortSignal,
): Promise<{ mentions: PersonRef[]; error?: string }> {
	const mentions: PersonRef[] = [];
	for (const reference of references ?? []) {
		const person = await resolveUserId(conn, reference, signal);
		if (!person) {
			return {
				mentions,
				error: `Could not resolve "${reference}" to a person. Use teams_find_user to look them up.`,
			};
		}
		assertAccess(conn, "write", "people", person.displayName, [
			person.displayName,
			person.upn,
			person.mail,
			person.id,
		]);
		mentions.push(person);
	}
	return { mentions };
}

const sharedParams = {
	channel: Type.String({ description: "Channel name or ID, or a 'Team/Channel' path" }),
	team: Type.Optional(Type.String({ description: "Team name or ID (omit when using a path)" })),
	body: Type.String({
			description:
				"Message text as lightweight markdown (bold, italics, code, bullets, numbered lists, links)",
		}),
	account: AccountParam,
	tenant: TenantParam,
	mentions: Type.Optional(
		Type.Array(Type.String(), { description: "People to @-mention, by name, UPN or e-mail" }),
	),
	importance: Type.Optional(Type.String({ description: "'normal' (default), 'high' or 'urgent'" })),
	html: Type.Optional(Type.Boolean({ description: "Send 'body' as raw HTML instead of converting markdown" })),
};

export const teamsSendChannelMessageTool = {
	name: "teams_send_channel_message",
	description:
		"Post a new message in a Microsoft Teams channel as the signed-in user. It appears exactly as if the " +
		"user posted it. Give the channel as a 'Team/Channel' path or pass 'team' separately. " +
		"'subject' starts a titled post; use teams_reply_channel_message to answer inside an existing thread.",
	parameters: Type.Object({
		...sharedParams,
		subject: Type.Optional(Type.String({ description: "Optional post title" })),
	}),
	promptSnippet: "Post a message in a Teams channel as the user",
	promptGuidelines: [
		"Write in the user's voice — the post is attributed to them, not to an assistant.",
		"Show the exact text and the target channel before posting when there is any ambiguity.",
		"Format for a chat, not for a document: lead with the answer, three short paragraphs at most, bullets for lists. Bold, italics, code, links and bullets are rendered; headings become bold lines and tables are not supported — keep those out.",
		"Answer in the language of the conversation you are writing into, and match its register.",
	],

	async execute(
		_toolCallId: string,
		params: {
			channel: string;
			team?: string;
			body: string;
			subject?: string;
			account?: string;
			tenant?: string;
			mentions?: string[];
			importance?: string;
			html?: boolean;
		},
		signal: AbortSignal | undefined,
		_onUpdate: undefined,
		ctx: ToolContext,
	): Promise<ToolResult> {
		return run(async () => {
			const conn = connectionFor(ctx, params.account, params.tenant);
			if (!params.body.trim()) return errorResult("Refusing to post an empty message.");

			const { team, channel } = await requireChannel(conn, params.team, params.channel, "write", signal);
			const { mentions, error } = await resolveMentions(conn, params.mentions, signal);
			if (error) return errorResult(error);

			const me = await currentUser(conn, signal).catch(() => undefined);
			const target = `${team.displayName}/${channel.displayName}`;
			// The disclosure is added here, not asked for: what leaves the account
			// has to carry it whether or not the model remembered.
			const body = applyAiFooter(params.body, conn.aiFooter, { html: params.html });

			try {
				const sent = await sendChannelMessage(
					conn,
					channel,
					{
						body,
						html: params.html,
						subject: params.subject,
						importance: params.importance,
						mentions,
					},
					signal,
				);

				auditWrite(conn.audit, {
					tool: "teams_send_channel_message",
					account: conn.account,
					tenant: conn.tenant,
					actor: me?.upn,
					target: `channel:${target}`,
					summary: truncate(body, 200),
				});

				return textResult(
					[
						`✅ Posted in **${target}** as ${me?.displayName ?? "you"} (${connectionLabel(conn)}).`,
						"",
						`> ${truncate(body, 300)}`,
						"",
						`messageId: ${sent.id}`,
					].join("\n"),
					{ teamId: team.id, channelId: channel.id, messageId: sent.id },
				);
			} catch (err) {
				auditWrite(conn.audit, {
					tool: "teams_send_channel_message",
					account: conn.account,
					tenant: conn.tenant,
					actor: me?.upn,
					target: `channel:${target}`,
					summary: truncate(body, 200),
					error: err instanceof Error ? err.message : String(err),
				});
				throw err;
			}
		});
	},
};

export const teamsReplyChannelMessageTool = {
	name: "teams_reply_channel_message",
	description:
		"Reply inside an existing Microsoft Teams channel thread as the signed-in user. " +
		"Get the messageId from teams_read_channel, teams_read_thread or teams_search_messages. " +
		"Replying keeps the conversation in one thread instead of starting a new post.",
	parameters: Type.Object({
		...sharedParams,
		messageId: Type.String({ description: "ID of the thread's opening message" }),
	}),
	promptSnippet: "Reply in a Teams channel thread as the user",

	async execute(
		_toolCallId: string,
		params: {
			channel: string;
			team?: string;
			messageId: string;
			body: string;
			account?: string;
			tenant?: string;
			mentions?: string[];
			importance?: string;
			html?: boolean;
		},
		signal: AbortSignal | undefined,
		_onUpdate: undefined,
		ctx: ToolContext,
	): Promise<ToolResult> {
		return run(async () => {
			const conn = connectionFor(ctx, params.account, params.tenant);
			if (!params.body.trim()) return errorResult("Refusing to post an empty reply.");

			const { team, channel } = await requireChannel(conn, params.team, params.channel, "write", signal);
			const { mentions, error } = await resolveMentions(conn, params.mentions, signal);
			if (error) return errorResult(error);

			const me = await currentUser(conn, signal).catch(() => undefined);
			const target = `${team.displayName}/${channel.displayName}`;
			// The disclosure is added here, not asked for: what leaves the account
			// has to carry it whether or not the model remembered.
			const body = applyAiFooter(params.body, conn.aiFooter, { html: params.html });

			try {
				const sent = await replyToChannelMessage(
					conn,
					channel,
					params.messageId,
					{
						body,
						html: params.html,
						importance: params.importance,
						mentions,
					},
					signal,
				);

				auditWrite(conn.audit, {
					tool: "teams_reply_channel_message",
					account: conn.account,
					tenant: conn.tenant,
					actor: me?.upn,
					target: `channel:${target}#${params.messageId}`,
					summary: truncate(body, 200),
				});

				return textResult(
					[
						`✅ Replied in **${target}** (thread ${params.messageId}) as ${me?.displayName ?? "you"}.`,
						"",
						`> ${truncate(body, 300)}`,
						"",
						`messageId: ${sent.id}`,
					].join("\n"),
					{ teamId: team.id, channelId: channel.id, messageId: sent.id, threadId: params.messageId },
				);
			} catch (err) {
				auditWrite(conn.audit, {
					tool: "teams_reply_channel_message",
					account: conn.account,
					tenant: conn.tenant,
					actor: me?.upn,
					target: `channel:${target}#${params.messageId}`,
					summary: truncate(body, 200),
					error: err instanceof Error ? err.message : String(err),
				});
				throw err;
			}
		});
	},
};
