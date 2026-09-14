/**
 * Presence tools — read who is around, set the user's own status.
 */

import { Type } from "typebox";
import {
	AVAILABILITY_VALUES,
	clearPreferredPresence,
	getChatPresence,
	getMyPresence,
	getPresenceForUser,
	setPreferredPresence,
	setStatusMessage,
	type Availability,
} from "../graph/presence.ts";
import { resolveUserId } from "../graph/me.ts";
import { auditWrite } from "../safety/audit.ts";
import { assertAccess } from "../safety/index.ts";
import { formatPresence, formatPresenceList } from "../utils/formatting.ts";
import { requireChat } from "./resolve.ts";
import {
	AccountParam,
	TenantParam,
	connectionFor,
	currentUser,
	errorResult,
	run,
	textResult,
	type ToolContext,
	type ToolResult,
} from "./shared.ts";

export const teamsGetPresenceTool = {
	name: "teams_get_presence",
	description:
		"Show Microsoft Teams presence: the signed-in user's own status by default, a colleague's when 'person' " +
		"is given, or everyone in a chat when 'chat' is given. Use this before pinging someone — or to answer " +
		"'is Anna available right now?'.",
	parameters: Type.Object({
		person: Type.Optional(Type.String({ description: "Name, e-mail or UPN of the person to check" })),
		chat: Type.Optional(Type.String({ description: "Chat ID or name — shows presence for every member" })),
		account: AccountParam,
		tenant: TenantParam,
	}),
	promptSnippet: "Check Teams presence",

	async execute(
		_toolCallId: string,
		params: { person?: string; chat?: string; account?: string; tenant?: string },
		signal: AbortSignal | undefined,
		_onUpdate: undefined,
		ctx: ToolContext,
	): Promise<ToolResult> {
		return run(async () => {
			const conn = connectionFor(ctx, params.account, params.tenant);

			if (params.chat) {
				const chat = await requireChat(conn, params.chat, "read", signal);
				const entries = await getChatPresence(conn, chat.id, signal);
				return textResult(
					[`## Presence in ${chat.label}`, "", formatPresenceList(entries)].join("\n"),
					{ chatId: chat.id, count: entries.length },
				);
			}

			if (params.person) {
				const person = await resolveUserId(conn, params.person, signal);
				if (!person?.id) {
					return errorResult(
						`Could not resolve "${params.person}" to a person. Use teams_find_user to look them up.`,
					);
				}
				assertAccess(conn, "read", "people", person.displayName, [
					person.displayName,
					person.upn,
					person.mail,
					person.id,
				]);
				const presence = await getPresenceForUser(conn, person.id, person.displayName, signal);
				return textResult(formatPresence(presence), { userId: person.id, ...presence });
			}

			const mine = await getMyPresence(conn, signal);
			return textResult(formatPresence(mine), { ...mine });
		});
	},
};

export const teamsSetPresenceTool = {
	name: "teams_set_presence",
	description:
		"Set the signed-in user's own Microsoft Teams presence — the same thing as picking a status in the " +
		`Teams app. Allowed values: ${AVAILABILITY_VALUES.join(", ")}. ` +
		"Optionally give 'expiresIn' as an ISO 8601 duration (PT2H, PT30M) after which Teams reverts to " +
		"automatic presence. Set reset: true to hand control back to Teams immediately.",
	parameters: Type.Object({
		availability: Type.Optional(
			Type.String({ description: `One of: ${AVAILABILITY_VALUES.join(", ")}` }),
		),
		expiresIn: Type.Optional(
			Type.String({ description: "ISO 8601 duration, e.g. 'PT2H' for two hours" }),
		),
		reset: Type.Optional(
			Type.Boolean({ description: "Clear the preferred status and let Teams calculate it again" }),
		),
		account: AccountParam,
		tenant: TenantParam,
	}),
	promptSnippet: "Set the user's Teams presence",

	async execute(
		_toolCallId: string,
		params: {
			availability?: string;
			expiresIn?: string;
			reset?: boolean;
			account?: string;
			tenant?: string;
		},
		signal: AbortSignal | undefined,
		_onUpdate: undefined,
		ctx: ToolContext,
	): Promise<ToolResult> {
		return run(async () => {
			const conn = connectionFor(ctx, params.account, params.tenant);
			const me = await currentUser(conn, signal).catch(() => undefined);

			if (params.reset) {
				await clearPreferredPresence(conn, signal);
				auditWrite(conn.audit, {
					tool: "teams_set_presence",
					account: conn.account,
					tenant: conn.tenant,
					actor: me?.upn,
					target: "self",
					summary: "cleared preferred presence",
				});
				return textResult("✅ Teams presence is back under Teams' own control.", { reset: true });
			}

			const value = AVAILABILITY_VALUES.find(
				(candidate) => candidate.toLowerCase() === (params.availability ?? "").toLowerCase(),
			);
			if (!value) {
				return errorResult(
					`availability must be one of: ${AVAILABILITY_VALUES.join(", ")}.`,
				);
			}

			await setPreferredPresence(conn, value as Availability, {
				expirationDuration: params.expiresIn,
				signal,
			});

			auditWrite(conn.audit, {
				tool: "teams_set_presence",
				account: conn.account,
				tenant: conn.tenant,
				actor: me?.upn,
				target: "self",
				summary: `presence → ${value}${params.expiresIn ? ` for ${params.expiresIn}` : ""}`,
			});

			return textResult(
				`✅ Your Teams presence is now **${value}**${params.expiresIn ? ` (until ${params.expiresIn} from now)` : ""}.`,
				{ availability: value, expiresIn: params.expiresIn },
			);
		});
	},
};

export const teamsSetStatusMessageTool = {
	name: "teams_set_status_message",
	description:
		"Set the status message shown under the signed-in user's name in Microsoft Teams " +
		"(e.g. 'In a workshop until Thursday'). Pass an empty message to clear it.",
	parameters: Type.Object({
		message: Type.String({ description: "Status text; empty string clears it" }),
		expiresAt: Type.Optional(
			Type.String({ description: "ISO 8601 UTC timestamp when the message should expire" }),
		),
		account: AccountParam,
		tenant: TenantParam,
	}),
	promptSnippet: "Set the user's Teams status message",

	async execute(
		_toolCallId: string,
		params: { message: string; expiresAt?: string; account?: string; tenant?: string },
		signal: AbortSignal | undefined,
		_onUpdate: undefined,
		ctx: ToolContext,
	): Promise<ToolResult> {
		return run(async () => {
			const conn = connectionFor(ctx, params.account, params.tenant);
			await setStatusMessage(conn, params.message, { expiresAt: params.expiresAt, signal });

			const me = await currentUser(conn, signal).catch(() => undefined);
			auditWrite(conn.audit, {
				tool: "teams_set_status_message",
				account: conn.account,
				tenant: conn.tenant,
				actor: me?.upn,
				target: "self",
				summary: params.message ? `status message → "${params.message}"` : "cleared status message",
			});

			return textResult(
				params.message
					? `✅ Status message set to: "${params.message}"${params.expiresAt ? ` (until ${params.expiresAt})` : ""}`
					: "✅ Status message cleared.",
				{ message: params.message },
			);
		});
	},
};
