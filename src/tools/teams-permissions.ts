/**
 * teams_permissions — inspect, test and change the guardrails.
 *
 * "Test" is the interesting action: it answers "would pi be allowed to post in
 * Engineering/Announcements?" without posting anything, which is how a user
 * builds confidence in a rule set before turning safetyLevel to "open".
 */

import { Type } from "typebox";
import { setPermissionBlock } from "../config/index.ts";
import {
	checkScope,
	formatPermissions,
	SCOPE_CATEGORIES,
	type PermissionBlock,
	type ScopeCategory,
	type ScopeList,
	type ScopeMode,
} from "../config/scope.ts";
import {
	AccountParam,
	TenantParam,
	connectionFor,
	connectionLabel,
	errorResult,
	run,
	textResult,
	type ToolContext,
	type ToolResult,
} from "./shared.ts";

interface PermissionsParams {
	action?: string;
	account?: string;
	tenant?: string;
	category?: string;
	mode?: string;
	target?: string;
	allow?: string[];
	deny?: string[];
}

const isCategory = (value: string): value is ScopeCategory =>
	(SCOPE_CATEGORIES as string[]).includes(value);

export const teamsPermissionsTool = {
	name: "teams_permissions",
	description:
		"Inspect or change what pi is allowed to do in Teams. " +
		"action 'show' (default) prints the effective allow/deny rules; " +
		"action 'test' checks whether a specific team, channel, chat or person is in scope without touching it; " +
		"action 'set' writes an allow/deny list for one category. " +
		"Rules apply per category (teams, channels, chats, people) and per mode (read, write).",
	parameters: Type.Object({
		action: Type.Optional(Type.String({ description: "'show' (default), 'test' or 'set'" })),
		account: AccountParam,
		tenant: TenantParam,
		category: Type.Optional(
			Type.String({ description: "'teams', 'channels', 'chats' or 'people'" }),
		),
		mode: Type.Optional(Type.String({ description: "'read' or 'write' (default 'write')" })),
		target: Type.Optional(
			Type.String({
				description: "For action 'test': what to check, e.g. 'Engineering/Announcements' or 'anna@contoso.com'",
			}),
		),
		allow: Type.Optional(Type.Array(Type.String(), { description: "For action 'set': allow patterns" })),
		deny: Type.Optional(Type.Array(Type.String(), { description: "For action 'set': deny patterns" })),
	}),
	promptSnippet: "Show or change what pi may do in Teams",
	promptGuidelines: [
		"Use teams_permissions with action 'test' before telling the user that pi can or cannot reach a channel.",
		"When a tool fails with a scope error, show the user the matching rule instead of retrying.",
	],

	async execute(
		_toolCallId: string,
		params: PermissionsParams,
		_signal: AbortSignal | undefined,
		_onUpdate: undefined,
		ctx: ToolContext,
	): Promise<ToolResult> {
		return run(async () => {
			const action = (params.action ?? "show").toLowerCase();
			const conn = connectionFor(ctx, params.account, params.tenant);
			const label = connectionLabel(conn);

			if (action === "show") {
				return textResult(
					[
						`## What pi may do in Teams (${label})`,
						"",
						`Safety level: **${conn.safetyLevel}** · auth mode: **${conn.authMode}**`,
						"",
						formatPermissions(conn.permissions),
						"",
						"_Patterns are case-insensitive globs. Channels match on `Team/Channel`, their name, or their ID; " +
							"chats match on topic, participant names and e-mail addresses._",
					].join("\n"),
					{ safetyLevel: conn.safetyLevel, permissions: conn.permissions },
				);
			}

			if (action === "test") {
				if (!params.target) return errorResult("action 'test' requires a target.");
				const category = (params.category ?? "channels").toLowerCase();
				if (!isCategory(category)) {
					return errorResult(`Unknown category "${category}". Use teams, channels, chats or people.`);
				}
				const mode = ((params.mode ?? "write").toLowerCase() === "read" ? "read" : "write") as ScopeMode;

				const decision = checkScope(conn.permissions, mode, category, [params.target]);
				return textResult(
					[
						`${decision.allowed ? "✅ Allowed" : "🚫 Blocked"}: ${mode} ${category} "${params.target}" (${label})`,
						"",
						decision.reason,
					].join("\n"),
					{ allowed: decision.allowed, reason: decision.reason, pattern: decision.pattern },
				);
			}

			if (action === "set") {
				const category = (params.category ?? "").toLowerCase();
				if (!isCategory(category)) {
					return errorResult(
						`action 'set' requires a category: teams, channels, chats or people.`,
					);
				}
				if (!params.allow?.length && !params.deny?.length) {
					return errorResult("action 'set' requires at least one allow or deny pattern.");
				}

				const list: ScopeList = {};
				if (params.allow?.length) list.allow = params.allow;
				if (params.deny?.length) list.deny = params.deny;

				const mode = params.mode?.toLowerCase();
				const block: PermissionBlock = {};
				if (mode === "read" || mode === "write") {
					block[mode] = { [category]: list };
				} else {
					block[category] = list;
				}

				// The block replaces the rules at this level; anything the user wants
				// to keep must be passed again — stated plainly in the result.
				setPermissionBlock(block, params.account, params.tenant);

				return textResult(
					[
						`✅ Scope rules for **${category}** updated ` +
							`${params.account ? (params.tenant ? `on ${params.account}/${params.tenant}` : `on account ${params.account}`) : "globally"}.`,
						"",
						params.allow?.length ? `allow: ${params.allow.join(", ")}` : "",
						params.deny?.length ? `deny: ${params.deny.join(", ")}` : "",
						"",
						"_This replaces the permission block at that level. Re-run teams_permissions with action 'show' to check the result._",
					]
						.filter(Boolean)
						.join("\n"),
					{ category, allow: params.allow, deny: params.deny },
				);
			}

			return errorResult(`Unknown action "${action}". Use 'show', 'test' or 'set'.`);
		});
	},
};
