/**
 * teams_status — who pi is acting as, and what it is allowed to do.
 */

import { Type } from "typebox";
import { buildConnectionCard, formatStatusText } from "../status.ts";
import { readCacheSummary } from "../auth/cache-plugin.ts";
import { readRootConfig } from "../config/index.ts";
import {
	AccountParam,
	TenantParam,
	connectionFor,
	run,
	textResult,
	type ToolContext,
	type ToolResult,
} from "./shared.ts";

export const teamsStatusTool = {
	name: "teams_status",
	description:
		"Show the current Microsoft Teams connection: which account and tenant pi is acting as, whether it is " +
		"signed in, the safety level, and the allow/deny rules in force. Call this when the user asks what pi " +
		"can do in Teams, or to check a connection before a write.",
	parameters: Type.Object({
		account: AccountParam,
		tenant: TenantParam,
	}),
	promptSnippet: "Show the Teams connection and permissions",

	async execute(
		_toolCallId: string,
		params: { account?: string; tenant?: string },
		_signal: AbortSignal | undefined,
		_onUpdate: undefined,
		ctx: ToolContext,
	): Promise<ToolResult> {
		return run(async () => {
			const conn = connectionFor(ctx, params.account, params.tenant);
			const card = buildConnectionCard(conn);
			const root = readRootConfig();

			const others: string[] = [];
			for (const account of root.accounts) {
				const tenants = [
					{ name: account.name, tenantId: account.tenantId },
					...(account.tenants ?? []).map((t) => ({ name: `${account.name}/${t.name}`, tenantId: t.tenantId })),
				];
				for (const entry of tenants) {
					const cache = readCacheSummary(account.name, entry.tenantId);
					const state = !cache.present
						? "not signed in"
						: cache.fresh
							? "signed in"
							: "session saved (renews on next use)";
					const who = cache.username ? ` — ${cache.username}` : "";
					others.push(`- ${entry.name}: ${state}${who}`);
				}
			}

			return textResult(
				[
					formatStatusText(card),
					"",
					"### All configured accounts",
					"",
					others.join("\n") || "(none)",
				].join("\n"),
				{ account: conn.account, tenant: conn.tenant, signedIn: card?.signedIn ?? false },
			);
		});
	},
};
