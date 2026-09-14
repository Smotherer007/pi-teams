/**
 * teams_accounts — list, switch between, and remove accounts and tenants.
 */

import { Type } from "typebox";
import {
	getConfigPath,
	readRootConfig,
	removeAccount,
	removeTenant,
	setDefaultAccount,
	setSafetyLevel,
	validateSafetyLevel,
} from "../config/index.ts";
import { readCacheSummary, removeCache } from "../auth/cache-plugin.ts";
import { resetApps } from "../auth/index.ts";
import { clearUserCache, errorResult, run, textResult, type ToolResult } from "./shared.ts";

interface AccountsParams {
	action?: string;
	account?: string;
	tenant?: string;
	safetyLevel?: string;
}

export const teamsAccountsTool = {
	name: "teams_accounts",
	description:
		"Manage configured Teams accounts and their tenants: list them, switch the default, change a safety level, " +
		"or delete one. Use this for multi-account and multi-company setups. " +
		"Actions: 'list' (default), 'use', 'delete', 'safety'.",
	parameters: Type.Object({
		action: Type.Optional(
			Type.String({ description: "'list' (default), 'use', 'delete', or 'safety'" }),
		),
		account: Type.Optional(Type.String({ description: "Account name the action applies to" })),
		tenant: Type.Optional(
			Type.String({ description: "Tenant beneath that account; omit to target the account itself" }),
		),
		safetyLevel: Type.Optional(
			Type.String({ description: "For action 'safety': 'open', 'confirm' or 'readonly'" }),
		),
	}),
	promptSnippet: "List or switch Microsoft Teams accounts",
	promptGuidelines: [
		"Use teams_accounts when the user mentions a different company, tenant, or Teams identity.",
		"Prefer passing account/tenant directly to the tool you need instead of switching the default.",
	],

	async execute(
		_toolCallId: string,
		params: AccountsParams,
		_signal: AbortSignal | undefined,
		_onUpdate: undefined,
		_ctx: unknown,
	): Promise<ToolResult> {
		return run(async () => {
			const action = (params.action ?? "list").toLowerCase();
			const config = readRootConfig();

			if (config.accounts.length === 0) {
				return textResult(
					`No Teams accounts configured yet. Run teams_setup, or edit ${getConfigPath()}.`,
					{ accounts: [] },
				);
			}

			switch (action) {
				case "list": {
					const lines = ["## Configured Teams accounts", ""];
					for (const account of config.accounts) {
						const isDefault = config.defaultAccount === account.name;
						const cache = readCacheSummary(account.name, account.tenantId);
						const state = !cache.present
							? "not signed in"
							: cache.fresh
								? "signed in"
								: "session saved";

						lines.push(
							`### ${account.displayName ?? account.name}${isDefault ? " (default)" : ""}`,
							"",
							`- name: \`${account.name}\``,
							`- tenant: ${account.tenantId}`,
							`- auth mode: ${account.authMode ?? "device-code"}`,
							`- safety: ${account.safetyLevel ?? config.safetyLevel ?? "confirm"}`,
							`- status: ${state}${cache.username ? ` as ${cache.username}` : ""}`,
						);

						for (const tenant of account.tenants ?? []) {
							const tenantCache = readCacheSummary(account.name, tenant.tenantId);
							const tenantState = !tenantCache.present
								? "not signed in"
								: tenantCache.fresh
									? "signed in"
									: "session saved";
							lines.push(
								`  - tenant \`${tenant.name}\` (${tenant.tenantId}) — ${tenantState}` +
									`${tenant.safetyLevel ? `, safety: ${tenant.safetyLevel}` : ""}`,
							);
						}
						lines.push("");
					}
					lines.push(`Config: ${getConfigPath()}`);
					return textResult(lines.join("\n"), {
						accounts: config.accounts.map((a) => a.name),
						defaultAccount: config.defaultAccount,
					});
				}

				case "use": {
					if (!params.account) return errorResult("action 'use' requires an account name.");
					setDefaultAccount(params.account, params.tenant);
					clearUserCache();
					return textResult(
						`✅ Default Teams account is now "${params.account}"${params.tenant ? ` (tenant ${params.tenant})` : ""}.`,
						{ defaultAccount: params.account, defaultTenant: params.tenant },
					);
				}

				case "delete": {
					if (!params.account) return errorResult("action 'delete' requires an account name.");

					if (params.tenant) {
						const tenantConfig = config.accounts
							.find((a) => a.name === params.account)
							?.tenants?.find((t) => t.name === params.tenant);
						const removed = removeTenant(params.account, params.tenant);
						if (removed && tenantConfig) {
							removeCache(params.account, tenantConfig.tenantId);
							resetApps();
						}
						return removed
							? textResult(`✅ Removed tenant "${params.tenant}" from account "${params.account}".`, {})
							: errorResult(`Tenant "${params.tenant}" not found in account "${params.account}".`);
					}

					const accountConfig = config.accounts.find((a) => a.name === params.account);
					const removed = removeAccount(params.account);
					if (removed && accountConfig) {
						removeCache(params.account, accountConfig.tenantId);
						for (const tenant of accountConfig.tenants ?? []) {
							removeCache(params.account, tenant.tenantId);
						}
						resetApps();
					}
					clearUserCache();
					return removed
						? textResult(`✅ Removed account "${params.account}" and its cached sessions.`, {})
						: errorResult(`Account "${params.account}" not found.`);
				}

				case "safety": {
					const level = validateSafetyLevel(params.safetyLevel);
					if (!level) {
						return errorResult(
							`action 'safety' requires safetyLevel to be "open", "confirm" or "readonly".`,
						);
					}
					setSafetyLevel(level, params.account, params.tenant);
					const target = params.account
						? params.tenant
							? `${params.account}/${params.tenant}`
							: params.account
						: "globally";
					return textResult(`✅ Safety level set to "${level}" ${target}.`, { safetyLevel: level });
				}

				default:
					return errorResult(`Unknown action "${action}". Use 'list', 'use', 'delete' or 'safety'.`);
			}
		});
	},
};
