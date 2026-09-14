/**
 * teams_setup — write an account (or a tenant beneath one) to pi-teams.json.
 *
 * Exists so the extension can be configured by talking to pi, instead of
 * hand-editing JSON. Everything it writes can equally be written by hand; this
 * is a convenience, not a separate source of truth.
 */

import { Type } from "typebox";
import {
	getConfigPath,
	upsertAccount,
	upsertTenant,
	setDefaultAccount,
	validateSafetyLevel,
	type AccountConfig,
	type TenantConfig,
} from "../config/index.ts";
import type { PermissionBlock, ScopeList } from "../config/scope.ts";
import { formatPermissions, resolvePermissions } from "../config/scope.ts";
import { errorResult, run, textResult, type ToolResult } from "./shared.ts";

interface SetupParams {
	name: string;
	tenantId: string;
	clientId?: string;
	displayName?: string;
	clientSecret?: string;
	authMode?: string;
	safetyLevel?: string;
	parentAccount?: string;
	setDefault?: boolean;
	scopes?: string[];
	allowTeams?: string[];
	denyTeams?: string[];
	allowChannels?: string[];
	denyChannels?: string[];
	allowChats?: string[];
	denyChats?: string[];
	allowPeople?: string[];
	denyPeople?: string[];
	writeAllowChannels?: string[];
	writeDenyChannels?: string[];
	writeAllowChats?: string[];
	writeDenyChats?: string[];
	writeDenyPeople?: string[];
}

const StringList = (description: string) =>
	Type.Optional(Type.Array(Type.String(), { description }));

function list(allow?: string[], deny?: string[]): ScopeList | undefined {
	if (!allow?.length && !deny?.length) return undefined;
	const result: ScopeList = {};
	if (allow?.length) result.allow = allow;
	if (deny?.length) result.deny = deny;
	return result;
}

/** Turn the flat parameters into the nested permission block stored in JSON. */
export function buildPermissionBlock(params: SetupParams): PermissionBlock | undefined {
	const block: PermissionBlock = {};

	const teams = list(params.allowTeams, params.denyTeams);
	const channels = list(params.allowChannels, params.denyChannels);
	const chats = list(params.allowChats, params.denyChats);
	const people = list(params.allowPeople, params.denyPeople);

	if (teams) block.teams = teams;
	if (channels) block.channels = channels;
	if (chats) block.chats = chats;
	if (people) block.people = people;

	const writeChannels = list(params.writeAllowChannels, params.writeDenyChannels);
	const writeChats = list(params.writeAllowChats, params.writeDenyChats);
	const writePeople = list(undefined, params.writeDenyPeople);

	if (writeChannels || writeChats || writePeople) {
		block.write = {};
		if (writeChannels) block.write.channels = writeChannels;
		if (writeChats) block.write.chats = writeChats;
		if (writePeople) block.write.people = writePeople;
	}

	return Object.keys(block).length > 0 ? block : undefined;
}

export const teamsSetupTool = {
	name: "teams_setup",
	description:
		"Configure a Microsoft Teams account (or an additional tenant beneath an existing account) and write it to " +
		"~/.pi/agent/pi-teams.json. Set clientId and tenantId from your Entra ID app registration. " +
		"Use parentAccount to add a guest/customer tenant to an account that already exists. " +
		"Allow/deny lists restrict which teams, channels, chats and people pi may touch. " +
		"After setup, run teams_login to sign in.",
	parameters: Type.Object({
		name: Type.String({
			description: "Short account (or tenant) name used in every other tool, e.g. 'work', 'customer-alpha'",
		}),
		tenantId: Type.String({
			description: "Entra ID tenant ID or domain, e.g. 'contoso.onmicrosoft.com' or a GUID",
		}),
		clientId: Type.Optional(
			Type.String({
				description:
					"Application (client) ID of the Entra ID app registration. Required for a new account; a tenant inherits the account's.",
			}),
		),
		displayName: Type.Optional(Type.String({ description: "Friendly name shown in status output" })),
		clientSecret: Type.Optional(
			Type.String({
				description:
					"Client secret — only for app-only (client-credentials) access. Not needed, and not recommended, for acting as the user.",
			}),
		),
		authMode: Type.Optional(
			Type.String({
				description:
					"'device-code' (act as the signed-in user, default), 'client-credentials' (app-only, cannot post as a person), or 'auto'",
			}),
		),
		safetyLevel: Type.Optional(
			Type.String({
				description: "'open' (no prompts), 'confirm' (ask before every write, default), or 'readonly' (never write)",
			}),
		),
		parentAccount: Type.Optional(
			Type.String({
				description: "Add this as a tenant beneath an existing account instead of creating a new account",
			}),
		),
		setDefault: Type.Optional(
			Type.Boolean({ description: "Make this the default account for tools that omit 'account'" }),
		),
		scopes: StringList("Override the requested Microsoft Graph scopes (advanced)"),
		allowTeams: StringList("Teams pi may see at all. Glob patterns, e.g. ['Engineering', 'Project *']"),
		denyTeams: StringList("Teams pi must never touch"),
		allowChannels: StringList("Channels pi may see. Use 'Team/Channel' paths, e.g. ['Engineering/*']"),
		denyChannels: StringList("Channels pi must never touch, e.g. ['*/Announcements']"),
		allowChats: StringList("Chats pi may see — match on chat topic or participant name/e-mail"),
		denyChats: StringList("Chats pi must never touch"),
		allowPeople: StringList("People pi may interact with (name, UPN or e-mail patterns)"),
		denyPeople: StringList("People pi must never message, e.g. ['ceo@contoso.com']"),
		writeAllowChannels: StringList("Channels pi may POST in (narrower than allowChannels)"),
		writeDenyChannels: StringList("Channels pi may read but never post in"),
		writeAllowChats: StringList("Chats pi may send messages to"),
		writeDenyChats: StringList("Chats pi may read but never send to"),
		writeDenyPeople: StringList("People pi may read about but never message"),
	}),
	promptSnippet: "Configure a Microsoft Teams account for pi",
	promptGuidelines: [
		"Use teams_setup when the user wants to connect a Teams account, add a second company/tenant, or change which channels pi may use.",
		"Ask for the clientId and tenantId of their Entra ID app registration if they have not provided them.",
		"After teams_setup, tell the user to run teams_login (or /teams-login) to sign in as themselves.",
	],

	async execute(
		_toolCallId: string,
		params: SetupParams,
		_signal: AbortSignal | undefined,
		_onUpdate: undefined,
		_ctx: unknown,
	): Promise<ToolResult> {
		return run(async () => {
			const safetyLevel = params.safetyLevel ? validateSafetyLevel(params.safetyLevel) : undefined;
			if (params.safetyLevel && !safetyLevel) {
				return errorResult(
					`Invalid safetyLevel "${params.safetyLevel}". Use "open", "confirm" or "readonly".`,
				);
			}

			const permissions = buildPermissionBlock(params);

			if (params.parentAccount) {
				const tenant: TenantConfig = {
					name: params.name,
					tenantId: params.tenantId,
					clientId: params.clientId,
					clientSecret: params.clientSecret,
					authMode: params.authMode as TenantConfig["authMode"],
					scopes: params.scopes,
					safetyLevel,
					permissions,
				};
				upsertTenant(params.parentAccount, tenant);

				return textResult(
					[
						`✅ Tenant "${params.name}" added to account "${params.parentAccount}".`,
						"",
						`Config: ${getConfigPath()}`,
						"",
						permissions ? formatPermissions(resolvePermissions([permissions])) : "No scope restrictions configured — pi may use everything this identity can reach.",
						"",
						`Next: run teams_login with account "${params.parentAccount}" and tenant "${params.name}".`,
					].join("\n"),
					{ account: params.parentAccount, tenant: params.name },
				);
			}

			if (!params.clientId) {
				return errorResult(
					"clientId is required for a new account. Create an Entra ID app registration " +
						"(public client, 'Allow public client flows' enabled) and pass its Application (client) ID.",
				);
			}

			const account: AccountConfig = {
				name: params.name,
				displayName: params.displayName,
				tenantId: params.tenantId,
				clientId: params.clientId,
				clientSecret: params.clientSecret,
				authMode: (params.authMode as AccountConfig["authMode"]) ?? "device-code",
				scopes: params.scopes,
				safetyLevel: safetyLevel ?? "confirm",
				permissions,
			};

			upsertAccount(account);
			if (params.setDefault) setDefaultAccount(params.name);

			return textResult(
				[
					`✅ Teams account "${params.name}" saved.`,
					"",
					`- Tenant: ${params.tenantId}`,
					`- Client: ${params.clientId}`,
					`- Auth mode: ${account.authMode}`,
					`- Safety level: ${account.safetyLevel}`,
					`- Config file: ${getConfigPath()} (mode 0600)`,
					"",
					"### Effective scope rules",
					"",
					permissions
						? formatPermissions(resolvePermissions([permissions]))
						: "No restrictions configured — pi may use everything this identity can reach.",
					"",
					`Next: run teams_login${params.setDefault ? "" : ` with account "${params.name}"`} to sign in as yourself.`,
				].join("\n"),
				{ account: params.name, configPath: getConfigPath() },
			);
		});
	},
};
