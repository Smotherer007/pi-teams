/**
 * Configuration — accounts, tenants, safety levels and scope rules.
 *
 * Single source of truth: `~/.pi/agent/pi-teams.json` (override the directory
 * with PI_CODING_AGENT_DIR). No environment variables are read for account
 * data, so what the file says is what pi may do.
 *
 * Shape
 * -----
 *   accounts[]            one entry per Teams identity you sign in as
 *     .tenants[]          additional tenants that identity reaches
 *                         (guest access, customer tenants, sub-companies)
 *
 * Both levels may narrow `safetyLevel` and `permissions`; the most specific
 * setting wins for safety, and scope rules are layered global → account →
 * tenant (see ./scope.ts).
 *
 * The file can hold a client secret, so it is written 0600 via an atomic
 * temp-file rename — the same treatment the token cache gets.
 */

import {
	existsSync,
	readFileSync,
	mkdirSync,
	writeFileSync,
	chmodSync,
	statSync,
	unlinkSync,
	renameSync,
} from "node:fs";
import { dirname, join } from "node:path";
import { homedir } from "node:os";
import {
	defaultPermissions,
	resolvePermissions,
	type PermissionBlock,
	type ResolvedPermissions,
} from "./scope.ts";
import { canOpenBrowser } from "../utils/environment.ts";

// ---------------------------------------------------------------------------
// Types
// ---------------------------------------------------------------------------

/**
 * How pi obtains a token.
 *
 * - `interactive`: delegated. The browser opens at the Microsoft sign-in page
 *   and MSAL catches the redirect on a loopback port. pi then acts **as the
 *   signed-in user** — messages come from their account, and everything they
 *   can see, pi can see.
 * - `device-code`: delegated as well, but the user types a short code on
 *   another device. For SSH sessions, containers, and anything headless.
 * - `client-credentials`: app-only. No user identity; Microsoft Graph does not
 *   allow posting chat or channel messages this way (outside of migration
 *   scenarios), so writes are refused with an explanation.
 * - `auto`: `client-credentials` when a secret is configured, otherwise
 *   `interactive` where a browser is reachable and `device-code` where it is
 *   not.
 */
export type AuthMode = "interactive" | "device-code" | "client-credentials" | "auto";

/** An auth mode after "auto" has been resolved away. */
export type ResolvedAuthMode = "interactive" | "device-code" | "client-credentials";

/** Safety level for mutation operations. */
export type SafetyLevel = "open" | "confirm" | "readonly";

/** A tenant reachable by an account — guest access, a customer, a sub-company. */
export interface TenantConfig {
	/** Short label used in tool parameters, e.g. "kunde-alpha" */
	name: string;
	/** Tenant ID or domain, e.g. "contoso.onmicrosoft.com" */
	tenantId: string;
	/** Client ID override (defaults to the account's) */
	clientId?: string;
	/** Client secret for app-only access in this tenant */
	clientSecret?: string;
	/** Auth mode override */
	authMode?: AuthMode;
	/** Requested Graph scopes override */
	scopes?: string[];
	/** Fixed loopback port for the browser sign-in redirect */
	loopbackPort?: number;
	/** Safety level override — most specific wins */
	safetyLevel?: SafetyLevel;
	/** Scope rules layered on top of the account's */
	permissions?: PermissionBlock;
}

/** A Teams identity pi can sign in as. */
export interface AccountConfig {
	/** Short label used in tool parameters, e.g. "neoimpulse" */
	name: string;
	/** Friendly name for status output */
	displayName?: string;
	/** Home tenant ID or domain */
	tenantId: string;
	/** Entra ID application (client) ID */
	clientId: string;
	/** Client secret — only for client-credentials mode */
	clientSecret?: string;
	/** Auth mode (default: auto) */
	authMode?: AuthMode;
	/** Requested Graph scopes (default: DEFAULT_SCOPES) */
	scopes?: string[];
	/**
	 * Fixed loopback port for the browser sign-in redirect.
	 *
	 * Only needed where the app registration lists an exact redirect URI such
	 * as `http://localhost:3000`; by default MSAL picks a free port and Entra
	 * accepts any port on localhost for a public client.
	 */
	loopbackPort?: number;
	/** Safety level override */
	safetyLevel?: SafetyLevel;
	/** Scope rules layered on top of the global ones */
	permissions?: PermissionBlock;
	/** Additional tenants this identity reaches */
	tenants?: TenantConfig[];
}

/** Root shape of pi-teams.json. */
export interface TeamsRootConfig {
	accounts: AccountConfig[];
	/** Account used when a tool omits `account` (default: first) */
	defaultAccount?: string;
	/** Tenant used when a tool omits `tenant` (default: the account's home tenant) */
	defaultTenant?: string;
	/** Global safety level (default: confirm) */
	safetyLevel?: SafetyLevel;
	/** Global scope rules */
	permissions?: PermissionBlock;
	/** Default page size for message listings (default: 25) */
	maxMessages?: number;
	/** Append every write to ~/.pi/agent/pi-teams-audit.jsonl (default: true) */
	audit?: boolean;
	/** Graph base URL override (sovereign clouds) */
	graphBaseUrl?: string;
	/** Login authority host override (sovereign clouds) */
	authorityHost?: string;
}

/** A single resolved account+tenant connection. */
export interface TeamsConnection {
	/** Account label */
	account: string;
	/** Friendly account name */
	accountDisplayName: string;
	/** Tenant label — equals the account label for the home tenant */
	tenant: string;
	/** Tenant ID or domain used as the authority */
	tenantId: string;
	clientId: string;
	clientSecret?: string;
	/** Effective auth mode — never "auto" after resolution */
	authMode: ResolvedAuthMode;
	scopes: string[];
	/** Fixed loopback port for the browser sign-in, when one is configured */
	loopbackPort?: number;
	safetyLevel: SafetyLevel;
	permissions: ResolvedPermissions;
	maxMessages: number;
	audit: boolean;
	graphBaseUrl: string;
	authorityHost: string;
	/** Every configured account, for multi-account awareness */
	allAccounts: AccountConfig[];
}

/** Thrown when the configuration cannot produce a usable connection. */
export class ConfigError extends Error {
	readonly missing: string[];
	constructor(missing: string[], message?: string) {
		super(message ?? `Missing required Teams configuration: ${missing.join(", ")}`);
		this.name = "ConfigError";
		this.missing = missing;
	}
}

// ---------------------------------------------------------------------------
// Defaults
// ---------------------------------------------------------------------------

/**
 * Delegated scopes requested by default.
 *
 * `openid`, `profile` and `offline_access` are absent on purpose: MSAL always
 * requests them and rejects them in an explicit scope list. The lasting
 * sign-in they buy is therefore automatic.
 */
export const DEFAULT_SCOPES: string[] = [
	"User.Read",
	"User.ReadBasic.All",
	"Team.ReadBasic.All",
	"Channel.ReadBasic.All",
	"ChannelMessage.Read.All",
	"ChannelMessage.Send",
	"ChannelMember.Read.All",
	"Chat.ReadWrite",
	"ChatMessage.Send",
	"Presence.ReadWrite",
	"Presence.Read.All",
	"Calendars.ReadWrite",
	"OnlineMeetings.ReadWrite",
	"Files.Read.All",
	"Sites.Read.All",
];

/** Scope set for app-only tokens (the `.default` of the app registration). */
export const CLIENT_CREDENTIALS_SCOPE = "https://graph.microsoft.com/.default";

const DEFAULTS = {
	safetyLevel: "confirm" as SafetyLevel,
	authMode: "auto" as AuthMode,
	maxMessages: 25,
	audit: true,
	graphBaseUrl: "https://graph.microsoft.com/v1.0",
	authorityHost: "https://login.microsoftonline.com",
};

const VALID_AUTH_MODES = new Set<string>([
	"interactive",
	"device-code",
	"client-credentials",
	"auto",
]);
const VALID_SAFETY_LEVELS = new Set<string>(["open", "confirm", "readonly"]);

// ---------------------------------------------------------------------------
// Paths
// ---------------------------------------------------------------------------

export function getAgentDir(): string {
	return process.env.PI_CODING_AGENT_DIR ?? join(homedir(), ".pi", "agent");
}

export function getConfigPath(): string {
	return join(getAgentDir(), "pi-teams.json");
}

// ---------------------------------------------------------------------------
// File I/O
// ---------------------------------------------------------------------------

export function readRootConfig(): TeamsRootConfig {
	const configPath = getConfigPath();
	if (!existsSync(configPath)) {
		ensureConfigTemplate();
		return { accounts: [] };
	}
	try {
		// Older versions may have written the file world-readable; it can hold a
		// client secret, so tighten it whenever we touch it.
		try {
			const mode = statSync(configPath).mode & 0o777;
			if (mode !== 0o600) chmodSync(configPath, 0o600);
		} catch {
			/* ignore */
		}
		const parsed = JSON.parse(readFileSync(configPath, "utf-8")) as TeamsRootConfig;
		if (!parsed || typeof parsed !== "object") return { accounts: [] };
		return { ...parsed, accounts: Array.isArray(parsed.accounts) ? parsed.accounts : [] };
	} catch {
		return { accounts: [] };
	}
}

/** Write the config atomically with 0600 permissions. */
export function writeRootConfig(config: TeamsRootConfig): void {
	const configPath = getConfigPath();
	const dir = dirname(configPath);
	if (!existsSync(dir)) mkdirSync(dir, { recursive: true, mode: 0o700 });

	const tmpPath = `${configPath}.${process.pid}.tmp`;
	try {
		writeFileSync(tmpPath, `${JSON.stringify(config, null, 2)}\n`, { encoding: "utf-8", mode: 0o600 });
		chmodSync(tmpPath, 0o600);
		// renameSync is atomic within a filesystem, so a crash mid-write cannot
		// leave a truncated config behind.
		renameSync(tmpPath, configPath);
	} catch (err) {
		try {
			unlinkSync(tmpPath);
		} catch {
			/* ignore */
		}
		throw err;
	}
}

// ---------------------------------------------------------------------------
// Validation helpers
// ---------------------------------------------------------------------------

function validateAuthMode(value: string | undefined): AuthMode | undefined {
	if (!value) return undefined;
	const normalized = value.toLowerCase().trim();
	return VALID_AUTH_MODES.has(normalized) ? (normalized as AuthMode) : undefined;
}

export function validateSafetyLevel(value: string | undefined): SafetyLevel | undefined {
	if (!value) return undefined;
	const normalized = value.toLowerCase().trim();
	return VALID_SAFETY_LEVELS.has(normalized) ? (normalized as SafetyLevel) : undefined;
}

/**
 * Effective safety level: tenant > account > global > default.
 * Exported for tests — the cascade is easy to get subtly wrong.
 */
export function resolveEffectiveSafetyLevel(
	globalLevel: SafetyLevel | undefined,
	account: AccountConfig | undefined,
	tenant: TenantConfig | undefined,
): SafetyLevel {
	return (
		validateSafetyLevel(tenant?.safetyLevel) ??
		validateSafetyLevel(account?.safetyLevel) ??
		validateSafetyLevel(globalLevel) ??
		DEFAULTS.safetyLevel
	);
}

/**
 * Effective auth mode, resolving "auto" against the available credentials and
 * the environment.
 *
 * `browserAvailable` is injected rather than detected here so the cascade stays
 * a pure function — the detection itself lives in ../auth/msal.ts.
 */
export function resolveAuthMode(
	account: AccountConfig | undefined,
	tenant: TenantConfig | undefined,
	hasSecret: boolean,
	browserAvailable = true,
): ResolvedAuthMode {
	const configured =
		validateAuthMode(tenant?.authMode) ??
		validateAuthMode(account?.authMode) ??
		DEFAULTS.authMode;

	if (configured !== "auto") return configured;
	if (hasSecret) return "client-credentials";
	return browserAvailable ? "interactive" : "device-code";
}

// ---------------------------------------------------------------------------
// Lookup
// ---------------------------------------------------------------------------

export function findAccount(
	config: TeamsRootConfig,
	name: string | undefined,
): AccountConfig | undefined {
	if (config.accounts.length === 0) return undefined;
	if (!name) return config.accounts[0];
	const lower = name.toLowerCase();
	return config.accounts.find(
		(a) =>
			a.name?.toLowerCase() === lower ||
			a.displayName?.toLowerCase() === lower ||
			a.tenantId?.toLowerCase() === lower,
	);
}

export function findTenant(
	account: AccountConfig,
	name: string | undefined,
): TenantConfig | undefined {
	if (!name) return undefined;
	const lower = name.toLowerCase();
	// The account label also addresses the home tenant.
	if (account.name.toLowerCase() === lower) return undefined;
	return (account.tenants ?? []).find(
		(t) => t.name?.toLowerCase() === lower || t.tenantId?.toLowerCase() === lower,
	);
}

// ---------------------------------------------------------------------------
// Resolution
// ---------------------------------------------------------------------------

/**
 * Resolve one account+tenant into a usable connection.
 *
 * @throws {ConfigError} when no account matches or required fields are missing
 */
export function resolveConnection(
	accountName?: string,
	tenantName?: string,
	root?: TeamsRootConfig,
): TeamsConnection {
	const config = root ?? readRootConfig();

	if (config.accounts.length === 0) {
		throw new ConfigError(
			["accounts"],
			`No Teams accounts configured. Run the teams_setup tool, or edit ${getConfigPath()}.`,
		);
	}

	const account = findAccount(config, accountName ?? config.defaultAccount);
	if (!account) {
		throw new ConfigError(
			[`account "${accountName}"`],
			`Account "${accountName}" not found. Available: ${config.accounts.map((a) => a.name).join(", ")}`,
		);
	}

	const requestedTenant = tenantName ?? (accountName ? undefined : config.defaultTenant);
	const tenant = findTenant(account, requestedTenant);

	if (requestedTenant && !tenant && account.name.toLowerCase() !== requestedTenant.toLowerCase()) {
		const available = [account.name, ...(account.tenants ?? []).map((t) => t.name)].join(", ");
		throw new ConfigError(
			[`tenant "${requestedTenant}"`],
			`Tenant "${requestedTenant}" not found in account "${account.name}". Available: ${available}`,
		);
	}

	const missing: string[] = [];
	const tenantId = tenant?.tenantId ?? account.tenantId;
	const clientId = tenant?.clientId ?? account.clientId;
	if (!tenantId) missing.push(`tenantId for account "${account.name}"`);
	if (!clientId) missing.push(`clientId for account "${account.name}"`);
	if (missing.length > 0) throw new ConfigError(missing);

	const clientSecret = tenant?.clientSecret ?? account.clientSecret ?? undefined;
	const authMode = resolveAuthMode(account, tenant, !!clientSecret, canOpenBrowser());

	if (authMode === "client-credentials" && !clientSecret) {
		throw new ConfigError(
			["clientSecret"],
			`Account "${account.name}" is set to client-credentials but has no clientSecret.`,
		);
	}

	const scopes =
		tenant?.scopes ??
		account.scopes ??
		(authMode === "client-credentials" ? [CLIENT_CREDENTIALS_SCOPE] : DEFAULT_SCOPES);

	return {
		account: account.name,
		accountDisplayName: account.displayName ?? account.name,
		tenant: tenant?.name ?? account.name,
		tenantId,
		clientId,
		clientSecret,
		authMode,
		loopbackPort: tenant?.loopbackPort ?? account.loopbackPort,
		scopes,
		safetyLevel: resolveEffectiveSafetyLevel(config.safetyLevel, account, tenant),
		permissions: resolvePermissions([config.permissions, account.permissions, tenant?.permissions]),
		maxMessages: config.maxMessages ?? DEFAULTS.maxMessages,
		audit: config.audit ?? DEFAULTS.audit,
		graphBaseUrl: (config.graphBaseUrl ?? DEFAULTS.graphBaseUrl).replace(/\/+$/, ""),
		authorityHost: (config.authorityHost ?? DEFAULTS.authorityHost).replace(/\/+$/, ""),
		allAccounts: config.accounts,
	};
}

export function tryResolveConnection(
	accountName?: string,
	tenantName?: string,
): TeamsConnection | undefined {
	try {
		return resolveConnection(accountName, tenantName);
	} catch {
		return undefined;
	}
}

/**
 * Every configured account × tenant combination.
 * Used by `teams_doctor` to validate the whole file at once.
 */
export function resolveAllConnections(): { connections: TeamsConnection[]; errors: string[] } {
	const config = readRootConfig();
	const errors: string[] = [];
	const connections: TeamsConnection[] = [];

	if (config.accounts.length === 0) {
		errors.push("No accounts configured in pi-teams.json");
		return { connections, errors };
	}

	for (const account of config.accounts) {
		const tenantNames: (string | undefined)[] = [undefined, ...(account.tenants ?? []).map((t) => t.name)];
		for (const tenantName of tenantNames) {
			try {
				connections.push(resolveConnection(account.name, tenantName, config));
			} catch (err) {
				const label = tenantName ? `${account.name}/${tenantName}` : account.name;
				errors.push(`${label}: ${err instanceof Error ? err.message : String(err)}`);
			}
		}
	}

	return { connections, errors };
}

/** Default permissions, exported so callers never have to build them by hand. */
export { defaultPermissions };

// ---------------------------------------------------------------------------
// Mutations — used by teams_setup / teams_accounts
// ---------------------------------------------------------------------------

/** Insert or update an account, keeping existing tenants unless replaced. */
export function upsertAccount(account: AccountConfig): TeamsRootConfig {
	const config = readRootConfig();
	const index = config.accounts.findIndex((a) => a.name.toLowerCase() === account.name.toLowerCase());

	if (index >= 0) {
		const existing = config.accounts[index]!;
		config.accounts[index] = {
			...existing,
			...account,
			tenants: account.tenants ?? existing.tenants,
			permissions: account.permissions ?? existing.permissions,
		};
	} else {
		config.accounts.push(account);
	}

	if (!config.defaultAccount) config.defaultAccount = config.accounts[0]!.name;
	if (!config.safetyLevel) config.safetyLevel = DEFAULTS.safetyLevel;

	writeRootConfig(config);
	return config;
}

/** Insert or update a tenant beneath an account. */
export function upsertTenant(accountName: string, tenant: TenantConfig): TeamsRootConfig {
	const config = readRootConfig();
	const account = config.accounts.find((a) => a.name.toLowerCase() === accountName.toLowerCase());
	if (!account) {
		throw new ConfigError(
			[`account "${accountName}"`],
			`Account "${accountName}" not found. Available: ${config.accounts.map((a) => a.name).join(", ") || "none"}`,
		);
	}

	account.tenants = account.tenants ?? [];
	const index = account.tenants.findIndex((t) => t.name.toLowerCase() === tenant.name.toLowerCase());
	if (index >= 0) {
		account.tenants[index] = { ...account.tenants[index]!, ...tenant };
	} else {
		account.tenants.push(tenant);
	}

	writeRootConfig(config);
	return config;
}

export function removeAccount(name: string): boolean {
	const config = readRootConfig();
	const index = config.accounts.findIndex((a) => a.name.toLowerCase() === name.toLowerCase());
	if (index < 0) return false;
	config.accounts.splice(index, 1);
	if (config.defaultAccount?.toLowerCase() === name.toLowerCase()) {
		config.defaultAccount = config.accounts[0]?.name;
	}
	writeRootConfig(config);
	return true;
}

export function removeTenant(accountName: string, tenantName: string): boolean {
	const config = readRootConfig();
	const account = config.accounts.find((a) => a.name.toLowerCase() === accountName.toLowerCase());
	if (!account?.tenants) return false;
	const index = account.tenants.findIndex((t) => t.name.toLowerCase() === tenantName.toLowerCase());
	if (index < 0) return false;
	account.tenants.splice(index, 1);
	writeRootConfig(config);
	return true;
}

export function setDefaultAccount(name: string, tenantName?: string): void {
	const config = readRootConfig();
	const account = config.accounts.find((a) => a.name.toLowerCase() === name.toLowerCase());
	if (!account) {
		throw new ConfigError(
			[`account "${name}"`],
			`Account "${name}" not found. Available: ${config.accounts.map((a) => a.name).join(", ") || "none"}`,
		);
	}
	config.defaultAccount = account.name;
	config.defaultTenant = tenantName;
	writeRootConfig(config);
}

/** Set the safety level globally, on an account, or on a tenant. */
export function setSafetyLevel(
	level: SafetyLevel,
	accountName?: string,
	tenantName?: string,
): void {
	const config = readRootConfig();

	if (!accountName) {
		config.safetyLevel = level;
		writeRootConfig(config);
		return;
	}

	const account = config.accounts.find((a) => a.name.toLowerCase() === accountName.toLowerCase());
	if (!account) {
		throw new ConfigError([`account "${accountName}"`], `Account "${accountName}" not found.`);
	}

	if (!tenantName) {
		account.safetyLevel = level;
	} else {
		const tenant = (account.tenants ?? []).find(
			(t) => t.name.toLowerCase() === tenantName.toLowerCase(),
		);
		if (!tenant) {
			throw new ConfigError(
				[`tenant "${tenantName}"`],
				`Tenant "${tenantName}" not found in account "${accountName}".`,
			);
		}
		tenant.safetyLevel = level;
	}

	writeRootConfig(config);
}

/** Replace the permission block globally, on an account, or on a tenant. */
export function setPermissionBlock(
	permissions: PermissionBlock,
	accountName?: string,
	tenantName?: string,
): void {
	const config = readRootConfig();

	if (!accountName) {
		config.permissions = permissions;
		writeRootConfig(config);
		return;
	}

	const account = config.accounts.find((a) => a.name.toLowerCase() === accountName.toLowerCase());
	if (!account) {
		throw new ConfigError([`account "${accountName}"`], `Account "${accountName}" not found.`);
	}

	if (!tenantName) {
		account.permissions = permissions;
	} else {
		const tenant = (account.tenants ?? []).find(
			(t) => t.name.toLowerCase() === tenantName.toLowerCase(),
		);
		if (!tenant) {
			throw new ConfigError(
				[`tenant "${tenantName}"`],
				`Tenant "${tenantName}" not found in account "${accountName}".`,
			);
		}
		tenant.permissions = permissions;
	}

	writeRootConfig(config);
}

// ---------------------------------------------------------------------------
// Template
// ---------------------------------------------------------------------------

const TEMPLATE_JSON = `{
  "accounts": [
    {
      "name": "work",
      "displayName": "Work account",
      "tenantId": "contoso.onmicrosoft.com",
      "clientId": "00000000-0000-0000-0000-000000000000",
      "authMode": "interactive",
      "safetyLevel": "confirm",
      "permissions": {
        "read": {
          "teams": { "allow": ["*"] },
          "chats": { "allow": ["*"] }
        },
        "write": {
          "channels": { "allow": ["*"], "deny": ["*/Announcements"] },
          "chats": { "allow": ["*"] },
          "people": { "deny": [] }
        }
      },
      "tenants": []
    }
  ],
  "defaultAccount": "work",
  "safetyLevel": "confirm",
  "maxMessages": 25,
  "audit": true
}
`;

/**
 * Create `~/.pi/agent/pi-teams.json` from a template when it does not exist.
 * Returns true when a file was created.
 */
export function ensureConfigTemplate(): boolean {
	const configPath = getConfigPath();
	if (existsSync(configPath)) return false;
	try {
		mkdirSync(dirname(configPath), { recursive: true, mode: 0o700 });
		writeFileSync(configPath, TEMPLATE_JSON, { encoding: "utf-8", mode: 0o600 });
		chmodSync(configPath, 0o600);
		return true;
	} catch {
		return false;
	}
}
