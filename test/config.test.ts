/**
 * Config resolution: the safety cascade, auth-mode resolution, and lookup.
 *
 * These run against a temporary PI_CODING_AGENT_DIR so no real config or token
 * file is ever touched.
 */

import { test, describe, before, after } from "node:test";
import assert from "node:assert/strict";
import { mkdtempSync, rmSync, existsSync, readFileSync, statSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";

const tempDir = mkdtempSync(join(tmpdir(), "pi-teams-test-"));
process.env.PI_CODING_AGENT_DIR = tempDir;

const {
	ensureConfigTemplate,
	findAccount,
	findTenant,
	getConfigPath,
	resolveAuthMode,
	resolveConnection,
	resolveEffectiveSafetyLevel,
	upsertAccount,
	upsertTenant,
	writeRootConfig,
	readRootConfig,
	removeAccount,
	setSafetyLevel,
	ConfigError,
} = await import("../src/config/index.ts");

const { canOpenBrowser } = await import("../src/utils/environment.ts");

const baseAccount = {
	name: "work",
	tenantId: "contoso.onmicrosoft.com",
	clientId: "11111111-1111-1111-1111-111111111111",
};

after(() => {
	rmSync(tempDir, { recursive: true, force: true });
});

describe("safety cascade", () => {
	test("tenant beats account beats global beats default", () => {
		const account = { ...baseAccount, safetyLevel: "open" as const };
		const tenant = { name: "t", tenantId: "t", safetyLevel: "readonly" as const };

		assert.equal(resolveEffectiveSafetyLevel("confirm", account, tenant), "readonly");
		assert.equal(resolveEffectiveSafetyLevel("confirm", account, undefined), "open");
		assert.equal(resolveEffectiveSafetyLevel("readonly", baseAccount, undefined), "readonly");
		assert.equal(resolveEffectiveSafetyLevel(undefined, baseAccount, undefined), "confirm");
	});

	test("an invalid level falls through instead of being trusted", () => {
		const account = { ...baseAccount, safetyLevel: "yolo" as never };
		assert.equal(resolveEffectiveSafetyLevel("readonly", account, undefined), "readonly");
	});
});

describe("auth mode", () => {
	test("auto picks client-credentials only when a secret exists", () => {
		assert.equal(resolveAuthMode(baseAccount, undefined, true, true), "client-credentials");
		assert.equal(resolveAuthMode(baseAccount, undefined, false, true), "interactive");
	});

	test("auto falls back to device code where no browser can be reached", () => {
		assert.equal(resolveAuthMode(baseAccount, undefined, false, false), "device-code");
	});

	test("a secret wins over the browser check", () => {
		// An app-only account has no user to put in front of a browser.
		assert.equal(resolveAuthMode(baseAccount, undefined, true, false), "client-credentials");
	});

	test("an explicit mode is respected, browser or not", () => {
		const account = { ...baseAccount, authMode: "device-code" as const };
		assert.equal(resolveAuthMode(account, undefined, true, true), "device-code");

		const interactive = { ...baseAccount, authMode: "interactive" as const };
		assert.equal(resolveAuthMode(interactive, undefined, false, false), "interactive");
	});

	test("a tenant may override the account's mode", () => {
		const account = { ...baseAccount, authMode: "interactive" as const };
		const tenant = { name: "t", tenantId: "t", authMode: "client-credentials" as const };
		assert.equal(resolveAuthMode(account, tenant, true, true), "client-credentials");
	});
});

describe("lookup", () => {
	const config = {
		accounts: [
			{ ...baseAccount, tenants: [{ name: "customer", tenantId: "customer.onmicrosoft.com" }] },
			{ name: "private", tenantId: "other.com", clientId: "2" },
		],
	};

	test("finds an account by name, display name or tenant id", () => {
		assert.equal(findAccount(config, "work")?.name, "work");
		assert.equal(findAccount(config, "WORK")?.name, "work");
		assert.equal(findAccount(config, "other.com")?.name, "private");
	});

	test("falls back to the first account when no name is given", () => {
		assert.equal(findAccount(config, undefined)?.name, "work");
	});

	test("the account's own name addresses its home tenant, not a sub-tenant", () => {
		const account = config.accounts[0]!;
		assert.equal(findTenant(account, "work"), undefined);
		assert.equal(findTenant(account, "customer")?.tenantId, "customer.onmicrosoft.com");
	});
});

describe("resolveConnection", () => {
	before(() => {
		writeRootConfig({
			accounts: [
				{
					...baseAccount,
					safetyLevel: "open",
					permissions: { channels: { deny: ["HR/*"] } },
					tenants: [
						{
							name: "customer",
							tenantId: "customer.onmicrosoft.com",
							safetyLevel: "readonly",
							permissions: { channels: { allow: ["Project/*"] } },
						},
					],
				},
			],
			defaultAccount: "work",
			safetyLevel: "confirm",
		});
	});

	test("resolves the home tenant with the account's settings", () => {
		const conn = resolveConnection();
		assert.equal(conn.account, "work");
		assert.equal(conn.tenant, "work");
		assert.equal(conn.safetyLevel, "open");
		// The default auth mode follows the environment: a browser where one is
		// reachable, the device code flow where it is not (SSH, no display).
		assert.equal(conn.authMode, canOpenBrowser() ? "interactive" : "device-code");
		assert.ok(conn.scopes.includes("ChannelMessage.Send"));
		assert.deepEqual(conn.permissions.write.channels.deny, ["HR/*"]);
	});

	test("resolves a sub-tenant with layered rules and its own safety level", () => {
		const conn = resolveConnection("work", "customer");
		assert.equal(conn.tenantId, "customer.onmicrosoft.com");
		assert.equal(conn.safetyLevel, "readonly");
		// The account's deny survives; the tenant's allow narrows.
		assert.deepEqual(conn.permissions.write.channels.deny, ["HR/*"]);
		assert.deepEqual(conn.permissions.write.channels.allow, ["Project/*"]);
	});

	test("an unknown tenant is an error, not a silent fallback", () => {
		assert.throws(() => resolveConnection("work", "nope"), ConfigError);
	});

	test("an unknown account is an error", () => {
		assert.throws(() => resolveConnection("nope"), ConfigError);
	});

	test("client-credentials without a secret is refused", () => {
		writeRootConfig({
			accounts: [{ ...baseAccount, name: "apponly", authMode: "client-credentials" }],
		});
		assert.throws(() => resolveConnection("apponly"), ConfigError);
	});
});

describe("mutations", () => {
	test("upsert adds, then updates in place", () => {
		writeRootConfig({ accounts: [] });
		upsertAccount({ ...baseAccount });
		assert.equal(readRootConfig().accounts.length, 1);

		upsertAccount({ ...baseAccount, displayName: "Work" });
		const config = readRootConfig();
		assert.equal(config.accounts.length, 1);
		assert.equal(config.accounts[0]!.displayName, "Work");
	});

	test("upsertTenant attaches a tenant to an existing account", () => {
		upsertTenant("work", { name: "customer", tenantId: "c.onmicrosoft.com" });
		assert.equal(readRootConfig().accounts[0]!.tenants?.[0]?.name, "customer");
	});

	test("setSafetyLevel writes at the level asked for", () => {
		setSafetyLevel("readonly", "work", "customer");
		assert.equal(readRootConfig().accounts[0]!.tenants?.[0]?.safetyLevel, "readonly");

		setSafetyLevel("open");
		assert.equal(readRootConfig().safetyLevel, "open");
	});

	test("removeAccount clears the default it pointed at", () => {
		assert.equal(removeAccount("work"), true);
		assert.equal(readRootConfig().accounts.length, 0);
		assert.equal(removeAccount("work"), false);
	});

	test("the config file is written owner-only", () => {
		upsertAccount({ ...baseAccount });
		assert.equal(statSync(getConfigPath()).mode & 0o777, 0o600);
	});
});

describe("template", () => {
	test("is created once and is valid JSON", () => {
		rmSync(getConfigPath(), { force: true });
		assert.equal(ensureConfigTemplate(), true);
		assert.equal(existsSync(getConfigPath()), true);

		const parsed = JSON.parse(readFileSync(getConfigPath(), "utf-8"));
		assert.ok(Array.isArray(parsed.accounts));
		assert.equal(parsed.safetyLevel, "confirm");

		// Existing files are never overwritten.
		assert.equal(ensureConfigTemplate(), false);
	});
});
