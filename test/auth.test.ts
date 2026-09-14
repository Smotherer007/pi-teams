/**
 * Auth wiring.
 *
 * No network: these prove that MSAL accepts our configuration, that the cache
 * plugin round-trips a blob, and that an unauthenticated connection fails with
 * the error that tells the user to sign in — rather than some MSAL internal.
 */

import { test, describe, after, beforeEach } from "node:test";
import assert from "node:assert/strict";
import { mkdtempSync, rmSync, existsSync, statSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";

const tempDir = mkdtempSync(join(tmpdir(), "pi-teams-auth-"));
process.env.PI_CODING_AGENT_DIR = tempDir;

const { writeRootConfig, resolveConnection } = await import("../src/config/index.ts");
const { getAccessToken, NotSignedInError, hasCachedSession, signOutAll } = await import(
	"../src/auth/index.ts"
);
const { createCachePlugin, getCachePath, readCacheSummary, cacheKey, removeCache } = await import(
	"../src/auth/cache-plugin.ts"
);
const { getPublicApp, requestScopes } = await import("../src/auth/msal.ts");

after(() => {
	rmSync(tempDir, { recursive: true, force: true });
});

beforeEach(() => {
	writeRootConfig({
		accounts: [
			{
				name: "work",
				tenantId: "contoso.onmicrosoft.com",
				clientId: "00000000-0000-0000-0000-000000000001",
				authMode: "interactive",
			},
		],
		defaultAccount: "work",
	});
});

describe("MSAL application", () => {
	test("accepts our configuration and exposes a token cache", () => {
		const conn = resolveConnection();
		const app = getPublicApp(conn);
		assert.ok(app.getTokenCache(), "expected a token cache");
	});

	test("the same connection reuses one application instance", () => {
		const conn = resolveConnection();
		assert.equal(getPublicApp(conn), getPublicApp(conn));
	});
});

describe("requestScopes", () => {
	test("strips the scopes MSAL insists on adding itself", () => {
		const conn = resolveConnection();
		const scopes = requestScopes({
			...conn,
			scopes: ["openid", "profile", "offline_access", "Chat.ReadWrite"],
		});
		assert.deepEqual(scopes, ["Chat.ReadWrite"]);
	});

	test("leaves a normal scope list untouched", () => {
		const conn = resolveConnection();
		assert.ok(requestScopes(conn).includes("ChannelMessage.Send"));
		assert.ok(!requestScopes(conn).some((s) => s === "offline_access"));
	});
});

describe("cache plugin", () => {
	const account = "work";
	const tenantId = "contoso.onmicrosoft.com";

	test("writes nothing when the cache did not change", async () => {
		const plugin = createCachePlugin(account, tenantId);
		await plugin.afterCacheAccess({
			cacheHasChanged: false,
			tokenCache: { serialize: () => "{}", deserialize: () => {} },
		} as never);
		assert.equal(existsSync(getCachePath(account, tenantId)), false);
	});

	test("persists a changed cache owner-only and reads it back", async () => {
		const blob = JSON.stringify({
			Account: { a: { username: "anna@contoso.com", name: "Anna Schmidt" } },
			AccessToken: { t: { expires_on: String(Math.floor(Date.now() / 1000) + 3600) } },
		});

		const plugin = createCachePlugin(account, tenantId);
		await plugin.afterCacheAccess({
			cacheHasChanged: true,
			tokenCache: { serialize: () => blob, deserialize: () => {} },
		} as never);

		const path = getCachePath(account, tenantId);
		assert.equal(existsSync(path), true);
		assert.equal(statSync(path).mode & 0o777, 0o600);

		let handedBack: string | undefined;
		await plugin.beforeCacheAccess({
			cacheHasChanged: false,
			tokenCache: { serialize: () => "", deserialize: (value: string) => (handedBack = value) },
		} as never);
		assert.equal(handedBack, blob);
	});

	test("readCacheSummary reports the identity and freshness", () => {
		const summary = readCacheSummary(account, tenantId);
		assert.equal(summary.present, true);
		assert.equal(summary.username, "anna@contoso.com");
		assert.equal(summary.fresh, true);
	});

	test("an expired access token is present but not fresh", async () => {
		const blob = JSON.stringify({
			Account: { a: { username: "anna@contoso.com" } },
			AccessToken: { t: { expires_on: String(Math.floor(Date.now() / 1000) - 60) } },
		});
		await createCachePlugin("stale", "t").afterCacheAccess({
			cacheHasChanged: true,
			tokenCache: { serialize: () => blob, deserialize: () => {} },
		} as never);

		const summary = readCacheSummary("stale", "t");
		assert.equal(summary.present, true);
		assert.equal(summary.fresh, false);
		removeCache("stale", "t");
	});

	test("a corrupt cache degrades to 'not signed in' instead of throwing", async () => {
		await createCachePlugin("broken", "t").afterCacheAccess({
			cacheHasChanged: true,
			tokenCache: { serialize: () => "{not json", deserialize: () => {} },
		} as never);
		assert.deepEqual(readCacheSummary("broken", "t"), { present: false, fresh: false });
		removeCache("broken", "t");
	});

	test("cache keys are filesystem-safe", () => {
		assert.equal(cacheKey("Work Account", "contoso.onmicrosoft.com"), "work_account__contoso.onmicrosoft.com");
		assert.ok(!cacheKey("a/b", "c:d").includes("/"));
	});
});

describe("getAccessToken", () => {
	test("asks the user to sign in when nothing is cached", async () => {
		signOutAll();
		const conn = resolveConnection();
		assert.equal(hasCachedSession(conn), false);

		await assert.rejects(() => getAccessToken(conn), NotSignedInError);
		await assert.rejects(() => getAccessToken(conn), /teams_login/);
	});

	test("client-credentials without a secret never reaches the network", async () => {
		writeRootConfig({
			accounts: [
				{
					name: "apponly",
					tenantId: "contoso.onmicrosoft.com",
					clientId: "00000000-0000-0000-0000-000000000002",
					authMode: "client-credentials",
				},
			],
		});
		// resolveConnection refuses it first — the misconfiguration is caught
		// before any token request is built.
		assert.throws(() => resolveConnection("apponly"), /clientSecret/);
	});
});
