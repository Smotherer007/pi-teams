/**
 * teams_login — sign in as the user via the device code flow.
 *
 * The tool streams the code and URL through `onUpdate` as soon as Microsoft
 * issues them, then keeps polling. That ordering matters: the user has a
 * couple of minutes to act, and they cannot act on a code that only appears
 * after the tool returns.
 */

import { Type } from "typebox";
import { pollForToken, requestDeviceCode } from "../auth/device-code.ts";
import { acquireAppToken } from "../auth/client-credentials.ts";
import { clearMemoryCache, forgetToken, storeToken } from "../auth/index.ts";
import { clearAllTokens, deleteToken } from "../auth/token-store.ts";
import { tokenScopes } from "../auth/jwt.ts";
import { getMe } from "../graph/me.ts";
import { formatGraphError } from "../utils/errors.ts";
import {
	AccountParam,
	TenantParam,
	clearUserCache,
	connectionFor,
	connectionLabel,
	errorResult,
	run,
	textResult,
	type ToolContext,
	type ToolResult,
} from "./shared.ts";

interface LoginParams {
	account?: string;
	tenant?: string;
}

export const teamsLoginTool = {
	name: "teams_login",
	description:
		"Sign in to Microsoft Teams as yourself using the device code flow. Shows a short code and a URL; " +
		"open the URL, enter the code, and sign in with your normal credentials. The session is saved " +
		"(including a refresh token) so this is only needed once per account/tenant. " +
		"Use the 'account' and 'tenant' parameters to sign in to a second company or a guest tenant.",
	parameters: Type.Object({
		account: AccountParam,
		tenant: TenantParam,
	}),
	promptSnippet: "Sign in to Microsoft Teams as the user",
	promptGuidelines: [
		"Use teams_login when a Teams tool reports that pi is not signed in, or when the user wants to add a second account.",
		"Show the user code and the verification URL to the user verbatim — they cannot complete sign-in without them.",
	],

	async execute(
		_toolCallId: string,
		params: LoginParams,
		signal: AbortSignal | undefined,
		onUpdate: ((update: { content: Array<{ type: "text"; text: string }> }) => void) | undefined,
		ctx: ToolContext,
	): Promise<ToolResult> {
		return run(async () => {
			const conn = connectionFor(ctx, params.account, params.tenant);
			const label = connectionLabel(conn);

			// App-only needs no user interaction at all.
			if (conn.authMode === "client-credentials") {
				const token = await acquireAppToken(conn, signal);
				storeToken(conn, token);
				return textResult(
					[
						`✅ Acquired an app-only token for ${label}.`,
						"",
						`Granted app roles: ${tokenScopes(token.accessToken).join(", ") || "(none visible in the token)"}`,
						"",
						"Note: app-only tokens cannot post messages as a person. Switch this account to " +
							'authMode "device-code" if pi should act as you.',
					].join("\n"),
					{ account: conn.account, tenant: conn.tenant, authMode: "client-credentials" },
				);
			}

			const challenge = await requestDeviceCode(conn, signal);

			const instructions = [
				`## Sign in to Teams (${label})`,
				"",
				`1. Open **${challenge.verificationUri}**`,
				`2. Enter the code: **${challenge.userCode}**`,
				"3. Sign in with the account pi should act as.",
				"",
				`The code expires at ${new Date(challenge.expiresAt).toLocaleTimeString()}. Waiting…`,
			].join("\n");

			onUpdate?.({ content: [{ type: "text", text: instructions }] });

			let lastNotice = 0;
			const token = await pollForToken(conn, challenge, {
				signal,
				onPending: (secondsLeft) => {
					// One nudge every 30s — enough to show progress, not enough to spam.
					if (secondsLeft > 0 && Date.now() - lastNotice > 30_000) {
						lastNotice = Date.now();
						onUpdate?.({
							content: [
								{
									type: "text",
									text: `${instructions}\n\nStill waiting for sign-in… (${secondsLeft}s left)`,
								},
							],
						});
					}
				},
			});

			storeToken(conn, token);
			clearUserCache();

			// Confirm the token actually works and report who pi now is.
			let identity = token.user?.upn ?? token.user?.displayName ?? "(unknown)";
			let jobTitle: string | undefined;
			try {
				const me = await getMe(conn, signal);
				identity = `${me.displayName} <${me.upn}>`;
				jobTitle = me.jobTitle;
			} catch (err) {
				return textResult(
					[
						`✅ Signed in to ${label}, but the first Graph call failed: ${formatGraphError(err)}`,
						"",
						"Run teams_doctor to see which permissions are missing.",
					].join("\n"),
					{ account: conn.account, tenant: conn.tenant, warning: true },
				);
			}

			const granted = tokenScopes(token.accessToken);
			const requested = conn.scopes.filter((s) => !["openid", "profile", "offline_access"].includes(s));
			const missing = requested.filter(
				(scope) => !granted.some((g) => g.toLowerCase() === scope.toLowerCase()),
			);

			return textResult(
				[
					`✅ Signed in to ${label} as **${identity}**${jobTitle ? ` (${jobTitle})` : ""}.`,
					"",
					`pi now acts as this user in Teams. Safety level: **${conn.safetyLevel}**.`,
					missing.length > 0
						? `\n⚠️ Not granted: ${missing.join(", ")}\nSome tools will fail until an admin consents to these scopes.`
						: "\nAll requested scopes were granted.",
				].join("\n"),
				{
					account: conn.account,
					tenant: conn.tenant,
					user: identity,
					grantedScopes: granted,
					missingScopes: missing,
				},
			);
		});
	},
};

export const teamsLogoutTool = {
	name: "teams_logout",
	description:
		"Sign out of a Teams account: removes the saved access and refresh tokens from " +
		"~/.pi/agent/pi-teams-tokens.json. The account stays configured, so teams_login can sign in again.",
	parameters: Type.Object({
		account: AccountParam,
		tenant: TenantParam,
		all: Type.Optional(Type.Boolean({ description: "Remove the cached tokens of every account" })),
	}),
	promptSnippet: "Sign out of a Teams account",

	async execute(
		_toolCallId: string,
		params: { account?: string; tenant?: string; all?: boolean },
		_signal: AbortSignal | undefined,
		_onUpdate: undefined,
		ctx: ToolContext,
	): Promise<ToolResult> {
		return run(async () => {
			clearUserCache();

			if (params.all) {
				const count = clearAllTokens();
				clearMemoryCache();
				return textResult(`✅ Removed ${count} cached Teams session(s).`, { removed: count });
			}

			const conn = connectionFor(ctx, params.account, params.tenant);
			const removed = deleteToken(conn.account, conn.tenantId);
			forgetToken(conn);

			return removed
				? textResult(`✅ Signed out of ${connectionLabel(conn)}.`, { account: conn.account })
				: errorResult(`No saved session for ${connectionLabel(conn)}.`);
		});
	},
};
