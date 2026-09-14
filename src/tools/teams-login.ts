/**
 * teams_login — sign in as the user.
 * teams_logout — forget a saved session.
 *
 * The browser flow is the default: pi opens the Microsoft sign-in page, the
 * user signs in with their own credentials and MFA, and MSAL catches the
 * redirect on a loopback port. Where no browser can be reached — SSH, a
 * container, a headless box — it falls back to the device code flow and says
 * why.
 */

import { Type } from "typebox";
import {
	browserUnavailableReason,
	canOpenBrowser,
	defaultSignInMode,
	signIn,
	signOut,
	signOutAll,
	type DeviceCodeInfo,
	type SignInMode,
} from "../auth/index.ts";
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
	mode?: string;
	port?: number;
}

type Update = (update: { content: Array<{ type: "text"; text: string }> }) => void;

/** What the user has to do, rendered the moment MSAL tells us. */
function deviceCodeInstructions(label: string, info: DeviceCodeInfo, why?: string): string {
	return [
		`## Sign in to Teams (${label})`,
		"",
		why ? `_Using the device code flow because ${why}._` : "",
		"",
		`1. Open **${info.verificationUri}**`,
		`2. Enter the code: **${info.userCode}**`,
		"3. Sign in with the account pi should act as.",
		"",
		`The code expires in ${Math.round(info.expiresIn / 60)} minutes. Waiting…`,
	]
		.filter((line) => line !== "")
		.join("\n");
}

export const teamsLoginTool = {
	name: "teams_login",
	description:
		"Sign in to Microsoft Teams as yourself. By default this opens your browser at the Microsoft sign-in " +
		"page and completes automatically — nothing to type. On a machine with no browser (SSH, a container) " +
		"it falls back to the device code flow and shows a short code to enter on another device. " +
		"The session is saved and renewed silently, so this is normally needed once per account. " +
		"Use 'account' and 'tenant' to sign in to a second company or a guest tenant.",
	parameters: Type.Object({
		account: AccountParam,
		tenant: TenantParam,
		mode: Type.Optional(
			Type.String({
				description:
					"Force a flow: 'interactive' (browser) or 'device-code'. Omit to let pi choose for this machine.",
			}),
		),
		port: Type.Optional(
			Type.Number({
				description:
					"Fixed loopback port for the browser redirect. Only needed when the app registration lists an exact redirect URI such as http://localhost:3000.",
			}),
		),
	}),
	promptSnippet: "Sign in to Microsoft Teams as the user",
	promptGuidelines: [
		"Use teams_login when a Teams tool reports that pi is not signed in, or when the user wants to add a second account or tenant.",
		"If the device code flow is used, show the user code and URL to the user verbatim — they cannot finish signing in without them.",
	],

	async execute(
		_toolCallId: string,
		params: LoginParams,
		signal: AbortSignal | undefined,
		onUpdate: Update | undefined,
		ctx: ToolContext,
	): Promise<ToolResult> {
		return run(async () => {
			const conn = connectionFor(ctx, params.account, params.tenant);
			const label = connectionLabel(conn);

			// App-only needs no user interaction at all.
			if (conn.authMode === "client-credentials") {
				const token = await signIn(conn, { signal });
				return textResult(
					[
						`✅ Acquired an app-only token for ${label}.`,
						"",
						`Granted app roles: ${tokenScopes(token.accessToken).join(", ") || "(none visible in the token)"}`,
						"",
						"Note: app-only tokens cannot post messages as a person. Switch this account to " +
							'authMode "interactive" if pi should act as you.',
					].join("\n"),
					{ account: conn.account, tenant: conn.tenant, authMode: "client-credentials" },
				);
			}

			const requested = params.mode?.toLowerCase();
			if (requested && requested !== "interactive" && requested !== "device-code") {
				return errorResult(`mode must be 'interactive' or 'device-code', not "${params.mode}".`);
			}

			let mode: SignInMode = (requested as SignInMode | undefined) ?? defaultSignInMode();
			let fallbackReason: string | undefined;

			// An explicit request for the browser on a machine that has none would
			// hang until the timeout; say so and use the code flow instead.
			if (mode === "interactive" && !canOpenBrowser()) {
				fallbackReason = browserUnavailableReason();
				mode = "device-code";
			} else if (mode === "device-code" && !requested && !canOpenBrowser()) {
				fallbackReason = browserUnavailableReason();
			}

			if (mode === "interactive") {
				onUpdate?.({
					content: [
						{
							type: "text",
							text: [
								`## Signing in to Teams (${label})`,
								"",
								"Opening your browser. Sign in there with the account pi should act as —",
								"the page closes itself when you are done.",
							].join("\n"),
						},
					],
				});
			}

			const result = await signIn(conn, {
				mode,
				signal,
				loopbackPort: params.port ?? conn.loopbackPort,
				onDeviceCode: (info) => {
					onUpdate?.({
						content: [{ type: "text", text: deviceCodeInstructions(label, info, fallbackReason) }],
					});
				},
			});

			clearUserCache();

			// Confirm the token actually works, and report who pi now is.
			let identity = result.account?.username ?? result.account?.name ?? "(unknown)";
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

			const granted = tokenScopes(result.accessToken);
			const missing = conn.scopes.filter(
				(scope) => !granted.some((entry) => entry.toLowerCase() === scope.toLowerCase()),
			);

			return textResult(
				[
					`✅ Signed in to ${label} as **${identity}**${jobTitle ? ` (${jobTitle})` : ""}.`,
					"",
					`Flow: ${result.mode === "interactive" ? "browser sign-in" : "device code"}` +
						`${fallbackReason ? ` (${fallbackReason})` : ""}`,
					`pi now acts as this user in Teams. Safety level: **${conn.safetyLevel}**.`,
					missing.length > 0
						? `\n⚠️ Not granted: ${missing.join(", ")}\nSome tools will fail until an admin consents to these scopes.`
						: "\nAll requested scopes were granted.",
				].join("\n"),
				{
					account: conn.account,
					tenant: conn.tenant,
					user: identity,
					mode: result.mode,
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
		"Sign out of a Teams account: removes the saved session from ~/.pi/agent/pi-teams-tokens/. " +
		"The account stays configured, so teams_login can sign in again.",
	parameters: Type.Object({
		account: AccountParam,
		tenant: TenantParam,
		all: Type.Optional(Type.Boolean({ description: "Remove the saved sessions of every account" })),
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
				const count = signOutAll();
				return textResult(`✅ Removed ${count} saved Teams session(s).`, { removed: count });
			}

			const conn = connectionFor(ctx, params.account, params.tenant);
			const removed = await signOut(conn);

			return removed
				? textResult(`✅ Signed out of ${connectionLabel(conn)}.`, { account: conn.account })
				: errorResult(`No saved session for ${connectionLabel(conn)}.`);
		});
	},
};
