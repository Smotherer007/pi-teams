/**
 * teams_doctor — diagnose configuration, sign-in, consent and connectivity.
 *
 * Written to be the first thing anyone runs when something does not work, and
 * to answer the question a 403 never answers: *which* permission is missing,
 * and who has to grant it.
 */

import { Type } from "typebox";
import {
	getConfigPath,
	resolveAllConnections,
	DEFAULT_SCOPES,
	type TeamsConnection,
} from "../config/index.ts";
import { getToken, getTokenPath, isFresh } from "../auth/token-store.ts";
import { tokenScopes } from "../auth/jwt.ts";
import { getAccessToken } from "../auth/index.ts";
import { getMe } from "../graph/me.ts";
import { listJoinedTeams } from "../graph/teams.ts";
import { formatGraphError, describeScope } from "../utils/errors.ts";
import { formatPermissions } from "../config/scope.ts";
import { getAuditPath } from "../safety/audit.ts";
import { run, textResult, type ToolResult } from "./shared.ts";

interface DoctorParams {
	account?: string;
	/** Skip the live Graph calls */
	offline?: boolean;
}

async function checkConnection(
	conn: TeamsConnection,
	offline: boolean,
	signal?: AbortSignal,
): Promise<string[]> {
	const label = conn.tenant === conn.account ? conn.account : `${conn.account}/${conn.tenant}`;
	const lines = [`### ${label}`, ""];

	lines.push(`- tenant: \`${conn.tenantId}\``);
	lines.push(`- client: \`${conn.clientId}\``);
	lines.push(`- auth mode: ${conn.authMode}`);
	lines.push(`- safety level: ${conn.safetyLevel}`);

	if (conn.authMode === "client-credentials") {
		lines.push(
			"- ⚠️ app-only: pi cannot post messages as a person with this mode " +
				"(Microsoft Graph restricts app-only message writes to migration scenarios)",
		);
	}

	const token = getToken(conn.account, conn.tenantId);
	if (!token) {
		lines.push("- ❌ not signed in — run `teams_login`");
		return lines;
	}

	lines.push(
		isFresh(token)
			? `- ✅ token valid until ${new Date(token.expiresAt).toLocaleString()}`
			: token.refreshToken
				? "- ⚠️ access token expired, refresh token present (will renew on the next call)"
				: "- ❌ token expired and no refresh token — run `teams_login`",
	);

	// Consent check works offline: the scopes are inside the token.
	const granted = tokenScopes(token.accessToken);
	if (granted.length > 0) {
		const wanted = (conn.scopes.length > 0 ? conn.scopes : DEFAULT_SCOPES).filter(
			(scope) => !["openid", "profile", "offline_access"].includes(scope),
		);
		const missing = wanted.filter(
			(scope) => !granted.some((g) => g.toLowerCase() === scope.toLowerCase()),
		);
		if (missing.length === 0) {
			lines.push(`- ✅ all ${wanted.length} requested scopes granted`);
		} else {
			lines.push("- ⚠️ missing consent:");
			for (const scope of missing) {
				const purpose = describeScope(scope);
				lines.push(`  - \`${scope}\`${purpose ? ` — needed for ${purpose}` : ""}`);
			}
			lines.push(
				"  Sign in again after an admin grants consent, or remove the unused scopes from the account's `scopes`.",
			);
		}
	}

	if (offline) return lines;

	try {
		await getAccessToken(conn, signal);
	} catch (err) {
		lines.push(`- ❌ could not obtain a token: ${formatGraphError(err)}`);
		return lines;
	}

	try {
		const me = await getMe(conn, signal);
		lines.push(`- ✅ acting as **${me.displayName}** <${me.upn}>`);
	} catch (err) {
		lines.push(`- ❌ /me failed: ${formatGraphError(err)}`);
		return lines;
	}

	try {
		const teams = await listJoinedTeams(conn, 5, signal);
		lines.push(`- ✅ can list teams (${teams.length === 5 ? "5+" : teams.length} visible)`);
	} catch (err) {
		lines.push(`- ⚠️ listing teams failed: ${formatGraphError(err)}`);
	}

	return lines;
}

export const teamsDoctorTool = {
	name: "teams_doctor",
	description:
		"Diagnose the Microsoft Teams setup: configuration, sign-in state, which Graph scopes were actually " +
		"granted, and whether live calls succeed — for every configured account and tenant. " +
		"Run this first whenever a Teams tool fails.",
	parameters: Type.Object({
		account: Type.Optional(Type.String({ description: "Check only this account" })),
		offline: Type.Optional(
			Type.Boolean({ description: "Skip live Graph calls and only check config and cached tokens" }),
		),
	}),
	promptSnippet: "Check the Teams configuration and sign-in health",
	promptGuidelines: [
		"Run teams_doctor when any Teams tool returns an auth, permission or configuration error.",
		"Report the missing scopes verbatim — the user may need an admin to consent to them.",
	],

	async execute(
		_toolCallId: string,
		params: DoctorParams,
		signal: AbortSignal | undefined,
		_onUpdate: undefined,
		_ctx: unknown,
	): Promise<ToolResult> {
		return run(async () => {
			const { connections, errors } = resolveAllConnections();
			const selected = params.account
				? connections.filter((c) => c.account.toLowerCase() === params.account!.toLowerCase())
				: connections;

			const lines = ["## Teams doctor", ""];
			lines.push(`- config: \`${getConfigPath()}\``);
			lines.push(`- tokens: \`${getTokenPath()}\``);
			lines.push(`- audit log: \`${getAuditPath()}\``);
			lines.push("");

			if (errors.length > 0) {
				lines.push("### Configuration problems", "");
				for (const error of errors) lines.push(`- ❌ ${error}`);
				lines.push("");
			}

			if (selected.length === 0) {
				lines.push(
					params.account
						? `No connection found for account "${params.account}".`
						: "No usable connections configured. Run `teams_setup` to add one.",
				);
				return textResult(lines.join("\n"), { ok: false });
			}

			for (const conn of selected) {
				lines.push(...(await checkConnection(conn, params.offline === true, signal)));
				lines.push("");
			}

			const first = selected[0]!;
			lines.push("### Scope rules of the default connection", "", formatPermissions(first.permissions));

			return textResult(lines.join("\n"), {
				accounts: selected.map((c) => `${c.account}/${c.tenant}`),
				configErrors: errors,
			});
		});
	},
};
