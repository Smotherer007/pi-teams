/**
 * Connection status helpers — pure, testable, TUI-free.
 *
 * Consumed by three surfaces: the footer status line, the persistent
 * connection card in the transcript, and the `/teams-status` command.
 */

import type { TeamsConnection } from "./config/index.ts";
import { getToken, isFresh } from "./auth/token-store.ts";
import { formatPermissions } from "./config/scope.ts";

/** Immutable snapshot of the active Teams connection. */
export interface ConnectionCard {
	account: string;
	accountDisplayName: string;
	tenant: string;
	tenantId: string;
	authMode: string;
	safetyLevel: string;
	/** Signed-in user, when a token is cached */
	user?: string;
	/** Whether a usable token exists right now */
	signedIn: boolean;
	/** ISO timestamp of token expiry, when known */
	expiresAt?: string;
	/** Other configured account names */
	otherAccounts: string[];
	/** Rendered scope rules */
	permissionSummary: string;
}

export function buildConnectionCard(
	conn: TeamsConnection | undefined,
): ConnectionCard | undefined {
	if (!conn) return undefined;

	const token = getToken(conn.account, conn.tenantId);

	return {
		account: conn.account,
		accountDisplayName: conn.accountDisplayName,
		tenant: conn.tenant,
		tenantId: conn.tenantId,
		authMode: conn.authMode,
		safetyLevel: conn.safetyLevel,
		user: token?.user?.upn ?? token?.user?.displayName,
		signedIn: isFresh(token) || !!token?.refreshToken,
		expiresAt: token ? new Date(token.expiresAt).toISOString() : undefined,
		otherAccounts: conn.allAccounts.map((a) => a.name).filter((name) => name !== conn.account),
		permissionSummary: formatPermissions(conn.permissions),
	};
}

/** Short footer label, e.g. "✓ Teams · work · anna@contoso.com". */
export function buildConnectionLabel(card: ConnectionCard | undefined): string {
	if (!card) return "Teams · not configured";
	if (!card.signedIn) return `Teams · ${card.account} · not signed in`;

	const scope = card.tenant === card.account ? card.account : `${card.account}/${card.tenant}`;
	return `✓ Teams · ${scope}${card.user ? ` · ${card.user}` : ""}`;
}

/** Rich markdown for `/teams-status` and the expanded card. */
export function formatStatusText(card: ConnectionCard | undefined): string {
	if (!card) {
		return [
			"## Microsoft Teams",
			"",
			"⚠️ **Not configured.**",
			"",
			"A template was created at `~/.pi/agent/pi-teams.json`. Fill in your Entra ID " +
				"`tenantId` and `clientId` — or run the `teams_setup` tool — then run `teams_login`.",
		].join("\n");
	}

	const lines: string[] = ["## Microsoft Teams", ""];
	lines.push(`- **Account:** ${card.accountDisplayName}${card.accountDisplayName !== card.account ? ` (${card.account})` : ""}`);
	if (card.tenant !== card.account) lines.push(`- **Tenant:** ${card.tenant}`);
	lines.push(`- **Tenant ID:** ${card.tenantId}`);
	lines.push(`- **Acting as:** ${card.signedIn ? (card.user ?? "(signed in)") : "not signed in — run teams_login"}`);
	lines.push(`- **Auth mode:** ${card.authMode}${card.authMode === "client-credentials" ? " (app-only — cannot post as a person)" : ""}`);
	lines.push(`- **Safety level:** ${card.safetyLevel}`);
	if (card.expiresAt) lines.push(`- **Token expires:** ${new Date(card.expiresAt).toLocaleString()}`);
	if (card.otherAccounts.length > 0) {
		lines.push(`- **Other accounts:** ${card.otherAccounts.join(", ")}`);
	}

	lines.push("", "### What pi may do", "", card.permissionSummary);
	return lines.join("\n");
}
