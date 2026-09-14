/**
 * Connection status helpers — pure, testable, TUI-free.
 *
 * Consumed by three surfaces: the footer status line, the persistent
 * connection card in the transcript, and the `/teams-status` command.
 */

import type { TeamsConnection } from "./config/index.ts";
import { readCacheSummary } from "./auth/cache-plugin.ts";
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
	/** Listen mode is switched on */
	watching?: boolean;
	/** One line on what the watcher is doing, when it runs */
	watchSummary?: string;
	/** The AI disclosure in force, rendered as one line; undefined when off */
	aiFooter?: string;
}

export function buildConnectionCard(
	conn: TeamsConnection | undefined,
): ConnectionCard | undefined {
	if (!conn) return undefined;

	const cache = readCacheSummary(conn.account, conn.tenantId);

	return {
		account: conn.account,
		accountDisplayName: conn.accountDisplayName,
		tenant: conn.tenant,
		tenantId: conn.tenantId,
		authMode: conn.authMode,
		safetyLevel: conn.safetyLevel,
		user: cache.username ?? cache.name,
		// A cached session counts as signed in even when the access token has
		// expired: MSAL renews it silently on the next call.
		signedIn: cache.present,
		expiresAt: cache.expiresAt ? new Date(cache.expiresAt).toISOString() : undefined,
		otherAccounts: conn.allAccounts.map((a) => a.name).filter((name) => name !== conn.account),
		permissionSummary: formatPermissions(conn.permissions),
		// What the configuration asks for; a caller that knows whether the watcher
		// is actually running overrides this with the runtime truth.
		watching: conn.watch.enabled,
		watchSummary: conn.watch.enabled
			? [
					`every ${conn.watch.intervalSeconds} s`,
					conn.watch.chats.length > 0 ? conn.watch.chats.join(", ") : "all recent chats",
					conn.watch.from.length > 0 ? `from ${conn.watch.from.join(", ")}` : undefined,
					`max ${conn.watch.maxTriggersPerHour}/h`,
				]
					.filter(Boolean)
					.join(" · ")
			: undefined,
		// Off is the normal state and needs no explanation; on gets the wording,
		// because that is what every recipient is about to read.
		aiFooter: conn.aiFooter.enabled ? `on — “${conn.aiFooter.text}”` : undefined,
	};
}

/** Short footer label, e.g. "✓ Teams · work · anna@contoso.com". */
export function buildConnectionLabel(card: ConnectionCard | undefined): string {
	if (!card) return "Teams · not configured";
	if (!card.signedIn) return `Teams · ${card.account} · not signed in`;

	const scope = card.tenant === card.account ? card.account : `${card.account}/${card.tenant}`;
	const listening = card.watching ? " · 👂 listening" : "";
	return `✓ Teams · ${scope}${card.user ? ` · ${card.user}` : ""}${listening}`;
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
	lines.push(
		`- **Auth mode:** ${card.authMode}${
			card.authMode === "client-credentials"
				? " (app-only — cannot post as a person)"
				: card.authMode === "interactive"
					? " (browser sign-in)"
					: " (device code)"
		}`,
	);
	lines.push(`- **Safety level:** ${card.safetyLevel}`);
	if (card.expiresAt) {
		lines.push(
			`- **Access token expires:** ${new Date(card.expiresAt).toLocaleString()} (renewed silently)`,
		);
	}
	if (card.otherAccounts.length > 0) {
		lines.push(`- **Other accounts:** ${card.otherAccounts.join(", ")}`);
	}
	lines.push(
		`- **Listen mode:** ${card.watching ? `on — ${card.watchSummary ?? "running"}` : "off"}`,
	);
	lines.push(`- **AI footer:** ${card.aiFooter ?? "off"}`);

	lines.push("", "### What pi may do", "", card.permissionSummary);
	return lines.join("\n");
}
