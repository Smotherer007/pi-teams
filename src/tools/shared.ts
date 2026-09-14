/**
 * Shared tool plumbing — parameter schemas, result helpers, connection
 * resolution and the guards every tool runs before touching Graph.
 */

import { Type } from "typebox";
import {
	resolveConnection,
	type TeamsConnection,
} from "../config/index.ts";
import { getMe } from "../graph/me.ts";
import { formatGraphError } from "../utils/errors.ts";
import type { SignedInUser } from "../types.ts";

// ---------------------------------------------------------------------------
// Common parameters
// ---------------------------------------------------------------------------

/** Which configured account to act as. */
export const AccountParam = Type.Optional(
	Type.String({
		description:
			"Configured account name to act as (see teams_accounts). Uses the default account if omitted.",
	}),
);

/** Which tenant beneath that account — for guest access and customer tenants. */
export const TenantParam = Type.Optional(
	Type.String({
		description:
			"Tenant name configured beneath the account (guest or customer tenant). Uses the account's home tenant if omitted.",
	}),
);

export const LimitParam = Type.Optional(
	Type.Number({ description: "Maximum number of items to return" }),
);

// ---------------------------------------------------------------------------
// Results
// ---------------------------------------------------------------------------

export interface ToolResult {
	content: Array<{ type: "text"; text: string }>;
	details: Record<string, unknown>;
}

export function textResult(text: string, details: Record<string, unknown> = {}): ToolResult {
	return { content: [{ type: "text", text }], details };
}

export function errorResult(message: string): ToolResult {
	return { content: [{ type: "text", text: `❌ ${message}` }], details: { error: true } };
}

/** Wrap a tool body so every failure comes back as a readable message. */
export async function run(fn: () => Promise<ToolResult>): Promise<ToolResult> {
	try {
		return await fn();
	} catch (err) {
		return errorResult(formatGraphError(err));
	}
}

// ---------------------------------------------------------------------------
// Connection
// ---------------------------------------------------------------------------

/** The context every tool receives from the extension. */
export interface ToolContext {
	cwd: string;
	connection?: TeamsConnection;
}

/**
 * Resolve the connection for this call.
 *
 * The session-level connection is reused when no override is given, so the
 * common case costs nothing; naming an account or tenant re-resolves against
 * the config file, which also picks up edits made during the session.
 */
export function connectionFor(
	ctx: ToolContext,
	account?: string,
	tenant?: string,
): TeamsConnection {
	if (!account && !tenant && ctx.connection) return ctx.connection;
	return resolveConnection(account, tenant);
}

// ---------------------------------------------------------------------------
// Consent gates
// ---------------------------------------------------------------------------

/**
 * Whether this account asks for a delegated scope.
 *
 * A delegated scope cannot be granted unless it was requested, so the
 * configuration is a reliable gate: if it is missing here, the Graph call
 * cannot succeed, and its failure would be a consent error that says nothing
 * about the one-line fix.
 */
export function hasScope(conn: TeamsConnection, scope: string): boolean {
	return conn.scopes.some((entry) => entry.toLowerCase() === scope.toLowerCase());
}

/** The message to return when a tool needs a scope the account does not request. */
export function missingScopeError(scope: string, feature: string): string {
	return (
		`${feature} needs the Microsoft Graph scope "${scope}", which this account does not request. ` +
		`Add it to "scopes" in pi-teams.json (or have an admin grant it in Entra ID) and run teams_login again.`
	);
}

// ---------------------------------------------------------------------------
// Identity cache
// ---------------------------------------------------------------------------

const meCache = new Map<string, SignedInUser>();

/**
 * The signed-in user for a connection, cached for the session.
 *
 * Several tools need the user's own ID — to label chats from their point of
 * view, to check "is this my message" — and none of them should pay for a
 * `/me` round-trip each time.
 */
export async function currentUser(
	conn: TeamsConnection,
	signal?: AbortSignal,
): Promise<SignedInUser> {
	const key = `${conn.account}::${conn.tenantId}`;
	const cached = meCache.get(key);
	if (cached) return cached;
	const me = await getMe(conn, signal);
	meCache.set(key, me);
	return me;
}

/** Forget cached identities (after a logout or account switch). */
export function clearUserCache(): void {
	meCache.clear();
}

/** Short label for a connection, e.g. "work" or "work/customer-alpha". */
export function connectionLabel(conn: TeamsConnection): string {
	return conn.tenant === conn.account ? conn.account : `${conn.account}/${conn.tenant}`;
}
