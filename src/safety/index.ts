/**
 * The safety model — three independent gates, all of which must pass before
 * pi says or changes anything as the user.
 *
 *  1. **Auth mode.** An app-only (client-credentials) token cannot post as a
 *     person; Microsoft Graph only allows that for migration. Rather than
 *     letting the call fail at Graph, writes are refused here with an
 *     explanation.
 *  2. **Safety level.** `readonly` blocks every mutation, `confirm` asks the
 *     user first, `open` lets them through. Configurable globally, per account
 *     and per tenant, most specific winning.
 *  3. **Scope rules.** Allow/deny lists decide *which* teams, channels, chats
 *     and people are in play at all — for reading as well as writing. See
 *     ../config/scope.ts.
 *
 * Gates 1 and 2 live in the extension's `tool_call` interceptor because they
 * need only the tool name. Gate 3 lives in the tools, because only they know
 * which channel a message is actually headed for.
 */

import type { ResolvedAuthMode, SafetyLevel, TeamsConnection } from "../config/index.ts";
import { checkScope, type ScopeCategory, type ScopeMode } from "../config/scope.ts";
import { ScopeDeniedError } from "../utils/errors.ts";
import { MUTATION_TOOLS } from "../tools/tool-names.ts";

// ---------------------------------------------------------------------------
// Gate 1 + 2 — tool-level
// ---------------------------------------------------------------------------

export function isMutationTool(toolName: string): boolean {
	return MUTATION_TOOLS.has(toolName);
}

/**
 * Tools that only change local files — the config and the token cache.
 *
 * They still ask for confirmation, but `readonly` must not block them: that
 * level means "do not change anything in Teams", and a user who set it still
 * needs a way to add an account or sign out. They also work before anything is
 * configured at all, which is the whole point of `teams_setup`.
 */
export const LOCAL_CONFIG_TOOLS = new Set<string>(["teams_setup", "teams_logout"]);

/**
 * Why a mutation must be refused outright, or undefined when it may proceed
 * (possibly after confirmation).
 */
export function blockReason(
	safetyLevel: SafetyLevel,
	authMode: ResolvedAuthMode,
	toolName: string,
): string | undefined {
	if (!isMutationTool(toolName)) return undefined;
	if (LOCAL_CONFIG_TOOLS.has(toolName)) return undefined;

	if (safetyLevel === "readonly") {
		return (
			`"${toolName}" is blocked: safetyLevel is "readonly". ` +
			`Set safetyLevel to "confirm" or "open" in pi-teams.json (globally, on the account, or on the tenant).`
		);
	}

	if (authMode === "client-credentials" && APP_ONLY_FORBIDDEN.has(toolName)) {
		return (
			`"${toolName}" cannot run on an app-only token. Microsoft Graph only permits app-only writes to ` +
			`chats and channels for migration scenarios, so pi cannot post as a person this way. ` +
			`Set authMode to "interactive" for this account and run teams_login.`
		);
	}

	return undefined;
}

/**
 * Tools that need a real user behind them.
 * Directory reads and calendar work still function app-only.
 */
const APP_ONLY_FORBIDDEN = new Set<string>([
	"teams_send_chat_message",
	"teams_send_channel_message",
	"teams_reply_channel_message",
	"teams_create_chat",
	"teams_delete_message",
	"teams_react",
	"teams_set_presence",
	"teams_set_status_message",
]);

// ---------------------------------------------------------------------------
// Confirmation summaries
// ---------------------------------------------------------------------------

/**
 * One line describing what is about to happen, shown in the confirm dialog.
 *
 * This is the last thing a user reads before pi speaks for them, so it shows
 * the destination and the actual text — never just the tool name.
 */
export function formatMutationSummary(
	toolName: string,
	params: Record<string, unknown>,
): string {
	const text = (value: unknown, max = 160): string => {
		const str = String(value ?? "").replace(/\s*\n\s*/g, " ").trim();
		return str.length > max ? `${str.slice(0, max - 1)}…` : str;
	};

	switch (toolName) {
		case "teams_send_chat_message":
			return `Send a chat message to "${text(params.chat, 60)}":\n\n${text(params.body, 400)}`;
		case "teams_send_channel_message": {
			const where = params.team ? `${text(params.team, 40)}/${text(params.channel, 40)}` : text(params.channel, 60);
			return `Post in channel ${where}:\n\n${text(params.body, 400)}`;
		}
		case "teams_reply_channel_message": {
			const where = params.team ? `${text(params.team, 40)}/${text(params.channel, 40)}` : text(params.channel, 60);
			return `Reply in thread ${text(params.messageId, 40)} (${where}):\n\n${text(params.body, 400)}`;
		}
		case "teams_create_chat": {
			const people = Array.isArray(params.participants) ? params.participants.join(", ") : text(params.participants);
			return `Create a new chat with: ${text(people, 200)}`;
		}
		case "teams_create_channel":
			return `Create channel "${text(params.name, 60)}" in team "${text(params.team, 60)}"`;
		case "teams_delete_message":
			return `Delete your message ${text(params.messageId, 60)}`;
		case "teams_react":
			return `${params.remove ? "Remove" : "Add"} reaction ${text(params.reaction, 10)} on message ${text(params.messageId, 50)}`;
		case "teams_set_presence":
			return `Set your Teams presence to ${text(params.availability, 30)}${params.expiresIn ? ` for ${text(params.expiresIn, 20)}` : ""}`;
		case "teams_set_status_message":
			return `Set your Teams status message to: "${text(params.message, 200)}"`;
		case "teams_create_meeting":
			return `Create meeting "${text(params.subject, 60)}" (${text(params.start, 30)} – ${text(params.end, 30)}) with ${text(Array.isArray(params.attendees) ? params.attendees.join(", ") : params.attendees, 160)}`;
		case "teams_update_meeting":
			return `Update meeting ${text(params.eventId, 50)}`;
		case "teams_cancel_meeting":
			return `Cancel meeting ${text(params.eventId, 50)}`;
		case "teams_setup":
			return `Write Teams account "${text(params.name, 40)}" to pi-teams.json`;
		case "teams_logout":
			return `Sign out of Teams account "${text(params.account, 40) || "(default)"}"`;
		default:
			return `${toolName}: ${text(JSON.stringify(params), 200)}`;
	}
}

// ---------------------------------------------------------------------------
// Gate 3 — scope enforcement
// ---------------------------------------------------------------------------

/**
 * Assert that a target is in scope, or throw.
 *
 * @param conn - the resolved connection carrying the rules
 * @param mode - read or write
 * @param category - which rule set applies
 * @param label - how the target is named in the error message
 * @param candidates - identifiers to match (id, name, path, e-mail, …)
 *
 * @throws {ScopeDeniedError}
 */
export function assertAccess(
	conn: TeamsConnection,
	mode: ScopeMode,
	category: ScopeCategory,
	label: string,
	candidates: (string | undefined)[],
): void {
	const decision = checkScope(conn.permissions, mode, category, candidates);
	if (!decision.allowed) {
		throw new ScopeDeniedError(`${mode} ${category.replace(/s$/, "")} "${label}"`, decision.reason);
	}
}

/** Non-throwing variant, for filtering lists. */
export function hasAccess(
	conn: TeamsConnection,
	mode: ScopeMode,
	category: ScopeCategory,
	candidates: (string | undefined)[],
): boolean {
	return checkScope(conn.permissions, mode, category, candidates).allowed;
}

export { ScopeDeniedError };
