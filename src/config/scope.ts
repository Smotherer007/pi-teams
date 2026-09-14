/**
 * Scope rules — which teams, channels, chats and people pi may touch.
 *
 * Pure functions only. No I/O, no Graph calls — everything here is testable
 * with plain objects, which matters because these rules are the last line of
 * defence before pi speaks as the user.
 *
 * Semantics
 * ---------
 * - Rules exist per category (teams / channels / chats / people) and per mode
 *   (read / write).
 * - `deny` accumulates across all levels (global → account → tenant) and always
 *   wins. A deny anywhere blocks the target.
 * - `allow` narrows: the most specific level that defines a non-empty allow
 *   list wins. If no level defines one, everything the signed-in user can
 *   reach is allowed (`*`).
 * - A target is matched against several candidate strings (id, display name,
 *   "Team/Channel" path, e-mail, UPN). A match on any candidate counts.
 * - Patterns are case-insensitive globs: `*` matches any run of characters
 *   (including `/`), `?` matches exactly one. An `id:` prefix is stripped so
 *   `id:19:abc@thread.tacv2` reads naturally in a config file.
 */

// ---------------------------------------------------------------------------
// Types
// ---------------------------------------------------------------------------

/** Categories a rule can apply to. */
export type ScopeCategory = "teams" | "channels" | "chats" | "people";

/** Read or write access. */
export type ScopeMode = "read" | "write";

/** Allow/deny lists for a single category. */
export interface ScopeList {
	allow?: string[];
	deny?: string[];
}

/** Per-category rules. */
export interface ScopeRules {
	teams?: ScopeList;
	channels?: ScopeList;
	chats?: ScopeList;
	people?: ScopeList;
}

/**
 * A permission block as written in pi-teams.json.
 *
 * Category keys at the top level apply to both modes; `read` and `write`
 * blocks refine one mode. A `write` allow list overrides the shorthand allow
 * list at the same level; deny lists from both are unioned.
 */
export interface PermissionBlock extends ScopeRules {
	read?: ScopeRules;
	write?: ScopeRules;
}

/** Fully resolved rules for every category and mode. */
export type ResolvedPermissions = {
	[M in ScopeMode]: { [C in ScopeCategory]: Required<ScopeList> };
};

/** The outcome of a scope check. */
export interface ScopeDecision {
	allowed: boolean;
	/** Which rule produced the decision, for error messages and `/teams-permissions`. */
	reason: string;
	/** The pattern that matched, when one did. */
	pattern?: string;
}

export const SCOPE_CATEGORIES: ScopeCategory[] = ["teams", "channels", "chats", "people"];
export const SCOPE_MODES: ScopeMode[] = ["read", "write"];

// ---------------------------------------------------------------------------
// Glob matching
// ---------------------------------------------------------------------------

/** Escape regex metacharacters except `*` and `?`, which become wildcards. */
function globToRegExp(pattern: string): RegExp {
	const escaped = pattern
		.replace(/[.+^${}()|[\]\\]/g, "\\$&")
		.replace(/\*/g, "[\\s\\S]*")
		.replace(/\?/g, "[\\s\\S]");
	return new RegExp(`^${escaped}$`, "i");
}

/** Normalize a pattern or candidate: trim, drop an `id:` prefix, lowercase. */
function normalize(value: string): string {
	return value.trim().replace(/^id:/i, "").toLowerCase();
}

/**
 * Match a single pattern against a single candidate string.
 * Exported for tests and for `teams_permissions` explanations.
 */
export function matchesPattern(pattern: string, candidate: string): boolean {
	const p = normalize(pattern);
	const c = normalize(candidate);
	if (!p || !c) return false;
	if (p === "*") return true;
	if (p === c) return true;
	return globToRegExp(p).test(c);
}

/** Find the first pattern in `patterns` that matches any of `candidates`. */
export function findMatch(
	patterns: readonly string[],
	candidates: readonly string[],
): string | undefined {
	for (const pattern of patterns) {
		for (const candidate of candidates) {
			if (matchesPattern(pattern, candidate)) return pattern;
		}
	}
	return undefined;
}

// ---------------------------------------------------------------------------
// Resolution
// ---------------------------------------------------------------------------

function emptyList(): Required<ScopeList> {
	return { allow: [], deny: [] };
}

/** Rules that allow everything — the default when nothing is configured. */
export function defaultPermissions(): ResolvedPermissions {
	const forMode = () => ({
		teams: { allow: ["*"], deny: [] as string[] },
		channels: { allow: ["*"], deny: [] as string[] },
		chats: { allow: ["*"], deny: [] as string[] },
		people: { allow: ["*"], deny: [] as string[] },
	});
	return { read: forMode(), write: forMode() };
}

/**
 * Collapse a list of permission blocks (least specific first) into resolved
 * rules. Later blocks are more specific and win on `allow`; `deny` is unioned.
 */
export function resolvePermissions(
	blocks: readonly (PermissionBlock | undefined)[],
): ResolvedPermissions {
	const resolved: ResolvedPermissions = {
		read: { teams: emptyList(), channels: emptyList(), chats: emptyList(), people: emptyList() },
		write: { teams: emptyList(), channels: emptyList(), chats: emptyList(), people: emptyList() },
	};

	for (const block of blocks) {
		if (!block) continue;

		for (const mode of SCOPE_MODES) {
			const modeRules = block[mode];

			for (const category of SCOPE_CATEGORIES) {
				const shorthand = block[category];
				const specific = modeRules?.[category];
				const target = resolved[mode][category];

				// deny accumulates from every level and both shapes
				for (const pattern of shorthand?.deny ?? []) {
					if (!target.deny.includes(pattern)) target.deny.push(pattern);
				}
				for (const pattern of specific?.deny ?? []) {
					if (!target.deny.includes(pattern)) target.deny.push(pattern);
				}

				// allow: the most specific non-empty list wins
				if (specific?.allow && specific.allow.length > 0) {
					target.allow = [...specific.allow];
				} else if (shorthand?.allow && shorthand.allow.length > 0) {
					target.allow = [...shorthand.allow];
				}
			}
		}
	}

	// Nothing configured for a slot means "everything the user can reach"
	for (const mode of SCOPE_MODES) {
		for (const category of SCOPE_CATEGORIES) {
			if (resolved[mode][category].allow.length === 0) {
				resolved[mode][category].allow = ["*"];
			}
		}
	}

	return resolved;
}

// ---------------------------------------------------------------------------
// Checking
// ---------------------------------------------------------------------------

/**
 * Decide whether a target may be read or written.
 *
 * @param permissions - resolved rules
 * @param mode - read or write
 * @param category - which rule set applies
 * @param candidates - identifiers for the target (id, name, path, e-mail…);
 *                     empty or blank entries are ignored
 */
export function checkScope(
	permissions: ResolvedPermissions,
	mode: ScopeMode,
	category: ScopeCategory,
	candidates: readonly (string | undefined)[],
): ScopeDecision {
	const list = permissions[mode][category];
	const values = candidates.filter((c): c is string => !!c && c.trim().length > 0);

	if (values.length === 0) {
		return { allowed: false, reason: `no identifier available to check against the ${category} rules` };
	}

	const denied = findMatch(list.deny, values);
	if (denied) {
		return {
			allowed: false,
			reason: `blocked by ${mode}.${category}.deny rule "${denied}"`,
			pattern: denied,
		};
	}

	const allowed = findMatch(list.allow, values);
	if (allowed) {
		return {
			allowed: true,
			reason: allowed === "*"
				? `allowed by ${mode}.${category} default (*)`
				: `allowed by ${mode}.${category}.allow rule "${allowed}"`,
			pattern: allowed,
		};
	}

	return {
		allowed: false,
		reason:
			`not covered by any ${mode}.${category}.allow rule ` +
			`(configured: ${list.allow.join(", ") || "none"})`,
	};
}

/** Convenience: keep only the entries a check lets through. */
export function filterAllowed<T>(
	permissions: ResolvedPermissions,
	mode: ScopeMode,
	category: ScopeCategory,
	items: readonly T[],
	candidatesOf: (item: T) => (string | undefined)[],
): T[] {
	return items.filter((item) => checkScope(permissions, mode, category, candidatesOf(item)).allowed);
}

/** Render resolved rules as markdown — used by `teams_permissions` and the doctor. */
export function formatPermissions(permissions: ResolvedPermissions): string {
	const lines: string[] = [];
	for (const mode of SCOPE_MODES) {
		lines.push(`**${mode === "read" ? "Read" : "Write"}**`);
		for (const category of SCOPE_CATEGORIES) {
			const list = permissions[mode][category];
			const allow = list.allow.join(", ") || "none";
			const deny = list.deny.length > 0 ? ` · deny: ${list.deny.join(", ")}` : "";
			lines.push(`- ${category}: allow: ${allow}${deny}`);
		}
		lines.push("");
	}
	return lines.join("\n").trimEnd();
}
