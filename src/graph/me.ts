/**
 * The signed-in user and people lookup.
 */

import type { TeamsConnection } from "../config/index.ts";
import type { PersonRef, SignedInUser } from "../types.ts";
import { graphList, graphRequest } from "./client.ts";
import { mapPerson, mapUser } from "./mappers.ts";

const USER_SELECT = "id,displayName,userPrincipalName,mail,jobTitle,department,officeLocation";

/** Who pi is currently acting as. */
export async function getMe(conn: TeamsConnection, signal?: AbortSignal): Promise<SignedInUser> {
	const raw = await graphRequest<Record<string, unknown>>(conn, "GET", "/me", {
		query: { $select: USER_SELECT },
		signal,
	});
	const user = mapUser(raw as Record<string, never> | undefined);
	if (!user) throw new Error("Could not read the signed-in user from Microsoft Graph.");
	return { ...user, tenantId: conn.tenantId };
}

/**
 * Find people by name, UPN or e-mail.
 *
 * Graph's `$filter startsWith` is the only broadly available option here:
 * `$search` on /users needs ConsistencyLevel=eventual and is not granted in
 * every tenant, so the query is escaped and matched against the three fields
 * people actually type.
 */
export async function findUsers(
	conn: TeamsConnection,
	query: string,
	max = 15,
	signal?: AbortSignal,
): Promise<PersonRef[]> {
	const escaped = query.replace(/'/g, "''");
	const filter =
		`startswith(displayName,'${escaped}') or ` +
		`startswith(userPrincipalName,'${escaped}') or ` +
		`startswith(mail,'${escaped}') or ` +
		`startswith(givenName,'${escaped}') or ` +
		`startswith(surname,'${escaped}')`;

	const raw = await graphList<Record<string, unknown>>(conn, "/users", {
		query: { $filter: filter, $select: USER_SELECT, $top: Math.min(max, 50) },
		max,
		signal,
	});

	return raw
		.map((entry) => mapPerson(entry as Record<string, never>))
		.filter((p): p is PersonRef => !!p);
}

/** Resolve a person reference (UPN, e-mail, display name or id) to a user id. */
export async function resolveUserId(
	conn: TeamsConnection,
	reference: string,
	signal?: AbortSignal,
): Promise<PersonRef | undefined> {
	// A UPN or object id can be fetched directly — one call instead of a search.
	if (reference.includes("@") || /^[0-9a-f-]{36}$/i.test(reference)) {
		try {
			const raw = await graphRequest<Record<string, unknown>>(
				conn,
				"GET",
				`/users/${encodeURIComponent(reference)}`,
				{ query: { $select: USER_SELECT }, signal },
			);
			const person = mapPerson(raw as Record<string, never> | undefined);
			if (person) return person;
		} catch {
			/* fall through to search */
		}
	}

	const matches = await findUsers(conn, reference, 5, signal);
	return matches[0];
}
