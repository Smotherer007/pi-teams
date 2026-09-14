/**
 * Minimal JWT claim reader.
 *
 * Used only to show *who* pi is signed in as. The token is never validated
 * here — it came straight from the token endpoint over TLS, and Graph is the
 * authority on whether it is accepted. Nothing security-relevant depends on
 * these claims.
 */

export interface IdTokenClaims {
	oid?: string;
	tid?: string;
	name?: string;
	preferred_username?: string;
	upn?: string;
	scp?: string;
	roles?: string[];
	exp?: number;
	[key: string]: unknown;
}

/** Decode the payload of a JWT, or undefined when it is not decodable. */
export function decodeJwtPayload(token: string): IdTokenClaims | undefined {
	try {
		const parts = token.split(".");
		if (parts.length < 2) return undefined;
		const payload = parts[1]!.replace(/-/g, "+").replace(/_/g, "/");
		const padded = payload.padEnd(payload.length + ((4 - (payload.length % 4)) % 4), "=");
		return JSON.parse(Buffer.from(padded, "base64").toString("utf-8")) as IdTokenClaims;
	} catch {
		return undefined;
	}
}

export function decodeIdToken(idToken: string): IdTokenClaims | undefined {
	return decodeJwtPayload(idToken);
}

/**
 * Scopes actually present in an access token (`scp`), or app roles (`roles`)
 * for app-only tokens. `teams_doctor` uses this to tell the user which consent
 * is still missing instead of letting them guess from a 403.
 */
export function tokenScopes(accessToken: string): string[] {
	const claims = decodeJwtPayload(accessToken);
	if (!claims) return [];
	if (typeof claims.scp === "string" && claims.scp.length > 0) return claims.scp.split(" ");
	if (Array.isArray(claims.roles)) return claims.roles;
	return [];
}
