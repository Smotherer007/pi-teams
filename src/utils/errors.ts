/**
 * Microsoft Graph errors → messages a person can act on.
 *
 * Graph is unusually good at returning precise error codes and unusually bad
 * at explaining what to do about them. The mapping below turns the common ones
 * into the next step ("consent to X", "run teams_login", "ask an admin").
 */

/** An error returned by the Graph API, with its code preserved. */
export class GraphError extends Error {
	readonly status: number;
	readonly code: string;
	readonly requestId?: string;
	readonly raw?: unknown;

	constructor(status: number, code: string, message: string, requestId?: string, raw?: unknown) {
		super(message);
		this.name = "GraphError";
		this.status = status;
		this.code = code;
		this.requestId = requestId;
		this.raw = raw;
	}
}

/** Raised when a scope rule blocks an operation. */
export class ScopeDeniedError extends Error {
	readonly target: string;
	readonly reason: string;
	constructor(target: string, reason: string) {
		super(`Blocked by configuration: ${target} — ${reason}.`);
		this.name = "ScopeDeniedError";
		this.target = target;
		this.reason = reason;
	}
}

const CONSENT_HINTS: Record<string, string> = {
	"ChannelMessage.Read.All": "reading channel messages",
	"ChannelMessage.Send": "posting in channels",
	"Chat.ReadWrite": "reading and sending chat messages",
	"ChatMessage.Send": "sending chat messages",
	"ChatMember.ReadWrite": "removing someone from a chat",
	"Presence.ReadWrite": "reading and setting your presence",
	"Calendars.ReadWrite": "reading and creating calendar events",
	"OnlineMeetings.ReadWrite": "creating Teams meetings",
};

/** Normalize any thrown value into a readable, actionable message. */
export function formatGraphError(error: unknown): string {
	if (error instanceof ScopeDeniedError) return error.message;

	if (error instanceof GraphError) {
		switch (error.code) {
			case "InvalidAuthenticationToken":
				return "The access token was rejected. Run teams_login to sign in again.";
			case "Unauthorized":
				return "Authentication failed. Run teams_login to sign in again.";
			case "Forbidden":
			case "AccessDenied":
			case "Authorization_RequestDenied":
				return (
					"Permission denied. The signed-in user (or the app registration) lacks access to this " +
					"resource. Run teams_doctor to see which Graph scopes were actually granted."
				);
			case "NotFound":
			case "itemNotFound":
			case "ResourceNotFound":
				return "Not found. Check the team, channel, chat or message ID.";
			case "ErrorItemNotFound":
				return "Not found — the item may have been deleted or you no longer have access.";
			case "activityLimitReached":
			case "TooManyRequests":
				return "Throttled by Microsoft Graph. Wait a moment and try again.";
			case "BadRequest":
				return `Graph rejected the request: ${error.message}`;
			case "UnknownError":
				return `Graph returned an unspecified error: ${error.message}`;
			default:
				return `${error.code}: ${error.message}`;
		}
	}

	if (error instanceof Error) {
		const msg = error.message;
		if (msg.includes("ENOTFOUND")) return "DNS error: could not resolve the Graph endpoint.";
		if (msg.includes("ECONNREFUSED") || msg.includes("ETIMEDOUT")) {
			return "Network error: could not reach Microsoft Graph. Check your connection or proxy.";
		}
		if (msg.includes("AADSTS7000218")) {
			return (
				"The app registration rejected the device code flow. Enable 'Allow public client flows' " +
				"in Entra ID → App registrations → Authentication."
			);
		}
		if (msg.includes("AADSTS65001")) {
			return "Consent is missing for one or more scopes. Sign in again, or ask an admin to grant consent.";
		}
		if (msg.includes("AADSTS50076") || msg.includes("AADSTS50079")) {
			return "Multi-factor authentication is required. Complete the sign-in in the browser and retry.";
		}
		if (msg.includes("AADSTS50020")) {
			return (
				"The signed-in user does not exist in that tenant. For a guest/customer tenant, make sure the " +
				"account is invited there and the app registration is multi-tenant."
			);
		}
		return msg.split("\n")[0] ?? "Unknown error";
	}

	return String(error);
}

export function isAuthError(error: unknown): boolean {
	if (error instanceof GraphError) {
		return error.status === 401 || error.code === "InvalidAuthenticationToken";
	}
	return false;
}

export function isNotFoundError(error: unknown): boolean {
	if (error instanceof GraphError) return error.status === 404;
	return false;
}

export function isThrottled(error: unknown): boolean {
	if (error instanceof GraphError) return error.status === 429 || error.status === 503;
	return false;
}

/** Human-readable purpose of a Graph scope, for doctor output. */
export function describeScope(scope: string): string | undefined {
	return CONSENT_HINTS[scope];
}
