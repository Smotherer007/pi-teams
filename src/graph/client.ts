/**
 * Microsoft Graph HTTP client.
 *
 * A thin, dependency-free wrapper around fetch that handles the three things
 * every Graph call needs and every hand-rolled call forgets: a fresh bearer
 * token, `@odata.nextLink` paging, and 429/503 back-off honouring
 * `Retry-After`.
 */

import type { TeamsConnection } from "../config/index.ts";
import { getAccessToken } from "../auth/index.ts";
import { GraphError } from "../utils/errors.ts";

// ---------------------------------------------------------------------------
// Types
// ---------------------------------------------------------------------------

export type HttpMethod = "GET" | "POST" | "PATCH" | "PUT" | "DELETE";

export interface RequestOptions {
	/** Query string parameters — `$select`, `$filter`, `$top`, … */
	query?: Record<string, string | number | boolean | undefined>;
	/** JSON request body */
	body?: unknown;
	signal?: AbortSignal;
	/** Use the beta endpoint instead of v1.0 */
	beta?: boolean;
	/** Extra headers, e.g. Prefer: include-unknown-enum-members */
	headers?: Record<string, string>;
	/** Max retries for throttling responses (default 3) */
	maxRetries?: number;
}

interface GraphErrorBody {
	error?: {
		code?: string;
		message?: string;
		innerError?: { "request-id"?: string; date?: string };
	};
}

interface GraphCollection<T> {
	value: T[];
	"@odata.nextLink"?: string;
	"@odata.count"?: number;
}

// ---------------------------------------------------------------------------
// URL building
// ---------------------------------------------------------------------------

function buildUrl(conn: TeamsConnection, path: string, options: RequestOptions): string {
	const base = options.beta
		? conn.graphBaseUrl.replace(/\/v1\.0$/, "/beta")
		: conn.graphBaseUrl;
	const url = new URL(path.startsWith("http") ? path : `${base}${path.startsWith("/") ? "" : "/"}${path}`);

	for (const [key, value] of Object.entries(options.query ?? {})) {
		if (value === undefined) continue;
		url.searchParams.set(key, String(value));
	}

	return url.toString();
}

// ---------------------------------------------------------------------------
// Core request
// ---------------------------------------------------------------------------

function sleep(ms: number): Promise<void> {
	return new Promise((resolve) => setTimeout(resolve, ms));
}

/**
 * Perform a Graph request and return the parsed body.
 *
 * Returns `undefined` for 204 No Content, which Graph uses for most writes
 * (reactions, deletes, presence).
 *
 * @throws {GraphError} for any non-2xx response that survives the retries
 */
export async function graphRequest<T = unknown>(
	conn: TeamsConnection,
	method: HttpMethod,
	path: string,
	options: RequestOptions = {},
): Promise<T | undefined> {
	const maxRetries = options.maxRetries ?? 3;
	const url = buildUrl(conn, path, options);

	let attempt = 0;
	for (;;) {
		const token = await getAccessToken(conn, options.signal);

		const headers: Record<string, string> = {
			Authorization: `Bearer ${token.accessToken}`,
			Accept: "application/json",
			...options.headers,
		};
		if (options.body !== undefined) headers["Content-Type"] = "application/json";

		const response = await fetch(url, {
			method,
			headers,
			body: options.body === undefined ? undefined : JSON.stringify(options.body),
			signal: options.signal,
		});

		if (response.status === 204 || response.status === 202) {
			return undefined;
		}

		if (response.ok) {
			const text = await response.text();
			if (!text) return undefined;
			return JSON.parse(text) as T;
		}

		// Throttling and transient service errors: back off and retry.
		if ((response.status === 429 || response.status === 503) && attempt < maxRetries) {
			const retryAfter = Number(response.headers.get("Retry-After") ?? "0");
			// Exponential fallback when the header is missing, capped so a tool
			// call cannot hang for minutes.
			const waitMs = retryAfter > 0 ? retryAfter * 1000 : Math.min(2 ** attempt * 1000, 8000);
			attempt += 1;
			await sleep(waitMs);
			continue;
		}

		throw await toGraphError(response);
	}
}

/**
 * Download a Graph resource as bytes.
 *
 * `graphRequest` parses JSON, so the two things that are not JSON — the images
 * embedded in a message and the files attached to it — come through here. Same
 * token handling and throttling back-off; the body is returned as it arrived,
 * because its content type is the only hint at what the file is.
 */
export async function graphDownload(
	conn: TeamsConnection,
	path: string,
	options: RequestOptions = {},
): Promise<{ data: Buffer; contentType?: string }> {
	const maxRetries = options.maxRetries ?? 3;
	// An absolute URL is allowed here too: Graph answers a hostedContent with a
	// redirect to a pre-authenticated location, and `fetch` follows it.
	const url = buildUrl(conn, path, options);

	let attempt = 0;
	for (;;) {
		const token = await getAccessToken(conn, options.signal);

		const response = await fetch(url, {
			method: "GET",
			headers: { Authorization: `Bearer ${token.accessToken}`, ...options.headers },
			signal: options.signal,
		});

		if (response.ok) {
			return {
				data: Buffer.from(await response.arrayBuffer()),
				contentType: response.headers.get("content-type") ?? undefined,
			};
		}

		if ((response.status === 429 || response.status === 503) && attempt < maxRetries) {
			const retryAfter = Number(response.headers.get("Retry-After") ?? "0");
			const waitMs = retryAfter > 0 ? retryAfter * 1000 : Math.min(2 ** attempt * 1000, 8000);
			attempt += 1;
			await sleep(waitMs);
			continue;
		}

		throw await toGraphError(response);
	}
}

async function toGraphError(response: Response): Promise<GraphError> {
	let code = `HTTP${response.status}`;
	let message = response.statusText || `Request failed with status ${response.status}`;
	let requestId: string | undefined;
	let raw: unknown;

	try {
		const text = await response.text();
		if (text) {
			raw = JSON.parse(text) as GraphErrorBody;
			const body = raw as GraphErrorBody;
			if (body.error?.code) code = body.error.code;
			if (body.error?.message) message = body.error.message;
			requestId = body.error?.innerError?.["request-id"];
		}
	} catch {
		/* keep the status-based defaults */
	}

	return new GraphError(response.status, code, message, requestId, raw);
}

// ---------------------------------------------------------------------------
// Collections
// ---------------------------------------------------------------------------

/**
 * Fetch a collection, following `@odata.nextLink` until `max` items are
 * collected. `max` exists because a chat history can be tens of thousands of
 * messages and nobody wants that in a tool result.
 */
export async function graphList<T>(
	conn: TeamsConnection,
	path: string,
	options: RequestOptions & { max?: number } = {},
): Promise<T[]> {
	const max = options.max ?? 100;
	const items: T[] = [];
	let nextPath: string | undefined = path;
	let query: RequestOptions["query"] | undefined = options.query;

	while (nextPath && items.length < max) {
		const page: GraphCollection<T> | undefined = await graphRequest<GraphCollection<T>>(
			conn,
			"GET",
			nextPath,
			{ ...options, query },
		);
		if (!page?.value) break;

		items.push(...page.value);
		nextPath = page["@odata.nextLink"];
		// nextLink already carries the query string.
		query = undefined;
	}

	return items.slice(0, max);
}

/** GET a single resource; `undefined` on 404 instead of throwing. */
export async function graphGetOptional<T>(
	conn: TeamsConnection,
	path: string,
	options: RequestOptions = {},
): Promise<T | undefined> {
	try {
		return await graphRequest<T>(conn, "GET", path, options);
	} catch (err) {
		if (err instanceof GraphError && err.status === 404) return undefined;
		throw err;
	}
}

/** POST a JSON body. */
export function graphPost<T = unknown>(
	conn: TeamsConnection,
	path: string,
	body: unknown,
	options: RequestOptions = {},
): Promise<T | undefined> {
	return graphRequest<T>(conn, "POST", path, { ...options, body });
}

/** PATCH a JSON body. */
export function graphPatch<T = unknown>(
	conn: TeamsConnection,
	path: string,
	body: unknown,
	options: RequestOptions = {},
): Promise<T | undefined> {
	return graphRequest<T>(conn, "PATCH", path, { ...options, body });
}

/** DELETE a resource. */
export function graphDelete(
	conn: TeamsConnection,
	path: string,
	options: RequestOptions = {},
): Promise<unknown> {
	return graphRequest(conn, "DELETE", path, options);
}
