/**
 * Calendar and Teams meetings.
 *
 * Creating a meeting goes through a calendar event with `isOnlineMeeting:
 * true` rather than `/me/onlineMeetings`: that way the meeting lands in
 * everyone's calendar with invitations, which is what a person means when they
 * say "set up a Teams meeting". `/me/onlineMeetings` is still exposed for the
 * ad-hoc "just give me a join link" case.
 */

import type { TeamsConnection } from "../config/index.ts";
import type { EventSummary, OnlineMeetingSummary } from "../types.ts";
import { graphDelete, graphGetOptional, graphList, graphPatch, graphPost } from "./client.ts";
import { mapEvent } from "./mappers.ts";

/* eslint-disable @typescript-eslint/no-explicit-any */
type Raw = Record<string, any>;

export interface EventInput {
	subject: string;
	/** ISO 8601 local date-time, e.g. "2026-09-15T10:00:00" */
	start: string;
	end: string;
	/** IANA or Windows time zone name; defaults to the mailbox setting */
	timeZone?: string;
	/** SMTP addresses */
	attendees?: string[];
	/** Optional attendees */
	optionalAttendees?: string[];
	body?: string;
	location?: string;
	/** Create a Teams meeting link (default true) */
	onlineMeeting?: boolean;
}

/** Events in a time range, ordered by start. */
export async function listEvents(
	conn: TeamsConnection,
	options: { start?: string; end?: string; max?: number; signal?: AbortSignal } = {},
): Promise<EventSummary[]> {
	const start = options.start ?? new Date().toISOString();
	const end = options.end ?? new Date(Date.now() + 7 * 24 * 3600 * 1000).toISOString();

	// calendarView expands recurring series into real occurrences; /me/events
	// would return the series master and miss today's instance.
	const raw = await graphList<Raw>(conn, "/me/calendarView", {
		query: {
			startDateTime: start,
			endDateTime: end,
			$orderby: "start/dateTime",
			$top: Math.min(options.max ?? 25, 50),
			$select:
				"id,subject,start,end,isAllDay,isOnlineMeeting,onlineMeeting,organizer,attendees,location,bodyPreview,webLink",
		},
		headers: { Prefer: 'outlook.timezone="UTC"' },
		max: options.max ?? 25,
		signal: options.signal,
	});

	return raw.map((entry) => mapEvent(entry));
}

export async function getEvent(
	conn: TeamsConnection,
	eventId: string,
	signal?: AbortSignal,
): Promise<EventSummary | undefined> {
	const raw = await graphGetOptional<Raw>(conn, `/me/events/${encodeURIComponent(eventId)}`, {
		signal,
	});
	return raw ? mapEvent(raw) : undefined;
}

function attendeeEntry(address: string, type: "required" | "optional") {
	return { emailAddress: { address }, type };
}

export async function createEvent(
	conn: TeamsConnection,
	input: EventInput,
	signal?: AbortSignal,
): Promise<EventSummary> {
	const timeZone = input.timeZone ?? "UTC";

	const body: Record<string, unknown> = {
		subject: input.subject,
		start: { dateTime: input.start, timeZone },
		end: { dateTime: input.end, timeZone },
		isOnlineMeeting: input.onlineMeeting !== false,
		onlineMeetingProvider: "teamsForBusiness",
		attendees: [
			...(input.attendees ?? []).map((a) => attendeeEntry(a, "required")),
			...(input.optionalAttendees ?? []).map((a) => attendeeEntry(a, "optional")),
		],
	};
	if (input.body) body.body = { contentType: "html", content: input.body };
	if (input.location) body.location = { displayName: input.location };

	const raw = await graphPost<Raw>(conn, "/me/events", body, { signal });
	return mapEvent(raw ?? {});
}

export async function updateEvent(
	conn: TeamsConnection,
	eventId: string,
	changes: Partial<EventInput>,
	signal?: AbortSignal,
): Promise<EventSummary> {
	const timeZone = changes.timeZone ?? "UTC";
	const body: Record<string, unknown> = {};

	if (changes.subject) body.subject = changes.subject;
	if (changes.start) body.start = { dateTime: changes.start, timeZone };
	if (changes.end) body.end = { dateTime: changes.end, timeZone };
	if (changes.body) body.body = { contentType: "html", content: changes.body };
	if (changes.location) body.location = { displayName: changes.location };
	if (changes.attendees || changes.optionalAttendees) {
		body.attendees = [
			...(changes.attendees ?? []).map((a) => attendeeEntry(a, "required")),
			...(changes.optionalAttendees ?? []).map((a) => attendeeEntry(a, "optional")),
		];
	}

	const raw = await graphPatch<Raw>(conn, `/me/events/${encodeURIComponent(eventId)}`, body, {
		signal,
	});
	return mapEvent(raw ?? {});
}

/** Cancel an event you organize (sends cancellations), or decline/delete otherwise. */
export async function cancelEvent(
	conn: TeamsConnection,
	eventId: string,
	comment: string | undefined,
	signal?: AbortSignal,
): Promise<void> {
	try {
		await graphPost(conn, `/me/events/${encodeURIComponent(eventId)}/cancel`, { comment }, { signal });
	} catch {
		// Only the organizer may cancel; for everyone else removing it from the
		// calendar is the closest equivalent.
		await graphDelete(conn, `/me/events/${encodeURIComponent(eventId)}`, { signal });
	}
}

/** Ad-hoc meeting link with no calendar entry. */
export async function createOnlineMeeting(
	conn: TeamsConnection,
	input: { subject?: string; start?: string; end?: string },
	signal?: AbortSignal,
): Promise<OnlineMeetingSummary> {
	const body: Record<string, unknown> = {};
	if (input.subject) body.subject = input.subject;
	if (input.start) body.startDateTime = input.start;
	if (input.end) body.endDateTime = input.end;

	const raw = await graphPost<Raw>(conn, "/me/onlineMeetings", body, { signal });
	return {
		id: raw?.id ?? "",
		subject: raw?.subject ?? input.subject,
		joinUrl: raw?.joinWebUrl ?? raw?.joinUrl ?? "",
		start: raw?.startDateTime,
		end: raw?.endDateTime,
	};
}
