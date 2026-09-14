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

/** RSVP responses Graph accepts on an invitation. */
export type EventResponse = "accept" | "decline" | "tentativelyAccept";

/**
 * Answer a meeting invitation.
 *
 * This is the calendar's equivalent of saying yes out loud: with
 * `sendResponse` the organizer is notified, which is why the tool asks for
 * confirmation rather than doing it quietly.
 */
export async function respondToEvent(
	conn: TeamsConnection,
	eventId: string,
	response: EventResponse,
	options: { comment?: string; sendResponse?: boolean; signal?: AbortSignal } = {},
): Promise<void> {
	await graphPost(
		conn,
		`/me/events/${encodeURIComponent(eventId)}/${response}`,
		{ comment: options.comment, sendResponse: options.sendResponse !== false },
		{ signal: options.signal },
	);
}

// ---------------------------------------------------------------------------
// Free / busy
// ---------------------------------------------------------------------------

/** One person's availability over the requested window. */
export interface FreeBusyWindow {
	/** The address Graph matched the schedule to */
	scheduleId: string;
	/**
	 * One character per slot, starting at the window start:
	 * `0` free, `1` tentative, `2` busy, `3` out of office, `4` working elsewhere.
	 */
	availabilityView: string;
	/** Everything that was not free, with subject omitted on purpose */
	busy: Array<{ status: string; start: string; end: string }>;
}

/**
 * Free/busy for several people in one call.
 *
 * `getSchedule` is what a delegated token may read about another person's
 * calendar — free/busy, not the appointments — and it is the reason
 * availability works without the `.Shared` scopes that `findMeetingTimes`
 * insists on. Subjects are reported by Graph but dropped here: knowing that
 * somebody is busy is the point, knowing why is their business.
 */
export async function getSchedule(
	conn: TeamsConnection,
	input: {
		schedules: string[];
		start: string;
		end: string;
		timeZone: string;
		/** Slot width in minutes; Graph allows 5–1440 */
		intervalMinutes?: number;
		signal?: AbortSignal;
	},
): Promise<FreeBusyWindow[]> {
	const raw = await graphPost<Raw>(conn, "/me/calendar/getSchedule", {
		schedules: input.schedules,
		startTime: { dateTime: input.start, timeZone: input.timeZone },
		endTime: { dateTime: input.end, timeZone: input.timeZone },
		availabilityViewInterval: Math.min(Math.max(input.intervalMinutes ?? 30, 5), 1440),
	}, {
		signal: input.signal,
		// Without this the busy blocks come back in UTC while the view is indexed
		// in the requested zone, and the two would not line up.
		headers: { Prefer: `outlook.timezone="${input.timeZone}"` },
	});

	return (raw?.value ?? []).map((entry: Raw) => ({
		scheduleId: entry.scheduleId ?? "",
		availabilityView: entry.availabilityView ?? "",
		busy: (entry.scheduleItems ?? [])
			.filter((item: Raw) => item.status && item.status !== "free")
			.map((item: Raw) => ({
				status: item.status,
				start: item.start?.dateTime ?? "",
				end: item.end?.dateTime ?? "",
			})),
	}));
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
