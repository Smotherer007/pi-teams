/**
 * Calendar and meeting tools.
 *
 * "Create a Teams meeting" means a calendar event with a join link and real
 * invitations, so that is what `teams_create_meeting` does by default; passing
 * no attendees and `calendarEvent: false` gives the bare join link instead.
 */

import { Type } from "typebox";
import {
	cancelEvent,
	createEvent,
	createOnlineMeeting,
	getEvent,
	listEvents,
	updateEvent,
} from "../graph/calendar.ts";
import { auditWrite } from "../safety/audit.ts";
import { assertAccess } from "../safety/index.ts";
import { formatEventDetail, formatEventList } from "../utils/formatting.ts";
import {
	AccountParam,
	LimitParam,
	TenantParam,
	connectionFor,
	currentUser,
	errorResult,
	run,
	textResult,
	type ToolContext,
	type ToolResult,
} from "./shared.ts";

export const teamsListMeetingsTool = {
	name: "teams_list_meetings",
	description:
		"List the signed-in user's calendar events and Teams meetings in a time range (default: the next 7 days). " +
		"Recurring series are expanded into actual occurrences. Use this for 'what's on my calendar', " +
		"'when is my next meeting', or to find an event ID.",
	parameters: Type.Object({
		start: Type.Optional(
			Type.String({ description: "ISO 8601 start of the range; defaults to now" }),
		),
		end: Type.Optional(
			Type.String({ description: "ISO 8601 end of the range; defaults to 7 days from now" }),
		),
		account: AccountParam,
		tenant: TenantParam,
		limit: LimitParam,
	}),
	promptSnippet: "List the user's Teams meetings and calendar events",

	async execute(
		_toolCallId: string,
		params: { start?: string; end?: string; account?: string; tenant?: string; limit?: number },
		signal: AbortSignal | undefined,
		_onUpdate: undefined,
		ctx: ToolContext,
	): Promise<ToolResult> {
		return run(async () => {
			const conn = connectionFor(ctx, params.account, params.tenant);
			const events = await listEvents(conn, {
				start: params.start,
				end: params.end,
				max: params.limit ?? 25,
				signal,
			});
			return textResult(formatEventList(events), { count: events.length });
		});
	},
};

export const teamsGetMeetingTool = {
	name: "teams_get_meeting",
	description:
		"Show one calendar event or Teams meeting in full: time, organizer, attendees, join link and notes. " +
		"Get the eventId from teams_list_meetings.",
	parameters: Type.Object({
		eventId: Type.String({ description: "Calendar event ID" }),
		account: AccountParam,
		tenant: TenantParam,
	}),
	promptSnippet: "Show details of a Teams meeting",

	async execute(
		_toolCallId: string,
		params: { eventId: string; account?: string; tenant?: string },
		signal: AbortSignal | undefined,
		_onUpdate: undefined,
		ctx: ToolContext,
	): Promise<ToolResult> {
		return run(async () => {
			const conn = connectionFor(ctx, params.account, params.tenant);
			const event = await getEvent(conn, params.eventId, signal);
			if (!event) return errorResult(`Event ${params.eventId} not found.`);
			return textResult(formatEventDetail(event), { eventId: event.id });
		});
	},
};

export const teamsCreateMeetingTool = {
	name: "teams_create_meeting",
	description:
		"Schedule a Microsoft Teams meeting as the signed-in user: creates a calendar event with a join link " +
		"and sends invitations to the attendees. Times are ISO 8601 local date-times " +
		"(e.g. '2026-09-15T10:00:00') interpreted in 'timeZone'. " +
		"Set calendarEvent: false for a bare join link with no calendar entry or invitations.",
	parameters: Type.Object({
		subject: Type.String({ description: "Meeting title" }),
		start: Type.String({ description: "Start, ISO 8601 local date-time, e.g. '2026-09-15T10:00:00'" }),
		end: Type.String({ description: "End, ISO 8601 local date-time" }),
		attendees: Type.Optional(
			Type.Array(Type.String(), { description: "Required attendees, by e-mail address" }),
		),
		optionalAttendees: Type.Optional(
			Type.Array(Type.String(), { description: "Optional attendees, by e-mail address" }),
		),
		timeZone: Type.Optional(
			Type.String({ description: "IANA or Windows time zone, e.g. 'Europe/Berlin'. Default UTC." }),
		),
		body: Type.Optional(Type.String({ description: "Agenda or invitation text" })),
		location: Type.Optional(Type.String({ description: "Physical location, if any" })),
		calendarEvent: Type.Optional(
			Type.Boolean({ description: "Create a calendar event (default true). False = join link only." }),
		),
		account: AccountParam,
		tenant: TenantParam,
	}),
	promptSnippet: "Schedule a Teams meeting",
	promptGuidelines: [
		"Confirm date, time, time zone and attendee list with the user before creating a meeting — invitations cannot be un-sent.",
		"Pass attendees as e-mail addresses; use teams_find_user to resolve names first.",
	],

	async execute(
		_toolCallId: string,
		params: {
			subject: string;
			start: string;
			end: string;
			attendees?: string[];
			optionalAttendees?: string[];
			timeZone?: string;
			body?: string;
			location?: string;
			calendarEvent?: boolean;
			account?: string;
			tenant?: string;
		},
		signal: AbortSignal | undefined,
		_onUpdate: undefined,
		ctx: ToolContext,
	): Promise<ToolResult> {
		return run(async () => {
			const conn = connectionFor(ctx, params.account, params.tenant);
			const me = await currentUser(conn, signal).catch(() => undefined);

			// Inviting someone is contacting them, so the people rules apply.
			for (const address of [...(params.attendees ?? []), ...(params.optionalAttendees ?? [])]) {
				assertAccess(conn, "write", "people", address, [address]);
			}

			if (params.calendarEvent === false) {
				const meeting = await createOnlineMeeting(
					conn,
					{ subject: params.subject, start: params.start, end: params.end },
					signal,
				);
				auditWrite(conn.audit, {
					tool: "teams_create_meeting",
					account: conn.account,
					tenant: conn.tenant,
					actor: me?.upn,
					target: "calendar",
					summary: `created ad-hoc meeting "${params.subject}"`,
				});
				return textResult(
					[`✅ Meeting link created for "${params.subject}".`, "", `Join: ${meeting.joinUrl}`].join("\n"),
					{ meetingId: meeting.id, joinUrl: meeting.joinUrl },
				);
			}

			const event = await createEvent(
				conn,
				{
					subject: params.subject,
					start: params.start,
					end: params.end,
					timeZone: params.timeZone,
					attendees: params.attendees,
					optionalAttendees: params.optionalAttendees,
					body: params.body,
					location: params.location,
					onlineMeeting: true,
				},
				signal,
			);

			auditWrite(conn.audit, {
				tool: "teams_create_meeting",
				account: conn.account,
				tenant: conn.tenant,
				actor: me?.upn,
				target: "calendar",
				summary:
					`created "${params.subject}" ${params.start}–${params.end} ` +
					`with ${(params.attendees ?? []).join(", ") || "no attendees"}`,
			});

			return textResult(
				[
					`✅ Meeting **${event.subject}** scheduled as ${me?.displayName ?? "you"}.`,
					"",
					formatEventDetail(event),
				].join("\n"),
				{ eventId: event.id, joinUrl: event.joinUrl },
			);
		});
	},
};

export const teamsUpdateMeetingTool = {
	name: "teams_update_meeting",
	description:
		"Change an existing calendar event the signed-in user organizes: time, subject, attendees, location or " +
		"agenda. Attendees receive an update. Only the fields you pass are changed.",
	parameters: Type.Object({
		eventId: Type.String({ description: "Calendar event ID" }),
		subject: Type.Optional(Type.String({ description: "New title" })),
		start: Type.Optional(Type.String({ description: "New start, ISO 8601 local date-time" })),
		end: Type.Optional(Type.String({ description: "New end, ISO 8601 local date-time" })),
		timeZone: Type.Optional(Type.String({ description: "Time zone for the new times" })),
		attendees: Type.Optional(
			Type.Array(Type.String(), { description: "Replacement list of required attendees (e-mail)" }),
		),
		optionalAttendees: Type.Optional(
			Type.Array(Type.String(), { description: "Replacement list of optional attendees (e-mail)" }),
		),
		body: Type.Optional(Type.String({ description: "New agenda text" })),
		location: Type.Optional(Type.String({ description: "New location" })),
		account: AccountParam,
		tenant: TenantParam,
	}),
	promptSnippet: "Update a Teams meeting",

	async execute(
		_toolCallId: string,
		params: {
			eventId: string;
			subject?: string;
			start?: string;
			end?: string;
			timeZone?: string;
			attendees?: string[];
			optionalAttendees?: string[];
			body?: string;
			location?: string;
			account?: string;
			tenant?: string;
		},
		signal: AbortSignal | undefined,
		_onUpdate: undefined,
		ctx: ToolContext,
	): Promise<ToolResult> {
		return run(async () => {
			const conn = connectionFor(ctx, params.account, params.tenant);

			for (const address of [...(params.attendees ?? []), ...(params.optionalAttendees ?? [])]) {
				assertAccess(conn, "write", "people", address, [address]);
			}

			const event = await updateEvent(
				conn,
				params.eventId,
				{
					subject: params.subject,
					start: params.start,
					end: params.end,
					timeZone: params.timeZone,
					attendees: params.attendees,
					optionalAttendees: params.optionalAttendees,
					body: params.body,
					location: params.location,
				},
				signal,
			);

			const me = await currentUser(conn, signal).catch(() => undefined);
			auditWrite(conn.audit, {
				tool: "teams_update_meeting",
				account: conn.account,
				tenant: conn.tenant,
				actor: me?.upn,
				target: "calendar",
				summary: `updated event ${params.eventId}`,
			});

			return textResult([`✅ Meeting updated.`, "", formatEventDetail(event)].join("\n"), {
				eventId: event.id,
			});
		});
	},
};

export const teamsCancelMeetingTool = {
	name: "teams_cancel_meeting",
	description:
		"Cancel a meeting the signed-in user organizes — attendees receive a cancellation notice. " +
		"For a meeting organized by somebody else, the event is removed from the user's own calendar instead.",
	parameters: Type.Object({
		eventId: Type.String({ description: "Calendar event ID" }),
		comment: Type.Optional(Type.String({ description: "Note sent with the cancellation" })),
		account: AccountParam,
		tenant: TenantParam,
	}),
	promptSnippet: "Cancel a Teams meeting",

	async execute(
		_toolCallId: string,
		params: { eventId: string; comment?: string; account?: string; tenant?: string },
		signal: AbortSignal | undefined,
		_onUpdate: undefined,
		ctx: ToolContext,
	): Promise<ToolResult> {
		return run(async () => {
			const conn = connectionFor(ctx, params.account, params.tenant);
			await cancelEvent(conn, params.eventId, params.comment, signal);

			const me = await currentUser(conn, signal).catch(() => undefined);
			auditWrite(conn.audit, {
				tool: "teams_cancel_meeting",
				account: conn.account,
				tenant: conn.tenant,
				actor: me?.upn,
				target: "calendar",
				summary: `cancelled event ${params.eventId}`,
			});

			return textResult(`✅ Meeting ${params.eventId} cancelled.`, { eventId: params.eventId });
		});
	},
};
