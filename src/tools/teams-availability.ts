/**
 * teams_availability — when everybody is free.
 *
 * Reads *free/busy*, not calendars: `getSchedule` is what a delegated token may
 * see about a colleague's day, and it is also the reason this works with the
 * scopes already granted — `findMeetingTimes` demands `Calendars.Read.Shared`,
 * a scope nobody has consented to just to ask "Tuesday at two?".
 *
 * Scheduling is then a second step on purpose. This tool answers the question;
 * `teams_create_meeting` acts on it once the user has agreed to a slot.
 */

import { Type } from "typebox";
import { getSchedule } from "../graph/calendar.ts";
import { resolveUserId } from "../graph/me.ts";
import { assertAccess } from "../safety/index.ts";
import {
	formatNaiveDay,
	formatSlotRange,
	truncate,
} from "../utils/formatting.ts";
import {
	describeBusy,
	findCommonFreeSlots,
	localDayStart,
	localNow,
	localTimeZone,
	type BusyView,
} from "../utils/slots.ts";
import {
	AccountParam,
	LimitParam,
	TenantParam,
	connectionFor,
	currentUser,
	run,
	textResult,
	type ToolContext,
	type ToolResult,
} from "./shared.ts";

/** How far ahead the default search looks. */
const DEFAULT_SEARCH_DAYS = 7;

export const teamsAvailabilityTool = {
	name: "teams_availability",
	description:
		"Find when the signed-in user and other people are all free, using Microsoft Graph free/busy. " +
		"Only availability is read — never the appointments themselves. " +
		"By default it searches the next 7 days for slots on weekdays between 09:00 and 17:00; pass 'start' and " +
		"'end' (ISO 8601 local date-times) to look at an exact window instead. " +
		"Use it before teams_create_meeting when the goal is to find a time rather than to book one.",
	parameters: Type.Object({
		people: Type.Optional(
			Type.Array(Type.String(), {
				description:
					"Names, e-mail addresses or UPNs of everyone who has to be there; the signed-in user is always included",
			}),
		),
		start: Type.Optional(
			Type.String({
				description:
					"ISO 8601 local date-time, e.g. '2026-09-15T09:00:00'. Default: today, in 'timeZone'",
			}),
		),
		end: Type.Optional(
			Type.String({
				description:
					"ISO 8601 local date-time. Default: 7 days after 'start'",
			}),
		),
		durationMinutes: Type.Optional(
			Type.Number({ description: "Length of the meeting to look for, in minutes (default 30)" }),
		),
		timeZone: Type.Optional(
			Type.String({
				description:
					"IANA or Windows time zone the times are expressed in. Defaults to this machine's time zone",
			}),
		),
		account: AccountParam,
		tenant: TenantParam,
		limit: LimitParam,
	}),
	promptSnippet: "Find a free slot for several people",
	promptGuidelines: [
		"Use teams_availability to find a time before creating a meeting; propose the slots it returns and let the user pick.",
		"Availability is other people's data: pass it on as times, not as an account of how busy someone is.",
	],

	async execute(
		_toolCallId: string,
		params: {
			people?: string[];
			start?: string;
			end?: string;
			durationMinutes?: number;
			timeZone?: string;
			account?: string;
			tenant?: string;
			limit?: number;
		},
		signal: AbortSignal | undefined,
		_onUpdate: undefined,
		ctx: ToolContext,
	): Promise<ToolResult> {
		return run(async () => {
			const conn = connectionFor(ctx, params.account, params.tenant);
			const me = await currentUser(conn, signal);

			const timeZone = params.timeZone ?? localTimeZone();
			const duration = Math.min(Math.max(params.durationMinutes ?? 30, 5), 1440);

			// Default: the coming week, searched inside working hours. An explicit
			// window is honoured as-is — asking "are we free at 20:00?" is a
			// legitimate question that a 09:00–17:00 filter would answer wrongly.
			const explicit = !!params.start || !!params.end;
			const start = params.start ?? localDayStart(0);
			const end = params.end ?? localDayStart(DEFAULT_SEARCH_DAYS);

			/** scheduleId (lower case) → what to call that person */
			const names = new Map<string, string>();
			const schedules: string[] = [];

			const selfAddress = me.mail ?? me.upn;
			schedules.push(selfAddress);
			names.set(selfAddress.toLowerCase(), `${me.displayName} (you)`);

			const unresolved: string[] = [];
			for (const reference of params.people ?? []) {
				const person = await resolveUserId(conn, reference, signal);
				if (!person) {
					unresolved.push(reference);
					continue;
				}

				// Free/busy is other people's data, so the people rules apply.
				assertAccess(conn, "read", "people", person.displayName, [
					person.displayName,
					person.upn,
					person.mail,
					person.id,
				]);

				const address = person.mail ?? person.upn;
				if (!address) {
					unresolved.push(reference);
					continue;
				}
				if (!schedules.some((entry) => entry.toLowerCase() === address.toLowerCase())) {
					schedules.push(address);
					names.set(address.toLowerCase(), person.displayName);
				}
			}

			const windows = await getSchedule(conn, {
				schedules,
				start,
				end,
				timeZone,
				intervalMinutes: duration,
				signal,
			});

			const views: BusyView[] = windows.map((window) => ({
				name: names.get(window.scheduleId.toLowerCase()) ?? window.scheduleId,
				view: window.availabilityView,
			}));

			const slots = findCommonFreeSlots(
				views,
				{
					start,
					intervalMinutes: duration,
					notBefore: explicit ? undefined : localNow(),
					workingHoursOnly: !explicit,
				},
				params.limit ?? 5,
			);

			const who = [...names.values()].map((name) => truncate(name, 60)).join(", ");
			const windowLabel = explicit
				? `${formatSlotRange(start, end)}`
				: `${formatNaiveDay(start)} – ${formatNaiveDay(end)} on weekdays, 09:00–17:00`;

			const lines = [
				`## Free/busy — ${windowLabel} (${timeZone}), ${duration} min`,
				"",
				`Checked: ${who}`,
				"",
			];

			if (slots.length === 0) {
				lines.push("**No slot where everybody is free.**", "");
				if (windows.length > 0) {
					lines.push("Busy blocks:", ...describeBusy(windows, names), "");
				}
			} else {
				lines.push(`**${slots.length} slot(s) with everybody free:**`, "");
				for (const slot of slots) {
					lines.push(`- ${formatSlotRange(slot.start, slot.end)}`);
				}
				lines.push("");
			}

			if (unresolved.length > 0) {
				lines.push(`_Could not resolve: ${unresolved.join(", ")} — use teams_find_user._`, "");
			}

			// The individual blocks are only useful when a slot was proposed and the
			// reader wants to sanity-check it, or when nothing was found at all.
			if (slots.length > 0 && windows.some((window) => window.busy.length > 0)) {
				lines.push("Busy blocks in the window:", ...describeBusy(windows, names), "");
			}

			return textResult(lines.join("\n").trimEnd(), {
				slots: slots.length,
				people: schedules.length,
				unresolved: unresolved.length,
				timeZone,
			});
		});
	},
};
