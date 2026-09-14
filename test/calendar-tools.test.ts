/**
 * The calendar tools: answering an invitation, and finding a slot.
 *
 * The slot arithmetic is the part of `teams_availability` that fails silently —
 * an off-by-one between the availability view and the clock would suggest a
 * time somebody is busy in, and nothing downstream would notice. It is pure, so
 * it is tested here against hand-written views.
 */

import { test, describe } from "node:test";
import assert from "node:assert/strict";
import {
	addMinutes,
	describeBusy,
	findCommonFreeSlots,
	parseNaive,
} from "../src/utils/slots.ts";
import { formatNaiveDay, formatSlotRange, formatTimeRange } from "../src/utils/formatting.ts";
import { blockReason, formatMutationSummary, isMutationTool } from "../src/safety/index.ts";
import { teamsRespondInviteTool } from "../src/tools/teams-meetings.ts";
import { teamsAvailabilityTool } from "../src/tools/teams-availability.ts";

/** Tuesday 15 September 2026, 09:00, naive. */
const MONDAY_9 = "2026-09-15T09:00:00";

describe("parseNaive / addMinutes", () => {
	test("a naive date-time is read as UTC, so the machine's zone cannot shift it", () => {
		assert.equal(parseNaive(MONDAY_9), Date.UTC(2026, 8, 15, 9, 0, 0));
	});

	test("Graph's seven-digit fractional seconds are tolerated", () => {
		assert.equal(parseNaive("2026-09-15T09:00:00.0000000"), Date.UTC(2026, 8, 15, 9, 0, 0));
	});

	test("rubbish and missing values return undefined rather than 1970", () => {
		assert.equal(parseNaive(undefined), undefined);
		assert.equal(parseNaive(""), undefined);
		assert.equal(parseNaive("2026-09-15"), undefined);
	});

	test("adding minutes crosses the hour", () => {
		assert.equal(addMinutes("2026-09-15T09:45:00", 30), "2026-09-15T10:15:00");
	});
});

describe("findCommonFreeSlots", () => {
	const everyoneFree = "0000";

	test("index i maps to the i-th slot in time", () => {
		const slots = findCommonFreeSlots([{ name: "Anna", view: "0000" }], {
			start: MONDAY_9,
			intervalMinutes: 30,
		});
		assert.deepEqual(
			slots.map((slot) => `${slot.start}–${slot.end}`),
			[
				"2026-09-15T09:00:00–2026-09-15T09:30:00",
				"2026-09-15T09:30:00–2026-09-15T10:00:00",
				"2026-09-15T10:00:00–2026-09-15T10:30:00",
				"2026-09-15T10:30:00–2026-09-15T11:00:00",
			],
		);
	});

	test("a slot needs everybody free, and busy is not the only obstacle", () => {
		const slots = findCommonFreeSlots(
			[
				{ name: "Anna", view: everyoneFree },
				{ name: "Tom", view: "0220" },
			],
			{ start: MONDAY_9, intervalMinutes: 30 },
		);
		assert.deepEqual(
			slots.map((slot) => slot.start.slice(11, 16)),
			["09:00", "10:30"],
		);
	});

	test("tentative and out-of-office count as taken", () => {
		const slots = findCommonFreeSlots(
			[{ name: "Anna", view: "1234" }],
			{ start: MONDAY_9, intervalMinutes: 30 },
		);
		assert.deepEqual(slots, []);
	});

	test("holes in the view are not free either", () => {
		const slots = findCommonFreeSlots([{ name: "Anna", view: "0-0" }], {
			start: MONDAY_9,
			intervalMinutes: 30,
		});
		assert.deepEqual(
			slots.map((slot) => slot.start.slice(11, 16)),
			["09:00", "10:00"],
		);
	});

	test("a truncated view ends the search — the unknown tail is not free", () => {
		const slots = findCommonFreeSlots(
			[
				{ name: "Anna", view: "000000" },
				{ name: "Tom", view: "00" },
			],
			{ start: MONDAY_9, intervalMinutes: 30 },
		);
		assert.equal(slots.length, 2);
	});

	test("workingHoursOnly drops the weekend and everything outside 09:00–17:00", () => {
		// Friday 18 September 2026, 16:30 — one working slot left before the weekend,
		// then nothing until Monday 09:00.
		const friday = "2026-09-18T16:30:00";
		const slots = findCommonFreeSlots([{ name: "Anna", view: "0".repeat(131) }], {
			start: friday,
			intervalMinutes: 30,
			workingHoursOnly: true,
		}, 3);

		assert.deepEqual(
			slots.map((slot) => `${slot.start.slice(0, 10)} ${slot.start.slice(11, 16)}`),
			["2026-09-18 16:30", "2026-09-21 09:00", "2026-09-21 09:30"],
		);
		// The last slot before the weekend ends at 17:00 — 17:00 is the end of the
		// working day, not the start of another slot.
		assert.equal(slots[0]!.end.slice(11, 16), "17:00");
	});

	test("notBefore skips the past", () => {
		const slots = findCommonFreeSlots([{ name: "Anna", view: "0000" }], {
			start: MONDAY_9,
			intervalMinutes: 30,
			notBefore: "2026-09-15T10:00:00",
		});
		assert.deepEqual(
			slots.map((slot) => slot.start.slice(11, 16)),
			["10:00", "10:30"],
		);
	});

	test("the limit caps the result", () => {
		const slots = findCommonFreeSlots([{ name: "Anna", view: "00000000" }], {
			start: MONDAY_9,
			intervalMinutes: 30,
		}, 3);
		assert.equal(slots.length, 3);
	});

	test("an empty field or an unreadable start yields nothing, not a guess", () => {
		assert.deepEqual(findCommonFreeSlots([], { start: MONDAY_9, intervalMinutes: 30 }), []);
		assert.deepEqual(
			findCommonFreeSlots([{ name: "Anna", view: "0000" }], { start: "nonsense", intervalMinutes: 30 }),
			[],
		);
	});

	test("minutes, not hours: a 15-minute grid lands on :15 and :45", () => {
		const slots = findCommonFreeSlots([{ name: "Anna", view: "0000" }], {
			start: MONDAY_9,
			intervalMinutes: 15,
		});
		assert.deepEqual(
			slots.map((slot) => slot.start.slice(11, 16)),
			["09:00", "09:15", "09:30", "09:45"],
		);
	});
});

describe("slot formatting", () => {
	test("a slot reads as weekday, date and clock range", () => {
		assert.equal(formatSlotRange(MONDAY_9, "2026-09-15T09:30:00"), "Tue 15 Sep, 09:00–09:30");
	});

	test("the weekday comes from the date itself, not from the machine's clock", () => {
		assert.equal(formatNaiveDay("2026-09-18T00:00:00"), "Fri 18 Sep 2026");
		assert.equal(formatNaiveDay("2026-09-21T00:00:00"), "Mon 21 Sep 2026");
	});

	test("a block without times degrades instead of printing NaN", () => {
		assert.equal(formatTimeRange(undefined, undefined), "unknown time");
		assert.equal(formatSlotRange(undefined, undefined), "?–?");
	});
});

describe("describeBusy", () => {
	test("reports times and status, never the subject", () => {
		const names = new Map([["anna@contoso.com", "Anna Schmidt"]]);
		const lines = describeBusy(
			[
				{
					scheduleId: "anna@contoso.com",
					busy: [{ status: "busy", start: "2026-09-15T09:00:00.0000000", end: "2026-09-15T10:00:00.0000000" }],
				},
			],
			names,
		);
		assert.deepEqual(lines, ["- Anna Schmidt: 09:00–10:00 (busy)"]);
	});

	test("an unknown address falls back to the address itself", () => {
		const lines = describeBusy([{ scheduleId: "x@contoso.com", busy: [] }], new Map());
		assert.deepEqual(lines, ["- x@contoso.com: no busy blocks reported"]);
	});
});

describe("teams_respond_invite", () => {
	test("it is a mutation, so readonly blocks it", () => {
		assert.equal(isMutationTool("teams_respond_invite"), true);
		assert.match(blockReason("readonly", "device-code", "teams_respond_invite") ?? "", /readonly/);
	});

	test("app-only may still answer an invitation — the calendar is not messaging", () => {
		assert.equal(blockReason("open", "client-credentials", "teams_respond_invite"), undefined);
	});

	test("the dialog names the answer, not the tool", () => {
		assert.match(
			formatMutationSummary("teams_respond_invite", { eventId: "AAMk", response: "decline" }),
			/decline the invitation AAMk/,
		);
		assert.match(
			formatMutationSummary("teams_respond_invite", {
				eventId: "AAMk",
				response: "accept",
				comment: "Bin dabei",
			}),
			/Bin dabei/,
		);
	});

	test("the tool offers accept, decline and tentative", () => {
		assert.equal(teamsRespondInviteTool.name, "teams_respond_invite");
		assert.deepEqual(Object.keys(teamsRespondInviteTool.parameters.properties), [
			"eventId",
			"response",
			"comment",
			"sendResponse",
			"account",
			"tenant",
		]);
		assert.match(teamsRespondInviteTool.description, /organizer is notified/i);
	});
});

describe("teams_availability", () => {
	test("it is a read, so readonly does not block it", () => {
		assert.equal(isMutationTool("teams_availability"), false);
		assert.equal(blockReason("readonly", "device-code", "teams_availability"), undefined);
	});

	test("the tool says it reads availability, not appointments", () => {
		assert.equal(teamsAvailabilityTool.name, "teams_availability");
		assert.match(teamsAvailabilityTool.description, /free\/busy/i);
		assert.match(teamsAvailabilityTool.description, /never the appointments/i);
		assert.deepEqual(Object.keys(teamsAvailabilityTool.parameters.properties), [
			"people",
			"start",
			"end",
			"durationMinutes",
			"timeZone",
			"account",
			"tenant",
			"limit",
		]);
	});
});
