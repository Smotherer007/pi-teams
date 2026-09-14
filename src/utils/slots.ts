/**
 * Finding a slot everyone is free in.
 *
 * Graph returns availability as a *string*, one character per time slot since
 * the window start — so the one thing that can go silently wrong is the mapping
 * from character index back to a clock time. That mapping lives here, pure and
 * testable, instead of inside a tool where the only test would be a live
 * calendar.
 *
 * Naive date-times ("2026-09-15T09:00:00", no offset) are parsed as UTC on
 * purpose: they already carry the zone the caller asked Graph for, and reading
 * them as UTC keeps the arithmetic independent of the machine's own zone.
 */

/** One person's availability, as Graph's availability view expresses it. */
export interface BusyView {
	/** Name to report the person by */
	name: string;
	/** One character per slot: `0` free, `1` tentative, `2` busy, `3` out of office, `4` working elsewhere */
	view: string;
}

export interface SlotWindow {
	/** First slot, naive ISO, aligned with index 0 of every view */
	start: string;
	/** Slot width in minutes — the same value the Graph call used */
	intervalMinutes: number;
	/** Ignore slots that begin before this naive ISO (used to skip the past) */
	notBefore?: string;
	/** Also ignore slots outside Mon–Fri, 09:00–17:00 */
	workingHoursOnly?: boolean;
}

export interface FreeSlot {
	start: string;
	end: string;
}

const FREE = "0";
const WORKDAY_START = 9 * 60;
const WORKDAY_END = 17 * 60;
const MINUTE_MS = 60_000;

/**
 * Parse a naive Graph date-time into a UTC timestamp.
 *
 * Returns undefined for anything without a date and time rather than guessing:
 * a missing value must not become "1 January 1970, which looks free".
 */
export function parseNaive(value: string | undefined): number | undefined {
	if (!value) return undefined;
	const match = /^(\d{4})-(\d{2})-(\d{2})[T ](\d{2}):(\d{2})(?::(\d{2}))?/.exec(value.trim());
	if (!match) return undefined;
	const [, year, month, day, hour, minute, second] = match;
	return Date.UTC(
		Number(year),
		Number(month) - 1,
		Number(day),
		Number(hour),
		Number(minute),
		Number(second ?? 0),
	);
}

/** UTC timestamp → naive ISO without milliseconds. */
export function formatNaive(timestamp: number): string {
	return new Date(timestamp).toISOString().slice(0, 19);
}

/** Shift a naive date-time by whole minutes. */
export function addMinutes(naive: string, minutes: number): string | undefined {
	const parsed = parseNaive(naive);
	return parsed === undefined ? undefined : formatNaive(parsed + minutes * MINUTE_MS);
}

/** `YYYY-MM-DDT00:00:00` for today + offset, in the machine's time zone. */
export function localDayStart(offsetDays = 0): string {
	const date = new Date();
	date.setDate(date.getDate() + offsetDays);
	return `${localDate(date)}T00:00:00`;
}

/** The current time as a naive date-time, in the machine's time zone. */
export function localNow(): string {
	return `${localDate(new Date())}T${localTime(new Date())}`;
}

/** The machine's time zone, for calls that would otherwise default to UTC. */
export function localTimeZone(): string {
	try {
		return Intl.DateTimeFormat().resolvedOptions().timeZone || "UTC";
	} catch {
		return "UTC";
	}
}

function localDate(date: Date): string {
	const pad = (n: number) => String(n).padStart(2, "0");
	return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}`;
}

function localTime(date: Date): string {
	const pad = (n: number) => String(n).padStart(2, "0");
	return `${pad(date.getHours())}:${pad(date.getMinutes())}:00`;
}

function insideWorkingHours(start: number, end: number): boolean {
	const from = new Date(start);
	const to = new Date(end);
	const weekday = from.getUTCDay();
	if (weekday === 0 || weekday === 6) return false;
	// A slot has to fit inside one working day; a window across midnight is not
	// a 30-minute meeting anyway.
	if (from.toISOString().slice(0, 10) !== to.toISOString().slice(0, 10)) return false;
	const startMinutes = from.getUTCHours() * 60 + from.getUTCMinutes();
	const endMinutes = to.getUTCHours() * 60 + to.getUTCMinutes();
	return startMinutes >= WORKDAY_START && endMinutes <= WORKDAY_END;
}

/** The first slots in which nobody has anything. */
export function findCommonFreeSlots(
	views: readonly BusyView[],
	window: SlotWindow,
	limit = 5,
): FreeSlot[] {
	const base = parseNaive(window.start);
	if (base === undefined || views.length === 0) return [];

	// The shortest view decides how far the search goes: where Graph truncated
	// a person's availability the tail is unknown, and unknown is not free.
	const horizon = Math.min(...views.map((entry) => entry.view.length));
	const earliest = parseNaive(window.notBefore);
	const step = window.intervalMinutes * MINUTE_MS;

	const slots: FreeSlot[] = [];
	for (let index = 0; index < horizon && slots.length < limit; index += 1) {
		if (views.some((entry) => entry.view[index] !== FREE)) continue;

		const start = base + index * step;
		const end = start + step;
		if (earliest !== undefined && start < earliest) continue;
		if (window.workingHoursOnly && !insideWorkingHours(start, end)) continue;

		slots.push({ start: formatNaive(start), end: formatNaive(end) });
	}

	return slots;
}

/**
 * The statuses keeping everyone apart, for when no slot is free.
 *
 * Turns "nothing works" into something a person can act on, and reports only
 * times and statuses — never what the appointment is.
 */
export function describeBusy(
	busy: readonly {
		scheduleId: string;
		busy: Array<{ status: string; start: string; end: string }>;
	}[],
	names: Map<string, string>,
): string[] {
	return busy.map((entry) => {
		const name = names.get(entry.scheduleId.toLowerCase()) ?? entry.scheduleId;
		const blocks = entry.busy
			.slice(0, 6)
			.map((block) => `${block.start.slice(11, 16)}–${block.end.slice(11, 16)} (${block.status})`)
			.join(", ");
		return `- ${name}: ${blocks || "no busy blocks reported"}`;
	});
}
