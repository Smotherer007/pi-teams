/**
 * Pure transformation functions: domain data → display strings.
 *
 * No side effects, no I/O. Everything the agent reads about Teams goes through
 * here, so the rules are: never invent content, never silently drop a sender
 * or timestamp, and keep IDs visible so follow-up tool calls have something to
 * quote.
 */

import type {
	ChannelSummary,
	ChatSummary,
	DriveItemSummary,
	EventSummary,
	MemberSummary,
	MessageSummary,
	PresenceInfo,
	TeamSummary,
} from "../types.ts";

// ---------------------------------------------------------------------------
// HTML → text
// ---------------------------------------------------------------------------

const ENTITIES: Record<string, string> = {
	"&nbsp;": " ",
	"&amp;": "&",
	"&lt;": "<",
	"&gt;": ">",
	"&quot;": '"',
	"&#39;": "'",
	"&apos;": "'",
};

/**
 * Flatten Teams message HTML into readable plain text.
 *
 * Teams sends most messages as HTML even when the user typed plain text, and
 * an agent reading raw `<div>`s wastes tokens and misreads content. This keeps
 * line structure, mention text and link targets, and drops the rest.
 */
export function htmlToText(content: string, contentType = "html"): string {
	if (!content) return "";
	if (contentType.toLowerCase() !== "html") return content.trim();

	let text = content;

	// Keep the target of real links: "text (https://…)"
	text = text.replace(
		/<a\b[^>]*href=["']([^"']+)["'][^>]*>([\s\S]*?)<\/a>/gi,
		(_m, href: string, label: string) => {
			const clean = label.replace(/<[^>]+>/g, "").trim();
			if (!clean) return href;
			return clean === href ? href : `${clean} (${href})`;
		},
	);

	// Images carry no text; name them so a reader knows something was there.
	text = text.replace(/<img\b[^>]*alt=["']([^"']*)["'][^>]*>/gi, (_m, alt: string) =>
		alt ? `[image: ${alt}]` : "[image]",
	);
	text = text.replace(/<img\b[^>]*>/gi, "[image]");

	// Block elements become line breaks.
	text = text.replace(/<\/(p|div|li|tr|h[1-6])>/gi, "\n");
	text = text.replace(/<br\s*\/?>/gi, "\n");
	text = text.replace(/<li\b[^>]*>/gi, "- ");

	// Everything else goes.
	text = text.replace(/<[^>]+>/g, "");

	for (const [entity, replacement] of Object.entries(ENTITIES)) {
		text = text.split(entity).join(replacement);
	}
	text = text.replace(/&#(\d+);/g, (_m, code: string) => String.fromCharCode(Number(code)));

	return text
		.split("\n")
		.map((line) => line.replace(/[ \t]+/g, " ").trimEnd())
		.join("\n")
		.replace(/\n{3,}/g, "\n\n")
		.trim();
}

// ---------------------------------------------------------------------------
// Dates
// ---------------------------------------------------------------------------

/** ISO timestamp → "2026-09-14 08:31" in the local timezone. */
export function formatDate(iso: string | undefined): string {
	if (!iso) return "";
	const date = new Date(iso);
	if (Number.isNaN(date.getTime())) return iso;
	const pad = (n: number) => String(n).padStart(2, "0");
	return (
		`${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())} ` +
		`${pad(date.getHours())}:${pad(date.getMinutes())}`
	);
}

/** "3 min ago", "2 h ago", "5 d ago" — for scanning a list quickly. */
export function formatRelative(iso: string | undefined): string {
	if (!iso) return "";
	const then = new Date(iso).getTime();
	if (Number.isNaN(then)) return "";
	const diffMs = Date.now() - then;
	const minutes = Math.round(diffMs / 60000);
	if (minutes < 1) return "just now";
	if (minutes < 60) return `${minutes} min ago`;
	const hours = Math.round(minutes / 60);
	if (hours < 24) return `${hours} h ago`;
	const days = Math.round(hours / 24);
	if (days < 30) return `${days} d ago`;
	return formatDate(iso);
}

const WEEKDAYS = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
const MONTHS = [
	"Jan",
	"Feb",
	"Mar",
	"Apr",
	"May",
	"Jun",
	"Jul",
	"Aug",
	"Sep",
	"Oct",
	"Nov",
	"Dec",
];

/**
 * Split a naive Graph date-time into its parts, as UTC.
 *
 * Availability views and free/busy blocks arrive already localised to the zone
 * the caller asked Graph for, so re-interpreting them in the machine's zone
 * would shift every time by the offset between the two.
 */
function naiveParts(value: string | undefined):
	| { weekday: string; day: number; month: string; year: number; time: string }
	| undefined {
	if (!value) return undefined;
	const match = /^(\d{4})-(\d{2})-(\d{2})[T ](\d{2}):(\d{2})/.exec(value.trim());
	if (!match) return undefined;
	const [, year, month, day, hour, minute] = match;
	const asUtc = new Date(Date.UTC(Number(year), Number(month) - 1, Number(day)));
	return {
		weekday: WEEKDAYS[asUtc.getUTCDay()] ?? "",
		day: Number(day),
		month: MONTHS[Number(month) - 1] ?? "",
		year: Number(year),
		time: `${hour}:${minute}`,
	};
}

/** "Mon 15 Sep 2026" — a window's day, without the clock time. */
export function formatNaiveDay(value: string | undefined): string {
	const parts = naiveParts(value);
	if (!parts) return value ?? "";
	return `${parts.weekday} ${parts.day} ${parts.month} ${parts.year}`;
}

/** "09:00–09:30" — the clock part of a free/busy block. */
export function formatTimeRange(start: string | undefined, end: string | undefined): string {
	const from = naiveParts(start);
	const to = naiveParts(end);
	if (!from || !to) return "unknown time";
	return `${from.time}–${to.time}`;
}

/** "Mon 15 Sep, 09:00–09:30" — one proposed slot. */
export function formatSlotRange(start: string | undefined, end: string | undefined): string {
	const parts = naiveParts(start);
	if (!parts) return `${start ?? "?"}–${end ?? "?"}`;
	return `${parts.weekday} ${parts.day} ${parts.month}, ${formatTimeRange(start, end)}`;
}

/** Shorten a string for previews, without cutting mid-escape. */
export function truncate(text: string, max: number): string {
	const oneLine = text.replace(/\s*\n\s*/g, " ").trim();
	if (oneLine.length <= max) return oneLine;
	return `${oneLine.slice(0, Math.max(0, max - 1))}…`;
}

export function formatBytes(bytes: number | undefined): string {
	if (bytes === undefined) return "";
	const units = ["B", "KB", "MB", "GB"];
	let value = bytes;
	let unit = 0;
	while (value >= 1024 && unit < units.length - 1) {
		value /= 1024;
		unit += 1;
	}
	return `${value.toFixed(value < 10 && unit > 0 ? 1 : 0)} ${units[unit]}`;
}

// ---------------------------------------------------------------------------
// Teams & channels
// ---------------------------------------------------------------------------

export function formatTeamList(teams: readonly TeamSummary[]): string {
	if (teams.length === 0) return "No teams found (or none allowed by your configuration).";
	const lines = [`Found ${teams.length} team(s):`, ""];
	for (const team of teams) {
		const flags = [team.visibility, team.isArchived ? "archived" : undefined]
			.filter(Boolean)
			.join(", ");
		lines.push(`- **${team.displayName}**${flags ? ` (${flags})` : ""}`);
		if (team.description) lines.push(`  ${truncate(team.description, 120)}`);
		lines.push(`  id: ${team.id}`);
	}
	return lines.join("\n");
}

export function formatChannelList(channels: readonly ChannelSummary[]): string {
	if (channels.length === 0) return "No channels found (or none allowed by your configuration).";
	const lines = [`Found ${channels.length} channel(s):`, ""];
	for (const channel of channels) {
		const path = channel.teamName ? `${channel.teamName}/${channel.displayName}` : channel.displayName;
		lines.push(`- **${path}**${channel.membershipType ? ` (${channel.membershipType})` : ""}`);
		if (channel.description) lines.push(`  ${truncate(channel.description, 120)}`);
		lines.push(`  channelId: ${channel.id}${channel.teamId ? ` · teamId: ${channel.teamId}` : ""}`);
	}
	return lines.join("\n");
}

export function formatMemberList(members: readonly MemberSummary[], title: string): string {
	if (members.length === 0) return `${title}: no members visible.`;
	const lines = [`${title} — ${members.length} member(s):`, ""];
	for (const member of members) {
		const roles = member.roles.length > 0 ? ` [${member.roles.join(", ")}]` : "";
		const contact = member.upn ?? member.mail;
		lines.push(`- ${member.displayName}${contact ? ` <${contact}>` : ""}${roles}`);
	}
	return lines.join("\n");
}

// ---------------------------------------------------------------------------
// Chats
// ---------------------------------------------------------------------------

export function formatChatList(chats: readonly ChatSummary[]): string {
	if (chats.length === 0) return "No chats found (or none allowed by your configuration).";
	const lines = [`Found ${chats.length} chat(s):`, ""];
	for (const chat of chats) {
		lines.push(`- **${chat.label}** (${chat.chatType})`);
		if (chat.lastMessagePreview) {
			const who = chat.lastMessageFrom ? `${chat.lastMessageFrom}: ` : "";
			lines.push(`  ${who}${truncate(chat.lastMessagePreview, 120)}`);
		}
		const when = chat.lastUpdated ? ` · ${formatRelative(chat.lastUpdated)}` : "";
		lines.push(`  chatId: ${chat.id}${when}`);
	}
	return lines.join("\n");
}

export function formatChatDetail(chat: ChatSummary): string {
	const lines = [`## ${chat.label}`, ""];
	lines.push(`- **Type:** ${chat.chatType}`);
	if (chat.topic) lines.push(`- **Topic:** ${chat.topic}`);
	lines.push(`- **Members:** ${chat.members.map((m) => m.displayName).join(", ") || "(none visible)"}`);
	if (chat.lastUpdated) lines.push(`- **Last activity:** ${formatDate(chat.lastUpdated)}`);
	lines.push(`- **chatId:** ${chat.id}`);
	if (chat.webUrl) lines.push(`- **Open in Teams:** ${chat.webUrl}`);
	return lines.join("\n");
}

// ---------------------------------------------------------------------------
// Messages
// ---------------------------------------------------------------------------

/** One message, rendered for reading. */
export function formatMessage(message: MessageSummary, options: { showId?: boolean } = {}): string {
	const who = message.from?.displayName ?? "(system)";
	const when = message.createdDateTime ? formatDate(message.createdDateTime) : "";
	const header = `**${who}**${when ? ` · ${when}` : ""}`;

	const lines = [header];
	if (message.subject) lines.push(`_${message.subject}_`);

	if (message.deletedDateTime) {
		lines.push("(message deleted)");
	} else {
		lines.push(message.text || "(empty message)");
	}

	const meta: string[] = [];
	if (message.reactions.length > 0) {
		const grouped = new Map<string, number>();
		for (const reaction of message.reactions) {
			grouped.set(reaction.type, (grouped.get(reaction.type) ?? 0) + 1);
		}
		meta.push(
			[...grouped.entries()].map(([type, count]) => `${type}${count > 1 ? `×${count}` : ""}`).join(" "),
		);
	}
	if (message.attachments.length > 0) {
		meta.push(`attachments: ${message.attachments.map((a) => a.name ?? "file").join(", ")}`);
	}
	if (message.imageUrls.length > 0) {
		// The images themselves are not in the text. Saying how many there are is
		// what tells a reader that something is missing and where to get it.
		const count = message.imageUrls.length;
		meta.push(`${count} image${count === 1 ? "" : "s"} — use teams_download_files`);
	}
	if (message.replyCount !== undefined && message.replyCount > 0) {
		meta.push(`${message.replyCount} repl${message.replyCount === 1 ? "y" : "ies"}`);
	}
	if (options.showId !== false) meta.push(`messageId: ${message.id}`);
	if (meta.length > 0) lines.push(`  _${meta.join(" · ")}_`);

	return lines.join("\n");
}

export function formatMessageList(
	messages: readonly MessageSummary[],
	title: string,
	options: { oldestFirst?: boolean } = {},
): string {
	if (messages.length === 0) return `${title}: no messages.`;
	const ordered = options.oldestFirst === false ? [...messages] : [...messages].reverse();
	const lines = [`## ${title}`, "", `${messages.length} message(s), oldest first:`, ""];
	for (const message of ordered) {
		lines.push(formatMessage(message));
		lines.push("");
	}
	return lines.join("\n").trimEnd();
}

// ---------------------------------------------------------------------------
// Presence
// ---------------------------------------------------------------------------

const PRESENCE_ICONS: Record<string, string> = {
	Available: "🟢",
	AvailableIdle: "🟢",
	Away: "🟡",
	BeRightBack: "🟡",
	Busy: "🔴",
	BusyIdle: "🔴",
	DoNotDisturb: "⛔",
	Offline: "⚪",
	PresenceUnknown: "⚪",
};

export function formatPresence(presence: PresenceInfo): string {
	const icon = PRESENCE_ICONS[presence.availability] ?? "⚪";
	const name = presence.displayName ? `${presence.displayName}: ` : "";
	const status = presence.statusMessage ? ` — "${truncate(presence.statusMessage, 100)}"` : "";
	return `${icon} ${name}${presence.availability} (${presence.activity})${status}`;
}

export function formatPresenceList(entries: readonly PresenceInfo[]): string {
	if (entries.length === 0) return "No presence information available.";
	return entries.map((entry) => `- ${formatPresence(entry)}`).join("\n");
}

// ---------------------------------------------------------------------------
// Calendar
// ---------------------------------------------------------------------------

export function formatEventList(events: readonly EventSummary[]): string {
	if (events.length === 0) return "No events in that range.";
	const lines = [`${events.length} event(s):`, ""];
	for (const event of events) {
		const when = event.isAllDay
			? `${formatDate(event.start)} (all day)`
			: `${formatDate(event.start)} – ${formatDate(event.end)?.slice(-5)}`;
		lines.push(`- **${event.subject}** · ${when}`);
		if (event.organizer) lines.push(`  organizer: ${event.organizer.displayName}`);
		if (event.attendees.length > 0) {
			lines.push(`  attendees: ${truncate(event.attendees.map((a) => a.displayName).join(", "), 140)}`);
		}
		if (event.joinUrl) lines.push(`  join: ${event.joinUrl}`);
		lines.push(`  eventId: ${event.id}`);
	}
	return lines.join("\n");
}

export function formatEventDetail(event: EventSummary): string {
	const lines = [`## ${event.subject}`, ""];
	lines.push(`- **When:** ${formatDate(event.start)} – ${formatDate(event.end)}${event.timeZone ? ` (${event.timeZone})` : ""}`);
	if (event.location) lines.push(`- **Location:** ${event.location}`);
	if (event.organizer) lines.push(`- **Organizer:** ${event.organizer.displayName}`);
	if (event.attendees.length > 0) {
		lines.push(`- **Attendees:** ${event.attendees.map((a) => a.displayName).join(", ")}`);
	}
	if (event.joinUrl) lines.push(`- **Join:** ${event.joinUrl}`);
	lines.push(`- **eventId:** ${event.id}`);
	if (event.bodyPreview) {
		lines.push("", "### Notes", "", truncate(event.bodyPreview, 600));
	}
	return lines.join("\n");
}

// ---------------------------------------------------------------------------
// Files
// ---------------------------------------------------------------------------

export function formatFileList(items: readonly DriveItemSummary[], title: string): string {
	if (items.length === 0) return `${title}: no files.`;
	const lines = [`## ${title}`, ""];
	for (const item of items) {
		const icon = item.isFolder ? "📁" : "📄";
		const size = item.isFolder ? "" : ` · ${formatBytes(item.size)}`;
		const modified = item.lastModifiedDateTime ? ` · ${formatRelative(item.lastModifiedDateTime)}` : "";
		lines.push(`- ${icon} **${item.name}**${size}${modified}`);
		if (item.webUrl) lines.push(`  ${item.webUrl}`);
	}
	return lines.join("\n");
}

/**
 * The result of saving a message's files.
 *
 * Typed structurally on purpose: this module is a leaf and knows nothing about
 * Graph, so the caller passes the shape it already has.
 */
export function formatDownloadResult(
	label: string,
	dir: string,
	saved: readonly { path: string; name: string; bytes: number; kind: string }[],
	skipped: readonly { name: string; reason: string; url?: string }[],
): string {
	const lines = [`## Files from ${label}`, ""];

	if (saved.length === 0) {
		lines.push("Nothing to download.");
	} else {
		lines.push(`${saved.length} file${saved.length === 1 ? "" : "s"} saved to \`${dir}\`:`, "");
		for (const file of saved) {
			const icon = file.kind === "image" ? "🖼" : "📄";
			lines.push(`- ${icon} \`${file.path}\` (${formatBytes(file.bytes)})`);
		}
		lines.push("", "Open them with the file-reading tool.");
	}

	if (skipped.length > 0) {
		lines.push("", "Not saved:");
		for (const file of skipped) {
			// The URL is the fallback: it can be opened by hand where the download
			// was refused.
			lines.push(`- **${file.name}** — ${file.reason}${file.url ? `\n  ${file.url}` : ""}`);
		}
	}

	return lines.join("\n");
}
