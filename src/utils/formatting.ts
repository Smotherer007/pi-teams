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

/** Convert plain text to the minimal HTML Teams expects for a message body. */
export function textToHtml(text: string): string {
	const escaped = text
		.replace(/&/g, "&amp;")
		.replace(/</g, "&lt;")
		.replace(/>/g, "&gt;");
	return escaped
		.split(/\n{2,}/)
		.map((para) => `<p>${para.replace(/\n/g, "<br>")}</p>`)
		.join("");
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
