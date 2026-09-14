/**
 * Graph JSON → domain types.
 *
 * Pure functions, no I/O. Graph responses are wide, inconsistently populated
 * and occasionally change shape between endpoints; mapping them once here
 * keeps every tool and formatter working against the stable interfaces in
 * ../types.ts.
 */

import type {
	ChannelSummary,
	ChatSummary,
	DriveItemSummary,
	EventSummary,
	MemberSummary,
	MessageAttachment,
	MessageLocation,
	MessageReaction,
	MessageSummary,
	PersonRef,
	PresenceInfo,
	SignedInUser,
	TeamSummary,
} from "../types.ts";
import { htmlToText } from "../utils/formatting.ts";

/* eslint-disable @typescript-eslint/no-explicit-any */
type Raw = Record<string, any>;

// ---------------------------------------------------------------------------
// People
// ---------------------------------------------------------------------------

export function mapUser(raw: Raw | undefined): SignedInUser | undefined {
	if (!raw) return undefined;
	return {
		id: raw.id ?? "",
		displayName: raw.displayName ?? raw.userPrincipalName ?? "(unknown)",
		upn: raw.userPrincipalName ?? "",
		mail: raw.mail ?? undefined,
		tenantId: raw.tenantId ?? "",
		jobTitle: raw.jobTitle ?? undefined,
	};
}

export function mapPerson(raw: Raw | undefined): PersonRef | undefined {
	if (!raw) return undefined;
	// Identities appear as {user:{...}}, {application:{...}}, or bare.
	const identity = raw.user ?? raw.application ?? raw.device ?? raw;
	const name = identity.displayName ?? identity.userPrincipalName ?? identity.email;
	if (!name && !identity.id) return undefined;
	return {
		id: identity.id ?? undefined,
		displayName: name ?? "(unknown)",
		upn: identity.userPrincipalName ?? undefined,
		mail: identity.email ?? identity.mail ?? undefined,
	};
}

export function mapMember(raw: Raw): MemberSummary {
	return {
		id: raw.id ?? "",
		displayName: raw.displayName ?? "(unknown)",
		upn: raw.userPrincipalName ?? undefined,
		mail: raw.email ?? raw.mail ?? undefined,
		roles: Array.isArray(raw.roles) ? raw.roles : [],
	};
}

/** Identifiers a person can be matched against by the scope rules. */
export function personCandidates(person: PersonRef | MemberSummary | undefined): (string | undefined)[] {
	if (!person) return [];
	return [person.displayName, (person as PersonRef).upn, (person as PersonRef).mail, person.id];
}

// ---------------------------------------------------------------------------
// Teams & channels
// ---------------------------------------------------------------------------

export function mapTeam(raw: Raw): TeamSummary {
	return {
		id: raw.id ?? "",
		displayName: raw.displayName ?? "(unnamed team)",
		description: raw.description ?? undefined,
		visibility: raw.visibility ?? undefined,
		isArchived: raw.isArchived ?? undefined,
		webUrl: raw.webUrl ?? undefined,
	};
}

export function mapChannel(raw: Raw, team?: { id?: string; name?: string }): ChannelSummary {
	return {
		id: raw.id ?? "",
		displayName: raw.displayName ?? "(unnamed channel)",
		description: raw.description ?? undefined,
		membershipType: raw.membershipType ?? undefined,
		webUrl: raw.webUrl ?? undefined,
		teamId: team?.id ?? undefined,
		teamName: team?.name ?? undefined,
	};
}

/** Identifiers a channel can be matched against — including "Team/Channel". */
export function channelCandidates(channel: ChannelSummary): (string | undefined)[] {
	const path = channel.teamName ? `${channel.teamName}/${channel.displayName}` : undefined;
	return [path, channel.displayName, channel.id, channel.teamName];
}

export function teamCandidates(team: TeamSummary): (string | undefined)[] {
	return [team.displayName, team.id];
}

// ---------------------------------------------------------------------------
// Chats
// ---------------------------------------------------------------------------

/**
 * Build a chat summary.
 *
 * Teams leaves 1:1 and ad-hoc group chats without a topic, so the label is
 * derived from the other participants — that is also what the scope rules
 * match against, which is why `me` is excluded from it.
 */
export function mapChat(raw: Raw, meId?: string): ChatSummary {
	const members: PersonRef[] = (raw.members ?? [])
		.map((m: Raw) => mapPerson(m))
		.filter((p: PersonRef | undefined): p is PersonRef => !!p);

	const others = meId ? members.filter((m) => m.id !== meId) : members;
	const derived = others.map((m) => m.displayName).filter(Boolean).join(", ");

	const preview = raw.lastMessagePreview?.body?.content
		? htmlToText(
				raw.lastMessagePreview.body.content,
				raw.lastMessagePreview.body.contentType ?? "text",
			)
		: undefined;

	return {
		id: raw.id ?? "",
		chatType: raw.chatType ?? "unknown",
		topic: raw.topic ?? undefined,
		label: raw.topic || derived || raw.id || "(chat)",
		members,
		lastUpdated: raw.lastUpdatedDateTime ?? raw.lastMessagePreview?.createdDateTime ?? undefined,
		lastMessagePreview: preview,
		lastMessageFrom: mapPerson(raw.lastMessagePreview?.from)?.displayName,
		webUrl: raw.webUrl ?? undefined,
	};
}

/** Identifiers a chat can be matched against — label, topic, id, participants. */
export function chatCandidates(chat: ChatSummary): (string | undefined)[] {
	return [
		chat.label,
		chat.topic,
		chat.id,
		...chat.members.flatMap((m) => [m.displayName, m.upn, m.mail]),
	];
}

// ---------------------------------------------------------------------------
// Messages
// ---------------------------------------------------------------------------

function mapReactions(raw: Raw): MessageReaction[] {
	return (raw.reactions ?? []).map((r: Raw) => ({
		type: r.reactionType ?? "?",
		user: mapPerson(r.user)?.displayName,
		createdDateTime: r.createdDateTime ?? undefined,
	}));
}

function mapAttachments(raw: Raw): MessageAttachment[] {
	return (raw.attachments ?? []).map((a: Raw) => ({
		id: a.id ?? undefined,
		name: a.name ?? undefined,
		contentType: a.contentType ?? undefined,
		contentUrl: a.contentUrl ?? undefined,
	}));
}

function mapMentions(raw: Raw): PersonRef[] {
	return (raw.mentions ?? [])
		.map((m: Raw) => mapPerson(m.mentioned) ?? { displayName: m.mentionText ?? "(mention)" })
		.filter((p: PersonRef | undefined): p is PersonRef => !!p);
}

export function mapMessage(raw: Raw, location?: MessageLocation): MessageSummary {
	const contentType = raw.body?.contentType ?? "text";
	return {
		id: raw.id ?? "",
		text: htmlToText(raw.body?.content ?? "", contentType),
		contentType,
		from: mapPerson(raw.from),
		createdDateTime: raw.createdDateTime ?? undefined,
		lastModifiedDateTime: raw.lastModifiedDateTime ?? undefined,
		deletedDateTime: raw.deletedDateTime ?? undefined,
		subject: raw.subject ?? undefined,
		importance: raw.importance ?? undefined,
		replyToId: raw.replyToId ?? undefined,
		mentions: mapMentions(raw),
		reactions: mapReactions(raw),
		attachments: mapAttachments(raw),
		webUrl: raw.webUrl ?? undefined,
		location,
	};
}

// ---------------------------------------------------------------------------
// Presence
// ---------------------------------------------------------------------------

export function mapPresence(raw: Raw, displayName?: string): PresenceInfo {
	return {
		id: raw.id ?? "",
		displayName,
		availability: raw.availability ?? "Unknown",
		activity: raw.activity ?? "Unknown",
		statusMessage: raw.statusMessage?.message?.content
			? htmlToText(raw.statusMessage.message.content, raw.statusMessage.message.contentType ?? "text")
			: undefined,
	};
}

// ---------------------------------------------------------------------------
// Calendar
// ---------------------------------------------------------------------------

export function mapEvent(raw: Raw): EventSummary {
	return {
		id: raw.id ?? "",
		subject: raw.subject ?? "(no subject)",
		start: raw.start?.dateTime ?? undefined,
		end: raw.end?.dateTime ?? undefined,
		timeZone: raw.start?.timeZone ?? undefined,
		isAllDay: raw.isAllDay ?? undefined,
		isOnlineMeeting: raw.isOnlineMeeting ?? undefined,
		joinUrl: raw.onlineMeeting?.joinUrl ?? raw.onlineMeetingUrl ?? undefined,
		organizer: mapPerson(raw.organizer?.emailAddress
			? { displayName: raw.organizer.emailAddress.name, email: raw.organizer.emailAddress.address }
			: raw.organizer),
		attendees: (raw.attendees ?? [])
			.map((a: Raw) =>
				mapPerson({ displayName: a.emailAddress?.name, email: a.emailAddress?.address }),
			)
			.filter((p: PersonRef | undefined): p is PersonRef => !!p),
		location: raw.location?.displayName ?? undefined,
		bodyPreview: raw.bodyPreview ?? undefined,
		webLink: raw.webLink ?? undefined,
	};
}

// ---------------------------------------------------------------------------
// Files
// ---------------------------------------------------------------------------

export function mapDriveItem(raw: Raw): DriveItemSummary {
	return {
		id: raw.id ?? "",
		name: raw.name ?? "(unnamed)",
		size: raw.size ?? undefined,
		lastModifiedDateTime: raw.lastModifiedDateTime ?? undefined,
		lastModifiedBy: raw.lastModifiedBy?.user?.displayName ?? undefined,
		webUrl: raw.webUrl ?? undefined,
		isFolder: !!raw.folder,
	};
}
