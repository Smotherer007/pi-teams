/**
 * Domain data types.
 *
 * Plain, immutable interfaces. No behavior, no classes, no inheritance.
 * Everything the Graph layer returns is normalized into these shapes before
 * it reaches a formatter or a tool.
 */

// ---------------------------------------------------------------------------
// Identity
// ---------------------------------------------------------------------------

/** The signed-in user behind an access token. */
export interface SignedInUser {
	/** Object ID (oid claim) */
	id: string;
	/** Display name */
	displayName: string;
	/** User principal name / login */
	upn: string;
	/** Primary SMTP address, when known */
	mail?: string;
	/** Tenant ID (tid claim) */
	tenantId: string;
	/** Job title, when known */
	jobTitle?: string;
}

/** A person referenced in a message, member list, or search result. */
export interface PersonRef {
	id?: string;
	displayName: string;
	upn?: string;
	mail?: string;
}

// ---------------------------------------------------------------------------
// Teams & channels
// ---------------------------------------------------------------------------

export interface TeamSummary {
	id: string;
	displayName: string;
	description?: string;
	visibility?: string;
	isArchived?: boolean;
	webUrl?: string;
}

export interface ChannelSummary {
	id: string;
	displayName: string;
	description?: string;
	/** "standard" | "private" | "shared" */
	membershipType?: string;
	webUrl?: string;
	/** Owning team, filled in by the caller for display + scope checks */
	teamId?: string;
	teamName?: string;
}

export interface MemberSummary {
	id: string;
	displayName: string;
	upn?: string;
	mail?: string;
	roles: string[];
}

// ---------------------------------------------------------------------------
// Chats
// ---------------------------------------------------------------------------

export interface ChatSummary {
	id: string;
	/** "oneOnOne" | "group" | "meeting" */
	chatType: string;
	/** Explicit topic for group chats, otherwise derived from members */
	topic?: string;
	/** Human label used for display and scope matching */
	label: string;
	members: PersonRef[];
	lastUpdated?: string;
	lastMessagePreview?: string;
	lastMessageFrom?: string;
	webUrl?: string;
}

export interface ChatMemberSummary {
	/** conversationMember id — the identifier DELETE /chats/{id}/members/{id} needs */
	membershipId: string;
	/** Object id of the person behind the membership */
	userId?: string;
	displayName: string;
	upn?: string;
	mail?: string;
	roles: string[];
}

// ---------------------------------------------------------------------------
// Messages
// ---------------------------------------------------------------------------

export interface MessageSummary {
	id: string;
	/** Plain-text rendering of the message body */
	text: string;
	/** Original content type: "text" | "html" */
	contentType: string;
	from?: PersonRef;
	createdDateTime?: string;
	lastModifiedDateTime?: string;
	/** Set when the message was soft-deleted */
	deletedDateTime?: string;
	subject?: string;
	importance?: string;
	/** Parent message id for channel replies */
	replyToId?: string;
	/** Number of replies, when the caller resolved them */
	replyCount?: number;
	mentions: PersonRef[];
	reactions: MessageReaction[];
	attachments: MessageAttachment[];
	webUrl?: string;
	/** Where the message lives — used for follow-up calls and scope checks */
	location?: MessageLocation;
}

export interface MessageReaction {
	type: string;
	user?: string;
	createdDateTime?: string;
}

export interface MessageAttachment {
	id?: string;
	name?: string;
	contentType?: string;
	contentUrl?: string;
}

export interface MessageLocation {
	kind: "chat" | "channel";
	chatId?: string;
	teamId?: string;
	teamName?: string;
	channelId?: string;
	channelName?: string;
}

// ---------------------------------------------------------------------------
// Presence
// ---------------------------------------------------------------------------

export interface PresenceInfo {
	id: string;
	displayName?: string;
	availability: string;
	activity: string;
	statusMessage?: string;
}

// ---------------------------------------------------------------------------
// Calendar & meetings
// ---------------------------------------------------------------------------

export interface EventSummary {
	id: string;
	subject: string;
	start?: string;
	end?: string;
	timeZone?: string;
	isAllDay?: boolean;
	isOnlineMeeting?: boolean;
	joinUrl?: string;
	organizer?: PersonRef;
	attendees: PersonRef[];
	location?: string;
	bodyPreview?: string;
	webLink?: string;
}

export interface OnlineMeetingSummary {
	id: string;
	subject?: string;
	joinUrl: string;
	start?: string;
	end?: string;
}

// ---------------------------------------------------------------------------
// Files
// ---------------------------------------------------------------------------

export interface DriveItemSummary {
	id: string;
	name: string;
	size?: number;
	lastModifiedDateTime?: string;
	lastModifiedBy?: string;
	webUrl?: string;
	isFolder: boolean;
}
