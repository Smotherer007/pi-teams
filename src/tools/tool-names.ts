/**
 * Tool names in one place.
 *
 * The safety interceptor and the tool modules both need to agree on what
 * counts as a mutation; keeping the set here (instead of in shared.ts) avoids
 * an import cycle between safety/ and tools/.
 */

/** Every tool that changes something in Teams, Outlook, or the local config. */
export const MUTATION_TOOLS = new Set<string>([
	// Messaging
	"teams_send_chat_message",
	"teams_send_channel_message",
	"teams_reply_channel_message",
	"teams_create_chat",
	"teams_delete_message",
	"teams_react",
	"teams_mark_read",
	"teams_update_message",
	"teams_chat_members",
	// Presence
	"teams_set_presence",
	"teams_set_status_message",
	// Calendar
	"teams_create_meeting",
	"teams_update_meeting",
	"teams_respond_invite",
	"teams_cancel_meeting",
	// Local configuration and credentials
	"teams_setup",
	"teams_logout",
	"teams_watch",
]);
