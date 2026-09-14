/**
 * pi-teams extension entry point.
 *
 * Registers the tools, the status surfaces, the slash commands, and — most
 * importantly — the `tool_call` interceptor that enforces the safety level
 * before anything is said in the user's name.
 */

import type { ExtensionAPI } from "@earendil-works/pi-coding-agent";
import { Box, Text } from "@earendil-works/pi-tui";

import {
	ensureConfigTemplate,
	getConfigPath,
	tryResolveConnection,
	type TeamsConnection,
} from "../config/index.ts";
import { buildConnectionCard, buildConnectionLabel, formatStatusText, type ConnectionCard } from "../status.ts";
import {
	blockReason,
	formatMutationSummary,
	isMutationTool,
	LOCAL_CONFIG_TOOLS,
} from "../safety/index.ts";

import { teamsSetupTool } from "../tools/teams-setup.ts";
import { teamsLoginTool, teamsLogoutTool } from "../tools/teams-login.ts";
import { teamsAccountsTool } from "../tools/teams-accounts.ts";
import { teamsStatusTool } from "../tools/teams-status.ts";
import { teamsDoctorTool } from "../tools/teams-doctor.ts";
import { teamsPermissionsTool } from "../tools/teams-permissions.ts";
import { teamsListTeamsTool } from "../tools/teams-list-teams.ts";
import { teamsListChannelsTool } from "../tools/teams-list-channels.ts";
import { teamsListMembersTool } from "../tools/teams-list-members.ts";
import { teamsFindUserTool } from "../tools/teams-find-user.ts";
import { teamsListChatsTool } from "../tools/teams-list-chats.ts";
import { teamsReadChatTool } from "../tools/teams-read-chat.ts";
import { teamsSendChatMessageTool } from "../tools/teams-send-chat-message.ts";
import { teamsCreateChatTool } from "../tools/teams-create-chat.ts";
import { teamsReadChannelTool } from "../tools/teams-read-channel.ts";
import { teamsReadThreadTool } from "../tools/teams-read-thread.ts";
import {
	teamsReplyChannelMessageTool,
	teamsSendChannelMessageTool,
} from "../tools/teams-send-channel-message.ts";
import { teamsCreateChannelTool } from "../tools/teams-create-channel.ts";
import { teamsListFilesTool } from "../tools/teams-list-files.ts";
import { teamsSearchMessagesTool } from "../tools/teams-search-messages.ts";
import { teamsReactTool } from "../tools/teams-react.ts";
import { teamsDeleteMessageTool } from "../tools/teams-delete-message.ts";
import { teamsInboxTool } from "../tools/teams-inbox.ts";
import {
	teamsGetPresenceTool,
	teamsSetPresenceTool,
	teamsSetStatusMessageTool,
} from "../tools/teams-presence.ts";
import {
	teamsCancelMeetingTool,
	teamsCreateMeetingTool,
	teamsGetMeetingTool,
	teamsListMeetingsTool,
	teamsUpdateMeetingTool,
} from "../tools/teams-meetings.ts";

/* eslint-disable @typescript-eslint/no-explicit-any */

/** Every tool, in the order they appear in the documentation. */
const tools = [
	// Setup & session
	teamsSetupTool,
	teamsLoginTool,
	teamsLogoutTool,
	teamsAccountsTool,
	teamsStatusTool,
	teamsPermissionsTool,
	teamsDoctorTool,
	// Discovery
	teamsListTeamsTool,
	teamsListChannelsTool,
	teamsListMembersTool,
	teamsFindUserTool,
	// Chats
	teamsListChatsTool,
	teamsReadChatTool,
	teamsSendChatMessageTool,
	teamsCreateChatTool,
	// Channels
	teamsReadChannelTool,
	teamsReadThreadTool,
	teamsSendChannelMessageTool,
	teamsReplyChannelMessageTool,
	teamsCreateChannelTool,
	teamsListFilesTool,
	// Cross-cutting
	teamsSearchMessagesTool,
	teamsReactTool,
	teamsDeleteMessageTool,
	teamsInboxTool,
	// Presence
	teamsGetPresenceTool,
	teamsSetPresenceTool,
	teamsSetStatusMessageTool,
	// Calendar
	teamsListMeetingsTool,
	teamsGetMeetingTool,
	teamsCreateMeetingTool,
	teamsUpdateMeetingTool,
	teamsCancelMeetingTool,
];

export default function (pi: ExtensionAPI) {
	// The session's default connection. Re-resolved when the config changes, so
	// teams_setup takes effect without a restart.
	let connection: TeamsConnection | undefined;

	const refresh = (): TeamsConnection | undefined => {
		connection = tryResolveConnection();
		return connection;
	};

	// -----------------------------------------------------------------------
	// Status card — rendered in the transcript, never sent to the model
	// -----------------------------------------------------------------------

	pi.registerEntryRenderer<ConnectionCard>("teams-connection", (entry, { expanded }, theme) => {
		const card = entry.data;
		const box = new Box(1, 1, (text) => theme.bg("customMessageBg", text));
		box.addChild(new Text(theme.fg("accent", theme.bold("Microsoft Teams")), 0, 0));

		if (!card) {
			box.addChild(new Text(theme.fg("warning", "not configured"), 0, 0));
			return box;
		}

		const scope = card.tenant === card.account ? card.account : `${card.account}/${card.tenant}`;
		box.addChild(new Text(theme.fg("dim", `${scope}${card.user ? ` · ${card.user}` : ""}`), 0, 0));
		box.addChild(
			new Text(
				card.signedIn
					? theme.fg("muted", `safety: ${card.safetyLevel}`)
					: theme.fg("warning", "not signed in — run /teams-login"),
				0,
				0,
			),
		);
		if (expanded) {
			box.addChild(new Text(theme.fg("muted", formatStatusText(card)), 0, 0));
		}
		return box;
	});

	// -----------------------------------------------------------------------
	// Commands
	// -----------------------------------------------------------------------

	pi.registerCommand("teams-status", {
		description: "Show the current Microsoft Teams connection and permissions",
		handler: async (_args, ctx) => {
			const card = buildConnectionCard(refresh());
			if (!card) {
				ctx.ui.notify(`Teams: no configuration found. See ${getConfigPath()}`, "warning");
				return;
			}
			pi.appendEntry<ConnectionCard>("teams-connection", card);
			ctx.ui.setStatus("teams", ctx.ui.theme.fg("success", buildConnectionLabel(card)));
			ctx.ui.notify(buildConnectionLabel(card), "info");
		},
	});

	pi.registerCommand("teams-login", {
		description: "Sign in to Microsoft Teams as yourself (device code flow)",
		handler: async (args, ctx) => {
			const account = args.trim();
			if (!refresh()) {
				ctx.ui.notify(
					`Teams: no account configured yet. Use the teams_setup tool, or edit ${getConfigPath()}.`,
					"warning",
				);
				return;
			}
			pi.sendUserMessage(
				account
					? `Please use the teams_login tool to sign in to the Teams account "${account}".`
					: "Please use the teams_login tool to sign me in to Microsoft Teams.",
				{ deliverAs: "steer" },
			);
			ctx.ui.notify("Starting Teams sign-in…", "info");
		},
	});

	pi.registerCommand("teams-inbox", {
		description: "Show what needs your attention in Microsoft Teams",
		handler: async (_args, ctx) => {
			if (!refresh()) {
				ctx.ui.notify("Teams is not configured. Use the teams_setup tool first.", "error");
				return;
			}
			pi.sendUserMessage(
				"Please use teams_inbox to show me what needs my attention in Teams right now, and summarize it.",
				{ deliverAs: "steer" },
			);
			ctx.ui.notify("Checking Teams…", "info");
		},
	});

	pi.registerCommand("teams-doctor", {
		description: "Diagnose the Microsoft Teams configuration and sign-in",
		handler: async (_args, ctx) => {
			pi.sendUserMessage(
				"Please run the teams_doctor tool and summarize what is wrong and how to fix it.",
				{ deliverAs: "steer" },
			);
			ctx.ui.notify("Running Teams doctor…", "info");
		},
	});

	pi.registerCommand("teams-permissions", {
		description: "Show what pi is allowed to do in Microsoft Teams",
		handler: async (_args, ctx) => {
			pi.sendUserMessage(
				"Please run teams_permissions with action 'show' and explain in plain language what you may and may not do in Teams.",
				{ deliverAs: "steer" },
			);
			ctx.ui.notify("Reading Teams guardrails…", "info");
		},
	});

	// -----------------------------------------------------------------------
	// Session start
	// -----------------------------------------------------------------------

	pi.on("session_start", async (_event, ctx) => {
		const created = ensureConfigTemplate();
		refresh();

		if (!connection) {
			ctx.ui.notify(
				created
					? `@patimweb/pi-teams: a config template was created at ${getConfigPath()} — add your Entra ID tenantId and clientId, or run the teams_setup tool.`
					: `@patimweb/pi-teams: no usable Teams account configured. See ${getConfigPath()} or run the teams_setup tool.`,
				"warning",
			);
			return;
		}

		const card = buildConnectionCard(connection);
		if (!card) return;

		ctx.ui.setStatus("teams", ctx.ui.theme.fg(card.signedIn ? "success" : "warning", buildConnectionLabel(card)));

		if (!card.signedIn) {
			ctx.ui.notify(
				`@patimweb/pi-teams: account "${card.account}" is configured but not signed in. Run /teams-login.`,
				"warning",
			);
		} else {
			ctx.ui.notify(
				`@patimweb/pi-teams loaded (${card.account}${card.user ? ` as ${card.user}` : ""}, safety: ${card.safetyLevel})`,
				"info",
			);
		}

		// Survives /reload, so the transcript keeps saying who pi is acting as.
		const existing = ctx.sessionManager
			.getEntries()
			.some((entry: any) => entry.type === "custom" && entry.customType === "teams-connection");
		if (!existing) pi.appendEntry<ConnectionCard>("teams-connection", card);
	});

	// -----------------------------------------------------------------------
	// Tools
	// -----------------------------------------------------------------------

	for (const tool of tools) {
		pi.registerTool({
			name: tool.name,
			label: tool.name.replace(/_/g, " ").replace(/\b\w/g, (c) => c.toUpperCase()),
			description: tool.description,
			parameters: tool.parameters as any,
			promptSnippet: "promptSnippet" in tool ? (tool as any).promptSnippet : undefined,
			promptGuidelines: "promptGuidelines" in tool ? (tool as any).promptGuidelines : undefined,
			async execute(
				toolCallId: string,
				params: any,
				signal: AbortSignal | undefined,
				onUpdate: any,
				ctx: any,
			) {
				return (tool as any).execute(toolCallId, params, signal, onUpdate, {
					cwd: ctx.cwd,
					connection,
				});
			},
		});
	}

	// -----------------------------------------------------------------------
	// Safety interceptor
	// -----------------------------------------------------------------------

	pi.on("tool_call", async (event, ctx) => {
		if (!isMutationTool(event.toolName)) return;

		// A tool may target a different account than the session default, so the
		// rules are resolved against the connection that will actually be used.
		const params = (event.input ?? {}) as Record<string, unknown>;
		const target =
			tryResolveConnection(
				typeof params.account === "string" ? params.account : undefined,
				typeof params.tenant === "string" ? params.tenant : undefined,
			) ?? connection;

		if (!target) {
			// teams_setup has to work before anything is configured — that is what
			// it is for.
			if (LOCAL_CONFIG_TOOLS.has(event.toolName)) return;
			return {
				block: true,
				reason: `Microsoft Teams is not configured. Run the teams_setup tool, or edit ${getConfigPath()}.`,
			};
		}

		const blocked = blockReason(target.safetyLevel, target.authMode, event.toolName);
		if (blocked) return { block: true, reason: blocked };

		if (target.safetyLevel === "confirm") {
			const scope = target.tenant === target.account ? target.account : `${target.account}/${target.tenant}`;
			const summary = formatMutationSummary(event.toolName, params);
			const approved = await ctx.ui.confirm(
				"Microsoft Teams — acting as you",
				`${summary}\n\nAccount: ${scope}\n\nAllow this?`,
			);
			if (!approved) return { block: true, reason: `User declined: ${summary}` };
		}
	});
}
