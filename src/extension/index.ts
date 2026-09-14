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
	setWatchConfig,
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
import {
	setActiveWatchLoop,
	startWatchLoop,
	stopActiveWatchLoop,
	type WatchLoop,
} from "../watch/loop.ts";
import { composeWatchPrompt } from "../watch/prompt.ts";

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
import { teamsListFilesTool } from "../tools/teams-list-files.ts";
import { teamsSearchMessagesTool } from "../tools/teams-search-messages.ts";
import { teamsReactTool } from "../tools/teams-react.ts";
import { teamsMarkReadTool } from "../tools/teams-mark-read.ts";
import { teamsUpdateMessageTool } from "../tools/teams-update-message.ts";
import { teamsChatMembersTool } from "../tools/teams-chat-members.ts";
import { teamsAvailabilityTool } from "../tools/teams-availability.ts";
import { teamsDeleteMessageTool } from "../tools/teams-delete-message.ts";
import { teamsInboxTool } from "../tools/teams-inbox.ts";
import { teamsWatchTool } from "../tools/teams-watch.ts";
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
	teamsRespondInviteTool,
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
	teamsWatchTool,
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
	teamsMarkReadTool,
	teamsChatMembersTool,
	// Channels
	teamsReadChannelTool,
	teamsReadThreadTool,
	teamsSendChannelMessageTool,
	teamsReplyChannelMessageTool,
	teamsListFilesTool,
	// Cross-cutting
	teamsSearchMessagesTool,
	teamsReactTool,
	teamsUpdateMessageTool,
	teamsDeleteMessageTool,
	teamsInboxTool,
	// Presence
	teamsGetPresenceTool,
	teamsSetPresenceTool,
	teamsSetStatusMessageTool,
	// Calendar
	teamsListMeetingsTool,
	teamsGetMeetingTool,
	teamsAvailabilityTool,
	teamsCreateMeetingTool,
	teamsUpdateMeetingTool,
	teamsRespondInviteTool,
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
	// Listen mode
	// -----------------------------------------------------------------------

	/** The watcher running in this session, when listen mode is on. */
	let watchLoop: WatchLoop | undefined;

	/**
	 * The settings the running watcher was started from.
	 *
	 * Used to tell a real config change from a tool that merely *read* the
	 * settings. Restarting resets the watcher's baseline, which discards every
	 * chat it has already looked at and swallows whatever arrived just before —
	 * so a `teams_watch status` must not be treated like an edit.
	 */
	let watchSignature: string | undefined;

	/** The last polling error the user was told about, so it is said once. */
	let reportedWatchError: string | undefined;

	const stopWatch = () => {
		watchLoop?.stop();
		watchLoop = undefined;
		watchSignature = undefined;
		setActiveWatchLoop(undefined);
	};

	/**
	 * Start the watcher for the session's account, if the config asks for it.
	 *
	 * Restarting is the same call: the old loop is stopped first, so a config
	 * change cannot leave two pollers running.
	 */
	const startWatch = (ctx: any) => {
		const conn = refresh();

		// Already watching this exact configuration: leave the loop alone, so its
		// baseline and its cooldowns survive a look at the status.
		const signature = conn?.watch.enabled ? JSON.stringify(conn.watch) : undefined;
		if (watchSignature !== undefined && signature === watchSignature && watchLoop) return;

		stopWatch();

		if (!conn?.watch.enabled) return;

		reportedWatchError = undefined;
		const loop = startWatchLoop({
			connection: conn,
			onWake: (event) => {
				pi.sendUserMessage(composeWatchPrompt(event, event.me), { deliverAs: "followUp" });
			},
			onTick: (status) => {
				if (!status.lastError) reportedWatchError = undefined;
			},
			onError: (message) => {
				// A watcher that fails every tick would otherwise say so every minute.
				if (message === reportedWatchError) return;
				reportedWatchError = message;
				ctx.ui.notify(`Teams listen mode: ${message}`, "warning");
			},
		});

		watchLoop = loop;
		watchSignature = signature;
		setActiveWatchLoop(loop);
	};

	/**
	 * Paint the footer status line from the current connection.
	 *
	 * The only place the status line is written, so the three surfaces —
	 * session start, `/teams-status`, and a tool result — cannot drift apart.
	 */
	const paintStatus = (ctx: any) => {
		const card = buildConnectionCard(refresh());
		if (!card) return;
		const watching = !!watchLoop?.status().running;
		ctx.ui.setStatus(
			"teams",
			ctx.ui.theme.fg(
				card.signedIn ? "success" : "warning",
				buildConnectionLabel({ ...card, watching }),
			),
		);
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
			paintStatus(ctx);
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

	pi.registerCommand("teams-listen", {
		description: "Let pi listen to Teams: on, off, or status (default)",
		handler: async (args, ctx) => {
			const conn = refresh();
			if (!conn) {
				ctx.ui.notify(
					`Teams: no account configured yet. Use the teams_setup tool, or edit ${getConfigPath()}.`,
					"warning",
				);
				return;
			}

			const action = args.trim().toLowerCase() || "status";

			if (action === "on" || action === "off") {
				// Written to the account, not globally: the watcher listens as one
				// identity, and "which account" must not be a guess.
				const watch = setWatchConfig({ enabled: action === "on" }, conn.account);
				startWatch(ctx);
				paintStatus(ctx);
				ctx.ui.notify(
					watch.enabled
						? `Teams listen mode on — polling every ${watch.intervalSeconds} s, max ${watch.maxTriggersPerHour} wakes/hour.`
						: "Teams listen mode off.",
					"info",
				);
				return;
			}

			if (action !== "status") {
				ctx.ui.notify(`Unknown argument "${action}". Use on, off, or status.`, "warning");
				return;
			}

			const runtime = watchLoop?.status();
			pi.appendEntry<ConnectionCard>("teams-connection", { ...buildConnectionCard(conn)!, watching: !!runtime?.running });
			ctx.ui.notify(
				[
					`Listen mode: ${conn.watch.enabled ? "on" : "off"}`,
					`every ${conn.watch.intervalSeconds} s`,
					conn.watch.chats.length > 0 ? `chats: ${conn.watch.chats.join(", ")}` : "all recent chats",
					conn.watch.from.length > 0 ? `people: ${conn.watch.from.join(", ")}` : "any sender",
					conn.watch.mentionOnly ? "mentions only" : "every message",
					runtime ? `${runtime.wakesThisHour} wake(s) this hour` : "not running in this session",
				].join(" · "),
				"info",
			);
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

		paintStatus(ctx);
		startWatch(ctx);

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
	// Listen mode lifecycle
	// -----------------------------------------------------------------------

	pi.on("session_shutdown", () => {
		// The timer would be unref'd anyway, but a poll in flight during shutdown
		// would still be a Graph call for a session that is gone.
		stopWatch();
	});

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

	// -----------------------------------------------------------------------
	// Status after a tool ran
	// -----------------------------------------------------------------------

	/**
	 * The status line is a snapshot of session start — but every change to the
	 * connection happens afterwards, in a tool, and tools have no UI context of
	 * their own. Without this the footer would still read "not signed in" after a
	 * successful login, and "listening" after listen mode was switched off.
	 */
	const STATUS_TOOLS = new Set([
		"teams_login",
		"teams_logout",
		"teams_setup",
		"teams_accounts",
		"teams_watch",
	]);

	/** Tools after which the watcher has to be re-read from the config. */
	const WATCH_CONFIG_TOOLS = new Set(["teams_watch", "teams_setup", "teams_accounts"]);

	pi.on("tool_result", async (event, ctx) => {
		if (!STATUS_TOOLS.has(event.toolName) || event.isError) return;

		if (event.toolName === "teams_logout") {
			// A watcher polling a session that was just removed only produces
			// errors; leaving it running would be noise, not service.
			stopWatch();
		} else if (WATCH_CONFIG_TOOLS.has(event.toolName)) {
			startWatch(ctx);
		}

		paintStatus(ctx);
	});
}
