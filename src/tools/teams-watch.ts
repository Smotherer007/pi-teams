/**
 * teams_watch — switch listen mode on and off, and report what it is doing.
 *
 * The actual polling lives in the extension, because only it can inject a
 * message into the session and update the status line. This tool owns the
 * configuration and the explanation; the extension restarts the loop when the
 * config changes.
 */

import { Type } from "typebox";
import {
	getConfigPath,
	resolveConnection,
	setWatchConfig,
	type MentionOnlyConfig,
	type ResolvedMentionOnly,
	type TeamsConnection,
	type WatchConfig,
	type ResolvedDispatchConfig,
} from "../config/index.ts";
import { getActiveWatchLoop, type WatchRuntimeStatus } from "../watch/loop.ts";
import { getWatchCursorPath } from "../watch/cursor.ts";
import { errorResult, run, textResult, type ToolContext, type ToolResult } from "./shared.ts";

interface WatchParams {
	action?: string;
	account?: string;
	tenant?: string;
	intervalSeconds?: number;
	chats?: string[];
	from?: string[];
	mentionOnly?: boolean | MentionOnlyConfig;
	cooldownSeconds?: number;
	maxTriggersPerHour?: number;
}

/**
 * What the cursor means, in the status output.
 *
 * The single most surprising thing about listen mode is what happens when it is
 * switched on again, so the answer is part of the status rather than only in
 * the docs.
 */
const CURSOR_NOTES = [
	"### What wakes pi",
	"",
	"A chat wakes pi when both are true: it has moved since pi last looked (a cursor per account, kept "
		+ "on disk), and its newest message is still unread for you in Teams. Switching listen mode back "
		+ "on therefore answers what you missed, and a chat you have already read in Teams stays quiet.",
	"",
	"A backlog is drained over several ticks and capped by the hourly wake limit, so a long absence does "
		+ "not produce a burst of answers. A chat held back by its cooldown or by that limit is delayed, "
		+ "never dropped: it stays open and is answered as soon as the wait is over.",
].join("\n");

/** Renders the mention-only switch, which is a default plus optional overrides. */
function describeMentionOnly(mentionOnly: ResolvedMentionOnly): string {
	const rules = [
		...mentionOnly.chats.map((rule) => `chat \`${rule.pattern}\` → ${rule.value}`),
		...mentionOnly.people.map((rule) => `person \`${rule.pattern}\` → ${rule.value}`),
	];
	const overrides = rules.length > 0 ? ` (${rules.join(", ")})` : "";
	return `- mentions only: ${mentionOnly.default}${overrides}`;
}

/** Renders the effective settings the same way the config file would. */
function describe(watch: {
	enabled: boolean;
	autoStart: boolean;
	intervalSeconds: number;
	chats: string[];
	from: string[];
	mentionOnly: ResolvedMentionOnly;
	cooldownSeconds: number;
	maxTriggersPerHour: number;
	dispatch?: ResolvedDispatchConfig;
}): string {
	return [
		`- enabled: ${watch.enabled}`,
		`- starts with a session: ${watch.autoStart}`,
		`- poll every: ${watch.intervalSeconds} s`,
		`- chats: ${watch.chats.length > 0 ? watch.chats.map((c) => `\`${c}\``).join(", ") : "every chat with recent activity"}`,
		`- people: ${watch.from.length > 0 ? watch.from.map((p) => `\`${p}\``).join(", ") : "any sender"}`,
		describeMentionOnly(watch.mentionOnly),
		`- cooldown per chat: ${watch.cooldownSeconds} s`,
		`- wake limit: ${watch.maxTriggersPerHour} per hour`,
		watch.dispatch?.mode === "process"
			? `- dispatch: one pi process per chat, ${watch.dispatch.maxConcurrent} in parallel`
			: "- dispatch: every wake is a turn in this session",
	].join("\n");
}

/**
 * The live facts about a running watcher, for the status action.
 *
 * Distinguishes "configured on" from "actually polling": after an edit to the
 * config file the two can differ until the extension picks it up.
 */
function runtimeLines(runtime: WatchRuntimeStatus | undefined, conn: TeamsConnection): string[] {
	if (!runtime) {
		return [
			"Nothing is polling in this session.",
			"",
			"Listen mode starts only when it is switched on here — `/teams-listen on`, or "
				+ "`teams_watch action: enable`. A stored `enabled: true` does not start it by itself: "
				+ "set `watch.autoStart` if a session is meant to start listening on its own.",
		];
	}

	const lines = [
		"### Running in this session",
		"",
		`- running: ${runtime.running}`,
		`- chats tracked: ${runtime.trackedChats}`,
		`- woken this hour: ${runtime.wakesThisHour}`,
	];
	if (runtime.waiting > 0) {
		lines.push(
			`- waiting: ${runtime.waiting} chat(s) — held back by their cooldown or by the wake limit, "
				+ "and answered once that clears`,
		);
	}
	if (runtime.lastTickAt) {
		lines.push(`- last poll: ${new Date(runtime.lastTickAt).toLocaleTimeString()}`);
	}
	if (runtime.lastWakeChat) {
		lines.push(`- last wake: ${runtime.lastWakeChat}`);
	}
	if (runtime.catchUp !== undefined) {
		const read = runtime.catchUpRead ? `, ${runtime.catchUpRead} of them already read` : "";
		lines.push(`- catch-up on start: ${runtime.catchUp} chat(s) had moved${read}`);
	}
	if (runtime.lastError) {
		lines.push(`- ⚠️ last error: ${runtime.lastError}`);
	}

	lines.push(`- cursor: \`${getWatchCursorPath(conn.account, conn.tenantId)}\``);
	return lines;
}

export const teamsWatchTool = {
	name: "teams_watch",
	description:
		"Control listen mode: pi polls your Teams chats and turns an incoming message into a prompt it answers. " +
		"A chat wakes it when it has moved since pi last looked, so switching listen mode on also answers "
			+ "what arrived while pi was not running. " +
		"Use action 'status' to see the current settings, 'enable' or 'disable' to switch it, and the optional " +
		"fields to narrow what may wake pi (chat patterns, mentions only, interval, cooldown, hourly limit). " +
		"The watcher runs in the extension session; this tool changes the configuration it reads.",
	parameters: Type.Object({
		action: Type.Optional(
			Type.String({ description: "'status' (default), 'enable' or 'disable'" }),
		),
		account: Type.Optional(
			Type.String({ description: "Account the settings apply to; omit to set them globally" }),
		),
		tenant: Type.Optional(
			Type.String({
				description:
					"Tenant beneath that account. Used by action 'status' to report the settings in force for it — listen mode itself is configured per account, not per tenant.",
			}),
		),
		intervalSeconds: Type.Optional(
			Type.Number({ description: "Seconds between polls (minimum 15, default 60)" }),
		),
		chats: Type.Optional(
			Type.Array(Type.String(), {
				description:
					"Glob patterns for the chats to watch — matched against topic, label, chat ID and participant names",
			}),
		),
		from: Type.Optional(
			Type.Array(Type.String(), {
				description:
					"Glob patterns for the people pi listens to — matched against display name, UPN and e-mail. Omit for any sender.",
			}),
		),
		mentionOnly: Type.Optional(
			Type.Union(
				[
					Type.Boolean(),
					Type.Object({
						default: Type.Optional(Type.Boolean()),
						chats: Type.Optional(Type.Record(Type.String(), Type.Boolean())),
						people: Type.Optional(Type.Record(Type.String(), Type.Boolean())),
					}),
				],
				{
					description:
						"Whether a message must address you before it wakes pi. `true`/`false` applies everywhere; "
						+ "`{ default, chats, people }` adds per-chat and per-person overrides, and a rule may be false "
						+ "to exempt someone who then does not have to mention you. 1:1 chats always count as addressed. "
						+ "First matching pattern wins.",
				},
			),
		),
		cooldownSeconds: Type.Optional(
			Type.Number({ description: "Seconds to stay quiet in a chat after waking pi for it (default 300)" }),
		),
		maxTriggersPerHour: Type.Optional(
			Type.Number({ description: "Hard cap on wakes per hour (default 10)" }),
		),
	}),
	promptSnippet: "Turn Teams listen mode on or off",
	promptGuidelines: [
		"Use teams_watch when the user wants pi to notice incoming Teams messages on its own.",
		"Tell the user plainly that listen mode spends a model turn per incoming message, and that pi answers in their name.",
		"Enabling listen mode takes effect immediately; the settings survive a restart.",
		"Enabling listen mode answers what arrived since the last run, so say how much is waiting rather than implying only new messages count.",
		"`mentionOnly` takes a boolean or `{ default, chats, people }`. A 1:1 chat already counts as a direct address, so the overrides are for group and meeting chats — for example a global `true` with one person set to `false`, so that person never has to mention you.",
	],

	async execute(
		_toolCallId: string,
		params: WatchParams,
		_signal: AbortSignal | undefined,
		_onUpdate: undefined,
		_ctx: ToolContext,
	): Promise<ToolResult> {
		return run(async () => {
			const action = (params.action ?? "status").toLowerCase();

			// Reported from the resolved connection, so the answer shows what is in
			// force for that account and tenant — not just what the file says.
			const inspect = (): TeamsConnection | undefined => {
				try {
					return resolveConnection(params.account, params.tenant);
				} catch {
					return undefined;
				}
			};

			const patch: WatchConfig = {};
			if (params.intervalSeconds !== undefined) patch.intervalSeconds = params.intervalSeconds;
			if (params.chats !== undefined) patch.chats = params.chats;
		if (params.from !== undefined) patch.from = params.from;
			if (params.mentionOnly !== undefined) patch.mentionOnly = params.mentionOnly;
			if (params.cooldownSeconds !== undefined) patch.cooldownSeconds = params.cooldownSeconds;
			if (params.maxTriggersPerHour !== undefined) {
				patch.maxTriggersPerHour = params.maxTriggersPerHour;
			}

			switch (action) {
				case "status": {
					const conn = inspect();
					if (!conn) {
						return errorResult(
							`No usable Teams account for these settings. Check ${getConfigPath()}.`,
						);
					}
					const watch = conn.watch;
					const loop = getActiveWatchLoop();
					const runtime = loop?.status();
					return textResult(
						[
							`## Listen mode — ${watch.enabled ? "on" : "off"}`,
							"",
							describe(watch),
							"",
							...runtimeLines(runtime, conn),
							"",
							CURSOR_NOTES,
						].join("\n"),
						{ watch, runtime },
					);
				}

				case "enable": {
					const watch = setWatchConfig({ ...patch, enabled: true }, params.account);
					return textResult(
						[
							"## Listen mode — on",
							"",
							describe(watch),
							"",
							"pi now polls these chats and treats anything that has moved since its last look as a prompt — " +
								"including what arrived while it was not running, so the first tick may answer several messages. " +
								"It answers in your name and under the safety level in force, so anything it decides to send " +
								"still follows the usual confirmation rules.",
							"",
							`Written to ${getConfigPath()}.`,
						].join("\n"),
						{ watch },
					);
				}

				case "disable": {
					const watch = setWatchConfig({ ...patch, enabled: false }, params.account);
					return textResult(
						[
							"## Listen mode — off",
							"",
							describe(watch),
							"",
							"Polling stops now; the settings stay in the file for the next time.",
						].join("\n"),
						{ watch },
					);
				}

				default:
					return errorResult(`Unknown action "${action}". Use 'status', 'enable' or 'disable'.`);
			}
		});
	},
};
