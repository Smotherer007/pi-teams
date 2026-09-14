/**
 * teams_find_user — look someone up in the directory.
 */

import { Type } from "typebox";
import { findUsers } from "../graph/me.ts";
import { hasAccess } from "../safety/index.ts";
import {
	AccountParam,
	LimitParam,
	TenantParam,
	connectionFor,
	run,
	textResult,
	type ToolContext,
	type ToolResult,
} from "./shared.ts";

export const teamsFindUserTool = {
	name: "teams_find_user",
	description:
		"Find people in the organization's directory by name, e-mail or user principal name. " +
		"Use this to get the exact identity before starting a chat, @-mentioning someone, or " +
		"inviting them to a meeting.",
	parameters: Type.Object({
		query: Type.String({ description: "Name, e-mail or UPN to search for" }),
		account: AccountParam,
		tenant: TenantParam,
		limit: LimitParam,
	}),
	promptSnippet: "Find a person in the Teams directory",
	promptGuidelines: [
		"Use teams_find_user when the user refers to a colleague by first name only and an exact identity is needed.",
	],

	async execute(
		_toolCallId: string,
		params: { query: string; account?: string; tenant?: string; limit?: number },
		signal: AbortSignal | undefined,
		_onUpdate: undefined,
		ctx: ToolContext,
	): Promise<ToolResult> {
		return run(async () => {
			const conn = connectionFor(ctx, params.account, params.tenant);
			const people = await findUsers(conn, params.query, params.limit ?? 15, signal);

			const visible = people.filter((person) =>
				hasAccess(conn, "read", "people", [person.displayName, person.upn, person.mail, person.id]),
			);

			if (visible.length === 0) {
				return textResult(`No directory match for "${params.query}".`, { count: 0 });
			}

			const lines = [`Found ${visible.length} person(s) for "${params.query}":`, ""];
			for (const person of visible) {
				lines.push(`- **${person.displayName}**${person.upn ? ` <${person.upn}>` : ""}`);
				if (person.id) lines.push(`  userId: ${person.id}`);
			}

			return textResult(lines.join("\n"), { count: visible.length, people: visible });
		});
	},
};
