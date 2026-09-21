/**
 * Safety gates and the message bodies that pass through them.
 */

import { test, describe } from "node:test";
import assert from "node:assert/strict";
import {
	blockReason,
	formatMutationSummary,
	isMutationTool,
	LOCAL_CONFIG_TOOLS,
} from "../src/safety/index.ts";
import { buildMessageBody, messagePath } from "../src/graph/messages.ts";
import { MUTATION_TOOLS } from "../src/tools/tool-names.ts";

describe("mutation classification", () => {
	test("writes are mutations, reads are not", () => {
		assert.equal(isMutationTool("teams_send_chat_message"), true);
		assert.equal(isMutationTool("teams_send_channel_message"), true);
		assert.equal(isMutationTool("teams_create_meeting"), true);
		assert.equal(isMutationTool("teams_read_chat"), false);
		assert.equal(isMutationTool("teams_list_teams"), false);
		assert.equal(isMutationTool("teams_search_messages"), false);
	});

	test("every mutation tool name is prefixed and unique", () => {
		for (const name of MUTATION_TOOLS) {
			assert.match(name, /^teams_/);
		}
	});
});

describe("blockReason", () => {
	test("readonly blocks every mutation", () => {
		const reason = blockReason("readonly", "device-code", "teams_send_chat_message");
		assert.match(reason ?? "", /readonly/);
	});

	test("readonly does not block reads", () => {
		assert.equal(blockReason("readonly", "device-code", "teams_read_chat"), undefined);
	});

	test("confirm and open let mutations through to the confirmation step", () => {
		assert.equal(blockReason("confirm", "interactive", "teams_send_chat_message"), undefined);
		assert.equal(blockReason("open", "device-code", "teams_send_chat_message"), undefined);
	});

	test("both delegated flows may write — only app-only is special", () => {
		assert.equal(blockReason("open", "interactive", "teams_send_channel_message"), undefined);
		assert.equal(blockReason("open", "device-code", "teams_send_channel_message"), undefined);
	});

	test("app-only tokens cannot post as a person", () => {
		const reason = blockReason("open", "client-credentials", "teams_send_channel_message");
		assert.match(reason ?? "", /app-only/);
		assert.match(reason ?? "", /interactive/);
	});

	test("app-only may still manage the calendar", () => {
		assert.equal(blockReason("open", "client-credentials", "teams_create_meeting"), undefined);
	});

	test("readonly does not lock the user out of their own configuration", () => {
		// Otherwise a user who set readonly could never add an account or sign out.
		assert.equal(blockReason("readonly", "device-code", "teams_setup"), undefined);
		assert.equal(blockReason("readonly", "device-code", "teams_logout"), undefined);
	});

	test("config tools are still mutations, so they still ask for confirmation", () => {
		assert.equal(isMutationTool("teams_setup"), true);
		assert.equal(LOCAL_CONFIG_TOOLS.has("teams_setup"), true);
	});
});

describe("messagePath", () => {
	test("the chat form uses an explicit user id — Graph rejects /me here", () => {
		const path = messagePath("19:abc", "1700", "user-1");
		assert.equal(path, "/users/user-1/chats/19%3Aabc/messages/1700");
	});

	test("every segment is escaped, so a chat id cannot break out of the path", () => {
		assert.equal(
			messagePath("19:abc@thread.v2", "1700/../evil", "user-1"),
			"/users/user-1/chats/19%3Aabc%40thread.v2/messages/1700%2F..%2Fevil",
		);
	});
});

describe("formatMutationSummary", () => {
	test("shows the destination and the actual text", () => {
		const summary = formatMutationSummary("teams_send_channel_message", {
			team: "Engineering",
			channel: "General",
			body: "Deployment is done.",
		});
		assert.match(summary, /Engineering\/General/);
		assert.match(summary, /Deployment is done\./);
	});

	test("truncates a long body instead of flooding the dialog", () => {
		const summary = formatMutationSummary("teams_send_chat_message", {
			chat: "Anna",
			body: "x".repeat(1000),
		});
		assert.ok(summary.length < 600);
		assert.match(summary, /…/);
	});

	test("falls back to a generic line for an unknown tool", () => {
		assert.match(formatMutationSummary("teams_future_tool", { a: 1 }), /teams_future_tool/);
	});
});

describe("buildMessageBody", () => {
	test("escapes plain text into HTML", () => {
		const body = buildMessageBody({ body: "a <b> c" });
		assert.deepEqual(body.body, { contentType: "html", content: "<p>a &lt;b&gt; c</p>" });
	});

	test("passes raw HTML through when asked", () => {
		const body = buildMessageBody({ body: "<p>raw</p>", html: true });
		assert.equal((body.body as { content: string }).content, "<p>raw</p>");
	});

	test("replaces @Name in the text with a real mention tag", () => {
		const body = buildMessageBody({
			body: "Hi @Anna Schmidt, please review.",
			mentions: [{ id: "u1", displayName: "Anna Schmidt" }],
		});
		const content = (body.body as { content: string }).content;
		assert.match(content, /<at id="0">Anna Schmidt<\/at>/);
		assert.ok(!content.includes("@Anna Schmidt"));
		assert.equal((body.mentions as unknown[]).length, 1);
	});

	test("prepends the mention when the body does not reference it", () => {
		const body = buildMessageBody({
			body: "please review",
			mentions: [{ id: "u1", displayName: "Anna Schmidt" }],
		});
		assert.match((body.body as { content: string }).content, /^<at id="0">Anna Schmidt<\/at> /);
	});

	test("carries subject and importance when given", () => {
		const body = buildMessageBody({ body: "x", subject: "Release", importance: "high" });
		assert.equal(body.subject, "Release");
		assert.equal(body.importance, "high");
	});

	test("omits the mentions array entirely when there are none", () => {
		assert.equal("mentions" in buildMessageBody({ body: "x" }), false);
	});

	test("attaches images as hosted contents the body points at", () => {
		const body = buildMessageBody({
			body: "Hier das Diagramm",
			images: [{ name: "diagramm.png", contentType: "image/png", contentBytes: "AAA=" }],
		});

		const content = (body.body as { content: string }).content;
		assert.match(content, /<img src="\.\.\/hostedContents\/1\/\$value" alt="diagramm\.png">/);
		assert.deepEqual(body.hostedContents, [
			{
				"@microsoft.graph.temporaryId": "1",
				contentBytes: "AAA=",
				contentType: "image/png",
			},
		]);
	});

	test("numbers several images so body and hostedContents agree", () => {
		const body = buildMessageBody({
			body: "zwei Bilder",
			images: [
				{ name: "a.png", contentType: "image/png", contentBytes: "AAA=" },
				{ name: "b.jpg", contentType: "image/jpeg", contentBytes: "BBB=" },
			],
		});

		const content = (body.body as { content: string }).content;
		assert.match(content, /hostedContents\/1\/\$value/);
		assert.match(content, /hostedContents\/2\/\$value/);

		const hosted = body.hostedContents as Record<string, unknown>[];
		assert.deepEqual(
			hosted.map((entry) => entry["@microsoft.graph.temporaryId"]),
			["1", "2"],
		);
	});

	test("omits hostedContents entirely when there are no images", () => {
		assert.equal("hostedContents" in buildMessageBody({ body: "x" }), false);
	});
});

describe("formatMutationSummary with images", () => {
	test("names the images that go out with the message", () => {
		const summary = formatMutationSummary("teams_send_chat_message", {
			chat: "Holodeck",
			body: "Hier das Bild",
			images: ["/tmp/screenshots/diagramm.png", "/tmp/zwei.jpg"],
		});

		assert.match(summary, /2 images: diagramm\.png, zwei\.jpg/);
	});

	test("leaves the summary alone when nothing is attached", () => {
		const summary = formatMutationSummary("teams_send_chat_message", {
			chat: "Holodeck",
			body: "Nur Text",
		});

		assert.doesNotMatch(summary, /image/);
	});
});
