/**
 * Markdown-ish text → the HTML subset Microsoft Teams renders in a message.
 *
 * Why this exists: a message body is HTML, but the model writes markdown. Sent
 * raw, `**wichtig**` arrives as literal asterisks and `- Punkt` as literal
 * dashes — which is exactly the "wall of text" a person sees in a chat when
 * nothing was formatted for the medium.
 *
 * Teams supports far less HTML than a browser. This translates only what
 * survives in a chat bubble:
 *
 *   `**bold**` → `<b>`            `- item` → `<ul><li>`
 *   `*italic*` / `_italic_` → `<i>`   `1. item` → `<ol><li>`
 *   `` `code` `` → `<code>`        ``` fences → `<pre>`
 *   `~~struck~~` → `<s>`          `# Heading` → a bold line
 *   `[label](url)` → `<a href>`
 *
 * Tables and images are deliberately unsupported: Teams renders neither in a
 * chat, so pretending otherwise would produce worse output than plain text.
 * Everything else is escaped and kept as written.
 */

const HTML_ESCAPES: Record<string, string> = {
	"&": "&amp;",
	"<": "&lt;",
	">": "&gt;",
};

function escapeHtml(text: string): string {
	return text.replace(/[&<>]/g, (char) => HTML_ESCAPES[char] ?? char);
}

/**
 * Inline markup of a single line.
 *
 * Order matters: code spans are protected first (so `**` inside them stays
 * literal), then links, then emphasis. The underscore rule requires a boundary
 * on both sides so `snake_case_name` survives untouched.
 */
function inline(raw: string): string {
	let text = escapeHtml(raw);

	text = text.replace(/`([^`]+)`/g, "<code>$1</code>");
	text = text.replace(
		/\[([^\]]+)\]\((https?:\/\/[^\s)]+|mailto:[^\s)]+)\)/g,
		'<a href="$2">$1</a>',
	);
	text = text.replace(/\*\*([^*]+)\*\*/g, "<b>$1</b>");
	text = text.replace(/(^|[^*\w])\*([^*\n]+)\*/g, "$1<i>$2</i>");
	text = text.replace(/(^|[\s(])_([^_\n]+)_(?=$|[\s).,;:!?])/g, "$1<i>$2</i>");
	text = text.replace(/~~([^~]+)~~/g, "<s>$1</s>");

	return text;
}

const BULLET = /^\s*[-*+]\s+(.*)$/;
const NUMBERED = /^\s*\d+[.)]\s+(.*)$/;
const HEADING = /^\s{0,3}#{1,6}\s+(.*)$/;

/**
 * Convert a message body to Teams HTML.
 *
 * Blank lines separate paragraphs, single newlines become `<br>` inside one —
 * which is what a person typing into the chat box gets, and the safest
 * assumption for text that was generated, not composed.
 *
 * Teams renders `<p>` (and lists) in a chat bubble with no margin at all, so
 * `<p>a</p><p>b</p>` looks exactly like `a<br>b` — every blank line the model
 * wrote disappears and a digest arrives as one block. The Teams client itself
 * writes an empty paragraph for a blank line, so blocks are joined with one.
 */
export function markdownToTeamsHtml(text: string): string {
	const lines = text.replace(/\r\n?/g, "\n").split("\n");
	const blocks: string[] = [];

	let paragraph: string[] = [];
	let list: { ordered: boolean; items: string[] } | undefined;
	let code: string[] | undefined;

	const flushParagraph = () => {
		if (paragraph.length === 0) return;
		blocks.push(`<p>${paragraph.map(inline).join("<br>")}</p>`);
		paragraph = [];
	};

	const flushList = () => {
		if (!list) return;
		const tag = list.ordered ? "ol" : "ul";
		const items = list.items.map((item) => `<li>${inline(item)}</li>`).join("");
		blocks.push(`<${tag}>${items}</${tag}>`);
		list = undefined;
	};

	const flushBlock = () => {
		flushParagraph();
		flushList();
	};

	for (const line of lines) {
		// A fence swallows everything until the next one, verbatim.
		if (/^\s*```/.test(line)) {
			if (code) {
				blocks.push(`<pre>${escapeHtml(code.join("\n"))}</pre>`);
				code = undefined;
			} else {
				flushBlock();
				code = [];
			}
			continue;
		}
		if (code) {
			code.push(line);
			continue;
		}

		if (line.trim() === "") {
			flushBlock();
			continue;
		}

		const heading = HEADING.exec(line);
		if (heading) {
			flushBlock();
			blocks.push(`<p><b>${inline(heading[1])}</b></p>`);
			continue;
		}

		const bullet = BULLET.exec(line);
		if (bullet) {
			flushParagraph();
			if (list && list.ordered) flushList();
			list ??= { ordered: false, items: [] };
			list.items.push(bullet[1]);
			continue;
		}

		const numbered = NUMBERED.exec(line);
		if (numbered) {
			flushParagraph();
			if (list && !list.ordered) flushList();
			list ??= { ordered: true, items: [] };
			list.items.push(numbered[1]);
			continue;
		}

		flushList();
		paragraph.push(line);
	}

	// An unterminated fence is still code, not lost text.
	if (code) blocks.push(`<pre>${escapeHtml(code.join("\n"))}</pre>`);
	flushBlock();

	return blocks.join(BLOCK_SPACER);
}

/** What the Teams client writes for an empty line between two blocks. */
export const BLOCK_SPACER = "<p>&nbsp;</p>";
