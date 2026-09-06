/**
 * A deliberately small Markdown subset: what the model actually emits in chat
 * replies (bold, italic, inline code, short headings, bullets, rules).
 *
 * The parser is pure so it can be unit-tested without a DOM, and it returns
 * structured spans rather than an HTML string, so the renderer never has to
 * inject model output as markup.
 */

export type Inline =
  | { type: "text"; text: string }
  | { type: "bold"; text: string }
  | { type: "italic"; text: string }
  | { type: "code"; text: string };

export type Block =
  | { type: "heading"; level: number; spans: Inline[] }
  | { type: "paragraph"; spans: Inline[] }
  | { type: "list"; ordered: boolean; start: number; items: Inline[][] }
  | { type: "table"; header: Inline[][]; rows: Inline[][][] }
  | { type: "rule" };

// Order matters: ** before * so bold wins.
const INLINE_PATTERN = /(\*\*[^\n]+?\*\*|`[^`\n]+`|\*[^\s*][^\n*]*?\*)/g;

/**
 * Italic only counts when the opening star starts a word. Without this,
 * "=B2*C2 is 5 * 3" reads as an italic span and the formula is mangled.
 */
function opensAWord(source: string, start: number): boolean {
  return start === 0 || /\s/.test(source[start - 1]);
}

export function parseInline(source: string): Inline[] {
  const spans: Inline[] = [];
  let index = 0;

  for (const match of source.matchAll(INLINE_PATTERN)) {
    const token = match[0];
    const start = match.index ?? 0;
    const isItalic = !token.startsWith("**") && !token.startsWith("`");
    if (isItalic && !opensAWord(source, start)) {
      continue; // leave it in place; it will be emitted as plain text
    }
    if (start > index) {
      spans.push({ type: "text", text: source.slice(index, start) });
    }
    if (token.startsWith("**")) {
      spans.push({ type: "bold", text: token.slice(2, -2) });
    } else if (token.startsWith("`")) {
      spans.push({ type: "code", text: token.slice(1, -1) });
    } else {
      spans.push({ type: "italic", text: token.slice(1, -1) });
    }
    index = start + token.length;
  }

  if (index < source.length) {
    spans.push({ type: "text", text: source.slice(index) });
  }
  return spans;
}

const HEADING = /^(#{1,6})\s+(.*)$/;
const BULLET = /^\s*[-*+]\s+(.*)$/;
// Bounded to three digits so a year ("2024. was strong") stays prose.
const ORDERED = /^\s*(\d{1,3})[.)]\s+(.*)$/;
const RULE = /^\s*([-*_])\1{2,}\s*$/;
const TABLE_ROW = /^\s*\|(.*)\|\s*$/;
const TABLE_DIVIDER = /^\s*\|[\s:|-]+\|\s*$/;

/** "| a | b |" -> ["a", "b"]. The outer pipes are already stripped by TABLE_ROW. */
function tableCells(line: string): string[] {
  const inner = TABLE_ROW.exec(line);
  return (inner ? inner[1] : line).split("|").map((cell) => cell.trim());
}

export function parseMarkdown(source: string): Block[] {
  const blocks: Block[] = [];
  let paragraph: string[] = [];
  let bullets: Inline[][] = [];
  let listOrdered: boolean | null = null;
  let listStart = 1;

  const flushParagraph = () => {
    if (paragraph.length > 0) {
      blocks.push({ type: "paragraph", spans: parseInline(paragraph.join(" ")) });
      paragraph = [];
    }
  };
  const flushList = () => {
    if (bullets.length > 0) {
      blocks.push({
        type: "list",
        ordered: listOrdered === true,
        start: listOrdered === true ? listStart : 1,
        items: bullets,
      });
      bullets = [];
      listOrdered = null;
      listStart = 1;
    }
  };
  const flush = () => {
    flushParagraph();
    flushList();
  };

  const lines = source.split("\n");

  for (let cursor = 0; cursor < lines.length; cursor += 1) {
    const line = lines[cursor].trimEnd();

    // A table is a header row plus a divider row; without the divider the
    // pipes are just punctuation and the line stays prose.
    if (TABLE_ROW.test(line) && cursor + 1 < lines.length && TABLE_DIVIDER.test(lines[cursor + 1])) {
      flush();
      const header = tableCells(line).map(parseInline);
      const rows: Inline[][][] = [];
      cursor += 2;
      while (cursor < lines.length && TABLE_ROW.test(lines[cursor])) {
        const cells = tableCells(lines[cursor]).map(parseInline);
        // Pad or trim so every cell stays under its heading.
        while (cells.length < header.length) cells.push([]);
        rows.push(cells.slice(0, header.length));
        cursor += 1;
      }
      cursor -= 1;
      blocks.push({ type: "table", header, rows });
      continue;
    }

    if (line.trim() === "") {
      flush();
      continue;
    }
    if (RULE.test(line)) {
      flush();
      blocks.push({ type: "rule" });
      continue;
    }

    const heading = HEADING.exec(line);
    if (heading) {
      flush();
      // Clamp: a pane this narrow has no room for h1/h2 scale.
      blocks.push({
        type: "heading",
        level: Math.min(3, Math.max(3, heading[1].length)),
        spans: parseInline(heading[2].trim()),
      });
      continue;
    }

    const bullet = BULLET.exec(line);
    if (bullet) {
      if (listOrdered === true) flushList();
      listOrdered = false;
      flushParagraph();
      bullets.push(parseInline(bullet[1].trim()));
      continue;
    }

    const ordered = ORDERED.exec(line);
    if (ordered) {
      if (listOrdered === false) flushList();
      if (bullets.length === 0) listStart = Number(ordered[1]);
      listOrdered = true;
      flushParagraph();
      bullets.push(parseInline(ordered[2].trim()));
      continue;
    }

    flushList();
    paragraph.push(line.trim());
  }

  flush();
  return blocks;
}
