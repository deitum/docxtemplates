/**
 * Shared docx plumbing for the skill's agent tools.
 *
 * Everything here is built on the library's own primitives — the sax-based
 * parser in `src/xml.ts`, the JSZip wrappers in `src/zip.ts` and the node
 * constructors in `src/reportUtils.ts` — so the tools see exactly the document
 * model that `createReport` walks at render time. That is the whole point: a
 * template these tools produce is one the engine can already parse.
 */
import fs from 'node:fs';
import path from 'node:path';
import { pathToFileURL } from 'node:url';

import type JSZip from 'jszip';

import { getMainDoc, readContentTypes } from '../../../src/main';
import { newNonTextNode, newTextNode } from '../../../src/reportUtils';
import { type Node, type NonTextNode, type TextNode } from '../../../src/types';
import { buildXml, parseXml } from '../../../src/xml';
import { zipGetText, zipLoad, zipSave, zipSetText } from '../../../src/zip';

// `buildXml` skips XML escaping between `literalXmlDelimiter` markers, which is
// how the engine lets report *data* inject raw markup. Nothing here wants that,
// and the default `||` is common enough in real prose to matter, so the
// delimiter is set to NUL — a character XML 1.0 forbids outright, and which
// therefore cannot occur in any document we are handed.
const NO_LITERAL_XML = '\u0000' as const;

// `indentXml: false` keeps the serializer from inserting newlines of its own,
// so untouched parts of the document round-trip byte-for-byte.
const XML_OPTIONS = {
  literalXmlDelimiter: NO_LITERAL_XML,
  indentXml: false,
} as const;

/** Same default `createReport` uses. */
export const DEFAULT_CMD_DELIMITER = '+++' as const;

// ==========================================
// Node model helpers
// ==========================================

export const isTag = (node: Node, tag: string): node is NonTextNode =>
  !node._fTextNode && node._tag === tag;

/**
 * Direct children with the given tag. Unlike a recursive search, this does not
 * reach into nested tables — which is what makes table indices predictable.
 */
export const directChildren = (node: Node, tag: string): NonTextNode[] =>
  node._children.filter((child): child is NonTextNode => isTag(child, tag));

/** First direct child with the given tag, if any. */
export const firstChild = (node: Node, tag: string): NonTextNode | undefined =>
  directChildren(node, tag)[0];

/** Depth-first walk over the whole subtree, `node` included. */
export function walk(node: Node, fn: (node: Node) => void): void {
  fn(node);
  for (const child of node._children) walk(child, fn);
}

/** Every descendant with the given tag, in document order. */
export function findAll(root: Node, tag: string): NonTextNode[] {
  const out: NonTextNode[] = [];
  walk(root, node => {
    if (isTag(node, tag)) out.push(node);
  });
  return out;
}

/** The first ancestor with the given tag, if any. */
export function closest(node: Node, tag: string): NonTextNode | undefined {
  let cur = node._parent;
  while (cur != null) {
    if (isTag(cur, tag)) return cur;
    cur = cur._parent;
  }
  return undefined;
}

/** Paragraphs that are not inside a table (top-level body text). */
export const paragraphsOutsideTables = (root: Node): NonTextNode[] =>
  findAll(root, 'w:p').filter(p => closest(p, 'w:tbl') == null);

/**
 * Tables in document order, skipping tables nested inside other tables. This
 * is the set — and the numbering — that `tableIndex` refers to everywhere in
 * the skill.
 */
export const topLevelTables = (root: Node): NonTextNode[] =>
  findAll(root, 'w:tbl').filter(table => closest(table, 'w:tbl') == null);

/** Index of `node` among its parent's children; -1 when it has no parent. */
export function indexInParent(node: Node): number {
  const parent = node._parent;
  if (parent == null) return -1;
  return parent._children.indexOf(node);
}

/** Inserts `newNode` right before `ref`. */
export function insertBefore(ref: Node, newNode: Node): void {
  const parent = ref._parent;
  if (parent == null) throw new Error('Cannot insert before a root node');
  newNode._parent = parent;
  parent._children.splice(parent._children.indexOf(ref), 0, newNode);
}

/** Inserts `newNode` right after `ref`. */
export function insertAfter(ref: Node, newNode: Node): void {
  const parent = ref._parent;
  if (parent == null) throw new Error('Cannot insert after a root node');
  newNode._parent = parent;
  parent._children.splice(parent._children.indexOf(ref) + 1, 0, newNode);
}

/** Detaches `node` from its parent. */
export function remove(node: Node): void {
  const parent = node._parent;
  if (parent == null) return;
  const idx = parent._children.indexOf(node);
  if (idx >= 0) parent._children.splice(idx, 1);
  delete node._parent;
}

/** Deep copy of a subtree, with `_parent` links rebuilt. */
export function cloneNode<T extends Node>(node: T, parent?: Node): T {
  const copy: Node = node._fTextNode
    ? { _fTextNode: true, _text: node._text, _children: [] }
    : {
        _fTextNode: false,
        _tag: node._tag,
        _attrs: { ...node._attrs },
        _children: [],
      };
  if (parent != null) copy._parent = parent;
  copy._children = node._children.map(child => cloneNode(child, copy));
  return copy as T;
}

// ==========================================
// Text
// ==========================================

/** The text nodes inside a subtree's `w:t` elements, in document order. */
export function textNodes(root: Node): TextNode[] {
  const out: TextNode[] = [];
  for (const t of findAll(root, 'w:t')) {
    for (const child of t._children) {
      if (child._fTextNode) out.push(child);
    }
  }
  return out;
}

/**
 * Visible text of a subtree. Word splits a sentence across as many runs as it
 * likes, so this is the only reliable thing to match against.
 */
export const textOf = (root: Node): string =>
  textNodes(root)
    .map(t => t._text)
    .join('');

/**
 * A flat, readable rendition of a document part, in document order, with table
 * rows as ` | `-separated cells. This is what to read first when comparing two
 * filled-out documents; the structured analysis is for pinpointing what to
 * change afterwards.
 */
export function plainTextLines(
  container: Node,
  lines: string[] = []
): string[] {
  for (const child of container._children) {
    if (child._fTextNode) continue;
    if (isTag(child, 'w:p')) {
      lines.push(textOf(child));
    } else if (isTag(child, 'w:tbl')) {
      for (const row of directChildren(child, 'w:tr')) {
        lines.push(
          directChildren(row, 'w:tc')
            .map(cell => textOf(cell).replace(/\n/g, ' '))
            .join(' | ')
        );
      }
    } else {
      // Content wrappers such as `w:body` or `w:sdt`: recurse to keep order.
      plainTextLines(child, lines);
    }
  }
  return lines;
}

/**
 * A subtree's text nodes plus the concatenated string they form, so that a
 * match found in the flat text can be mapped back to the nodes it spans.
 */
export type TextIndex = {
  nodes: TextNode[];
  /** Offset of each node's text within `full`. */
  offsets: number[];
  full: string;
};

export function indexText(root: Node): TextIndex {
  const nodes = textNodes(root);
  const offsets: number[] = [];
  let full = '';
  for (const node of nodes) {
    offsets.push(full.length);
    full += node._text;
  }
  return { nodes, offsets, full };
}

/**
 * Replaces `full[start:end]` with `replacement`, touching only the text nodes
 * that overlap the range. Every other run — and therefore every bit of
 * formatting the document already had — is left exactly as it was; the
 * replacement inherits the formatting of the run the range starts in.
 *
 * The index is stale afterwards. Apply several splices back-to-front, or
 * re-index between them.
 */
export function spliceText(
  index: TextIndex,
  start: number,
  end: number,
  replacement: string
): void {
  let placed = false;
  index.nodes.forEach((node, i) => {
    const nodeStart = index.offsets[i] ?? 0;
    const nodeEnd = nodeStart + node._text.length;
    if (nodeEnd <= start || nodeStart >= end) return;

    const before = node._text.slice(0, Math.max(0, start - nodeStart));
    const after = node._text.slice(
      Math.min(node._text.length, end - nodeStart)
    );
    node._text = placed ? before + after : before + replacement + after;
    placed = true;
  });
}

/**
 * Replaces the first occurrence of `search` with `replacement`, preserving the
 * formatting of every run it does not touch. Returns whether it found one.
 */
export function replaceText(
  root: Node,
  search: string,
  replacement: string
): boolean {
  if (search === '') return false;
  const index = indexText(root);
  const start = index.full.indexOf(search);
  if (start < 0) return false;
  spliceText(index, start, start + search.length, replacement);
  return true;
}

/** Replaces every occurrence of `search`; returns how many were replaced. */
export function replaceTextAll(
  root: Node,
  search: string,
  replacement: string
): number {
  if (search === '') return 0;
  const index = indexText(root);

  const starts: number[] = [];
  for (
    let at = index.full.indexOf(search);
    at >= 0;
    at = index.full.indexOf(search, at + search.length)
  ) {
    starts.push(at);
  }

  // Back to front, so that each splice leaves the offsets of the ones still to
  // come untouched.
  for (const start of starts.reverse()) {
    spliceText(index, start, start + search.length, replacement);
  }
  return starts.length;
}

// ==========================================
// Building command nodes
// ==========================================

/** Wraps `code` in the command delimiters, e.g. `INS name` -> `+++INS name+++`. */
export const command = (
  code: string,
  delimiter: [string, string] = [DEFAULT_CMD_DELIMITER, DEFAULT_CMD_DELIMITER]
): string => `${delimiter[0]}${code}${delimiter[1]}`;

/**
 * A new element node. Thin re-export of the engine's own constructor, so the
 * tools build nodes exactly the way `createReport` does.
 */
export const newElement = (
  tag: string,
  attrs: { [key: string]: string } = {},
  children: Node[] = []
): NonTextNode => newNonTextNode(tag, attrs, children);

const textRun = (text: string): NonTextNode =>
  newElement('w:r', {}, [
    newElement('w:t', { 'xml:space': 'preserve' }, [newTextNode(text)]),
  ]);

/**
 * A `w:p` holding nothing but `text`. When `text` is a command, the engine
 * deletes the whole paragraph after running it (a `w:p` whose buffer holds
 * commands and no inserted text is dropped), which is what makes FOR/IF
 * markers vanish from the report.
 */
export const commandParagraph = (text: string): NonTextNode =>
  newElement('w:p', {}, [textRun(text)]);

/**
 * A `w:tr` modelled on `refRow` — same cell count, same cell properties, hence
 * the same column widths and borders — whose first cell holds nothing but
 * `text` and whose other cells are empty. Same disappearing act as
 * `commandParagraph`, one level up: this is how FOR and END-FOR are written for
 * table loops in this library, each on a row of its own.
 */
export function commandRowLike(refRow: NonTextNode, text: string): NonTextNode {
  const row = cloneNode(refRow);
  directChildren(row, 'w:tc').forEach((cell, i) => {
    const tcPr = firstChild(cell, 'w:tcPr');
    const paragraph = newElement('w:p', {}, i === 0 ? [textRun(text)] : []);
    cell._children = tcPr != null ? [tcPr, paragraph] : [paragraph];
    for (const child of cell._children) child._parent = cell;
  });
  return row;
}

/**
 * Replaces a paragraph's contents with a single run holding `text`, keeping the
 * paragraph's own properties (`w:pPr`) so numbering and indentation survive.
 */
export function setParagraphText(paragraph: NonTextNode, text: string): void {
  const pPr = firstChild(paragraph, 'w:pPr');
  paragraph._children = pPr != null ? [pPr] : [];
  const run = textRun(text);
  run._parent = paragraph;
  paragraph._children.push(run);
}

/**
 * Locates a run of paragraphs by the text of its first and last one. `endText`
 * is searched from the start paragraph onwards, so a phrase that occurs both
 * before and inside the block still resolves to the right range.
 */
export function findParagraphRange(
  root: Node,
  startText: string,
  endText: string,
  what: string
): { start: NonTextNode; end: NonTextNode } {
  const paragraphs = paragraphsOutsideTables(root);
  const startIdx = paragraphs.findIndex(p => textOf(p).includes(startText));
  if (startIdx < 0) {
    throw new Error(`${what}: no paragraph contains "${startText}"`);
  }
  const endOffset = paragraphs
    .slice(startIdx)
    .findIndex(p => textOf(p).includes(endText));
  if (endOffset < 0) {
    throw new Error(
      `${what}: no paragraph at or after "${startText}" contains "${endText}"`
    );
  }
  const start = paragraphs[startIdx];
  const end = paragraphs[startIdx + endOffset];
  if (start == null || end == null) throw new Error(`${what}: range not found`);
  return { start, end };
}

/**
 * Wraps a block of paragraphs between two command-only paragraphs — the shape
 * every paragraph-level FOR and IF takes in this library.
 */
export function wrapParagraphs(
  root: Node,
  options: {
    startText: string;
    endText: string;
    openCmd: string;
    closeCmd: string;
    what: string;
  }
): void {
  const { start, end } = findParagraphRange(
    root,
    options.startText,
    options.endText,
    options.what
  );
  insertBefore(start, commandParagraph(options.openCmd));
  insertAfter(end, commandParagraph(options.closeCmd));
}

// ==========================================
// Docx parts
// ==========================================

export type DocPart = {
  /** Path inside the zip, e.g. `word/document.xml`. */
  filename: string;
  /** Parsed root node of that part. */
  root: Node;
  /** Whether this is the main document rather than a header or footer. */
  isMain: boolean;
};

const HEADER_FOOTER = /^word\/(header|footer)\d+\.xml$/;

/**
 * Reads a docx and parses its main document plus every header and footer.
 *
 * The main part is looked up through `[Content_Types].xml` rather than assumed
 * to be `word/document.xml`: Office 365 sometimes names it `document2.xml`
 * (upstream issue #131), and the engine itself resolves it the same way.
 */
export async function readDocParts(
  filePath: string
): Promise<{ zip: JSZip; parts: DocPart[] }> {
  const zip = await zipLoad(fs.readFileSync(filePath));
  const mainDoc = `word/${getMainDoc(await readContentTypes(zip))}`;
  const filenames = [
    mainDoc,
    ...Object.keys(zip.files)
      .filter(name => HEADER_FOOTER.test(name))
      .sort(),
  ];

  const parts: DocPart[] = [];
  for (const filename of filenames) {
    const xml = await zipGetText(zip, filename);
    if (xml == null) continue;
    parts.push({
      filename,
      root: await parseXml(xml),
      isMain: filename === mainDoc,
    });
  }
  if (parts.length === 0)
    throw new Error(`${filePath} does not contain ${mainDoc}`);
  return { zip, parts };
}

/** Serializes the (mutated) parts back into the zip and writes it out. */
export async function writeDocParts(
  zip: JSZip,
  parts: DocPart[],
  outputPath: string
): Promise<void> {
  for (const { filename, root } of parts) {
    zipSetText(zip, filename, buildXml(root, XML_OPTIONS));
  }
  const out = await zipSave(zip, 1);
  fs.mkdirSync(path.dirname(path.resolve(outputPath)), { recursive: true });
  fs.writeFileSync(outputPath, out);
}

// ==========================================
// CLI plumbing
// ==========================================

export type AliasOptions = {
  commandAliases?: { [alias: string]: string };
  operatorAliases?: { [alias: string]: string };
};

/** Reads and parses a JSON file, with the path in the error message. */
export function readJson<T>(filePath: string): T {
  const resolved = path.resolve(filePath);
  try {
    return JSON.parse(fs.readFileSync(resolved, 'utf-8')) as T;
  } catch (err) {
    throw new Error(`Could not read JSON from ${resolved}`, { cause: err });
  }
}

/** Exits with a message on stderr if `filePath` is not an existing file. */
export function requireFile(filePath: string): string {
  const resolved = path.resolve(filePath);
  if (!fs.existsSync(resolved)) {
    console.error(`File not found: ${resolved}`);
    process.exit(1);
  }
  return resolved;
}

/**
 * Splits `argv` into positional arguments and `--flag value` / `--flag=value`
 * pairs. Flags listed in `booleanFlags` take no value.
 */
export function parseArgs(
  argv: string[],
  booleanFlags: string[] = []
): { positional: string[]; flags: Record<string, string | true> } {
  const positional: string[] = [];
  const flags: Record<string, string | true> = {};

  for (let i = 0; i < argv.length; i += 1) {
    const arg = argv[i];
    if (arg == null) continue;
    if (!arg.startsWith('-')) {
      positional.push(arg);
      continue;
    }
    const name = arg.replace(/^--?/, '');
    const eq = name.indexOf('=');
    if (eq >= 0) {
      flags[name.slice(0, eq)] = name.slice(eq + 1);
      continue;
    }
    if (booleanFlags.includes(name)) {
      flags[name] = true;
      continue;
    }
    const next = argv[i + 1];
    if (next == null || next.startsWith('--')) {
      flags[name] = true;
      continue;
    }
    flags[name] = next;
    i += 1;
  }

  return { positional, flags };
}

/** Reads the `--aliases <file.json>` flag, if present. */
export function readAliases(
  flags: Record<string, string | true>
): AliasOptions {
  const file = flags.aliases;
  if (typeof file !== 'string') return {};
  return readJson<AliasOptions>(file);
}

/** Reads the `--delimiter <str>` flag, defaulting to `+++`. */
export function readDelimiter(
  flags: Record<string, string | true>
): [string, string] {
  const value = flags.delimiter;
  if (typeof value !== 'string') {
    return [DEFAULT_CMD_DELIMITER, DEFAULT_CMD_DELIMITER];
  }
  const parts = value.split(',');
  if (parts.length === 2 && parts[0] != null && parts[1] != null) {
    return [parts[0], parts[1]];
  }
  return [value, value];
}

/**
 * Whether the module with this `import.meta.url` is the one node was asked to
 * run. Each tool guards its `main()` with this so that the test suite — and
 * the other tools — can import its functions without triggering a CLI run.
 */
export const isEntryPoint = (metaUrl: string): boolean =>
  process.argv[1] != null &&
  pathToFileURL(path.resolve(process.argv[1])).href === metaUrl;

/** An error's message, followed by the messages of everything that caused it. */
function describe(err: unknown): string {
  if (!(err instanceof Error)) return String(err);
  return err.cause != null
    ? `${err.message}: ${describe(err.cause)}`
    : err.message;
}

/**
 * Wraps a tool's `main`, printing failures without a stack-trace wall. Template
 * errors arrive as an array when `failFast` is off, and every one of them is
 * worth reading — that is the whole point of collecting them.
 */
export function runMain(name: string, main: () => Promise<void>): void {
  main().catch((err: unknown) => {
    const errors = Array.isArray(err) ? err : [err];
    console.error(`${name} failed:`);
    for (const e of errors) console.error(`  ${describe(e)}`);
    process.exit(1);
  });
}
