/**
 * `analyze` — make a binary .docx readable.
 *
 * Claude cannot open a .docx, so every workflow in this skill starts here. The
 * output is deliberately structural rather than pretty: paragraph and table
 * indices printed here are the same ones `generate.ts` and `refine.ts` accept,
 * so the analysis doubles as the coordinate system for the rest of the skill.
 *
 * Usage:
 *   node analyze.mjs <file.docx> [file2.docx ...] [--aliases aliases.json]
 *                                [--delimiter +++] [--text-only]
 */
import fs from 'node:fs';
import path from 'node:path';

import { listCommands } from '../../../src/index';
import {
  type CommandSummary,
  type Node,
  type NonTextNode,
} from '../../../src/types';
import {
  type AliasOptions,
  directChildren,
  firstChild,
  isEntryPoint,
  parseArgs,
  paragraphsOutsideTables,
  plainTextLines,
  readAliases,
  readDelimiter,
  readDocParts,
  requireFile,
  runMain,
  textOf,
  topLevelTables,
} from './docxXml';

// ==========================================
// Shapes
// ==========================================

type RunInfo = {
  text: string;
  bold?: true;
  italic?: true;
  underline?: true;
  /** Half-points, as Word stores it: 24 means 12pt. */
  size?: string;
  font?: string;
  color?: string;
};

type ParagraphInfo = {
  index: number;
  text: string;
  runs: RunInfo[];
  style?: string;
  /** Set when the paragraph is part of a numbered or bulleted list. */
  listLevel?: string;
};

type TableInfo = {
  index: number;
  /** Whether the first row looks like a header (all cells bold). */
  headerRow: boolean;
  rows: { cells: { text: string; paragraphs: ParagraphInfo[] }[] }[];
};

type PartInfo = {
  filename: string;
  paragraphs: ParagraphInfo[];
  tables: TableInfo[];
};

type Analysis = {
  file: string;
  plainText: string;
  main: PartInfo;
  headersAndFooters: PartInfo[];
  /** Commands already present, i.e. the file is (or contains) a template. */
  commands: CommandSummary[];
  /** Set instead of `commands` when command parsing failed. */
  commandsError?: string;
};

// ==========================================
// Extraction
// ==========================================

/** Reads an attribute, tolerating both of sax's attribute shapes. */
function attr(node: NonTextNode, name: string): string | undefined {
  const value = node._attrs[name];
  if (value == null) return undefined;
  return typeof value === 'string' ? value : value.value;
}

/** A `w:b`-style toggle is on unless it carries an explicit falsy `w:val`. */
function toggleOn(rPr: NonTextNode, tag: string): boolean {
  const el = firstChild(rPr, tag);
  if (el == null) return false;
  const val = attr(el, 'w:val');
  return val !== '0' && val !== 'false';
}

/** The value of `<child w:val="...">` inside `parent`, if both are there. */
function childAttr(
  parent: NonTextNode,
  tag: string,
  attrName = 'w:val'
): string | undefined {
  const child = firstChild(parent, tag);
  return child != null ? attr(child, attrName) : undefined;
}

function runInfo(run: NonTextNode): RunInfo {
  const info: RunInfo = { text: textOf(run) };
  const rPr = firstChild(run, 'w:rPr');
  if (rPr == null) return info;

  if (toggleOn(rPr, 'w:b')) info.bold = true;
  if (toggleOn(rPr, 'w:i')) info.italic = true;
  if (toggleOn(rPr, 'w:u')) info.underline = true;

  const size = childAttr(rPr, 'w:sz');
  if (size != null) info.size = size;
  const font = childAttr(rPr, 'w:rFonts', 'w:ascii');
  if (font != null) info.font = font;
  const color = childAttr(rPr, 'w:color');
  if (color != null) info.color = color;

  return info;
}

function paragraphInfo(paragraph: NonTextNode, index: number): ParagraphInfo {
  const info: ParagraphInfo = {
    index,
    text: textOf(paragraph),
    runs: directChildren(paragraph, 'w:r')
      .map(runInfo)
      .filter(run => run.text !== ''),
  };

  const pPr = firstChild(paragraph, 'w:pPr');
  if (pPr != null) {
    const style = childAttr(pPr, 'w:pStyle');
    if (style != null) info.style = style;
    const numPr = firstChild(pPr, 'w:numPr');
    if (numPr != null) info.listLevel = childAttr(numPr, 'w:ilvl') ?? '0';
  }

  return info;
}

function tableInfo(table: NonTextNode, index: number): TableInfo {
  const rows = directChildren(table, 'w:tr').map(row => ({
    cells: directChildren(row, 'w:tc').map(cell => ({
      text: textOf(cell),
      paragraphs: directChildren(cell, 'w:p').map(paragraphInfo),
    })),
  }));

  // A header row is the usual signal that the rows below it are data — i.e.
  // that the table is a loop candidate.
  const firstRow = rows[0];
  const headerRow =
    rows.length > 1 &&
    firstRow != null &&
    firstRow.cells.length > 0 &&
    firstRow.cells.every(cell =>
      cell.paragraphs.some(p => p.runs.length > 0 && p.runs.every(r => r.bold))
    );

  return { index, headerRow, rows };
}

const partInfo = (filename: string, root: Node): PartInfo => ({
  filename,
  paragraphs: paragraphsOutsideTables(root).map(paragraphInfo),
  tables: topLevelTables(root).map(tableInfo),
});

// ==========================================
// Main
// ==========================================

export async function analyze(
  filePath: string,
  delimiter: [string, string],
  aliases: AliasOptions
): Promise<Analysis> {
  const { parts } = await readDocParts(filePath);
  const main = parts.find(part => part.isMain) ?? parts[0];
  if (main == null) throw new Error(`${filePath} has no readable content`);

  const analysis: Analysis = {
    file: path.basename(filePath),
    plainText: plainTextLines(main.root).join('\n'),
    main: partInfo(main.filename, main.root),
    headersAndFooters: parts
      .filter(part => !part.isMain)
      .map(part => partInfo(part.filename, part.root)),
    commands: [],
  };

  // A filled-out document has no commands, and that is not an error — but a
  // half-written template can fail to parse, and knowing that is the point.
  try {
    analysis.commands = await listCommands(
      fs.readFileSync(filePath),
      delimiter,
      aliases
    );
  } catch (err) {
    analysis.commandsError = err instanceof Error ? err.message : String(err);
  }

  return analysis;
}

async function main(): Promise<void> {
  const { positional, flags } = parseArgs(process.argv.slice(2), ['text-only']);
  if (positional.length === 0) {
    console.error(
      'Usage: node analyze.mjs <file.docx> [file2.docx ...] ' +
        '[--aliases aliases.json] [--delimiter +++] [--text-only]'
    );
    process.exit(1);
  }

  const delimiter = readDelimiter(flags);
  const aliases = readAliases(flags);
  const analyses: Analysis[] = [];
  for (const arg of positional) {
    analyses.push(await analyze(requireFile(arg), delimiter, aliases));
  }

  if (flags['text-only'] === true) {
    for (const analysis of analyses) {
      if (analyses.length > 1) console.log(`\n===== ${analysis.file} =====`);
      console.log(analysis.plainText);
    }
    return;
  }

  const first = analyses[0];
  console.log(
    JSON.stringify(
      analyses.length === 1 && first != null ? first : { files: analyses },
      null,
      2
    )
  );
}

if (isEntryPoint(import.meta.url)) runMain('Analysis', main);
