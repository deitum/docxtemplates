/**
 * `refine` — surgical changes to a template that already exists.
 *
 * `generate` rebuilds a template from a mapping; this one edits in place, which
 * is what you want once a human has adjusted the styling in Word and you only
 * need to rename an expression or wrap a block in a loop.
 *
 * Usage:
 *   node refine.mjs <template.docx> <modifications.json> [output.docx]
 */
import path from 'node:path';

import { type Node, type NonTextNode } from '../../../src/types';
import { normalizeDelimiter } from './generate';
import {
  type DocPart,
  type TextIndex,
  command,
  directChildren,
  firstChild,
  indexText,
  insertAfter,
  isEntryPoint,
  isTag,
  newElement,
  parseArgs,
  readDocParts,
  readJson,
  remove,
  replaceTextAll,
  requireFile,
  runMain,
  spliceText,
  textOf,
  topLevelTables,
  wrapParagraphs,
  writeDocParts,
} from './docxXml';

// ==========================================
// Modifications file
// ==========================================

type Modification =
  /** Swap one whole command for another: `INS a.b` -> `INS c.d`. */
  | { type: 'replaceCommand'; from: string; to: string }
  /** Rewrite a fragment inside every command that contains it. */
  | { type: 'renameExpression'; from: string; to: string }
  /** Drop a command, optionally leaving literal text behind. */
  | { type: 'removeCommand'; code: string; replaceWith?: string }
  /** Turn literal document text into an `INS` command. */
  | { type: 'addCommand'; text: string; code: string }
  /** Wrap a block of paragraphs in FOR / END-FOR. */
  | {
      type: 'wrapFor';
      var: string;
      over: string;
      startText: string;
      endText?: string;
    }
  /** Wrap a block of paragraphs in IF / END-IF. */
  | { type: 'wrapIf'; expr: string; startText: string; endText?: string }
  /** Merge absolutely-positioned tables into one inline table. */
  | { type: 'mergeFloatingTables'; tableIndices?: number[] };

type Modifications = {
  cmdDelimiter?: string | [string, string];
  modifications: Modification[];
};

// ==========================================
// Command-scoped text edits
// ==========================================

type CommandSpan = {
  /** Offset of the opening delimiter in the part's flat text. */
  start: number;
  /** Offset just past the closing delimiter. */
  end: number;
  /** The command body, delimiters excluded. */
  code: string;
};

/** Every delimited command in a part, in document order. */
function findCommandSpans(
  index: TextIndex,
  delimiter: [string, string]
): CommandSpan[] {
  const [open, close] = delimiter;
  const spans: CommandSpan[] = [];
  let at = 0;
  for (;;) {
    const start = index.full.indexOf(open, at);
    if (start < 0) break;
    const bodyStart = start + open.length;
    const closeAt = index.full.indexOf(close, bodyStart);
    if (closeAt < 0) break;
    spans.push({
      start,
      end: closeAt + close.length,
      code: index.full.slice(bodyStart, closeAt),
    });
    at = closeAt + close.length;
  }
  return spans;
}

/**
 * Rewrites every command in a part through `edit`, which is handed the command
 * body and returns the text that replaces the command *including* its
 * delimiters — `''` to delete it outright, or `undefined` to leave it alone.
 *
 * Working on the part's flat text rather than node by node is what makes this
 * safe on templates that have been through Word, which cheerfully splits
 * `+++INS client.name+++` across five runs.
 */
function editCommands(
  root: Node,
  delimiter: [string, string],
  edit: (code: string) => string | undefined
): number {
  const index = indexText(root);
  const spans = findCommandSpans(index, delimiter);
  let count = 0;

  // Back to front, so earlier offsets stay valid as we go.
  for (const span of spans.reverse()) {
    const replacement = edit(span.code);
    if (replacement === undefined) continue;
    spliceText(index, span.start, span.end, replacement);
    count += 1;
  }
  return count;
}

// ==========================================
// Floating tables
// ==========================================

const isFloating = (table: NonTextNode): boolean => {
  const tblPr = firstChild(table, 'w:tblPr');
  return tblPr != null && firstChild(tblPr, 'w:tblpPr') != null;
};

/** Removes the borders of a merged table, so the seams stay invisible. */
function clearBorders(table: NonTextNode): void {
  const tblPr = firstChild(table, 'w:tblPr');
  if (tblPr == null) return;
  const existing = firstChild(tblPr, 'w:tblBorders');
  if (existing != null) remove(existing);
  const sides = ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'];
  const borders = newElement(
    'w:tblBorders',
    {},
    sides.map(side => newElement(`w:${side}`, { 'w:val': 'none', 'w:sz': '0' }))
  );
  borders._parent = tblPr;
  tblPr._children.push(borders);
}

/**
 * Merges absolutely-positioned tables into the first one, side by side.
 *
 * Word uses floating tables for signature blocks and other two-column layouts.
 * They drift out of alignment at the slightest provocation and are miserable to
 * template — a single inline table with the same cells behaves.
 */
function mergeFloatingTables(root: Node, indices?: number[]): string {
  const all = topLevelTables(root);
  const targets =
    indices != null
      ? indices.map(i => {
          const table = all[i];
          if (table == null) throw new Error(`No table with index ${i}`);
          return table;
        })
      : all.filter(isFloating);

  const [base, ...rest] = targets;
  if (base == null || rest.length === 0) {
    return 'mergeFloatingTables: fewer than two tables matched, nothing to do';
  }

  for (const table of rest) {
    // The empty paragraphs Word leaves between floating tables would push the
    // merged table apart again.
    removeEmptyParagraphsBetween(base, table);

    const baseRows = directChildren(base, 'w:tr');
    const rows = directChildren(table, 'w:tr');
    rows.forEach((row, i) => {
      const target = baseRows[i];
      if (target == null) {
        // More rows than the base has: move the whole row across.
        remove(row);
        const last = baseRows[baseRows.length - 1];
        if (last != null) insertAfter(last, row);
        else {
          row._parent = base;
          base._children.push(row);
        }
        baseRows.push(row);
        return;
      }
      for (const cell of directChildren(row, 'w:tc')) {
        remove(cell);
        cell._parent = target;
        target._children.push(cell);
      }
    });

    // Columns come across too, so the original widths are preserved.
    const baseGrid = firstChild(base, 'w:tblGrid');
    const grid = firstChild(table, 'w:tblGrid');
    if (baseGrid != null && grid != null) {
      for (const col of directChildren(grid, 'w:gridCol')) {
        remove(col);
        col._parent = baseGrid;
        baseGrid._children.push(col);
      }
    }
    remove(table);
  }

  // The survivor becomes an ordinary inline table.
  const tblPr = firstChild(base, 'w:tblPr');
  const tblpPr = tblPr != null ? firstChild(tblPr, 'w:tblpPr') : undefined;
  if (tblpPr != null) remove(tblpPr);
  clearBorders(base);

  return `merged ${targets.length} floating tables into one inline table`;
}

/** Deletes the empty paragraphs sitting between two siblings. */
function removeEmptyParagraphsBetween(a: Node, b: Node): void {
  const parent = a._parent;
  if (parent == null || b._parent !== parent) return;
  const from = parent._children.indexOf(a);
  const to = parent._children.indexOf(b);
  if (from < 0 || to < 0 || to < from) return;
  for (const node of parent._children.slice(from + 1, to)) {
    if (isTag(node, 'w:p') && textOf(node).trim() === '') remove(node);
  }
}

// ==========================================
// Applying modifications
// ==========================================

export function applyModifications(
  parts: DocPart[],
  file: Modifications
): string[] {
  const delimiter = normalizeDelimiter(file.cmdDelimiter);
  const mainPart = parts.find(part => part.isMain) ?? parts[0];
  if (mainPart == null) throw new Error('Document has no main part');
  const mainRoot = mainPart.root;
  const log: string[] = [];

  /** Runs a command edit over every part and reports the total. */
  const overParts = (edit: (code: string) => string | undefined): number =>
    parts.reduce(
      (total, part) => total + editCommands(part.root, delimiter, edit),
      0
    );

  for (const mod of file.modifications ?? []) {
    switch (mod.type) {
      case 'replaceCommand': {
        const n = overParts(code =>
          code.trim() === mod.from.trim()
            ? command(mod.to, delimiter)
            : undefined
        );
        log.push(`${n}x ${mod.from} -> ${mod.to}`);
        break;
      }

      case 'renameExpression': {
        const n = overParts(code => {
          const renamed = code.split(mod.from).join(mod.to);
          return renamed === code ? undefined : command(renamed, delimiter);
        });
        log.push(`${n} command(s): ${mod.from} -> ${mod.to}`);
        break;
      }

      case 'removeCommand': {
        // `replaceWith` is literal document text, not a command — that is the
        // point of removing the command in the first place.
        const n = overParts(code =>
          code.trim() === mod.code.trim() ? (mod.replaceWith ?? '') : undefined
        );
        log.push(`${n}x removed ${mod.code}`);
        break;
      }

      case 'addCommand': {
        const cmd = command(`INS ${mod.code}`, delimiter);
        const n = parts.reduce(
          (total, part) => total + replaceTextAll(part.root, mod.text, cmd),
          0
        );
        if (n === 0) {
          throw new Error(`addCommand: "${mod.text}" is not in the document`);
        }
        log.push(`${n}x "${mod.text}" -> ${cmd}`);
        break;
      }

      case 'wrapFor':
        wrapParagraphs(mainRoot, {
          startText: mod.startText,
          endText: mod.endText ?? mod.startText,
          openCmd: command(`FOR ${mod.var} IN ${mod.over}`, delimiter),
          closeCmd: command(`END-FOR ${mod.var}`, delimiter),
          what: `wrapFor ${mod.var}`,
        });
        log.push(`wrapped block in FOR ${mod.var} IN ${mod.over}`);
        break;

      case 'wrapIf':
        wrapParagraphs(mainRoot, {
          startText: mod.startText,
          endText: mod.endText ?? mod.startText,
          openCmd: command(`IF ${mod.expr}`, delimiter),
          closeCmd: command('END-IF', delimiter),
          what: `wrapIf ${mod.expr}`,
        });
        log.push(`wrapped block in IF ${mod.expr}`);
        break;

      case 'mergeFloatingTables':
        log.push(mergeFloatingTables(mainRoot, mod.tableIndices));
        break;

      default: {
        const unknown = mod as { type: string };
        throw new Error(`Unknown modification type "${unknown.type}"`);
      }
    }
  }

  return log;
}

// ==========================================
// Main
// ==========================================

async function main(): Promise<void> {
  const { positional } = parseArgs(process.argv.slice(2));
  const [templateArg, modsArg, outputArg] = positional;
  if (templateArg == null || modsArg == null) {
    console.error(
      'Usage: node refine.mjs <template.docx> <modifications.json> [output.docx]'
    );
    process.exit(1);
  }

  const templatePath = requireFile(templateArg);
  const mods = readJson<Modifications>(requireFile(modsArg));
  const outputPath = path.resolve(
    outputArg ?? templatePath.replace(/\.docx$/i, '_refined.docx')
  );

  const { zip, parts } = await readDocParts(templatePath);
  for (const line of applyModifications(parts, mods)) console.log(`  ${line}`);
  await writeDocParts(zip, parts, outputPath);
  console.log(`\nTemplate written to ${outputPath}`);
  console.log('Next: run verify.mjs to confirm it still renders.');
}

if (isEntryPoint(import.meta.url)) runMain('Refinement', main);
