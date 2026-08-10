/**
 * `generate` — turn a filled-out .docx into a template.
 *
 * The mapping file says *what* to templatise; this tool does the XML surgery
 * and nothing else. Two things about it are specific to this library rather
 * than to docxtemplater-style engines, and both matter:
 *
 * - Mapping values are **JavaScript expressions**, not tag names. `client.name`
 *   and `items.filter(i => i.qty > 0).length` are equally valid.
 * - `FOR`/`END-FOR` and `IF`/`END-IF` go in a paragraph — or a table row — *of
 *   their own*. The engine drops any paragraph, row or cell that ends up
 *   holding only commands, so the markers leave no trace in the report.
 *
 * Usage:
 *   node generate.mjs <original.docx> <mapping.json> [output.docx]
 */
import fs from 'node:fs';
import path from 'node:path';

import { listCommands } from '../../../src/index';
import { type Node } from '../../../src/types';
import {
  type AliasOptions,
  type DocPart,
  DEFAULT_CMD_DELIMITER,
  command,
  commandRowLike,
  directChildren,
  insertAfter,
  insertBefore,
  isEntryPoint,
  parseArgs,
  readDocParts,
  readJson,
  remove,
  replaceText,
  replaceTextAll,
  requireFile,
  runMain,
  setParagraphText,
  textOf,
  topLevelTables,
  wrapParagraphs,
  writeDocParts,
} from './docxXml';

// ==========================================
// Mapping file
// ==========================================

type TableLoop = {
  /** Loop variable; referenced inside the loop as `$<var>`. */
  var: string;
  /** JS expression the loop iterates over, e.g. `order.items`. */
  over: string;
  /** 0-based index among the document's non-nested tables. */
  tableIndex: number;
  /** First data row; defaults to 1 (row 0 being the header). */
  startRow?: number;
  /** Last data row to absorb into the loop; defaults to the final row. */
  endRow?: number;
  /** Column index (as a string key) -> expression for that cell. */
  fields: { [columnIndex: string]: string };
};

type SectionLoop = {
  var: string;
  over: string;
  /** Text identifying the first paragraph of the repeated block. */
  startText: string;
  /** Text identifying its last paragraph; defaults to `startText`. */
  endText?: string;
};

type Conditional = {
  /** JS expression; the block is included when it is truthy. */
  expr: string;
  /** Text identifying the first paragraph of the conditional block. */
  paragraphText: string;
  /** Text identifying its last paragraph; defaults to `paragraphText`. */
  endParagraphText?: string;
};

export type Mapping = {
  /** `'+++'`, or a `['{', '}']`-style pair. Defaults to `'+++'`. */
  cmdDelimiter?: string | [string, string];
  /**
   * Localized command names, for templates rendered with `commandAliases`,
   * e.g. `{ "FOR": "ДЛЯ", "IN": "ИЗ", "END-FOR": "КОНЕЦ ДЛЯ" }`.
   */
  commandNames?: { [builtIn: string]: string };
  /** Exact document text -> JS expression to insert in its place. */
  variables?: { [text: string]: string };
  tableLoops?: TableLoop[];
  sectionLoops?: SectionLoop[];
  conditionals?: Conditional[];
};

const DEFAULT_COMMAND_NAMES = {
  INS: 'INS',
  FOR: 'FOR',
  IN: 'IN',
  'END-FOR': 'END-FOR',
  IF: 'IF',
  'END-IF': 'END-IF',
} as const;

type CommandNames = { [K in keyof typeof DEFAULT_COMMAND_NAMES]: string };

/** Renders the commands this tool emits, honouring any localized names. */
class Commands {
  private readonly names: CommandNames;
  private readonly delimiter: [string, string];

  constructor(mapping: Mapping, delimiter: [string, string]) {
    this.names = { ...DEFAULT_COMMAND_NAMES, ...mapping.commandNames };
    this.delimiter = delimiter;
  }

  ins(expr: string): string {
    return command(`${this.names.INS} ${expr}`, this.delimiter);
  }

  for(loopVar: string, over: string): string {
    return command(
      `${this.names.FOR} ${loopVar} ${this.names.IN} ${over}`,
      this.delimiter
    );
  }

  endFor(loopVar: string): string {
    return command(`${this.names['END-FOR']} ${loopVar}`, this.delimiter);
  }

  if(expr: string): string {
    return command(`${this.names.IF} ${expr}`, this.delimiter);
  }

  endIf(): string {
    return command(this.names['END-IF'], this.delimiter);
  }
}

/**
 * The alias options that make a template written with `commandNames` readable
 * again — the inverse of the mapping. `IN` is not a command but a keyword
 * inside a `FOR` expression, so it belongs with the operator aliases.
 *
 * Save the result as `aliases.json` next to the template: `analyze`, `verify`
 * and `createReport` all need it to recognise the localized commands.
 */
export function aliasesFromCommandNames(mapping: Mapping): AliasOptions {
  const commandAliases: { [alias: string]: string } = {};
  const operatorAliases: { [alias: string]: string } = {};

  for (const [builtIn, name] of Object.entries(mapping.commandNames ?? {})) {
    // `=` is already a built-in shorthand for INS, and an alias that spells a
    // command the way it is spelled anyway is just noise.
    if (name === builtIn || name === '=' || name === '!') continue;
    if (builtIn === 'IN') operatorAliases[name] = 'IN';
    else commandAliases[name] = builtIn;
  }

  return {
    ...(Object.keys(commandAliases).length > 0 ? { commandAliases } : {}),
    ...(Object.keys(operatorAliases).length > 0 ? { operatorAliases } : {}),
  };
}

export const normalizeDelimiter = (
  delimiter: Mapping['cmdDelimiter']
): [string, string] => {
  if (delimiter == null) return [DEFAULT_CMD_DELIMITER, DEFAULT_CMD_DELIMITER];
  return typeof delimiter === 'string' ? [delimiter, delimiter] : delimiter;
};

// ==========================================
// Table loops
// ==========================================

/**
 * Collapses the data rows of a table into a single templated row, wrapped by a
 * FOR row above and an END-FOR row below.
 */
function applyTableLoop(root: Node, loop: TableLoop, cmds: Commands): string[] {
  const log: string[] = [];
  const table = topLevelTables(root)[loop.tableIndex];
  if (table == null) {
    throw new Error(
      `The document has no table with index ${loop.tableIndex}. ` +
        `Table indices come from analyze.mjs and skip nested tables.`
    );
  }

  const rows = directChildren(table, 'w:tr');
  const startRow = loop.startRow ?? 1;
  const endRow = Math.min(loop.endRow ?? rows.length - 1, rows.length - 1);
  const templateRow = rows[startRow];
  if (templateRow == null) {
    throw new Error(
      `Table ${loop.tableIndex} has no row ${startRow} (it has ${rows.length})`
    );
  }

  // 1. Templatise the row that survives.
  const cells = directChildren(templateRow, 'w:tc');
  for (const [columnIndex, expr] of Object.entries(loop.fields)) {
    const cell = cells[Number(columnIndex)];
    if (cell == null) {
      throw new Error(
        `Table ${loop.tableIndex} row ${startRow} has no column ${columnIndex}`
      );
    }
    const paragraphs = directChildren(cell, 'w:p');
    const target =
      paragraphs.find(p => textOf(p).trim() !== '') ?? paragraphs[0];
    if (target == null) {
      throw new Error(
        `Table ${loop.tableIndex} row ${startRow} column ${columnIndex} ` +
          `contains no paragraph to replace`
      );
    }
    // Replacing the existing text (rather than rebuilding the paragraph) keeps
    // the run's formatting, so a bold total stays bold.
    const existing = textOf(target);
    if (existing !== '') replaceText(target, existing, cmds.ins(expr));
    else setParagraphText(target, cmds.ins(expr));
    log.push(`table ${loop.tableIndex} column ${columnIndex} -> ${expr}`);
  }

  // 2. Drop the remaining data rows; the loop regenerates them.
  for (let i = endRow; i > startRow; i -= 1) {
    const row = rows[i];
    if (row != null) remove(row);
  }

  // 3. Wrap in FOR / END-FOR rows, modelled on the row they surround.
  insertBefore(
    templateRow,
    commandRowLike(templateRow, cmds.for(loop.var, loop.over))
  );
  insertAfter(templateRow, commandRowLike(templateRow, cmds.endFor(loop.var)));
  log.push(
    `table ${loop.tableIndex} rows ${startRow}..${endRow} -> ` +
      `FOR ${loop.var} IN ${loop.over}`
  );
  return log;
}

// ==========================================
// Generation
// ==========================================

const truncate = (text: string, max = 40): string =>
  text.length <= max ? text : `${text.slice(0, max - 1)}…`;

/**
 * Applies `mapping` to the parsed document parts, in place. Returns a log of
 * what was changed, one line per edit.
 */
export function applyMapping(parts: DocPart[], mapping: Mapping): string[] {
  const cmds = new Commands(mapping, normalizeDelimiter(mapping.cmdDelimiter));
  const log: string[] = [];
  const mainPart = parts.find(part => part.isMain) ?? parts[0];
  if (mainPart == null) throw new Error('Document has no main part');
  const mainRoot = mainPart.root;

  // Order matters. Structural edits are located by document text, so they all
  // have to happen while that text is still the original prose — once a phrase
  // has become `+++INS ...+++`, nothing will match it again.
  for (const loop of mapping.tableLoops ?? []) {
    log.push(...applyTableLoop(mainRoot, loop, cmds));
  }

  for (const loop of mapping.sectionLoops ?? []) {
    wrapParagraphs(mainRoot, {
      startText: loop.startText,
      endText: loop.endText ?? loop.startText,
      openCmd: cmds.for(loop.var, loop.over),
      closeCmd: cmds.endFor(loop.var),
      what: `sectionLoop ${loop.var}`,
    });
    log.push(`section loop -> FOR ${loop.var} IN ${loop.over}`);
  }

  for (const cond of mapping.conditionals ?? []) {
    wrapParagraphs(mainRoot, {
      startText: cond.paragraphText,
      endText: cond.endParagraphText ?? cond.paragraphText,
      openCmd: cmds.if(cond.expr),
      closeCmd: cmds.endIf(),
      what: `conditional ${cond.expr}`,
    });
    log.push(`conditional -> IF ${cond.expr}`);
  }

  // Longest first: otherwise mapping both "Acme" and "Acme Corp" would leave a
  // stray "Corp" behind. Variables are the only edit applied to headers and
  // footers as well as to the body.
  const variables = Object.entries(mapping.variables ?? {}).sort(
    ([a], [b]) => b.length - a.length
  );
  for (const [text, expr] of variables) {
    let count = 0;
    for (const part of parts) {
      count += replaceTextAll(part.root, text, cmds.ins(expr));
    }
    if (count === 0) {
      throw new Error(
        `Variable "${text}" was not found in the document. Copy the text ` +
          `exactly as analyze.mjs reports it, including punctuation.`
      );
    }
    log.push(`${count}x "${truncate(text)}" -> ${expr}`);
  }

  return log;
}

// ==========================================
// Main
// ==========================================

async function main(): Promise<void> {
  const { positional } = parseArgs(process.argv.slice(2));
  const [originalArg, mappingArg, outputArg] = positional;
  if (originalArg == null || mappingArg == null) {
    console.error(
      'Usage: node generate.mjs <original.docx> <mapping.json> [output.docx]'
    );
    process.exit(1);
  }

  const originalPath = requireFile(originalArg);
  const mapping = readJson<Mapping>(requireFile(mappingArg));
  const outputPath = path.resolve(
    outputArg ?? originalPath.replace(/\.docx$/i, '_template.docx')
  );

  const { zip, parts } = await readDocParts(originalPath);
  for (const line of applyMapping(parts, mapping)) console.log(`  ${line}`);
  await writeDocParts(zip, parts, outputPath);
  console.log(`\nTemplate written to ${outputPath}`);

  // Parse the result the way `createReport` would, so a mapping that produced
  // nonsense fails here rather than at render time.
  const aliases = aliasesFromCommandNames(mapping);
  const commands = await listCommands(
    fs.readFileSync(outputPath),
    normalizeDelimiter(mapping.cmdDelimiter),
    aliases
  );
  console.log(`\n${commands.length} command(s) in the template:`);
  for (const cmd of commands) console.log(`  ${cmd.type}: ${cmd.code}`);

  if (Object.keys(aliases).length > 0) {
    // Without these, neither the other tools nor `createReport` can tell the
    // localized commands from ordinary prose.
    const aliasPath = outputPath.replace(/\.docx$/i, '_aliases.json');
    fs.writeFileSync(aliasPath, `${JSON.stringify(aliases, null, 2)}\n`);
    console.log(`\nCommand aliases written to ${aliasPath}`);
    console.log('Pass them with --aliases to analyze.mjs and verify.mjs.');
  }

  console.log(
    '\nNext: write sample data and run verify.mjs to render the template.'
  );
}

if (isEntryPoint(import.meta.url)) runMain('Generation', main);
