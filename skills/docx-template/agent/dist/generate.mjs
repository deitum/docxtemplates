import fs from 'node:fs';
import path from 'node:path';
import { i as isEntryPoint, r as runMain, e as parseArgs, j as requireFile, k as readJson, a as readDocParts, w as writeDocParts, l as listCommands, m as wrapParagraphs, o as replaceTextAll, D as DEFAULT_CMD_DELIMITER, n as command, t as topLevelTables, d as directChildren, c as textOf, z as replaceText, A as setParagraphText, q as remove, B as insertBefore, C as commandRowLike, s as insertAfter } from './shared-docxXml.mjs';
import 'node:url';
import 'stream';
import 'events';
import 'buffer';
import 'util';
import 'node:vm';

const DEFAULT_COMMAND_NAMES = {
  INS: "INS",
  FOR: "FOR",
  IN: "IN",
  "END-FOR": "END-FOR",
  IF: "IF",
  "END-IF": "END-IF"
};
class Commands {
  names;
  delimiter;
  constructor(mapping, delimiter) {
    this.names = { ...DEFAULT_COMMAND_NAMES, ...mapping.commandNames };
    this.delimiter = delimiter;
  }
  ins(expr) {
    return command(`${this.names.INS} ${expr}`, this.delimiter);
  }
  for(loopVar, over) {
    return command(
      `${this.names.FOR} ${loopVar} ${this.names.IN} ${over}`,
      this.delimiter
    );
  }
  endFor(loopVar) {
    return command(`${this.names["END-FOR"]} ${loopVar}`, this.delimiter);
  }
  if(expr) {
    return command(`${this.names.IF} ${expr}`, this.delimiter);
  }
  endIf() {
    return command(this.names["END-IF"], this.delimiter);
  }
}
function aliasesFromCommandNames(mapping) {
  const commandAliases = {};
  const operatorAliases = {};
  for (const [builtIn, name] of Object.entries(mapping.commandNames ?? {})) {
    if (name === builtIn || name === "=" || name === "!") continue;
    if (builtIn === "IN") operatorAliases[name] = "IN";
    else commandAliases[name] = builtIn;
  }
  return {
    ...Object.keys(commandAliases).length > 0 ? { commandAliases } : {},
    ...Object.keys(operatorAliases).length > 0 ? { operatorAliases } : {}
  };
}
const normalizeDelimiter = (delimiter) => {
  if (delimiter == null) return [DEFAULT_CMD_DELIMITER, DEFAULT_CMD_DELIMITER];
  return typeof delimiter === "string" ? [delimiter, delimiter] : delimiter;
};
function applyTableLoop(root, loop, cmds) {
  const log = [];
  const table = topLevelTables(root)[loop.tableIndex];
  if (table == null) {
    throw new Error(
      `The document has no table with index ${loop.tableIndex}. Table indices come from analyze.mjs and skip nested tables.`
    );
  }
  const rows = directChildren(table, "w:tr");
  const startRow = loop.startRow ?? 1;
  const endRow = Math.min(loop.endRow ?? rows.length - 1, rows.length - 1);
  const templateRow = rows[startRow];
  if (templateRow == null) {
    throw new Error(
      `Table ${loop.tableIndex} has no row ${startRow} (it has ${rows.length})`
    );
  }
  const cells = directChildren(templateRow, "w:tc");
  for (const [columnIndex, expr] of Object.entries(loop.fields)) {
    const cell = cells[Number(columnIndex)];
    if (cell == null) {
      throw new Error(
        `Table ${loop.tableIndex} row ${startRow} has no column ${columnIndex}`
      );
    }
    const paragraphs = directChildren(cell, "w:p");
    const target = paragraphs.find((p) => textOf(p).trim() !== "") ?? paragraphs[0];
    if (target == null) {
      throw new Error(
        `Table ${loop.tableIndex} row ${startRow} column ${columnIndex} contains no paragraph to replace`
      );
    }
    const existing = textOf(target);
    if (existing !== "") replaceText(target, existing, cmds.ins(expr));
    else setParagraphText(target, cmds.ins(expr));
    log.push(`table ${loop.tableIndex} column ${columnIndex} -> ${expr}`);
  }
  for (let i = endRow; i > startRow; i -= 1) {
    const row = rows[i];
    if (row != null) remove(row);
  }
  insertBefore(
    templateRow,
    commandRowLike(templateRow, cmds.for(loop.var, loop.over))
  );
  insertAfter(templateRow, commandRowLike(templateRow, cmds.endFor(loop.var)));
  log.push(
    `table ${loop.tableIndex} rows ${startRow}..${endRow} -> FOR ${loop.var} IN ${loop.over}`
  );
  return log;
}
const truncate = (text, max = 40) => text.length <= max ? text : `${text.slice(0, max - 1)}\u2026`;
function applyMapping(parts, mapping) {
  const cmds = new Commands(mapping, normalizeDelimiter(mapping.cmdDelimiter));
  const log = [];
  const mainPart = parts.find((part) => part.isMain) ?? parts[0];
  if (mainPart == null) throw new Error("Document has no main part");
  const mainRoot = mainPart.root;
  for (const loop of mapping.tableLoops ?? []) {
    log.push(...applyTableLoop(mainRoot, loop, cmds));
  }
  for (const loop of mapping.sectionLoops ?? []) {
    wrapParagraphs(mainRoot, {
      startText: loop.startText,
      endText: loop.endText ?? loop.startText,
      openCmd: cmds.for(loop.var, loop.over),
      closeCmd: cmds.endFor(loop.var),
      what: `sectionLoop ${loop.var}`
    });
    log.push(`section loop -> FOR ${loop.var} IN ${loop.over}`);
  }
  for (const cond of mapping.conditionals ?? []) {
    wrapParagraphs(mainRoot, {
      startText: cond.paragraphText,
      endText: cond.endParagraphText ?? cond.paragraphText,
      openCmd: cmds.if(cond.expr),
      closeCmd: cmds.endIf(),
      what: `conditional ${cond.expr}`
    });
    log.push(`conditional -> IF ${cond.expr}`);
  }
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
        `Variable "${text}" was not found in the document. Copy the text exactly as analyze.mjs reports it, including punctuation.`
      );
    }
    log.push(`${count}x "${truncate(text)}" -> ${expr}`);
  }
  return log;
}
async function main() {
  const { positional } = parseArgs(process.argv.slice(2));
  const [originalArg, mappingArg, outputArg] = positional;
  if (originalArg == null || mappingArg == null) {
    console.error(
      "Usage: node generate.mjs <original.docx> <mapping.json> [output.docx]"
    );
    process.exit(1);
  }
  const originalPath = requireFile(originalArg);
  const mapping = readJson(requireFile(mappingArg));
  const outputPath = path.resolve(
    outputArg ?? originalPath.replace(/\.docx$/i, "_template.docx")
  );
  const { zip, parts } = await readDocParts(originalPath);
  for (const line of applyMapping(parts, mapping)) console.log(`  ${line}`);
  await writeDocParts(zip, parts, outputPath);
  console.log(`
Template written to ${outputPath}`);
  const aliases = aliasesFromCommandNames(mapping);
  const commands = await listCommands(
    fs.readFileSync(outputPath),
    normalizeDelimiter(mapping.cmdDelimiter),
    aliases
  );
  console.log(`
${commands.length} command(s) in the template:`);
  for (const cmd of commands) console.log(`  ${cmd.type}: ${cmd.code}`);
  if (Object.keys(aliases).length > 0) {
    const aliasPath = outputPath.replace(/\.docx$/i, "_aliases.json");
    fs.writeFileSync(aliasPath, `${JSON.stringify(aliases, null, 2)}
`);
    console.log(`
Command aliases written to ${aliasPath}`);
    console.log("Pass them with --aliases to analyze.mjs and verify.mjs.");
  }
  console.log(
    "\nNext: write sample data and run verify.mjs to render the template."
  );
}
if (isEntryPoint(import.meta.url)) runMain("Generation", main);

export { aliasesFromCommandNames, applyMapping, normalizeDelimiter };
