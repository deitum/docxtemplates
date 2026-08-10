import fs from 'node:fs';
import path from 'node:path';
import { i as isEntryPoint, r as runMain, a as readDocParts, p as plainTextLines, l as listCommands, t as topLevelTables, b as paragraphsOutsideTables, d as directChildren, c as textOf, f as firstChild, e as parseArgs, g as readDelimiter, h as readAliases, j as requireFile } from './shared-docxXml.mjs';
import 'node:url';
import 'stream';
import 'events';
import 'buffer';
import 'util';
import 'node:vm';

function attr(node, name) {
  const value = node._attrs[name];
  if (value == null) return void 0;
  return typeof value === "string" ? value : value.value;
}
function toggleOn(rPr, tag) {
  const el = firstChild(rPr, tag);
  if (el == null) return false;
  const val = attr(el, "w:val");
  return val !== "0" && val !== "false";
}
function childAttr(parent, tag, attrName = "w:val") {
  const child = firstChild(parent, tag);
  return child != null ? attr(child, attrName) : void 0;
}
function runInfo(run) {
  const info = { text: textOf(run) };
  const rPr = firstChild(run, "w:rPr");
  if (rPr == null) return info;
  if (toggleOn(rPr, "w:b")) info.bold = true;
  if (toggleOn(rPr, "w:i")) info.italic = true;
  if (toggleOn(rPr, "w:u")) info.underline = true;
  const size = childAttr(rPr, "w:sz");
  if (size != null) info.size = size;
  const font = childAttr(rPr, "w:rFonts", "w:ascii");
  if (font != null) info.font = font;
  const color = childAttr(rPr, "w:color");
  if (color != null) info.color = color;
  return info;
}
function paragraphInfo(paragraph, index) {
  const info = {
    index,
    text: textOf(paragraph),
    runs: directChildren(paragraph, "w:r").map(runInfo).filter((run) => run.text !== "")
  };
  const pPr = firstChild(paragraph, "w:pPr");
  if (pPr != null) {
    const style = childAttr(pPr, "w:pStyle");
    if (style != null) info.style = style;
    const numPr = firstChild(pPr, "w:numPr");
    if (numPr != null) info.listLevel = childAttr(numPr, "w:ilvl") ?? "0";
  }
  return info;
}
function tableInfo(table, index) {
  const rows = directChildren(table, "w:tr").map((row) => ({
    cells: directChildren(row, "w:tc").map((cell) => ({
      text: textOf(cell),
      paragraphs: directChildren(cell, "w:p").map(paragraphInfo)
    }))
  }));
  const firstRow = rows[0];
  const headerRow = rows.length > 1 && firstRow != null && firstRow.cells.length > 0 && firstRow.cells.every(
    (cell) => cell.paragraphs.some((p) => p.runs.length > 0 && p.runs.every((r) => r.bold))
  );
  return { index, headerRow, rows };
}
const partInfo = (filename, root) => ({
  filename,
  paragraphs: paragraphsOutsideTables(root).map(paragraphInfo),
  tables: topLevelTables(root).map(tableInfo)
});
async function analyze(filePath, delimiter, aliases) {
  const { parts } = await readDocParts(filePath);
  const main2 = parts.find((part) => part.isMain) ?? parts[0];
  if (main2 == null) throw new Error(`${filePath} has no readable content`);
  const analysis = {
    file: path.basename(filePath),
    plainText: plainTextLines(main2.root).join("\n"),
    main: partInfo(main2.filename, main2.root),
    headersAndFooters: parts.filter((part) => !part.isMain).map((part) => partInfo(part.filename, part.root)),
    commands: []
  };
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
async function main() {
  const { positional, flags } = parseArgs(process.argv.slice(2), ["text-only"]);
  if (positional.length === 0) {
    console.error(
      "Usage: node analyze.mjs <file.docx> [file2.docx ...] [--aliases aliases.json] [--delimiter +++] [--text-only]"
    );
    process.exit(1);
  }
  const delimiter = readDelimiter(flags);
  const aliases = readAliases(flags);
  const analyses = [];
  for (const arg of positional) {
    analyses.push(await analyze(requireFile(arg), delimiter, aliases));
  }
  if (flags["text-only"] === true) {
    for (const analysis of analyses) {
      if (analyses.length > 1) console.log(`
===== ${analysis.file} =====`);
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
if (isEntryPoint(import.meta.url)) runMain("Analysis", main);

export { analyze };
