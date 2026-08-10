import path from 'node:path';
import { i as isEntryPoint, r as runMain, e as parseArgs, j as requireFile, k as readJson, a as readDocParts, w as writeDocParts, m as wrapParagraphs, n as command, o as replaceTextAll, t as topLevelTables, d as directChildren, q as remove, s as insertAfter, f as firstChild, u as isTag, c as textOf, v as newElement, x as indexText, y as spliceText } from './shared-docxXml.mjs';
import { normalizeDelimiter } from './generate.mjs';
import 'node:vm';
import 'node:fs';
import 'node:url';
import 'stream';
import 'events';
import 'buffer';
import 'util';

function findCommandSpans(index, delimiter) {
  const [open, close] = delimiter;
  const spans = [];
  let at = 0;
  for (; ; ) {
    const start = index.full.indexOf(open, at);
    if (start < 0) break;
    const bodyStart = start + open.length;
    const closeAt = index.full.indexOf(close, bodyStart);
    if (closeAt < 0) break;
    spans.push({
      start,
      end: closeAt + close.length,
      code: index.full.slice(bodyStart, closeAt)
    });
    at = closeAt + close.length;
  }
  return spans;
}
function editCommands(root, delimiter, edit) {
  const index = indexText(root);
  const spans = findCommandSpans(index, delimiter);
  let count = 0;
  for (const span of spans.reverse()) {
    const replacement = edit(span.code);
    if (replacement === void 0) continue;
    spliceText(index, span.start, span.end, replacement);
    count += 1;
  }
  return count;
}
const isFloating = (table) => {
  const tblPr = firstChild(table, "w:tblPr");
  return tblPr != null && firstChild(tblPr, "w:tblpPr") != null;
};
function clearBorders(table) {
  const tblPr = firstChild(table, "w:tblPr");
  if (tblPr == null) return;
  const existing = firstChild(tblPr, "w:tblBorders");
  if (existing != null) remove(existing);
  const sides = ["top", "left", "bottom", "right", "insideH", "insideV"];
  const borders = newElement(
    "w:tblBorders",
    {},
    sides.map((side) => newElement(`w:${side}`, { "w:val": "none", "w:sz": "0" }))
  );
  borders._parent = tblPr;
  tblPr._children.push(borders);
}
function mergeFloatingTables(root, indices) {
  const all = topLevelTables(root);
  const targets = indices != null ? indices.map((i) => {
    const table = all[i];
    if (table == null) throw new Error(`No table with index ${i}`);
    return table;
  }) : all.filter(isFloating);
  const [base, ...rest] = targets;
  if (base == null || rest.length === 0) {
    return "mergeFloatingTables: fewer than two tables matched, nothing to do";
  }
  for (const table of rest) {
    removeEmptyParagraphsBetween(base, table);
    const baseRows = directChildren(base, "w:tr");
    const rows = directChildren(table, "w:tr");
    rows.forEach((row, i) => {
      const target = baseRows[i];
      if (target == null) {
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
      for (const cell of directChildren(row, "w:tc")) {
        remove(cell);
        cell._parent = target;
        target._children.push(cell);
      }
    });
    const baseGrid = firstChild(base, "w:tblGrid");
    const grid = firstChild(table, "w:tblGrid");
    if (baseGrid != null && grid != null) {
      for (const col of directChildren(grid, "w:gridCol")) {
        remove(col);
        col._parent = baseGrid;
        baseGrid._children.push(col);
      }
    }
    remove(table);
  }
  const tblPr = firstChild(base, "w:tblPr");
  const tblpPr = tblPr != null ? firstChild(tblPr, "w:tblpPr") : void 0;
  if (tblpPr != null) remove(tblpPr);
  clearBorders(base);
  return `merged ${targets.length} floating tables into one inline table`;
}
function removeEmptyParagraphsBetween(a, b) {
  const parent = a._parent;
  if (parent == null || b._parent !== parent) return;
  const from = parent._children.indexOf(a);
  const to = parent._children.indexOf(b);
  if (from < 0 || to < 0 || to < from) return;
  for (const node of parent._children.slice(from + 1, to)) {
    if (isTag(node, "w:p") && textOf(node).trim() === "") remove(node);
  }
}
function applyModifications(parts, file) {
  const delimiter = normalizeDelimiter(file.cmdDelimiter);
  const mainPart = parts.find((part) => part.isMain) ?? parts[0];
  if (mainPart == null) throw new Error("Document has no main part");
  const mainRoot = mainPart.root;
  const log = [];
  const overParts = (edit) => parts.reduce(
    (total, part) => total + editCommands(part.root, delimiter, edit),
    0
  );
  for (const mod of file.modifications ?? []) {
    switch (mod.type) {
      case "replaceCommand": {
        const n = overParts(
          (code) => code.trim() === mod.from.trim() ? command(mod.to, delimiter) : void 0
        );
        log.push(`${n}x ${mod.from} -> ${mod.to}`);
        break;
      }
      case "renameExpression": {
        const n = overParts((code) => {
          const renamed = code.split(mod.from).join(mod.to);
          return renamed === code ? void 0 : command(renamed, delimiter);
        });
        log.push(`${n} command(s): ${mod.from} -> ${mod.to}`);
        break;
      }
      case "removeCommand": {
        const n = overParts(
          (code) => code.trim() === mod.code.trim() ? mod.replaceWith ?? "" : void 0
        );
        log.push(`${n}x removed ${mod.code}`);
        break;
      }
      case "addCommand": {
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
      case "wrapFor":
        wrapParagraphs(mainRoot, {
          startText: mod.startText,
          endText: mod.endText ?? mod.startText,
          openCmd: command(`FOR ${mod.var} IN ${mod.over}`, delimiter),
          closeCmd: command(`END-FOR ${mod.var}`, delimiter),
          what: `wrapFor ${mod.var}`
        });
        log.push(`wrapped block in FOR ${mod.var} IN ${mod.over}`);
        break;
      case "wrapIf":
        wrapParagraphs(mainRoot, {
          startText: mod.startText,
          endText: mod.endText ?? mod.startText,
          openCmd: command(`IF ${mod.expr}`, delimiter),
          closeCmd: command("END-IF", delimiter),
          what: `wrapIf ${mod.expr}`
        });
        log.push(`wrapped block in IF ${mod.expr}`);
        break;
      case "mergeFloatingTables":
        log.push(mergeFloatingTables(mainRoot, mod.tableIndices));
        break;
      default: {
        const unknown = mod;
        throw new Error(`Unknown modification type "${unknown.type}"`);
      }
    }
  }
  return log;
}
async function main() {
  const { positional } = parseArgs(process.argv.slice(2));
  const [templateArg, modsArg, outputArg] = positional;
  if (templateArg == null || modsArg == null) {
    console.error(
      "Usage: node refine.mjs <template.docx> <modifications.json> [output.docx]"
    );
    process.exit(1);
  }
  const templatePath = requireFile(templateArg);
  const mods = readJson(requireFile(modsArg));
  const outputPath = path.resolve(
    outputArg ?? templatePath.replace(/\.docx$/i, "_refined.docx")
  );
  const { zip, parts } = await readDocParts(templatePath);
  for (const line of applyModifications(parts, mods)) console.log(`  ${line}`);
  await writeDocParts(zip, parts, outputPath);
  console.log(`
Template written to ${outputPath}`);
  console.log("Next: run verify.mjs to confirm it still renders.");
}
if (isEntryPoint(import.meta.url)) runMain("Refinement", main);

export { applyModifications };
