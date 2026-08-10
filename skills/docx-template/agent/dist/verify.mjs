import fs from 'node:fs';
import path from 'node:path';
import { pathToFileURL } from 'node:url';
import { i as isEntryPoint, r as runMain, e as parseArgs, j as requireFile, g as readDelimiter, h as readAliases, l as listCommands, E as createReport, k as readJson, a as readDocParts, p as plainTextLines } from './shared-docxXml.mjs';
import 'node:vm';
import 'stream';
import 'events';
import 'buffer';
import 'util';

const SMART_QUOTES = /[‘’‚“”„]/;
const loopVarOf = (code) => code.trim().split(/\s+/)[0] ?? "";
function checkCommands(commands) {
  const problems = [];
  const stack = [];
  for (const cmd of commands) {
    switch (cmd.type) {
      case "FOR":
        stack.push({ type: "FOR", label: loopVarOf(cmd.code) });
        break;
      case "END-FOR": {
        const open = stack.pop();
        const closing = loopVarOf(cmd.code);
        if (open == null || open.type !== "FOR") {
          problems.push(
            `${["END-FOR", closing].join(" ").trim()} closes a loop that was never opened`
          );
        } else if (closing !== "" && closing !== open.label) {
          problems.push(
            `END-FOR ${closing} closes FOR ${open.label} \u2014 the variable names differ`
          );
        }
        break;
      }
      case "IF":
        stack.push({ type: "IF", label: cmd.code });
        break;
      case "END-IF": {
        const open = stack.pop();
        if (open == null || open.type !== "IF") {
          problems.push("END-IF closes a conditional that was never opened");
        }
        break;
      }
      case "ELSE":
      case "ELSE-IF": {
        const open = stack[stack.length - 1];
        if (open == null || open.type !== "IF") {
          problems.push(`${cmd.type} appears outside an IF \u2026 END-IF block`);
        }
        break;
      }
    }
    if (SMART_QUOTES.test(cmd.code)) {
      problems.push(
        `${cmd.type} ${cmd.code} contains a curly quote. Word autocorrects straight quotes; retype them, or render with fixSmartQuotes: true.`
      );
    }
  }
  for (const open of stack.reverse()) {
    problems.push(
      open.type === "FOR" ? `FOR ${open.label} is never closed by an END-FOR` : `IF ${open.label} is never closed by an END-IF`
    );
  }
  return problems;
}
async function loadContext(file) {
  const mod = await import(pathToFileURL(requireFile(file)).href);
  const context = mod.default;
  if (context == null || typeof context !== "object") {
    throw new Error(`${file} must default-export an object`);
  }
  return context;
}
async function main() {
  const { positional, flags } = parseArgs(process.argv.slice(2), [
    "allow-nullish",
    "quiet"
  ]);
  const [templateArg, dataArg] = positional;
  if (templateArg == null) {
    console.error(
      "Usage: node verify.mjs <template.docx> [data.json] [-o report.docx] [--delimiter +++] [--aliases aliases.json] [--context context.mjs] [--allow-nullish] [--quiet]"
    );
    process.exit(1);
  }
  const templatePath = requireFile(templateArg);
  const template = fs.readFileSync(templatePath);
  const delimiter = readDelimiter(flags);
  const aliases = readAliases(flags);
  const commands = await listCommands(template, delimiter, aliases);
  console.log(`${commands.length} command(s):`);
  for (const cmd of commands) console.log(`  ${cmd.type}: ${cmd.code}`);
  const problems = checkCommands(commands);
  if (problems.length > 0) {
    console.error(`
${problems.length} problem(s) in the template:`);
    for (const problem of problems) console.error(`  - ${problem}`);
    process.exit(1);
  }
  if (dataArg == null) {
    console.log("\nCommands are balanced. Pass a data JSON file to render.");
    return;
  }
  const dataPath = requireFile(dataArg);
  const outFlag = flags.o ?? flags.out;
  const outputPath = path.resolve(
    typeof outFlag === "string" ? outFlag : templatePath.replace(/\.docx$/i, "_report.docx")
  );
  const contextFlag = flags.context;
  const report = await createReport({
    template,
    data: readJson(dataPath),
    cmdDelimiter: delimiter,
    // Collect every bad command in one pass instead of stopping at the first;
    // `createReport` then rejects with an array, which `runMain` unpacks.
    failFast: false,
    // A command that evaluates to null is almost always a typo in the
    // expression or a missing key in the data, not an intentional blank.
    rejectNullish: flags["allow-nullish"] !== true,
    ...typeof contextFlag === "string" ? { additionalJsContext: await loadContext(contextFlag) } : {},
    ...aliases
  });
  fs.mkdirSync(path.dirname(outputPath), { recursive: true });
  fs.writeFileSync(outputPath, report);
  console.log(`
Rendered with ${path.basename(dataPath)} -> ${outputPath}`);
  if (flags.quiet === true) return;
  const { parts } = await readDocParts(outputPath);
  const mainPart = parts.find((part) => part.isMain) ?? parts[0];
  if (mainPart == null) return;
  console.log("\n----- rendered text -----");
  console.log(plainTextLines(mainPart.root).join("\n"));
}
if (isEntryPoint(import.meta.url)) runMain("Verification", main);

export { checkCommands };
