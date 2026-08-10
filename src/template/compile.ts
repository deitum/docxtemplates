/**
 * Reading a template's commands out of the tree, once, without rendering it.
 *
 * Three separate passes used to do variations of this. `extractQuery` walked
 * the tree with a `seekingQuery` flag threaded down into the hot path of
 * `processText`, where every command was resolved only to be thrown away.
 * `listCommands` ran the full `walkTemplate` with a stand-in command processor
 * that reached into private state to flush the collected command. And the
 * render itself did it a third time.
 *
 * The scan is the same in all three cases: go over the text nodes in document
 * order and split them on the delimiters, remembering across nodes whether the
 * cursor is inside a command — Word splits a command across runs freely, which
 * is why the state cannot be per-node.
 */
import { getCommand, splitCommand } from '../commands/parse';
import { WTag } from '../ooxml';
import { nextNodeInTree, tagOf } from '../reportUtils';
import {
  Command,
  type CreateReportOptions,
  type Node,
  type TextNode,
} from '../types';

/** One command, as written in the template. */
export type CommandSite = {
  /** The text node the command's closing delimiter was found in. */
  node: TextNode;
  /** The text between the delimiters, before any resolution. */
  raw: string;
};

/** What a template part contains, found without rendering it. */
export type TemplateProgram = {
  /** Every command in the part, in document order. */
  commands: CommandSite[];
};

const isTextNodeInsideWt = (node: Node): node is TextNode =>
  node._fTextNode && tagOf(node._parent) === WTag.t;

/**
 * Finds every command in a part.
 *
 * Deliberately does no resolution: `*shorthand` expansion depends on the
 * `ALIAS` commands that ran before it, and which of those have run differs
 * between listing a template's commands and rendering it. The caller decides.
 */
export function compileTemplate(
  template: Node,
  cmdDelimiter: [string, string]
): TemplateProgram {
  const commands: CommandSite[] = [];
  const [open, close] = cmdDelimiter;

  // Carried across nodes on purpose: Word splits a command over several runs
  // whenever it feels like it, and `preprocessTemplate` only rejoins the pieces
  // it can.
  let collecting = false;
  let raw = '';

  let node: Node | null = template;
  while ((node = nextNodeInTree(node)) != null) {
    if (!isTextNodeInsideWt(node)) continue;
    const text = node._text;
    if (text == null || text === '') continue;

    const segments = text.split(open).flatMap(s => s.split(close));
    for (let idx = 0; idx < segments.length; idx++) {
      if (collecting) raw += segments[idx] ?? '';
      // A delimiter follows: close the command being collected, then flip
      // between "inside a command" and "ordinary text".
      if (idx < segments.length - 1) {
        if (collecting) {
          commands.push({ node, raw });
          raw = '';
        }
        collecting = !collecting;
      }
    }
  }

  return { commands };
}

/**
 * No shorthands. Resolving a command outside of a render means no `ALIAS`
 * command has run, so `*shorthand` is the unknown alias it is at that point —
 * which is what both callers below want, and what they have always done.
 */
const NO_SHORTHANDS: { [shorthand: string]: string } = Object.create(null);

/** A command site, resolved into a command name and its payload. */
export function resolveSite(
  site: CommandSite,
  options: CreateReportOptions
): { raw: string; name: string | undefined; code: string } {
  const raw = getCommand(site.raw, NO_SHORTHANDS, options);
  const { cmdName, cmdRest } = splitCommand(raw, options.operatorAliases);
  return { raw, name: cmdName, code: cmdRest };
}

/**
 * The payload of the template's QUERY command, which `createReport` passes to a
 * `data` function so that it can fetch exactly what the report needs.
 *
 * Stops at the first one, as it always has: a second QUERY is ignored rather
 * than reported.
 */
export function extractQuery(
  template: Node,
  options: CreateReportOptions
): string | undefined {
  for (const site of compileTemplate(template, options.cmdDelimiter).commands) {
    const { name, code } = resolveSite(site, options);
    if (name === Command.QUERY) return code;
  }
  return undefined;
}
