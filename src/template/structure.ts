/**
 * The FOR/IF nesting of a template, worked out from its commands rather than
 * from running it.
 *
 * The engine discovers this shape as it goes, with a stack of open constructs
 * that is pushed and popped by the commands themselves. That works, but it
 * means the shape of a template cannot be looked at without rendering it, and
 * that a malformed one is only diagnosed at the point the walk trips over it.
 *
 * This computes the same shape up front. It reports what it finds instead of
 * throwing: the engine's own diagnosis stays the authority on what a template
 * is allowed to do, for the reasons set out in `analyzeStructure` below.
 */
import { Command, type CreateReportOptions } from '../types';
import { type CommandSite, resolveSite, type TemplateProgram } from './compile';

/** A FOR or IF construct, and where its parts are. */
export type Construct = {
  kind: Command.FOR | Command.IF;
  /** Index into `TemplateProgram.commands` of the opening command. */
  open: number;
  /** Index of the closing command; `null` if the template never closes it. */
  close: number | null;
  /**
   * Indices of the `ELSE-IF` / `ELSE` commands, in order. Together with `open`
   * and `close` these are the boundaries of the construct's branches.
   */
  branches: number[];
  /** How many constructs enclose this one. */
  depth: number;
  /** The loop variable, for a FOR. */
  varName?: string;
};

/** Something about the template's structure that does not add up. */
export type StructuralProblem = {
  id:
    | 'unclosed_for'
    | 'unclosed_if'
    | 'unexpected_end'
    | 'else_outside_if'
    | 'else_after_else';
  /** Index into `TemplateProgram.commands`. */
  at: number;
  /** The command, as written. */
  command: string;
};

export type TemplateStructure = {
  constructs: Construct[];
  problems: StructuralProblem[];
};

/** The name of the command at a site, or `undefined` if it cannot be read. */
const nameOf = (
  site: CommandSite,
  options: CreateReportOptions
): { name: string | undefined; raw: string } => {
  try {
    const { name, raw } = resolveSite(site, options);
    return { name, raw };
  } catch {
    // An unresolvable `*shorthand`: the ALIAS that defines it is only executed
    // during a render. It cannot open or close a construct here.
    return { name: undefined, raw: site.raw };
  }
};

/**
 * Pairs up the FOR/END-FOR and IF/ELSE-IF/ELSE/END-IF commands of a template.
 *
 * Note what this deliberately does *not* do: throw. It is tempting to move the
 * engine's structural errors here — the plan this work follows proposed exactly
 * that — but the errors a template produces are part of the contract, and with
 * `failFast: false` that includes the order they arrive in. Those errors are
 * currently interleaved with the runtime ones in the order the walk meets them:
 * a template with a bad expression, a stray END-FOR and an unterminated loop
 * reports the expression failures *first*. Diagnosing structure up front would
 * put the structural errors first instead, which is a different array.
 *
 * So this is an observer. It is what a future renderer that works from the
 * structure would consume, and it makes the nesting of a template testable on
 * its own; it does not decide what is an error.
 */
export function analyzeStructure(
  program: TemplateProgram,
  options: CreateReportOptions
): TemplateStructure {
  const constructs: Construct[] = [];
  const problems: StructuralProblem[] = [];
  /** Indices into `constructs` of the constructs currently open. */
  const open: number[] = [];
  /** Whether the innermost IF has already seen its unconditional ELSE. */
  const sawElse: boolean[] = [];

  const innermost = () => {
    const idx = open[open.length - 1];
    return idx == null ? undefined : constructs[idx];
  };

  program.commands.forEach((site, at) => {
    const { name, raw } = nameOf(site, options);

    switch (name) {
      case Command.FOR:
      case Command.IF: {
        const varName = /^(\S+)\s+IN\s+/i.exec(resolveRest(site, options))?.[1];
        constructs.push({
          kind: name === Command.FOR ? Command.FOR : Command.IF,
          open: at,
          close: null,
          branches: [],
          depth: open.length,
          ...(name === Command.FOR && varName != null ? { varName } : {}),
        });
        open.push(constructs.length - 1);
        sawElse.push(false);
        return;
      }

      case Command.ELSE_IF:
      case Command.ELSE: {
        const current = innermost();
        if (current == null || current.kind !== Command.IF) {
          problems.push({ id: 'else_outside_if', at, command: raw });
          return;
        }
        if (sawElse[sawElse.length - 1]) {
          problems.push({ id: 'else_after_else', at, command: raw });
          return;
        }
        current.branches.push(at);
        if (name === Command.ELSE) sawElse[sawElse.length - 1] = true;
        return;
      }

      case Command.END_FOR: {
        // `END-FOR name` closes the loop of that name. One that names an
        // enclosing loop rather than the innermost is a crossed pair; one that
        // names no open loop at all is ignored, because a construct that spans
        // several cells legitimately meets the END-FOR of an earlier part of
        // the same row.
        const varName = resolveRest(site, options).trim();
        const current = innermost();
        if (current == null) {
          // Nothing at all is open: there is no construct this could close.
          problems.push({ id: 'unexpected_end', at, command: raw });
          return;
        }
        if (current.kind === Command.FOR && current.varName === varName) {
          current.close = at;
          open.pop();
          sawElse.pop();
          return;
        }
        if (
          open.some(
            idx =>
              constructs[idx]?.kind === Command.FOR &&
              constructs[idx]?.varName === varName
          )
        ) {
          problems.push({ id: 'unexpected_end', at, command: raw });
        }
        return;
      }

      case Command.END_IF: {
        const current = innermost();
        if (current == null || current.kind !== Command.IF) {
          problems.push({ id: 'unexpected_end', at, command: raw });
          return;
        }
        current.close = at;
        open.pop();
        sawElse.pop();
        return;
      }

      default:
        return;
    }
  });

  // Whatever is still open at the end of the part was never closed.
  for (const idx of open) {
    const construct = constructs[idx];
    if (construct == null) continue;
    problems.push({
      id: construct.kind === Command.FOR ? 'unclosed_for' : 'unclosed_if',
      at: construct.open,
      command: program.commands[construct.open]?.raw ?? '',
    });
  }

  return { constructs, problems };
}

const resolveRest = (
  site: CommandSite,
  options: CreateReportOptions
): string => {
  try {
    return resolveSite(site, options).code;
  } catch {
    return '';
  }
};
