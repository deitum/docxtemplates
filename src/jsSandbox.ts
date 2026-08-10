import vm from 'node:vm';
import { getCurLoop } from './reportUtils';
import {
  type ReportData,
  type Context,
  type RunJsContext,
  type SandBox,
  type SandboxRuntime,
} from './types';
import {
  isError,
  CommandExecutionError,
  NullishCommandResultError,
} from './errors';
import { logger } from './debug';

/**
 * The names the sandbox reserves for itself. `code` is the snippet about to
 * run: it is read back *by name* from inside the evaluated source, so the two
 * must agree — hence the constant. Template variables are exposed alongside
 * under `varPrefix`, where they cannot collide with the user's own data.
 */
const SandboxKey = {
  code: '__code__',
  result: '__result__',
  varPrefix: '$',
  /** Index of the innermost loop's current iteration. */
  loopIndex: '$idx',
} as const;

/** Evaluates the snippet with the sandbox as scope, when sandboxing is off. */
const UNSANDBOXED_SOURCE = `with(this) { return eval(${SandboxKey.code}); }`;

/*
 * One `SandboxRuntime` per document part, built on first use and discarded with
 * the part.
 *
 * Contextifying an object is what `vm` charges for — two orders of magnitude
 * more than compiling or running the snippet itself — so it happens once per
 * part rather than once per command. The snippets of a part therefore share one
 * set of built-ins and one global scope, which is also what makes a function
 * defined in one snippet behave sanely when called from another.
 */

const newSandbox = (): SandBox => ({
  [SandboxKey.code]: undefined,
  [SandboxKey.result]: undefined,
});

const getRuntime = (ctx: Context): SandboxRuntime => {
  const existing = ctx.scope.jsRuntime;
  if (existing != null) return existing;
  // `createContext` hands the very object back, contextified; `vm` types it as
  // a bare dictionary, but it is the sandbox we just built.
  const context = vm.createContext(newSandbox()) as SandBox;
  const runtime: SandboxRuntime = { context, scripts: new Map() };
  ctx.scope.jsRuntime = runtime;
  return runtime;
};

/** `Object.assign`, but tolerant of `data` being anything at all. */
const assignInto = (target: SandBox, source: unknown): void => {
  if (source == null) return;
  Object.assign(target, source);
};

/**
 * Brings the sandbox up to date for the snippet about to run. The order matters
 * and is part of the contract: values carried over from earlier snippets are
 * overridden by `data`, which is overridden by `additionalJsContext`.
 */
function prepareSandbox(
  sandbox: SandBox,
  data: ReportData | undefined,
  code: string,
  ctx: Context
): SandBox {
  sandbox[SandboxKey.code] = code;
  sandbox[SandboxKey.result] = undefined;
  assignInto(sandbox, data);
  assignInto(sandbox, ctx.options.additionalJsContext);

  // Add currently defined vars, including loop vars and the index
  // of the innermost loop
  const curLoop = getCurLoop(ctx);
  if (curLoop) sandbox[SandboxKey.loopIndex] = curLoop.idx;
  for (const varName of Object.keys(ctx.scope.vars)) {
    sandbox[`${SandboxKey.varPrefix}${varName}`] = ctx.scope.vars[varName];
  }
  return sandbox;
}

/** The subset of the engine's context a custom sandbox is handed. */
const runJsContextOf = (ctx: Context): RunJsContext => ({
  options: ctx.options,
  vars: ctx.scope.vars,
  loops: ctx.scope.loops,
  jsSandbox: ctx.scope.jsSandbox,
});

/**
 * A `vm.Script` is independent of the context it runs in, so a template that
 * evaluates the same expression on every iteration of a loop compiles it once.
 */
function compile(ctx: Context, code: string): vm.Script {
  const { scripts } = getRuntime(ctx);
  let script = scripts.get(code);
  if (script == null) {
    script = new vm.Script(code);
    scripts.set(code, script);
  }
  return script;
}

// Runs a user snippet in a sandbox, and returns the result.
// The snippet can return a Promise, which is then awaited.
// The sandbox is kept for the execution of snippets later on
// in the template. Sandboxing can also be disabled via
// ctx.options.noSandbox.
export async function runUserJsAndGetRaw(
  data: ReportData | undefined,
  code: string,
  ctx: Context
): Promise<any> {
  // `runJs` and `noSandbox` hand the sandbox to code that may keep or replace
  // it, so they get a throwaway object seeded with the state carried over so
  // far. The `vm` path evaluates into the part's live context, which is where
  // that state already lives.
  const usesSharedContext = !ctx.options.runJs && !ctx.options.noSandbox;
  const sandbox = prepareSandbox(
    usesSharedContext
      ? getRuntime(ctx).context
      : { ...(ctx.scope.jsSandbox ?? newSandbox()) },
    data,
    code,
    ctx
  );

  // Run the JS snippet and extract the result
  let context;
  let result;
  try {
    if (ctx.options.runJs) {
      const temp = ctx.options.runJs({ sandbox, ctx: runJsContextOf(ctx) });
      context = temp.modifiedSandbox;
      result = await temp.result;
    } else if (ctx.options.noSandbox) {
      context = sandbox;
      const wrapper = new Function(UNSANDBOXED_SOURCE);
      result = await wrapper.call(context);
    } else {
      context = sandbox;
      const script = compile(ctx, sandbox[SandboxKey.code] ?? '');
      result = await script.runInContext(context);
    }
  } catch (err) {
    const e = isError(err) ? err : new Error(`${err}`);
    if (ctx.options.errorHandler != null) {
      context = sandbox;
      result = await ctx.options.errorHandler(e, code);
    } else {
      throw new CommandExecutionError(e, code);
    }
  }

  if (ctx.options.rejectNullish && result == null) {
    const nerr = new NullishCommandResultError(code);
    if (ctx.options.errorHandler != null) {
      result = await ctx.options.errorHandler(nerr, code);
    } else {
      throw nerr;
    }
  }

  // Save the sandbox for later use, omitting the reserved properties.
  ctx.scope.jsSandbox = {
    ...context,
    [SandboxKey.code]: undefined,
    [SandboxKey.result]: undefined,
  };
  logger.debug('Command returned: ', result);
  return result;
}
