import vm from 'node:vm';
import { getCurLoop } from './reportUtils';
import { type ReportData, type Context, type SandBox } from './types';
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
  // Retrieve the current JS sandbox contents (if any) and add
  // the code to be run, and a placeholder for the result,
  // as well as all data defined by the user
  const sandbox: SandBox = {
    ...(ctx.jsSandbox || {}),
    [SandboxKey.code]: code,
    [SandboxKey.result]: undefined,
    ...data,
    ...ctx.options.additionalJsContext,
  };

  // Add currently defined vars, including loop vars and the index
  // of the innermost loop
  const curLoop = getCurLoop(ctx);
  if (curLoop) sandbox[SandboxKey.loopIndex] = curLoop.idx;
  Object.keys(ctx.vars).forEach(varName => {
    sandbox[`${SandboxKey.varPrefix}${varName}`] = ctx.vars[varName];
  });

  // Run the JS snippet and extract the result
  let context;
  let result;
  try {
    if (ctx.options.runJs) {
      const temp = ctx.options.runJs({ sandbox, ctx });
      context = temp.modifiedSandbox;
      result = await temp.result;
    } else if (ctx.options.noSandbox) {
      context = sandbox;
      const wrapper = new Function(UNSANDBOXED_SOURCE);
      result = await wrapper.call(context);
    } else {
      const script = new vm.Script(sandbox[SandboxKey.code] ?? '');
      context = vm.createContext(sandbox);
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
  ctx.jsSandbox = {
    ...context,
    [SandboxKey.code]: undefined,
    [SandboxKey.result]: undefined,
  };
  logger.debug('Command returned: ', result);
  return result;
}
