import createReport from './main';
export { listCommands } from './main';
export { getMetadata } from './docx/metadata';
// Listed rather than `export *`, so that adding a helper to `errors.ts` does
// not quietly make it part of the public API.
export {
  CommandExecutionError,
  CommandSyntaxError,
  ErrorId,
  ImageError,
  IncompleteConditionalStatementError,
  InternalError,
  InvalidAliasError,
  InvalidCommandError,
  isError,
  NullishCommandResultError,
  ObjectCommandResultError,
  TemplateError,
  TemplateParseError,
  UnterminatedForLoopError,
} from './errors';
export type { ErrorProperties } from './errors';
export type { QueryResolver } from './types';
export { createReport };
export default createReport;
