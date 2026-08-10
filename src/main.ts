import { findHighestImgId } from './commands/media';
import { newContext } from './context';
import { logger } from './debug';
import {
  ensureContentTypes,
  parseTemplate,
  prepSecondaryXMLs,
  type TemplatePart,
} from './docx/parts';
import { writePartResources } from './docx/relationships';
import { withPart } from './errors';
import { PackagePath, partPathOf } from './ooxml';
import { resolveOptions } from './options';
import { validateTemplate } from './optionsSchema';
import preprocessTemplate from './preprocessTemplate';
import {
  type BuiltInCommand,
  Command,
  type CommandSummary,
  type CreateReportOptions,
  type Node,
  type ReportData,
  type UserOptions,
  type ZipInput,
} from './types';
import { compileTemplate, extractQuery, resolveSite } from './template/compile';
import { produceJsReport } from './walk';
import { buildXml } from './xml';
import { zipSave, zipSetText } from './zip';

/**
 * Create Report from docx template
 *
 * example:
 * ```js
 * const report = await createReport({
 *   template,
 *   data: query => graphqlServer.execute(query),
 *   additionalJsContext: {
 *     // all of these will be available to JS snippets in your template commands
 *     foo: 'bar',
 *     qrCode: async url => {
 *       // do stuff
 *     },
 *   },
 *   cmdDelimiter: '+++',
 *   literalXmlDelimiter: '||',
 *   processLineBreaks: true,
 *   noSandbox: false,
 * });
 * ```
 *
 * @param options Options for Report
 */
async function createReport(options: UserOptions): Promise<Uint8Array>;
/**
 * For development and testing purposes. Don't use _probe if you don't know what you are doing
 */
async function createReport(options: UserOptions, _probe: 'JS'): Promise<Node>;

/**
 * For development and testing purposes. Don't use _probe if you don't know what you are doing
 */
async function createReport(
  options: UserOptions,
  _probe: 'XML'
): Promise<string>;
async function createReport(
  options: UserOptions,
  _probe?: 'JS' | 'XML'
): Promise<Node | string | Uint8Array> {
  logger.debug('Report options:', { attach: options });
  const { template, data, queryVars } = options;
  validateTemplate(template);
  const createOptions = resolveOptions(options);
  const xmlOptions = {
    literalXmlDelimiter: createOptions.literalXmlDelimiter,
    indentXml: createOptions.indentXml,
  };

  const { jsTemplate, mainDocument, zip, contentTypes } =
    await parseTemplate(template);

  logger.debug('Preprocessing template...');
  const mainPart: TemplatePart = {
    name: mainDocument,
    template: preprocessTemplate(
      jsTemplate,
      createOptions.cmdDelimiter,
      createOptions.preserveSpace
    ),
  };

  // Fetch the data that will fill in the template
  let queryResult: ReportData;
  if (typeof data === 'function') {
    logger.debug('Looking for the query in the template...');
    const query = extractQuery(mainPart.template, createOptions);
    logger.debug(`Query: ${query || 'no query found'}`);
    queryResult = await data(query, queryVars);
  } else {
    queryResult = data;
  }

  const parts = [
    mainPart,
    ...(await prepSecondaryXMLs(zip, mainDocument, createOptions)),
  ];

  // Continue numbering images and shapes where the template left off, so that
  // generated ones can't collide with those already in the document.
  let lastImageAndShapeId = Math.max(
    ...parts.map(part => findHighestImgId(part.template))
  );
  let numImages = 0;
  let numHtmls = 0;

  for (const part of parts) {
    logger.debug(`Generating report for ${part.name}...`);
    // A fresh context per part: only the image/shape ids carry over.
    const ctx = newContext(createOptions, lastImageAndShapeId);
    // The one place that knows which part is being rendered, and so the one
    // place that can say which part an error came from. A report is built from
    // `document.xml` plus every header and footer, and with `failFast: false`
    // their errors all end up in the same array.
    let result;
    try {
      result = await produceJsReport(queryResult, part.template, ctx);
    } catch (err) {
      throw withPart(err, part.name);
    }
    if (result.status === 'errors') throw withPart(result.errors, part.name);
    lastImageAndShapeId = ctx.resources.lastShapeId;

    // The probes are a testing shortcut into the main document, and return
    // before the package is assembled.
    if (part === mainPart) {
      if (_probe === 'JS') return result.report;
      if (_probe === 'XML')
        return buildXml(result.report, xmlOptions).toString('utf-8');
    }

    logger.debug(`Writing ${part.name}...`);
    zipSetText(zip, partPathOf(part.name), buildXml(result.report, xmlOptions));
    numImages += Object.keys(result.images).length;
    numHtmls += Object.keys(result.htmls).length;
    await writePartResources(zip, part.name, result, createOptions.indentXml);
  }

  if (numImages || numHtmls) {
    logger.debug(`Completing ${PackagePath.contentTypes}...`);
    logger.debug('Content types', { attach: contentTypes });
    ensureContentTypes(contentTypes, {
      images: numImages > 0,
      htmls: numHtmls > 0,
    });
    zipSetText(
      zip,
      PackagePath.contentTypes,
      buildXml(contentTypes, xmlOptions)
    );
  }

  logger.debug('Zipping...');
  return zipSave(zip, createOptions.compressionLevel);
}

/**
 * Lists all the commands in a docx template.
 *
 * example:
 * ```js
 * const template_buffer = fs.readFileSync('template.docx');
 * const commands = await listCommands(template_buffer, ['{', '}']);
 * // `commands` will contain something like:
 * [
 *    { raw: 'INS some_variable', code: 'some_variable', type: 'INS' },
 *    { raw: 'IMAGE svgImgFile()', code: 'svgImgFile()', type: 'IMAGE' },
 * ]
 * ```
 *
 * @param template the docx template as a Buffer-like object
 * @param delimiter the command delimiter (defaults to ['+++', '+++'])
 * @param aliasOptions the command/operator aliases used by the template, if any
 */
export async function listCommands(
  template: ZipInput,
  delimiter?: string | [string, string],
  aliasOptions?: Pick<UserOptions, 'commandAliases' | 'operatorAliases'>
): Promise<CommandSummary[]> {
  const options: CreateReportOptions = resolveOptions({
    cmdDelimiter: delimiter,
    commandAliases: aliasOptions?.commandAliases,
    operatorAliases: aliasOptions?.operatorAliases,
    // Listing what a template contains is not the place to complain about an
    // unbalanced IF or FOR.
    failFast: false,
  });

  const { jsTemplate, mainDocument, zip } = await parseTemplate(template);
  const parts = [
    preprocessTemplate(jsTemplate, options.cmdDelimiter, options.preserveSpace),
    ...(await prepSecondaryXMLs(zip, mainDocument, options)).map(
      part => part.template
    ),
  ];

  const commands: CommandSummary[] = [];
  for (const part of parts) {
    for (const site of compileTemplate(part, options.cmdDelimiter).commands) {
      const { raw, name, code } = resolveSite(site, options);
      // `CMD_NODE` is scaffolding left behind by `preprocessTemplate`, not
      // something the template author wrote.
      if (name != null && name !== Command.CMD_NODE) {
        commands.push({ raw, type: name as BuiltInCommand, code });
      }
    }
  }
  return commands;
}

export default createReport;
