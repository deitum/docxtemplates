/**
 * Reading a .docx package: finding the parts that hold a template, parsing them
 * and keeping `[Content_Types].xml` up to date with what the report ended up
 * embedding.
 */
import type JSZip from 'jszip';
import { logger } from '../debug';
import { TemplateParseError } from '../errors';
import {
  CtAttr,
  HTML_EXTENSION,
  HTML_MEDIA_TYPE,
  IMAGE_MEDIA_TYPES,
  MainDocMediaType,
  PackagePath,
  PkgTag,
  partPathOf,
} from '../ooxml';
import preprocessTemplate from '../preprocessTemplate';
import { addChild, newNonTextNode } from '../reportUtils';
import {
  type CreateReportOptions,
  type Node,
  type NonTextNode,
  type ZipInput,
} from '../types';
import { parseXml } from '../xml';
import { zipGetText, zipLoad } from '../zip';

/** An XML part of the package that is rendered as a template. */
export type TemplatePart = {
  /** File name inside `word/`, e.g. `document.xml` or `header1.xml`. */
  name: string;
  /** Parsed and preprocessed contents. */
  template: Node;
};

/** `word/<something>.xml`: the parts that may carry template commands. */
const PART_PATH_REGEX = new RegExp(`^${PackagePath.wordDir}/[^/]+\\.xml$`);

/**
 * Parts this library generated itself on a previous run. Rendering a report
 * again would otherwise pick up its own output as a template.
 */
const GENERATED_PART_PREFIX = `${PackagePath.wordDir}/template`;

export async function parseTemplate(template: ZipInput) {
  logger.debug('Unzipping...');
  const zip = await zipLoad(template);

  // Office365 files may name the main template file document2.xml or something
  // else (see issue #131), so the content types 'manifest' is parsed first to
  // find out which part is the main document.
  logger.debug('finding main template file (e.g. document.xml)');
  const contentTypes = await readContentTypes(zip);
  const mainDocument = getMainDoc(contentTypes);

  const mainPath = partPathOf(mainDocument);
  logger.debug(`Reading ${mainPath}...`);
  const templateXml = await zipGetText(zip, mainPath);
  if (templateXml == null)
    throw new TemplateParseError(`${mainDocument} could not be found`);
  logger.debug(`${mainPath} file length: ${templateXml.length}`);

  logger.debug(`Parsing ${mainPath} XML...`);
  const tic = Date.now();
  const jsTemplate = await parseXml(templateXml);
  logger.debug(`${mainPath} parsed in ${Date.now() - tic} ms`);

  return { jsTemplate, mainDocument, zip, contentTypes };
}

/**
 * Parses every part other than the main document — the headers, footers and
 * so on, which may contain commands of their own.
 */
export async function prepSecondaryXMLs(
  zip: JSZip,
  mainDocument: string,
  options: CreateReportOptions
): Promise<TemplatePart[]> {
  const mainPath = partPathOf(mainDocument);
  const paths: string[] = [];
  zip.forEach(filePath => {
    if (
      PART_PATH_REGEX.test(filePath) &&
      filePath !== mainPath &&
      !filePath.startsWith(GENERATED_PART_PREFIX)
    ) {
      paths.push(filePath);
    }
  });

  // Sort to ensure deterministic processing.
  paths.sort();

  const parts: TemplatePart[] = [];
  for (const path of paths) {
    logger.debug(`Reading ${path}...`);
    const raw = await zipGetText(zip, path);
    if (raw == null) throw new TemplateParseError(`${path} could not be read`);
    logger.debug(`${path} file length: ${raw.length}`);
    logger.debug(`Parsing ${path} XML...`);
    const parsed = await parseXml(raw);
    parts.push({
      name: path.slice(`${PackagePath.wordDir}/`.length),
      template: preprocessTemplate(
        parsed,
        options.cmdDelimiter,
        options.preserveSpace
      ),
    });
  }
  return parts;
}

/** Parses one XML file of the package, by path. */
export async function parseZipPath(
  zip: JSZip,
  xmlPath: string
): Promise<NonTextNode> {
  const xmlFile = await zipGetText(zip, xmlPath);
  if (xmlFile == null)
    throw new TemplateParseError(`${xmlPath} could not be read`);
  const node = await parseXml(xmlFile);
  if (node._fTextNode)
    throw new TemplateParseError(`${xmlPath} is a text node when parsed`);
  return node;
}

export async function readContentTypes(zip: JSZip): Promise<NonTextNode> {
  return parseZipPath(zip, PackagePath.contentTypes);
}

export function getMainDoc(contentTypes: NonTextNode): string {
  const mainDocMimes: string[] = Object.values(MainDocMediaType);
  for (const t of contentTypes._children) {
    if (t._fTextNode) continue;
    const contentType = t._attrs[CtAttr.contentType];
    if (typeof contentType !== 'string' || !mainDocMimes.includes(contentType))
      continue;
    const path = t._attrs[CtAttr.partName];
    if (typeof path === 'string' && path)
      return path.replace(`/${PackagePath.wordDir}/`, '');
  }
  throw new TemplateParseError(
    `Could not find main document (e.g. document.xml) in ${PackagePath.contentTypes}`
  );
}

/**
 * Declares the media types of whatever the report embedded. Word refuses to
 * open a package that carries a part whose extension it has no content type
 * for.
 */
export function ensureContentTypes(
  contentTypes: NonTextNode,
  embedded: { images: boolean; htmls: boolean }
): void {
  const ensure = (extension: string, contentType: string) => {
    const alreadyDeclared = contentTypes._children.some(o => {
      if (o._fTextNode) return false;
      const declared = o._attrs[CtAttr.extension];
      return (
        typeof declared === 'string' &&
        declared.toLowerCase() === extension.toLowerCase()
      );
    });
    if (alreadyDeclared) return;
    addChild(
      contentTypes,
      newNonTextNode(PkgTag.defaultType, {
        [CtAttr.extension]: extension,
        [CtAttr.contentType]: contentType,
      })
    );
  };

  if (embedded.images) {
    logger.debug(`Completing ${PackagePath.contentTypes} for IMAGES...`);
    for (const [extension, mediaType] of Object.entries(IMAGE_MEDIA_TYPES)) {
      ensure(extension, mediaType);
    }
  }
  if (embedded.htmls) {
    logger.debug(`Completing ${PackagePath.contentTypes} for HTML...`);
    ensure(HTML_EXTENSION, HTML_MEDIA_TYPE);
  }
}
