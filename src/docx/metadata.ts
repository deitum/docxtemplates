/**
 * The document properties Word keeps in `docProps/`: page and word counts,
 * title, author, timestamps.
 */
import { MetaTag, PackagePath } from '../ooxml';
import { type Node, type ZipInput } from '../types';
import { zipLoad } from '../zip';
import { parseZipPath } from './parts';

/** The text of a node, if it has exactly a text node inside. */
const textOf = (node: Node): string | undefined => {
  const child = node._children[0];
  if (child == null) return undefined;
  if (child._fTextNode) return child._text;
  throw new Error(`Not a text node`);
};

/** The text of the first child element with the given tag. */
const findNodeText = (parent: Node, tag: MetaTag): string | undefined => {
  for (const child of parent._children) {
    if (!child._fTextNode && child._tag === tag) return textOf(child);
  }
  return undefined;
};

const numberize = (value: unknown): number | undefined => {
  const num = Number(value);
  return Number.isFinite(num) ? num : undefined;
};

/**
 * Extract metadata from a document, such as the number of pages or words.
 * @param template the docx template as a Buffer-like object
 */
export async function getMetadata(template: ZipInput) {
  const zip = await zipLoad(template);
  const appXml = await parseZipPath(zip, PackagePath.appProps);
  const coreXml = await parseZipPath(zip, PackagePath.coreProps);
  // TODO: extract custom.xml as well?

  return {
    pages: numberize(findNodeText(appXml, MetaTag.pages)),
    words: numberize(findNodeText(appXml, MetaTag.words)),
    characters: numberize(findNodeText(appXml, MetaTag.characters)),
    lines: numberize(findNodeText(appXml, MetaTag.lines)),
    paragraphs: numberize(findNodeText(appXml, MetaTag.paragraphs)),
    company: findNodeText(appXml, MetaTag.company),
    template: findNodeText(appXml, MetaTag.template),

    // from CoreXML
    title: findNodeText(coreXml, MetaTag.title),
    subject: findNodeText(coreXml, MetaTag.subject),
    creator: findNodeText(coreXml, MetaTag.creator),
    description: findNodeText(coreXml, MetaTag.description),
    lastModifiedBy: findNodeText(coreXml, MetaTag.lastModifiedBy),
    revision: findNodeText(coreXml, MetaTag.revision),
    lastPrinted: findNodeText(coreXml, MetaTag.lastPrinted),
    created: findNodeText(coreXml, MetaTag.created),
    modified: findNodeText(coreXml, MetaTag.modified),
    category: findNodeText(coreXml, MetaTag.category),
  };
}
