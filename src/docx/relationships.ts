/**
 * Writing out what a rendered part embedded: the image and HTML files
 * themselves, plus the `.rels` entries that tie them to the part.
 *
 * Relationship ids are only unique within one part, so every part gets its own
 * `word/_rels/<part>.rels` — a hyperlink referenced from `header1.xml` is
 * unresolvable if its relationship lives in `document.xml.rels`.
 */
import type JSZip from 'jszip';
import { logger } from '../debug';
import {
  Ns,
  PackagePath,
  PkgTag,
  RelAttr,
  RelType,
  TARGET_MODE_EXTERNAL,
  mediaPathOf,
  partPathOf,
  relsPathOf,
} from '../ooxml';
import { DEFAULT_LITERAL_XML_DELIMITER } from '../options';
import { addChild, newNonTextNode } from '../reportUtils';
import { type Htmls, type Images, type Links, type Node } from '../types';
import { buildXml, parseXml } from '../xml';
import { zipGetText, zipSetText } from '../zip';

/** Everything one rendered part wants written into the package. */
export type PartResources = {
  images: Images;
  links: Links;
  htmls: Htmls;
};

type Relationship = {
  id: string;
  type: RelType;
  target: string;
  /** Whether the target is outside the package, as a hyperlink's is. */
  external?: boolean;
};

/** The `.rels` part of a document part, created empty if it has none yet. */
const EMPTY_RELS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="${Ns.packageRelationships}">
        </Relationships>`;

/**
 * Writes the images, hyperlinks and HTML chunks a part produced, and registers
 * all of them in that part's `.rels` in one pass.
 */
export const writePartResources = async (
  zip: JSZip,
  partName: string,
  resources: PartResources,
  indentXml: boolean
): Promise<void> => {
  const rels: Relationship[] = [
    ...writeImageFiles(zip, partName, resources.images),
    ...linkRelationships(resources.links),
    ...writeHtmlFiles(zip, partName, resources.htmls),
  ];
  await addRelationships(zip, partName, rels, indentXml);
};

const writeImageFiles = (
  zip: JSZip,
  partName: string,
  images: Images
): Relationship[] => {
  const rels: Relationship[] = [];
  for (const [imageId, { extension, data }] of Object.entries(images)) {
    const imgName = `template_${partName}_${imageId}${extension}`;
    logger.debug(`Writing image ${imageId} (${imgName})...`);
    const imgPath = mediaPathOf(imgName);
    if (typeof data === 'string') zip.file(imgPath, data, { base64: true });
    else zip.file(imgPath, data, { binary: true });
    // Relationship targets are relative to the part's own directory (`word/`).
    rels.push({
      id: imageId,
      type: RelType.image,
      target: `${PackagePath.mediaDir}/${imgName}`,
    });
  }
  return rels;
};

const linkRelationships = (links: Links): Relationship[] =>
  Object.entries(links).map(([linkId, { url }]) => ({
    id: linkId,
    type: RelType.hyperlink,
    target: url,
    external: true,
  }));

const writeHtmlFiles = (
  zip: JSZip,
  partName: string,
  htmls: Htmls
): Relationship[] => {
  const rels: Relationship[] = [];
  for (const [htmlId, htmlData] of Object.entries(htmls)) {
    // Replace all period characters in the filename to play nice with more
    // picky parsers (like Docx4j)
    const htmlName = `template_${partName.replace(/\./g, '_')}_${htmlId}.html`;
    logger.debug(`Writing html ${htmlId} (${htmlName})...`);
    zipSetText(zip, partPathOf(htmlName), Buffer.from(htmlData));
    rels.push({ id: htmlId, type: RelType.altChunk, target: htmlName });
  }
  return rels;
};

const addRelationships = async (
  zip: JSZip,
  partName: string,
  rels: Relationship[],
  indentXml: boolean
): Promise<void> => {
  if (!rels.length) return;
  logger.debug(`Completing ${partName}.rels...`);
  const relsPath = relsPathOf(partName);
  const relsNode = await getRels(zip, relsPath);
  for (const { id, type, target, external } of rels) {
    addChild(
      relsNode,
      newNonTextNode(PkgTag.relationship, {
        [RelAttr.id]: id,
        [RelAttr.type]: type,
        [RelAttr.target]: target,
        ...(external ? { [RelAttr.targetMode]: TARGET_MODE_EXTERNAL } : {}),
      })
    );
  }
  zipSetText(
    zip,
    relsPath,
    buildXml(relsNode, {
      // A `.rels` part never holds the literal-XML markers a report body can.
      literalXmlDelimiter: DEFAULT_LITERAL_XML_DELIMITER,
      indentXml,
    })
  );
};

const getRels = async (zip: JSZip, relsPath: string): Promise<Node> => {
  const relsXml = await zipGetText(zip, relsPath);
  return parseXml(relsXml || EMPTY_RELS_XML);
};
