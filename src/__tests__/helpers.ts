import fs from 'node:fs';
import path from 'node:path';
import JSZip from 'jszip';
import { type Node } from '../types';

/** Absolute path of the `fixtures` directory. */
export const fixturesDir = path.join(import.meta.dirname, 'fixtures');

/** Absolute path of a file in the `fixtures` directory. */
export const fixturePath = (name: string) => path.join(fixturesDir, name);

/** Contents of a file in the `fixtures` directory. */
export const readFixture = (name: string) =>
  fs.promises.readFile(fixturePath(name));

// ==========================================
// In-memory templates
// ==========================================

const CONTENT_TYPES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`;

const W_NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

// Paragraph text is escaped, so that callers can write commands the way a user
// would type them in Word (`+++IF a > 1+++`, `?a=1&b=2`, ...).
const escapeXml = (text: string) =>
  text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

const paragraphs = (lines: string[]) =>
  lines
    .map(line => `<w:p><w:r><w:t>${escapeXml(line)}</w:t></w:r></w:p>`)
    .join('');

/**
 * Builds the XML of a table: one array per row, and one array of paragraph
 * lines per cell.
 *
 * ```js
 * tableXml([[['Name'], ['+++INS $person.name+++']]]);
 * ```
 */
export const tableXml = (rows: string[][][]) => {
  const cells = (row: string[][]) =>
    row
      .map(
        lines =>
          `<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>` +
          `${paragraphs(lines)}</w:tc>`
      )
      .join('');
  return `<w:tbl>${rows.map(row => `<w:tr>${cells(row)}</w:tr>`).join('')}</w:tbl>`;
};

type MakeDocxOptions = {
  /** Paragraphs of the main document (`word/document.xml`). */
  body?: string[];
  /**
   * Raw XML appended to the main document's body, for structures that `body`
   * cannot express (tables, ...). See `tableXml`.
   */
  bodyXml?: string;
  /** Paragraphs of a header part (`word/header1.xml`). */
  header?: string[];
  /** Paragraphs of a footer part (`word/footer1.xml`). */
  footer?: string[];
  /**
   * Extra (or replacement) entries in the zip, keyed by path. Use this to
   * inject malformed parts, or to drop a part by passing `null`.
   */
  files?: { [path: string]: string | null };
};

/**
 * Builds a minimal but valid .docx in memory, so that tests don't need a
 * binary fixture for every template variation. `[Content_Types].xml` plus
 * `word/document.xml` is all the library needs to render a report; header and
 * footer parts are picked up simply by being present in the zip.
 */
export const makeDocx = async ({
  body = [],
  bodyXml = '',
  header,
  footer,
  files,
}: MakeDocxOptions = {}): Promise<Buffer> => {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', CONTENT_TYPES_XML);
  zip.file(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${W_NS}><w:body>${paragraphs(body)}${bodyXml}</w:body></w:document>`
  );
  if (header) {
    zip.file(
      'word/header1.xml',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr ${W_NS}>${paragraphs(header)}</w:hdr>`
    );
  }
  if (footer) {
    zip.file(
      'word/footer1.xml',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr ${W_NS}>${paragraphs(footer)}</w:ftr>`
    );
  }
  for (const [filePath, contents] of Object.entries(files ?? {})) {
    if (contents == null) zip.remove(filePath);
    else zip.file(filePath, contents);
  }
  return zip.generateAsync({ type: 'nodebuffer' });
};

/** Reads a text file out of a generated report. */
export const readReportFile = async (
  report: Uint8Array,
  filePath: string
): Promise<string | null> => {
  const zip = await JSZip.loadAsync(report);
  const file = zip.file(filePath);
  if (!file) return null;
  return file.async('text');
};

/** Lists the entry names of a generated report. */
export const listReportFiles = async (
  report: Uint8Array
): Promise<string[]> => {
  const zip = await JSZip.loadAsync(report);
  return Object.keys(zip.files);
};

// ==========================================
// Reading rendered reports
// ==========================================

/** Concatenates the text of all the text nodes below the given node. */
export const nodeText = (node: Node): string =>
  node._fTextNode ? node._text : node._children.map(nodeText).join('');

/** Returns the text of the report, one line per (non-empty) paragraph. */
export const reportText = (node: Node): string => {
  const lines: string[] = [];
  const walk = (n: Node) => {
    if (!n._fTextNode && n._tag === 'w:p') {
      const text = nodeText(n).trim();
      if (text !== '') lines.push(text);
      return;
    }
    n._children.forEach(walk);
  };
  walk(node);
  return lines.join('\n');
};

/**
 * Returns the text of every cell of every table row of the report, one array
 * per row — so that tests can check how many cells a row ended up with, and not
 * just what was rendered in them.
 */
export const tableCells = (node: Node): string[][] => {
  const rows: string[][] = [];
  const walk = (n: Node) => {
    if (!n._fTextNode && n._tag === 'w:tr') {
      rows.push(
        n._children
          .filter(child => !child._fTextNode && child._tag === 'w:tc')
          .map(cell => nodeText(cell).trim())
      );
    }
    n._children.forEach(walk);
  };
  walk(node);
  return rows;
};

// ==========================================
// Errors
// ==========================================

class NoErrorThrownError extends Error {}

/**
 * Runs `call` and returns the error it threw, so that tests can make several
 * assertions on it. Fails the test (by returning a `NoErrorThrownError`) if
 * nothing was thrown.
 */
export const getError = async <TError>(
  call: () => unknown
): Promise<TError> => {
  try {
    await call();
    throw new NoErrorThrownError();
  } catch (error: unknown) {
    return error as TError;
  }
};
