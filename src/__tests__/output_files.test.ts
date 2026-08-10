import { describe, it, expect } from 'vitest';
import {
  listReportFiles,
  makeDocx,
  readFixture,
  readReportFile,
} from './helpers';
import { createReport } from '../index';
import { CommandSyntaxError } from '../errors';
import { setDebugLogSink } from '../debug';

if (process.env.DEBUG) setDebugLogSink(console.log);

// The `_probe` shortcut in `createReport` returns before the zip is assembled,
// so everything in this file renders a complete report and inspects the
// resulting .docx.

describe('relationships', () => {
  it('writes hyperlink relationships for LINK commands', async () => {
    const template = await makeDocx({
      body: [`+++LINK ({ url: 'https://example.test/a', label: 'A' })+++`],
    });
    const report = await createReport({ template, data: {} });

    const rels = await readReportFile(report, 'word/_rels/document.xml.rels');
    expect(rels).toContain(
      'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"'
    );
    expect(rels).toContain('Target="https://example.test/a"');
    expect(rels).toContain('TargetMode="External"');

    // The relationship id must match the one referenced from the document.
    const document = await readReportFile(report, 'word/document.xml');
    expect(document).toContain('<w:hyperlink r:id="link1"');
    expect(rels).toContain('Id="link1"');
  });

  it('escapes XML-hostile characters in the relationship target', async () => {
    const template = await makeDocx({
      body: [`+++LINK ({ url: 'https://example.test/?a=1&b=2' })+++`],
    });
    const report = await createReport({ template, data: {} });

    const rels = await readReportFile(report, 'word/_rels/document.xml.rels');
    expect(rels).toContain('Target="https://example.test/?a=1&amp;b=2"');
  });
});

describe('headers and footers', () => {
  it('renders commands in header and footer parts', async () => {
    const template = await makeDocx({
      body: ['body: +++body_text+++'],
      header: ['header: +++header_text+++'],
      footer: ['footer: +++footer_text+++'],
    });
    const report = await createReport({
      template,
      data: {
        body_text: 'BODY',
        header_text: 'HEADER',
        footer_text: 'FOOTER',
      },
    });

    expect(await readReportFile(report, 'word/document.xml')).toContain(
      '>BODY<'
    );
    expect(await readReportFile(report, 'word/header1.xml')).toContain(
      '>HEADER<'
    );
    expect(await readReportFile(report, 'word/footer1.xml')).toContain(
      '>FOOTER<'
    );
  });

  it('reports errors found in a header part', async () => {
    const template = await makeDocx({
      body: ['all good'],
      header: ['+++IF-NOT nonsense+++'],
    });
    await expect(createReport({ template, data: {} })).rejects.toThrow(
      CommandSyntaxError
    );
  });

  it('collects errors from a header part when failFast is false', async () => {
    const template = await makeDocx({
      body: ['all good'],
      header: ['+++IF-NOT nonsense+++'],
    });
    await expect(
      createReport({ template, data: {}, failFast: false })
    ).rejects.toEqual([expect.any(CommandSyntaxError)]);
  });

  it('writes images of a header into that header own relationships', async () => {
    const data = await readFixture('sample.png');
    const template = await makeDocx({
      body: ['no image here'],
      header: ['+++IMAGE img()+++'],
    });
    const report = await createReport({
      template,
      data: {},
      additionalJsContext: {
        img: () => ({ width: 6, height: 6, data, extension: '.png' }),
      },
    });

    expect(await listReportFiles(report)).toContain(
      'word/media/template_header1.xml_img1.png'
    );
    const rels = await readReportFile(report, 'word/_rels/header1.xml.rels');
    expect(rels).toContain('Target="media/template_header1.xml_img1.png"');

    // The image extension must have been registered in the content types.
    const contentTypes = await readReportFile(report, '[Content_Types].xml');
    expect(contentTypes).toContain('ContentType="image/png"');
  });

  it('numbers images continuously across the main document and the header', async () => {
    const data = await readFixture('sample.png');
    const template = await makeDocx({
      body: ['+++IMAGE img()+++'],
      header: ['+++IMAGE img()+++'],
    });
    const report = await createReport({
      template,
      data: {},
      additionalJsContext: {
        img: () => ({ width: 6, height: 6, data, extension: '.png' }),
      },
    });

    const files = await listReportFiles(report);
    expect(files).toContain('word/media/template_document.xml_img1.png');
    expect(files).toContain('word/media/template_header1.xml_img2.png');
  });

  it('writes hyperlinks of a header into that header own relationships', async () => {
    // Relationship ids are scoped to the part that references them, so a
    // hyperlink used in a header must be declared in `header1.xml.rels`;
    // declaring it in `document.xml.rels` leaves Word unable to resolve the
    // `r:id` and it refuses to open the file.
    const template = await makeDocx({
      body: ['no link here'],
      header: [`+++LINK ({ url: 'https://example.test/h' })+++`],
    });
    const report = await createReport({ template, data: {} });

    expect(await readReportFile(report, 'word/header1.xml')).toContain(
      '<w:hyperlink r:id="link1"'
    );
    expect(
      await readReportFile(report, 'word/_rels/header1.xml.rels')
    ).toContain('Target="https://example.test/h"');
    expect(await readReportFile(report, 'word/_rels/document.xml.rels')).toBe(
      null
    );
  });

  it('keeps the relationships of each part to itself', async () => {
    // Both parts number their links from 1: that is only correct as long as
    // each `link1` lives in the .rels of the part that uses it.
    const template = await makeDocx({
      body: [`+++LINK ({ url: 'https://example.test/body' })+++`],
      header: [`+++LINK ({ url: 'https://example.test/header' })+++`],
    });
    const report = await createReport({ template, data: {} });

    const documentRels = await readReportFile(
      report,
      'word/_rels/document.xml.rels'
    );
    const headerRels = await readReportFile(
      report,
      'word/_rels/header1.xml.rels'
    );
    expect(documentRels).toContain('Target="https://example.test/body"');
    expect(documentRels).not.toContain('example.test/header');
    expect(headerRels).toContain('Target="https://example.test/header"');
    expect(headerRels).not.toContain('example.test/body');
  });

  it('writes the html chunk of a header into its own file and relationships', async () => {
    const template = await makeDocx({
      body: [`+++HTML '<p>body</p>'+++`],
      header: [`+++HTML '<p>header</p>'+++`],
    });
    const report = await createReport({ template, data: {} });

    // One chunk per part: naming both after the main document would have the
    // header overwrite the body's chunk.
    expect(
      await readReportFile(report, 'word/template_header1_xml_html1.html')
    ).toBe('<p>header</p>');
    expect(
      await readReportFile(report, 'word/template_document_xml_html1.html')
    ).toBe('<p>body</p>');
    expect(
      await readReportFile(report, 'word/_rels/header1.xml.rels')
    ).toContain('Target="template_header1_xml_html1.html"');
  });
});

describe('content types', () => {
  it('registers the html content type when HTML commands are used', async () => {
    const template = await makeDocx({
      body: [`+++HTML '<p>hi</p>'+++`],
    });
    const report = await createReport({ template, data: {} });

    expect(await listReportFiles(report)).toContain(
      'word/template_document_xml_html1.html'
    );
    expect(await readReportFile(report, '[Content_Types].xml')).toContain(
      'ContentType="text/html"'
    );
  });

  it('leaves the content types untouched when there is nothing to register', async () => {
    const template = await makeDocx({ body: ['+++name+++'] });
    const before = await readReportFile(template, '[Content_Types].xml');
    const report = await createReport({ template, data: { name: 'John' } });
    expect(await readReportFile(report, '[Content_Types].xml')).toEqual(before);
  });
});

describe('encoding', () => {
  // `zipSetText` hands JSZip a `Buffer` that `buildXml` has already UTF-8
  // encoded. Declaring it as text would have JSZip encode it a second time,
  // which only shows up on characters outside ASCII.
  it('writes non-ASCII content as UTF-8', async () => {
    const template = await makeDocx({
      body: ['+++greeting+++', '+++emoji+++'],
      header: ['+++greeting+++'],
    });
    const report = await createReport({
      template,
      data: { greeting: 'Ёлка — «привет» 中文', emoji: '👋🏽' },
    });

    for (const part of ['word/document.xml', 'word/header1.xml']) {
      const xml = await readReportFile(report, part);
      expect(xml).toContain('Ёлка — «привет» 中文');
    }
    expect(await readReportFile(report, 'word/document.xml')).toContain('👋🏽');
  });

  it('writes non-ASCII hyperlink targets and image data intact', async () => {
    const template = await makeDocx({
      body: [`+++LINK ({ url: 'https://example.test/пример?q=中文' })+++`],
    });
    const report = await createReport({ template, data: {} });
    const rels = await readReportFile(report, 'word/_rels/document.xml.rels');
    expect(rels).toContain('Target="https://example.test/пример?q=中文"');
  });
});
