/**
 * `npm run bench` — a reproducible measurement of where report generation
 * spends its time.
 *
 * Three shapes of document, because they stress different parts of the engine:
 * a FOR loop (the JS sandbox), a large static document (parsing, serialising
 * and zipping), and a document embedding images (the media pipeline). Each
 * phase is timed separately, so that a change can be attributed rather than
 * guessed at.
 *
 * Numbers are only comparable against numbers taken on the same machine. Run it
 * before and after a change, not once.
 */
import JSZip from 'jszip';
import { findHighestImgId } from '../src/commands/media';
import { prepSecondaryXMLs } from '../src/docx/parts';
import { newContext } from '../src/context';
import { resolveOptions } from '../src/options';
import preprocessTemplate from '../src/preprocessTemplate';
import { type CreateReportOptions, type ImagePars } from '../src/types';
import { produceJsReport } from '../src/walk';
import { buildXml, parseXml } from '../src/xml';
import { zipLoad, zipSave, zipSetText } from '../src/zip';
import createReport from '../src/main';

// ==========================================
// Timing
// ==========================================

const now = () => Number(process.hrtime.bigint()) / 1e6;

type Phases = Record<string, number>;

/** Runs `fn` `repeats` times and keeps the fastest run of each phase. */
async function best(
  repeats: number,
  fn: () => Promise<Phases>
): Promise<Phases> {
  let out: Phases | null = null;
  for (let i = 0; i < repeats; i++) {
    const phases = await fn();
    if (out == null) out = phases;
    else {
      for (const key of Object.keys(phases)) {
        out[key] = Math.min(out[key] ?? Infinity, phases[key] ?? Infinity);
      }
    }
  }
  return out ?? {};
}

const fmt = (ms: number) => `${ms.toFixed(1)} ms`;

function report(title: string, phases: Phases) {
  console.log(`\n${title}`);
  const width = Math.max(...Object.keys(phases).map(k => k.length));
  let total = 0;
  for (const [phase, ms] of Object.entries(phases)) {
    if (phase === 'total') continue;
    total += ms;
    console.log(`  ${phase.padEnd(width)}  ${fmt(ms).padStart(10)}`);
  }
  const stated = phases.total ?? total;
  console.log(`  ${'—'.repeat(width)}  ${'—'.repeat(10)}`);
  console.log(`  ${'total'.padEnd(width)}  ${fmt(stated).padStart(10)}`);
}

// ==========================================
// Templates
// ==========================================

const W_NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

const CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`;

const paragraphs = (lines: string[]) =>
  lines
    .map(
      line =>
        `<w:p><w:r><w:t>${line
          .replace(/&/g, '&amp;')
          .replace(/</g, '&lt;')
          .replace(/>/g, '&gt;')}</w:t></w:r></w:p>`
    )
    .join('');

const documentXml = (lines: string[]) =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${W_NS}><w:body>${paragraphs(lines)}</w:body></w:document>`;

async function makeDocx(lines: string[]): Promise<Buffer> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', CONTENT_TYPES);
  zip.file('word/document.xml', documentXml(lines));
  return zip.generateAsync({ type: 'nodebuffer' });
}

// ==========================================
// The phase-by-phase run
// ==========================================

/**
 * Mirrors what `createReport` does, but with a clock between the phases. It has
 * to be kept in step with `main.ts` by hand; the end-to-end number below is the
 * check that it has not drifted too far.
 */
async function timedRun(
  template: Buffer,
  data: unknown,
  userOptions: Partial<CreateReportOptions> & {
    additionalJsContext?: object;
  } = {}
): Promise<Phases> {
  const options = resolveOptions(userOptions as never);
  const phases: Phases = {};

  let tic = now();
  const zip = await zipLoad(template);
  const rawXml = await zip.file('word/document.xml')!.async('text');
  phases.unzip = now() - tic;

  tic = now();
  const parsed = await parseXml(rawXml);
  phases.parse = now() - tic;

  tic = now();
  const prepared = preprocessTemplate(
    parsed,
    options.cmdDelimiter,
    options.preserveSpace
  );
  phases.preprocess = now() - tic;

  tic = now();
  await prepSecondaryXMLs(zip, 'document.xml', options);
  findHighestImgId(prepared);
  phases.scanParts = now() - tic;

  tic = now();
  const ctx = newContext(options);
  const result = await produceJsReport(data, prepared, ctx);
  phases.walk = now() - tic;
  if (result.status === 'errors') throw result.errors[0];

  tic = now();
  const xml = buildXml(result.report, {
    literalXmlDelimiter: options.literalXmlDelimiter,
    indentXml: options.indentXml,
  });
  phases.build = now() - tic;

  tic = now();
  zipSetText(zip, 'word/document.xml', xml);
  await zipSave(zip, options.compressionLevel);
  phases.zip = now() - tic;

  return phases;
}

/** End to end through the public API, as a user would see it. */
async function endToEnd(
  template: Buffer,
  data: unknown,
  extra: object = {}
): Promise<number> {
  const tic = now();
  await createReport({ template, data, ...extra } as never);
  return now() - tic;
}

// ==========================================
// Cases
// ==========================================

const LOOP_ITERATIONS = 2000;
const STATIC_PARAGRAPHS = 20_000;
const IMAGE_COUNT = 50;

const PNG_BASE64 =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';

async function benchLoop() {
  const template = await makeDocx([
    'Report header',
    '+++FOR row IN rows+++',
    '+++$row.name+++',
    '+++$row.value+++',
    '+++$idx+++',
    '+++END-FOR row+++',
  ]);
  const data = {
    rows: Array.from({ length: LOOP_ITERATIONS }, (_, i) => ({
      name: `row ${i}`,
      value: i * 3,
    })),
  };

  report(
    `FOR loop, ${LOOP_ITERATIONS} iterations x 3 INS ` +
      `(${LOOP_ITERATIONS * 3} evaluations), sandboxed`,
    await best(3, () => timedRun(template, data))
  );
  console.log(
    `  end-to-end (sandbox)     ${fmt(await endToEnd(template, data))}`
  );
  console.log(
    `  end-to-end (noSandbox)   ` +
      fmt(await endToEnd(template, data, { noSandbox: true }))
  );
}

async function benchStatic() {
  const lines = Array.from(
    { length: STATIC_PARAGRAPHS },
    (_, i) => `Paragraph ${i}: lorem ipsum dolor sit amet, consectetur.`
  );
  const template = await makeDocx(lines);
  report(
    `Static document, ${STATIC_PARAGRAPHS} paragraphs`,
    await best(3, () => timedRun(template, {}))
  );
  console.log(
    `  end-to-end               ${fmt(await endToEnd(template, {}))}`
  );
}

async function benchImages() {
  const template = await makeDocx([
    '+++FOR i IN list+++',
    '+++IMAGE img()+++',
    '+++END-FOR i+++',
  ]);
  const data = { list: Array.from({ length: IMAGE_COUNT }, (_, i) => i) };
  const additionalJsContext = {
    img: (): ImagePars => ({
      width: 2,
      height: 2,
      data: PNG_BASE64,
      extension: '.png',
    }),
  };
  report(
    `${IMAGE_COUNT} embedded images`,
    await best(3, () => timedRun(template, data, { additionalJsContext }))
  );
  console.log(
    `  end-to-end               ` +
      fmt(await endToEnd(template, data, { additionalJsContext }))
  );
}

async function main() {
  console.log(`node ${process.version}`);
  await benchLoop();
  await benchStatic();
  await benchImages();
  console.log('');
}

main().catch((err: unknown) => {
  console.error(err);
  process.exitCode = 1;
});
