import { BUFFER_TAGS } from './ooxml';
import { type Context, type CreateReportOptions } from './types';

/**
 * The mutable state of one pass over one XML part. `createReport` builds a
 * fresh context per part (main document, then each header and footer), carrying
 * over only the image/shape id counter so that ids stay unique across the file.
 */
export function newContext(
  options: CreateReportOptions,
  imageAndShapeIdIncrement = 0
): Context {
  return {
    gCntIf: 0,
    gCntEndIf: 0,
    level: 1,
    fCmd: false,
    cmd: '',
    fSeekQuery: false,
    buffers: Object.fromEntries(
      BUFFER_TAGS.map(tag => [
        tag,
        { text: '', cmds: '', fInsertedText: false },
      ])
    ) as Context['buffers'],
    imageAndShapeIdIncrement,
    images: {},
    linkId: 0,
    links: {},
    htmlId: 0,
    htmls: {},
    // Keyed by names taken from the template, hence the null prototype: a plain
    // object would report `toString` and friends as defined.
    vars: Object.create(null),
    loops: [],
    fJump: false,
    shorthands: Object.create(null),
    options,
    // To verify we don't have a nested IF within the same `w:p` or `w:tr` tag
    pIfCheckMap: new Map(),
    trIfCheckMap: new Map(),
  };
}
