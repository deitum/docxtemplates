import { BUFFER_TAGS } from './ooxml';
import { newResources } from './resources';
import {
  type Context,
  type CreateReportOptions,
  type Scope,
  type WalkState,
} from './types';

/**
 * The mutable state of one pass over one XML part. `createReport` builds a
 * fresh context per part (main document, then each header and footer), carrying
 * over only the image/shape id counter so that ids stay unique across the file.
 */
export function newContext(
  options: CreateReportOptions,
  lastImageAndShapeId = 0
): Context {
  return {
    options,
    walk: newWalkState(),
    scope: newScope(),
    resources: newResources(lastImageAndShapeId),
  };
}

function newWalkState(): WalkState {
  return {
    level: 1,
    jumpRequested: false,
    buffers: Object.fromEntries(
      BUFFER_TAGS.map(tag => [
        tag,
        { text: '', cmds: '', hasInsertedText: false },
      ])
    ) as WalkState['buffers'],
    isCollectingCommand: false,
    command: '',
    seekingQuery: false,
    openIfCount: 0,
    closedIfCount: 0,
    // To verify we don't have a nested IF within the same `w:p` or `w:tr` tag
    ifByParagraph: new Map(),
    ifByTableRow: new Map(),
  };
}

function newScope(): Scope {
  return {
    // Keyed by names taken from the template, hence the null prototype: a plain
    // object would report `toString` and friends as defined.
    vars: Object.create(null),
    loops: [],
    shorthands: Object.create(null),
  };
}
