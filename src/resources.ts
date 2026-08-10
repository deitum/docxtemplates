/**
 * What IMAGE, LINK and HTML accumulate while a part is rendered: the files and
 * relationships that get written alongside the XML, and the node each command
 * parks for the walk to splice in.
 *
 * These used to be ten fields on the engine's context — three collections,
 * three "pending node" slots, three id counters and the run properties a link
 * inherits — read and incremented from wherever. Behind this interface the
 * counters are nobody else's business: callers ask for a resource to be
 * registered and get back the relationship id to reference it by.
 */
import {
  type Htmls,
  type Image,
  type Images,
  type Links,
  type Node,
  type NonTextNode,
} from './types';

/** The kinds of node a command can park for the walk to pick up. */
export type PendingKind = 'image' | 'link' | 'html';

/** A node a command built, and whatever should follow it. */
export type PendingNode = { node: Node; extra?: Node[] };

export type Resources = {
  /** Images to write into `word/media/`, by relationship id. */
  readonly images: Images;
  /** Hyperlink targets, by relationship id. */
  readonly links: Links;
  /** HTML chunks to write as separate parts, by relationship id. */
  readonly htmls: Htmls;

  /** Registers an image and returns the relationship id to reference it by. */
  addImage(image: Image): string;
  /** Registers a hyperlink target and returns its relationship id. */
  addLink(url: string): string;
  /** Registers an HTML chunk and returns its relationship id. */
  addHtml(html: string): string;

  /**
   * The id of the last `wp:docPr` / `v:shape` handed out. Numbering continues
   * from whatever the template already used, so that copies made by a FOR loop
   * cannot collide with the shapes it shipped with.
   */
  readonly lastShapeId: number;
  /** Takes the next shape id. */
  nextShapeId(): number;

  /** Parks a node for the walk to splice in when it reaches the right level. */
  park(kind: PendingKind, pending: PendingNode): void;
  /** Takes the node parked under `kind`, if any, clearing it. */
  takePending(kind: PendingKind): PendingNode | undefined;

  /**
   * The last `w:rPr` seen, so that a LINK can inherit the formatting of the run
   * it replaces.
   */
  textRunProps: NonTextNode | undefined;
};

/** Relationship id prefixes, one counter per kind of embedded resource. */
const REL_ID_PREFIX: { readonly [kind in PendingKind]: string } = {
  image: 'img',
  link: 'link',
  html: 'html',
};

export function newResources(lastImageAndShapeId = 0): Resources {
  const images: Images = {};
  const links: Links = {};
  const htmls: Htmls = {};
  const pending = new Map<PendingKind, PendingNode>();

  let shapeId = lastImageAndShapeId;
  let linkId = 0;
  let htmlId = 0;

  return {
    images,
    links,
    htmls,

    addImage(image) {
      // Images share the shape counter: both end up as `wp:docPr` ids.
      shapeId += 1;
      const relId = `${REL_ID_PREFIX.image}${shapeId}`;
      images[relId] = image;
      return relId;
    },
    addLink(url) {
      linkId += 1;
      const relId = `${REL_ID_PREFIX.link}${linkId}`;
      links[relId] = { url };
      return relId;
    },
    addHtml(html) {
      htmlId += 1;
      const relId = `${REL_ID_PREFIX.html}${htmlId}`;
      htmls[relId] = html;
      return relId;
    },

    get lastShapeId() {
      return shapeId;
    },
    nextShapeId() {
      shapeId += 1;
      return shapeId;
    },

    park(kind, node) {
      pending.set(kind, node);
    },
    takePending(kind) {
      const node = pending.get(kind);
      if (node != null) pending.delete(kind);
      return node;
    },

    textRunProps: undefined,
  };
}
