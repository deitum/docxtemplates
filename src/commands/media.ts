/**
 * Builders for the OOXML that the IMAGE, LINK and HTML commands drop into the
 * report. Each one parks its node on the context; the walk splices it into the
 * output tree when it comes back up past the `w:t` / `w:r` / `w:p` it replaces.
 */
import {
  ATag,
  DrawAttr,
  DrawingExtUri,
  EMU_PER_CM,
  MsTag,
  Ns,
  PicTag,
  RAttr,
  ROTATION_UNITS_PER_DEGREE,
  WAttr,
  WpTag,
  WTag,
  XmlnsAttr,
} from '../ooxml';
import { newNonTextNode, newTextNode } from '../reportUtils';
import {
  type Context,
  type Image,
  ImageExtensions,
  type ImagePars,
  type LinkPars,
  type Node,
} from '../types';

const node = newNonTextNode;

/** The extension SVG images are recognised by. */
const SVG_EXTENSION = '.svg' as const;

/** `w:hyperlink@w:history` — mark the link as visited-tracking, as Word does. */
const LINK_KEEP_HISTORY = '1' as const;

/** `w:u@w:val` — the underline a link gets when it carries no formatting. */
const UNDERLINE_SINGLE = 'single' as const;

/**
 * A 1x1 transparent PNG, used when an SVG comes without a thumbnail. The
 * thumbnail is scaffolding the docx standard requires rather than something
 * anyone sees, but without one the SVG does not render at all.
 */
const PLACEHOLDER_THUMBNAIL: Image = {
  data: 'bm90aGluZwo=',
  extension: '.png',
};

/** Fixed attribute values of the picture XML below, named where not obvious. */
const PicValue = {
  /** `pic:cNvPr@id` — unused by Word for pictures, which are keyed by docPr. */
  unusedId: '0',
  /** Boolean-valued DrawingML attributes are the strings '0' and '1'. */
  on: '1',
  off: '0',
  /** `a:blip@cstate` — the image is stored at print quality. */
  printQuality: 'print',
  /** `a:prstGeom@prst` — a plain rectangle, i.e. no shape cropping. */
  rectangle: 'rect',
  /** `pic:spPr@bwMode` — let Word decide how to render in black and white. */
  autoBwMode: 'auto',
} as const;

// ==========================================
// Images
// ==========================================

function validateImage(img: Image) {
  const isBinary =
    img.data instanceof Uint8Array ||
    img.data instanceof ArrayBuffer ||
    typeof img.data === 'string';
  if (!isBinary) {
    throw new Error(
      'image .data property needs to be provided as Uint8Array (e.g. Buffer), ArrayBuffer, or as a base64-encoded string'
    );
  }
  if (!ImageExtensions.includes(img.extension)) {
    throw new Error(
      `An extension (one of ${ImageExtensions}) needs to be provided when providing an image or a thumbnail.`
    );
  }
}

function validateImagePars(pars: ImagePars) {
  if (!Number.isFinite(pars.width))
    throw new Error(`invalid image width: ${pars.width} (in cm)`);
  if (!Number.isFinite(pars.height))
    throw new Error(`invalid image height: ${pars.height} (in cm)`);
  validateImage(pars);
  if (pars.thumbnail) validateImage(pars.thumbnail);
}

/** Validates an image and registers it, returning its relationship id. */
const imageToContext = (ctx: Context, img: Image): string => {
  validateImage(img);
  return ctx.resources.addImage(img);
};

function getImageData(imagePars: ImagePars): Image {
  const { data, extension } = imagePars;
  if (!extension) {
    throw new Error(
      'If you return image `data`, make sure you return an extension as well!'
    );
  }
  return { extension, data };
}

/**
 * The `a:extLst` entries of the picture. SVGs get a second one pointing at the
 * SVG itself, with the (possibly synthetic) thumbnail taking the place of the
 * image proper — that is how Word renders SVGs, and how older versions that
 * cannot fall back gracefully.
 */
function buildImageExtensions(
  ctx: Context,
  imagePars: ImagePars,
  imgRelId: string
): { extNodes: Node[]; renderedRelId: string } {
  const extNodes: Node[] = [
    node(ATag.ext, { [DrawAttr.uri]: DrawingExtUri.useLocalDpi }, [
      node(MsTag.useLocalDpi, {
        [XmlnsAttr.a14]: Ns.msDrawing2010,
        val: PicValue.off,
      }),
    ]),
  ];

  if (ctx.resources.images[imgRelId]?.extension !== SVG_EXTENSION) {
    return { extNodes, renderedRelId: imgRelId };
  }

  // Default to an empty thumbnail: it is not critical, just part of the docx
  // standard's scaffolding. Without one the SVG won't render at all, even in
  // versions of Word that don't otherwise need it.
  const thumbnail: Image = imagePars.thumbnail ?? PLACEHOLDER_THUMBNAIL;
  const thumbRelId = imageToContext(ctx, thumbnail);
  extNodes.push(
    node(ATag.ext, { [DrawAttr.uri]: DrawingExtUri.svgBlip }, [
      node(MsTag.svgBlip, {
        [XmlnsAttr.asvg]: Ns.msSvg2016,
        [RAttr.embed]: imgRelId,
      }),
    ])
  );
  return { extNodes, renderedRelId: thumbRelId };
}

export const processImage = (ctx: Context, imagePars: ImagePars) => {
  validateImagePars(imagePars);
  const cx = (imagePars.width * EMU_PER_CM).toFixed(0);
  const cy = (imagePars.height * EMU_PER_CM).toFixed(0);

  const relId = imageToContext(ctx, getImageData(imagePars));
  const id = String(ctx.resources.lastShapeId);
  const alt = imagePars.alt || '';
  const rot = imagePars.rotation
    ? (imagePars.rotation * ROTATION_UNITS_PER_DEGREE).toString()
    : undefined;

  const { extNodes, renderedRelId } = buildImageExtensions(
    ctx,
    imagePars,
    relId
  );

  const pic = node(PicTag.pic, { [XmlnsAttr.pic]: Ns.drawingPicture }, [
    node(PicTag.nvPicPr, {}, [
      node(PicTag.cNvPr, {
        [DrawAttr.id]: PicValue.unusedId,
        [DrawAttr.name]: `Picture ${id}`,
        [DrawAttr.descr]: alt,
      }),
      node(PicTag.cNvPicPr, {}, [
        node(ATag.picLocks, {
          noChangeAspect: PicValue.on,
          noChangeArrowheads: PicValue.on,
        }),
      ]),
    ]),
    node(PicTag.blipFill, {}, [
      node(
        ATag.blip,
        { [RAttr.embed]: renderedRelId, cstate: PicValue.printQuality },
        [node(ATag.extLst, {}, extNodes)]
      ),
      node(ATag.srcRect),
      node(ATag.stretch, {}, [node(ATag.fillRect)]),
    ]),
    node(PicTag.spPr, { bwMode: PicValue.autoBwMode }, [
      node(ATag.xfrm, rot ? { rot } : {}, [
        node(ATag.off, { x: '0', y: '0' }),
        node(ATag.ext, { cx, cy }),
      ]),
      node(ATag.prstGeom, { prst: PicValue.rectangle }, [node(ATag.avLst)]),
      node(ATag.noFill),
      node(ATag.ln, {}, [node(ATag.noFill)]),
    ]),
  ]);

  // `distT/B/L/R` are the text wrap distances around an inline drawing; zero
  // keeps the image flush with the surrounding text.
  const drawing = node(WTag.drawing, {}, [
    node(WpTag.inline, { distT: '0', distB: '0', distL: '0', distR: '0' }, [
      node(WpTag.extent, { cx, cy }),
      node(WpTag.docPr, {
        [DrawAttr.id]: id,
        [DrawAttr.name]: `Picture ${id}`,
        [DrawAttr.descr]: alt,
      }),
      node(WpTag.cNvGraphicFramePr, {}, [
        node(ATag.graphicFrameLocks, {
          [XmlnsAttr.a]: Ns.drawingMain,
          noChangeAspect: PicValue.on,
        }),
      ]),
      node(ATag.graphic, { [XmlnsAttr.a]: Ns.drawingMain }, [
        node(ATag.graphicData, { [DrawAttr.uri]: Ns.drawingPicture }, [pic]),
      ]),
    ]),
  ]);

  ctx.resources.park('image', {
    node: drawing,
    ...(imagePars.caption
      ? {
          extra: [
            node(WTag.br),
            node(WTag.t, {}, [newTextNode(imagePars.caption)]),
          ],
        }
      : {}),
  });
};

/**
 * The highest `wp:docPr` id already used in a document. New images continue
 * from there, so that they don't collide with the ones the template shipped
 * with.
 */
export function findHighestImgId(mainDoc: Node): number {
  const ids: number[] = [];
  const search = (n: Node) => {
    for (const c of n._children) {
      if (c._fTextNode) continue;
      if (c._tag === WpTag.docPr) {
        const raw = c._attrs[DrawAttr.id];
        if (typeof raw === 'string') {
          const id = Number.parseInt(raw, 10);
          if (Number.isSafeInteger(id)) ids.push(id);
        }
      }
      if (c._children.length > 0) search(c);
    }
  };
  search(mainDoc);
  return ids.length > 0 ? Math.max(...ids) : 0;
}

// ==========================================
// Links and HTML
// ==========================================

export const processLink = (ctx: Context, linkPars: LinkPars) => {
  const { url, label = url } = linkPars;
  const relId = ctx.resources.addLink(url);
  const { textRunProps } = ctx.resources;
  ctx.resources.park('link', {
    node: node(
      WTag.hyperlink,
      { [RAttr.id]: relId, [WAttr.history]: LINK_KEEP_HISTORY },
      [
        node(WTag.r, {}, [
          // A link with no formatting of its own gets the conventional underline
          textRunProps ||
            node(WTag.rPr, {}, [
              node(WTag.u, { [WAttr.val]: UNDERLINE_SINGLE }),
            ]),
          node(WTag.t, {}, [newTextNode(label)]),
        ]),
      ]
    ),
  });
};

export const processHtml = (ctx: Context, data: string) => {
  const relId = ctx.resources.addHtml(data);
  ctx.resources.park('html', {
    node: node(WTag.altChunk, { [RAttr.id]: relId }),
  });
};
