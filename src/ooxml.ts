/**
 * The OOXML vocabulary this library reads and writes: element names, attribute
 * names, namespace URIs, relationship types, media types and package paths.
 *
 * Everything is grouped by XML prefix and named after the local part, so
 * `WTag.p` is `w:p` and stays greppable against the OOXML documentation. The
 * point is compile-time safety: a mistyped constant fails the build, whereas a
 * mistyped string literal produces a document Word silently refuses to open.
 *
 * Only what the library actually emits or looks for is listed here — this is a
 * working vocabulary, not a copy of the specification.
 */

// ==========================================
// Elements
// ==========================================

/** WordprocessingML — the document body itself. */
export enum WTag {
  /** Paragraph. */
  p = 'w:p',
  /** Text run. */
  r = 'w:r',
  /** Literal text inside a run. */
  t = 'w:t',
  /** Table. */
  tbl = 'w:tbl',
  /** Table row. */
  tr = 'w:tr',
  /** Table cell. */
  tc = 'w:tc',
  /** Run properties (formatting). */
  rPr = 'w:rPr',
  /** Underline. */
  u = 'w:u',
  /** Line break. */
  br = 'w:br',
  /** Hyperlink. */
  hyperlink = 'w:hyperlink',
  /** Embedded foreign content, e.g. an HTML fragment. */
  altChunk = 'w:altChunk',
  /** Anchor for a drawing. */
  drawing = 'w:drawing',
}

/** DrawingML picture placement inside a document. */
export enum WpTag {
  /** Floating drawing. */
  anchor = 'wp:anchor',
  /** Inline drawing. */
  inline = 'wp:inline',
  /** Non-visual drawing properties; carries the id images are numbered by. */
  docPr = 'wp:docPr',
  /** Drawing extent (size). */
  extent = 'wp:extent',
  cNvGraphicFramePr = 'wp:cNvGraphicFramePr',
}

/** DrawingML core. */
export enum ATag {
  ext = 'a:ext',
  extLst = 'a:extLst',
  blip = 'a:blip',
  srcRect = 'a:srcRect',
  stretch = 'a:stretch',
  fillRect = 'a:fillRect',
  xfrm = 'a:xfrm',
  off = 'a:off',
  prstGeom = 'a:prstGeom',
  avLst = 'a:avLst',
  noFill = 'a:noFill',
  ln = 'a:ln',
  picLocks = 'a:picLocks',
  graphic = 'a:graphic',
  graphicData = 'a:graphicData',
  graphicFrameLocks = 'a:graphicFrameLocks',
}

/** DrawingML picture. */
export enum PicTag {
  pic = 'pic:pic',
  nvPicPr = 'pic:nvPicPr',
  cNvPr = 'pic:cNvPr',
  cNvPicPr = 'pic:cNvPicPr',
  blipFill = 'pic:blipFill',
  spPr = 'pic:spPr',
}

/** Microsoft DrawingML extensions. */
export enum MsTag {
  /** DPI hint (2010 extension). */
  useLocalDpi = 'a14:useLocalDpi',
  /** SVG payload of a picture (2016 extension). */
  svgBlip = 'asvg:svgBlip',
}

/** VML — legacy shapes, still emitted by Word inside `mc:AlternateContent`. */
export enum VTag {
  shape = 'v:shape',
}

/** Open Packaging Conventions: `[Content_Types].xml` and `.rels` parts. */
export enum PkgTag {
  /** A content type bound to a file extension. */
  defaultType = 'Default',
  /** A relationship entry in a `.rels` part. */
  relationship = 'Relationship',
}

/** Document metadata parts (`docProps/`). */
export enum MetaTag {
  pages = 'Pages',
  words = 'Words',
  characters = 'Characters',
  lines = 'Lines',
  paragraphs = 'Paragraphs',
  company = 'Company',
  template = 'Template',
  title = 'dc:title',
  subject = 'dc:subject',
  creator = 'dc:creator',
  description = 'dc:description',
  lastModifiedBy = 'cp:lastModifiedBy',
  revision = 'cp:revision',
  lastPrinted = 'cp:lastPrinted',
  created = 'dcterms:created',
  modified = 'dcterms:modified',
  category = 'cp:category',
}

// ==========================================
// Attributes
// ==========================================

/** Relationship references. */
export enum RAttr {
  id = 'r:id',
  embed = 'r:embed',
}

export enum WAttr {
  /** Carries the value of most WordprocessingML property elements. */
  val = 'w:val',
  history = 'w:history',
}

export enum XmlAttr {
  space = 'xml:space',
}

/** `[Content_Types].xml` entries. */
export enum CtAttr {
  extension = 'Extension',
  contentType = 'ContentType',
  partName = 'PartName',
}

/** `.rels` entries. */
export enum RelAttr {
  id = 'Id',
  type = 'Type',
  target = 'Target',
  targetMode = 'TargetMode',
}

/** Shared by several drawing elements. */
export enum DrawAttr {
  id = 'id',
  uri = 'uri',
  name = 'name',
  descr = 'descr',
}

/** The only value `TargetMode` ever takes here: a link out of the package. */
export const TARGET_MODE_EXTERNAL = 'External' as const;

/** `xml:space="preserve"` — keeps Word from collapsing significant spaces. */
export const XML_SPACE_PRESERVE = 'preserve' as const;

// ==========================================
// Namespaces
// ==========================================

export enum Ns {
  drawingMain = 'http://schemas.openxmlformats.org/drawingml/2006/main',
  drawingPicture = 'http://schemas.openxmlformats.org/drawingml/2006/picture',
  packageRelationships = 'http://schemas.openxmlformats.org/package/2006/relationships',
  msDrawing2010 = 'http://schemas.microsoft.com/office/drawing/2010/main',
  msSvg2016 = 'http://schemas.microsoft.com/office/drawing/2016/SVG/main',
}

/** `xmlns:*` attribute names, for declaring the namespaces above inline. */
export enum XmlnsAttr {
  a = 'xmlns:a',
  a14 = 'xmlns:a14',
  asvg = 'xmlns:asvg',
  pic = 'xmlns:pic',
}

/**
 * `a:ext` is keyed by GUID rather than by name; these are the two Microsoft
 * publishes for the extensions above.
 */
export enum DrawingExtUri {
  useLocalDpi = '{28A0092B-C50C-407E-A947-70E740481C1C}',
  svgBlip = '{96DAC541-7B7A-43D3-8B79-37D633B846F1}',
}

// ==========================================
// Relationship and media types
// ==========================================

export enum RelType {
  image = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
  hyperlink = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
  /** Alternative-format chunk, i.e. the `w:altChunk` an HTML command inserts. */
  altChunk = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk',
}

/**
 * File extension -> media type, for the entries `[Content_Types].xml` needs.
 * A lookup table rather than an enum: two extensions share one media type, and
 * the keys are data (they are iterated over), not names of constants.
 */
export const IMAGE_MEDIA_TYPES: { readonly [extension: string]: string } = {
  png: 'image/png',
  jpg: 'image/jpeg',
  jpeg: 'image/jpeg',
  gif: 'image/gif',
  bmp: 'image/bmp',
  svg: 'image/svg+xml',
};

export const HTML_EXTENSION = 'html' as const;
export const HTML_MEDIA_TYPE = 'text/html' as const;

/** Media types the main document part can be declared as. */
export enum MainDocMediaType {
  document = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
  macroEnabled = 'application/vnd.ms-word.document.macroEnabled.main+xml',
}

// ==========================================
// Package layout
// ==========================================

export enum PackagePath {
  contentTypes = '[Content_Types].xml',
  /** Directory holding the document parts and their media. */
  wordDir = 'word',
  mediaDir = 'media',
  relsDir = '_rels',
  appProps = 'docProps/app.xml',
  coreProps = 'docProps/core.xml',
}

export const RELS_EXTENSION = '.rels' as const;

/** `word/_rels/<part>.rels` — where a part's relationships live. */
export const relsPathOf = (part: string): string =>
  `${PackagePath.wordDir}/${PackagePath.relsDir}/${part}${RELS_EXTENSION}`;

/** `word/media/<name>` — where an embedded image is written. */
export const mediaPathOf = (name: string): string =>
  `${PackagePath.wordDir}/${PackagePath.mediaDir}/${name}`;

/** `word/<part>` — a document part (`document.xml`, `header1.xml`, ...). */
export const partPathOf = (part: string): string =>
  `${PackagePath.wordDir}/${part}`;

// ==========================================
// Units
// ==========================================

/** English Metric Units per centimetre — how OOXML measures drawings. */
export const EMU_PER_CM = 360e3;

/**
 * OOXML states rotations in 60,000ths of a degree, positive angles moving
 * clockwise. See http://officeopenxml.com/drwSp-rotate.php
 */
export const ROTATION_UNITS_PER_DEGREE = 60e3;

// ==========================================
// Tags with buffered contents
// ==========================================

/**
 * The tags whose text and commands are buffered while walking, so that one
 * holding nothing but commands can be dropped from the report.
 */
export const BUFFER_TAGS = [WTag.p, WTag.tr, WTag.tc] as const;
export type BufferTag = (typeof BUFFER_TAGS)[number];

export const isBufferTag = (tag: string | null): tag is BufferTag =>
  tag === WTag.p || tag === WTag.tr || tag === WTag.tc;
