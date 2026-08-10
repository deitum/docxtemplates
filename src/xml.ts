import sax, { type QualifiedAttribute } from 'sax';
import { type Node } from './types';
import { logger } from './debug';

const parseXml = (templateXml: string): Promise<Node> => {
  const parser = sax.parser(true, {
    // true for XML-like (false for HTML-like)
    trim: false,
    normalize: false,
  });
  let template: Node;
  let curNode: Node | null | undefined = null;
  let numXmlElements = 0;
  return new Promise((resolve, reject) => {
    parser.onopentag = node => {
      const newNode: Node = {
        _children: [],
        _fTextNode: false,
        _tag: node.name,
        _attrs: node.attributes,
      };
      if (curNode) newNode._parent = curNode;
      if (curNode != null) curNode._children.push(newNode);
      else template = newNode;
      curNode = newNode;
      numXmlElements += 1;
    };
    parser.onclosetag = () => {
      curNode = curNode != null ? curNode._parent : null;
    };
    parser.ontext = text => {
      if (curNode == null) return;
      curNode._children.push({
        _parent: curNode,
        _children: [],
        _fTextNode: true,
        _text: text,
      });
    };
    parser.onend = () => {
      logger.debug(`Number of XML elements: ${numXmlElements}`);
      resolve(template);
    };
    parser.onerror = err => {
      reject(err);
    };
    parser.write(templateXml);
    parser.end();
  });
};

type XmlOptions = {
  literalXmlDelimiter: string;
  indentXml: boolean;
};

const XML_DECLARATION =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

/** Two spaces per level, when `indentXml` is on. */
const INDENT_STEP = '  ';

function buildXml(
  node: Node,
  options: XmlOptions,
  indent: string = '',
  firstRun: boolean = true
): Buffer {
  const xmlBuffers: Buffer[] = [
    Buffer.from(firstRun ? XML_DECLARATION : '', 'utf-8'),
  ];
  if (node._fTextNode)
    xmlBuffers.push(Buffer.from(sanitizeText(node._text, options)));
  else {
    let attrs = '';
    const nodeAttrs = node._attrs;
    Object.entries(nodeAttrs).forEach(([key, value]) => {
      if (value == null) return;
      attrs += ` ${key}="${sanitizeAttr(value)}"`;
    });
    const fHasChildren = node._children.length > 0;
    const suffix = fHasChildren ? '' : '/';

    // Conditionally add newlines and indentation based on indentXml option
    const newline = options.indentXml ? `\n${indent}` : '';
    xmlBuffers.push(Buffer.from(`${newline}<${node._tag}${attrs}${suffix}>`));

    let fLastChildIsNode = false;
    node._children.forEach(child => {
      xmlBuffers.push(
        buildXml(
          child,
          options,
          options.indentXml ? `${indent}${INDENT_STEP}` : '',
          false
        )
      );
      fLastChildIsNode = !child._fTextNode;
    });
    if (fHasChildren) {
      // Only add indentation if indentXml is true and last child is a node
      const indent2 =
        options.indentXml && fLastChildIsNode ? `\n${indent}` : '';
      xmlBuffers.push(Buffer.from(`${indent2}</${node._tag}>`));
    }
  }
  return Buffer.concat(xmlBuffers);
}

const XML_ENTITIES: { [char: string]: string } = {
  '&': '&amp;',
  '<': '&lt;',
  '>': '&gt;',
  "'": '&apos;',
  '"': '&quot;',
};

// One pass, so that the `&` of an entity this very call produced is not escaped
// again — which is why the order of the replacements used to matter.
const escape = (str: string, chars: RegExp) =>
  str.replace(chars, char => XML_ENTITIES[char] ?? char);

const TEXT_SPECIAL_CHARS = /[&<>]/g;
const ATTR_SPECIAL_CHARS = /[&<>'"]/g;

/**
 * Escapes the text of a node, except inside the `literalXmlDelimiter` markers:
 * that is how a command result gets to inject raw markup (a `w:br`, say).
 */
const sanitizeText = (str: string, options: XmlOptions) => {
  const segments = str.split(options.literalXmlDelimiter);
  // Every other segment is literal XML, starting with the second one.
  return segments
    .map((segment, idx) =>
      idx % 2 === 0 ? escape(segment, TEXT_SPECIAL_CHARS) : segment
    )
    .join('');
};

const sanitizeAttr = (attr: string | QualifiedAttribute) =>
  escape(typeof attr === 'string' ? attr : attr.value, ATTR_SPECIAL_CHARS);

// ==========================================
// Public API
// ==========================================
export { parseXml, buildXml };
