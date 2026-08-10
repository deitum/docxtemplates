import { WTag, XML_SPACE_PRESERVE, XmlAttr } from './ooxml';
import { insertTextSiblingAfter, nextNodeInTree, tagOf } from './reportUtils';
import { Command, type Node } from './types';

// In-place
// In case of split commands (or even split delimiters), joins all the pieces
// at the starting node
const preprocessTemplate = (
  template: Node,
  delimiter: [string, string],
  preserveSpace: boolean
) => {
  let node: Node | null = template;
  let fCmd = false;
  let openNode = null;
  let idxDelimiter = 0;
  const placeholderCmd = `${delimiter[0]}${Command.CMD_NODE}${delimiter[1]}`;

  while (node != null) {
    const tag = tagOf(node);

    // Add `xml:space` attr `preserve` to `w:t` tags
    if (preserveSpace && !node._fTextNode && tag === WTag.t) {
      node._attrs[XmlAttr.space] = XML_SPACE_PRESERVE;
    }

    // Add a space if we reach a new `w:p` tag and there's an open node (hence, in a command)
    if (tag === WTag.p && openNode) {
      openNode._text += ' ';
    }

    // Process text nodes inside `w:t` tags
    if (node._fTextNode && tagOf(node._parent) === WTag.t) {
      if (openNode == null) openNode = node;
      const textIn = node._text;
      node._text = '';
      for (let i = 0; i < textIn.length; i++) {
        const c = textIn[i];

        // What's the current expected delimiter
        const currentDelimiter = fCmd ? delimiter[1] : delimiter[0];

        // Matches the expected delimiter character
        if (c === currentDelimiter[idxDelimiter]) {
          idxDelimiter += 1;

          // Finished matching delimiter? Then toggle `fCmd`,
          // add a new `w:t` + text node (either before or after the delimiter),
          // depending on the case
          if (idxDelimiter === currentDelimiter.length) {
            fCmd = !fCmd;
            const fNodesMatch = node === openNode;
            if (fCmd && openNode._text.length) {
              openNode = insertTextSiblingAfter(openNode);
              if (fNodesMatch) node = openNode;
            }
            openNode._text += currentDelimiter;
            if (!fCmd && i < textIn.length - 1) {
              openNode = insertTextSiblingAfter(openNode);
              if (fNodesMatch) node = openNode;
            }
            idxDelimiter = 0;
            if (!fCmd) openNode = node; // may switch open node to the current one
          }

          // Doesn't match the delimiter, but we had some partial match
        } else if (idxDelimiter) {
          openNode._text += currentDelimiter.slice(0, idxDelimiter);
          idxDelimiter = 0;
          if (!fCmd) openNode = node;
          openNode._text += c;

          // General case
        } else {
          openNode._text += c;
        }
      }

      // Close the text node if nothing's pending
      if (!fCmd && !idxDelimiter) openNode = null;

      // If text was present but not any more, add a placeholder, so that this node
      // will be purged during report generation
      if (textIn.length && !node._text.length) node._text = placeholderCmd;
    }

    node = nextNodeInTree(node);
  }
  return template;
};

// ==========================================
// Public API
// ==========================================
export default preprocessTemplate;
