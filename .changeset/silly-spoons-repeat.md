---
'@deitum/docxtemplates': patch
---

Make report generation substantially faster, and fix two sandbox bugs that fell
out of the way it used to work.

Every command used to get a brand new JS evaluation context: `vm.createContext`
ran once per command, per loop iteration. Contextifying an object is by far the
most expensive thing `vm` does, so on a template with loops it accounted for
almost all of the generation time. Each document part now evaluates its snippets
in one context, and compiled snippets are cached. A report with a 2000-iteration
`FOR` and three insertions per row went from 784 ms to 35 ms.

Two consequences of the old behaviour are fixed by this, both cases where a
snippet could not see what an earlier snippet had done:

- A function defined in one snippet closed over the context it was defined in,
  so calling it from a later snippet updated variables nobody could read again.
  `EXEC bump = () => (counter += 1)` followed by `EXEC bump()` now increments the
  `counter` that the rest of the template sees.
- Values built in one snippet were made with that context's built-ins, so
  `items instanceof Array` answered `false` in the next snippet. It now answers
  `true`.

Serialising the output tree also stopped concatenating buffers per node, which
copied every byte once per level of nesting, and looking up a node's next
sibling no longer scans the parent's children — quadratic on the tens of
thousands of children a long document hangs off `w:body`. Together those cut a
20 000-paragraph document from 204 ms to 157 ms.

No change to the generated .docx: the whole fixture corpus renders byte for byte
as before.
