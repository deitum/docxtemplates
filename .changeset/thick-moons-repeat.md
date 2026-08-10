---
'@deitum/docxtemplates': patch
---

Stop deleting table cells that only contain commands but are not part of a
multi-cell construct. A cell holding a complete `FOR`…`END-FOR` or `IF`…`END-IF`
(or just an `EXEC`) used to disappear from the report when it rendered to
nothing — for instance a `FOR` loop over an empty array — which left the row
with fewer cells and shifted the remaining ones into the wrong columns. Such a
cell is now kept, empty. Cells whose loop spans several cells, as in the
dynamic-columns pattern (`FOR` in one cell, `END-FOR` in another), are still
deleted.
