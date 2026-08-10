---
'@deitum/docxtemplates': patch
---

Fix relationships generated for header and footer parts. `LINK` and `HTML`
commands used in a header or footer registered their relationship in
`word/_rels/document.xml.rels` instead of the part's own `.rels`, so Word could
not resolve the `r:id` and refused to open the report. Two headers (or a header
and the main document) using `HTML` also overwrote each other's chunk file,
because both were named after the main document.

Each part now writes its images, hyperlinks and HTML chunks into its own
relationships, as images already did.
