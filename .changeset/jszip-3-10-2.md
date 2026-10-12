---
'@deitum/docxtemplates': patch
---

Require jszip 3.10.2, which recognises binary data (`Uint8Array`,
`ArrayBuffer`) created in another realm — an iframe, say — and reads a `Blob`
through `blob.arrayBuffer()` where there is no `FileReader`, as in Node.
