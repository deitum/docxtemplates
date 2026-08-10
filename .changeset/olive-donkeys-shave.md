---
'@deitum/docxtemplates': patch
---

Fix `ImageError` so that it is recognised by `instanceof`. `CommandExecutionError`
reset the prototype of every instance to its own, which meant that errors thrown
by `IMAGE` commands were reported as plain `CommandExecutionError`s and
`err instanceof ImageError` was always `false`.
