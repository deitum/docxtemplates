---
'@deitum/docxtemplates': minor
---

Add a `transformCode` option, which rewrites the JavaScript of a command before
it is evaluated.

It receives the code of every command that runs some — the expression of `INS`,
`IF`, `ELSE-IF`, `FOR`, `IMAGE`, `LINK` and `HTML`, and the code of `EXEC` —
once `fixSmartQuotes` and `operatorAliases` have been applied, and returns the
code to run instead. That makes room for what no option covers, such as
comparing `dd.mm.yyyy` strings as dates rather than as strings:

```js
createReport({
  template,
  data,
  // `a > b` → `compare(a, '>', b)`
  transformCode: code => rewriteComparisons(code),
  additionalJsContext: { compare },
});
```

The rewrite applies whatever the sandbox, `runJs` and `noSandbox` included, and
is done once per distinct snippet and document part, so a loop does not redo it
on every iteration. Errors and `errorHandler` still quote the command as the
template has it. Templates rendered without the option behave exactly as before.
