---
'@deitum/docxtemplates': minor
---

Give every error a `properties` object saying where it came from.

A report is rendered from `word/document.xml` plus every header and footer, and
with `failFast: false` the errors come back as one flat array. Until now nothing
in an error said which of those files it came from, or carried a code you could
branch on without matching against the message text.

```js
try {
  await createReport({ template, data, failFast: false });
} catch (errors) {
  for (const e of errors) {
    console.log(e.properties.part); // 'header1.xml'
    console.log(e.properties.id); // 'command_execution'
    console.log(e.properties.command); // 'user.fullName'
    console.log(e.properties.explanation); // what to do about it
  }
}
```

`id` values are listed in the exported `ErrorId` and are stable across releases;
prefer them to `message`, which is prose. Every error class now descends from the
exported `TemplateError`, so they can all be caught as one type.

Nothing existing changes: the class names, the messages, the fields the classes
already had (`command`, `alias`, `result`, `err`) and the array thrown when
`failFast` is off are all exactly as they were.

`src/index.ts` also lists its error exports explicitly instead of re-exporting
the whole module, so internal helpers cannot drift into the public API. The
exported names are unchanged apart from the additions above.
