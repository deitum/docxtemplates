# DOCX template skill — agent instructions

You convert filled-out, real-world `.docx` documents into reusable templates for
[@deitum/docxtemplates](https://github.com/deitum/docxtemplates).

The tools handle the binary I/O and the XML surgery. **You are the analyst**: you
decide what should be templated, and you check that the result is right by
rendering it.

Paths below are relative to this skill's directory. Installed as a plugin, that
is `${CLAUDE_PLUGIN_ROOT}/skills/docx-template/`.

---

## Workflow

### 1. Extract

A `.docx` is a zip of XML; you cannot read it directly.

```bash
node agent/dist/analyze.mjs examples/contract.docx              # structured JSON
node agent/dist/analyze.mjs examples/*.docx                     # one block per file
node agent/dist/analyze.mjs --text-only examples/*.docx         # prose only
node agent/dist/analyze.mjs --aliases aliases.json tpl.docx     # localized commands
```

The JSON gives you `plainText`, `main.paragraphs` (with per-run formatting),
`main.tables` (rows and cells), `headersAndFooters`, and `commands` — anything
the file already contains as a command. Paragraph and table indices in that
output are the coordinates every other tool takes.

Start with `--text-only` when comparing documents; switch to the full JSON when
you need to pinpoint a cell or check whether a run is bold.

### 2. Decide

Read the extraction and work out three things.

**Expressions** — content that varies between instances. Values are arbitrary
JavaScript evaluated against the report data, not tag names:

```
client.name
order.total.toFixed(2)
new Date(order.date).toLocaleDateString('ru-RU')
people.filter(p => p.active).length
```

**Loops** — repeated structures:

- _Table rows_: a header row plus two or more similar data rows.
- _Paragraph blocks_: numbered clauses, one block per person, signature blocks.
  This is the common case in legal and formal documents.
- Comparing documents: a block appearing three times in one and five times in
  another is a loop. Fields that repeat per item in the input data are a loop.

**Conditionals** — blocks present in one document and absent from another.

Present your reading to the user and confirm before generating.

### 3. Generate

```bash
node agent/dist/generate.mjs examples/contract.docx field-mapping.json templates/contract.docx
```

### 4. Verify

Not optional. `verify` is the reason this skill can be trusted: it is the same
engine that will render the template in production.

```bash
node agent/dist/verify.mjs templates/contract.docx templates/contract_sample_data.json -o /tmp/report.docx
```

It lists every command, refuses unbalanced `FOR`/`IF` blocks and curly quotes,
renders with your sample data, and prints the resulting text. Read that text.
A template can render cleanly and still say the wrong thing.

Run it without a data file for a quick structural check:

```bash
node agent/dist/verify.mjs templates/contract.docx
```

Flags: `--delimiter +++`, `--aliases aliases.json`,
`--context context.mjs` (a module default-exporting `additionalJsContext`, for
`IMAGE`/`LINK` helpers), `--allow-nullish` (permit commands that evaluate to
`null`), `--quiet` (skip the rendered text).

### 5. Refine

For small changes to a template that already exists — especially one a human has
restyled in Word:

```bash
node agent/dist/refine.mjs templates/contract.docx modifications.json templates/contract.docx
```

---

## Commands in this library

| Syntax                                         | Purpose                                                     |
| ---------------------------------------------- | ----------------------------------------------------------- |
| `+++INS expr+++`                               | Insert the value of a JS expression                         |
| `+++= expr+++`, `+++expr+++`                   | The same, shorthand                                         |
| `+++FOR x IN expr+++` … `+++END-FOR x+++`      | Loop over an array                                          |
| `+++$x.field+++`                               | The current item — note the `$` prefix                      |
| `+++$idx+++`                                   | Index in the innermost loop, from 0                         |
| `+++IF expr+++` … `+++END-IF+++`               | Conditional block                                           |
| `+++ELSE-IF expr+++`, `+++ELSE+++`             | Further branches, `ELSE` last                               |
| `+++IMAGE expr+++`                             | Expression returns `{width, height, data, extension}`       |
| `+++LINK expr+++`                              | Expression returns `{url, label}`                           |
| `+++HTML expr+++`                              | HTML string (Word only, uses altchunk)                      |
| `+++EXEC code+++`, `+++! code+++`              | Run code, insert nothing                                    |
| `+++ALIAS name INS $x.field+++`, `+++*name+++` | Name a command, then reuse it — handy in narrow table cells |
| `+++QUERY ...+++`                              | Passed to the caller's data resolver                        |

---

## What breaks templates in this library

These are the differences from `{tag}`-style engines. Getting one wrong produces
a template that looks fine and fails at render time.

### 1. Block commands need a paragraph — or a row — to themselves

The engine deletes any `w:p`, `w:tr` or `w:tc` that ends up holding only
commands. That is what makes markers disappear from the report, and it means a
table loop is **three rows**, not tags stuffed into one:

```
| +++FOR person IN people+++ |                          |   <- row vanishes
| +++INS $person.name+++     | +++INS $person.since+++  |   <- repeats
| +++END-FOR person+++       |                          |   <- row vanishes
```

`generate.mjs` does this for you. If you ever hand-edit, keep the shape.

The same for paragraph blocks: `+++FOR clause IN clauses+++` on a line of its
own before the block, `+++END-FOR clause+++` on a line of its own after it.

### 2. Loop variables carry a `$`

Inside `+++FOR person IN project.people+++`, the item is `$person`, not
`person`. Forgetting the `$` gives you a "not defined" error at render time.

### 3. Section numbering is computed, not supplied

When a loop generates numbered sections, the numbers can no longer be literals.
Because commands are JavaScript, you compute them in the template rather than
making the caller precompute a variable per section:

```
1. Opening of the meeting                       <- fixed
2. Election of officers                         <- fixed

+++FOR clause IN transferClauses+++
+++= $idx + 3+++. Transfer of shares from +++= $clause.from+++
Resolution no. +++= $idx + 2+++.
+++END-FOR clause+++

+++= transferClauses.length + 3+++. Amendments to the articles
+++= transferClauses.length + 4+++. Closing
```

`$idx` counts from 0 in the innermost loop, so `$idx + 3` starts the loop's
sections at 3. After the loop, use `<array>.length + n`. The data stays clean:
just the array.

### 4. Word replaces straight quotes with curly ones

`'aubergine'` becomes `‘aubergine’` the moment someone edits the template in
Word, and the expression stops being valid JavaScript. Prefer expressions
without string literals. `verify` flags curly quotes it finds; the caller can
also pass `fixSmartQuotes: true` to `createReport`.

### 5. A whole `IF` block cannot nest inside one paragraph or row

An inline `IF … END-IF` may live in a single paragraph
(`Status: +++IF ok+++fine+++ELSE+++not so fine+++END-IF+++`), but two nested
ones in the same paragraph or table row are rejected.

### 6. `END-FOR` must name its loop variable

`+++END-FOR person+++`, not a bare `+++END-FOR+++`. `verify` checks that the
names match up.

### 7. Commands split across runs are fine

Word splits `+++INS client.name+++` across as many runs as it likes. The engine
reassembles them before parsing, and so do these tools — `analyze` shows you the
reassembled text, and `refine` matches against it. Never try to work around
splitting by hand.

### 8. `{` and `}` make poor delimiters

They collide with JavaScript object literals and arrow-function bodies. Stay
with `+++` unless the user asks otherwise; if they want something tidier,
suggest `{#` … `#}`.

---

## `field-mapping.json`

```json
{
  "cmdDelimiter": "+++",
  "variables": {
    "Acme Corporation": "client.name",
    "15 January 2025": "formatDate(order.date)",
    "1 234,50": "order.total.toFixed(2)"
  },
  "tableLoops": [
    {
      "var": "item",
      "over": "order.items",
      "tableIndex": 0,
      "startRow": 1,
      "endRow": 3,
      "fields": {
        "0": "$item.description",
        "1": "$item.quantity",
        "2": "$item.price.toFixed(2)"
      }
    }
  ],
  "sectionLoops": [
    {
      "var": "clause",
      "over": "transferClauses",
      "startText": "3. Transfer of shares",
      "endText": "signed by the parties"
    }
  ],
  "conditionals": [
    {
      "expr": "order.discount > 0",
      "paragraphText": "A discount of",
      "endParagraphText": "applies to this order"
    }
  ]
}
```

**Every value is a JavaScript expression.** No braces, no `#`, no `/` — the tool
writes the `+++INS …+++` around it.

`variables` — keys must match the document text **exactly**, as `analyze`
reports it (that is the reassembled text, so run splitting is not your problem).
Longest keys are applied first, so mapping both `Acme` and `Acme Corporation`
does the right thing. Applied to headers and footers as well as the body. A key
that matches nothing is an error, not a silent skip.

`tableLoops` — `tableIndex` counts tables in document order, skipping tables
nested inside other tables. `startRow` is the first data row (default 1, after
the header). `endRow` is the last row the loop absorbs (default: the last row) —
set it when the table ends with a totals row that must survive. `fields` maps a
0-based column index to the expression for that cell; the cell's existing
formatting is kept.

`sectionLoops` and `conditionals` — located by text. `startText` and `endText`
are substrings of the first and last paragraph of the block; `endText` is
searched from the start paragraph onwards, so a phrase occurring both before and
inside the block still resolves correctly. Omit the end and the block is a single
paragraph. These apply to the main document body only.

`commandNames` — for localized templates, see below.

### Order of application

`generate` applies table loops, then section loops, then conditionals, then
variables — because everything except variables is located by the document's
original prose. Once a phrase has become `+++INS …+++`, nothing will match it
again. Write your mapping with that in mind: `startText` and `paragraphText`
must be the text as it appears in the _filled-out_ document.

---

## `modifications.json`

```json
{
  "cmdDelimiter": "+++",
  "modifications": [
    {
      "type": "replaceCommand",
      "from": "INS client.name",
      "to": "INS customer.fullName"
    },
    { "type": "renameExpression", "from": "client.", "to": "customer." },
    { "type": "removeCommand", "code": "INS internalNote", "replaceWith": "" },
    { "type": "addCommand", "text": "Draft", "code": "document.status" },
    {
      "type": "wrapFor",
      "var": "clause",
      "over": "clauses",
      "startText": "3. Transfer",
      "endText": "signed by"
    },
    {
      "type": "wrapIf",
      "expr": "order.discount > 0",
      "startText": "A discount of"
    },
    { "type": "mergeFloatingTables" }
  ]
}
```

| Type                  | Effect                                                                                            |
| --------------------- | ------------------------------------------------------------------------------------------------- |
| `replaceCommand`      | Swaps a whole command whose code matches `from` exactly                                           |
| `renameExpression`    | Rewrites `from` → `to` inside every command containing it. Scoped to commands: prose is untouched |
| `removeCommand`       | Deletes the command; `replaceWith` is literal text left in its place, not a command               |
| `addCommand`          | Turns literal document text into `+++INS code+++`                                                 |
| `wrapFor` / `wrapIf`  | Inserts the command paragraphs around a block, exactly as `generate` does                         |
| `mergeFloatingTables` | See below                                                                                         |

Command edits run over the body, headers and footers. `wrapFor`, `wrapIf` and
`mergeFloatingTables` apply to the body.

---

## Working from several example files

The primary case. Two or more filled-out final documents; one template out.

1. Extract each with `--text-only` and diff them **semantically** — do not
   assume paragraph indices line up.
2. Text that differs → an expression. Text that is identical → boilerplate.
3. Look for structural differences, not just different values:
   - two shareholders in one document, three in the other → a loop
   - five numbered clauses versus three → a loop, plus computed section numbers
   - a discount paragraph in one and not the other → a conditional
4. Keep the document's formatting patterns. If an address reads
   `Street 1, 12345 Town`, the template says
   `+++= a.street+++, +++= a.zip+++ +++= a.town+++` — not one expression for the
   whole line, unless the data really does hold it that way.
5. Show the comparison to the user and get confirmation before generating.

## Working from input data plus a document

The user supplies a completed document and the data behind it — a form, PDF,
spreadsheet, JSON, email.

- **Trace every field.** One name may appear in the header, the body and the
  signature block; all of them become the same expression.
- **Let the data's shape drive the loops.** Fields that repeat per shareholder or
  per line item are an array, and that array is the loop.
- **Name expressions after the data**, converted to the document's own model:
  a JSON key `company_business_name` becomes `company.businessName` if you are
  also reshaping the data, or stays `company_business_name` if you are not.
  Decide once and be consistent.
- **Match the document, not the input.** An address split across five form fields
  may appear as one formatted line; the template mirrors the document.

## Converting a template written in another syntax

Documents already carrying `${var}`, `{{var}}` or `<%=var%>`.

**Use the coded template as the base.** Do not reverse-engineer from a filled-out
example when a coded template exists — the placeholders are already in the right
places, and converting them is a mechanical, reliable edit.

1. `node agent/dist/analyze.mjs coded_template.docx` and collect every
   placeholder from `plainText`.
2. Map each one in `variables`, with the placeholder as the key:

```json
{
  "variables": {
    "${company_businessName}": "company.businessName",
    "${seller_fullName}": "seller.fullName"
  }
}
```

3. Map **every** placeholder. A leftover `${...}` is not an error in this
   library — it renders as literal text — which makes it easy to miss. Read
   `verify`'s rendered output and search it for `${`.
4. If the source has loop or conditional markers (`${#items}` … `${/items}`),
   do not map them as variables. Delete them with `removeCommand`-style edits or
   a fresh mapping, and re-express the block as a `sectionLoop` or `tableLoop`,
   because this library needs the markers in paragraphs of their own.

---

## Localized commands

The library accepts alternative names for commands and operators, so a template
can be written in the user's own language. When the user's documents are in
Russian and they want the template to read that way too, put the alternative
names in the mapping:

```json
{
  "commandNames": {
    "INS": "=",
    "FOR": "ДЛЯ",
    "IN": "ИЗ",
    "END-FOR": "КОНЕЦ ДЛЯ",
    "IF": "ЕСЛИ",
    "END-IF": "КОНЕЦ ЕСЛИ"
  }
}
```

`generate` inverts that map and writes it beside the template as
`<name>_aliases.json`:

```json
{
  "commandAliases": { "ДЛЯ": "FOR", "КОНЕЦ ДЛЯ": "END-FOR" },
  "operatorAliases": { "ИЗ": "IN" }
}
```

**That file is part of the template.** Pass it with `--aliases` to `analyze` and
`verify`, and give it to whoever calls `createReport` — without it, the engine
cannot tell `+++ДЛЯ компания ИЗ компании+++` from an `INS` command, and the
template silently renders as gibberish.

Operators can be localized too. `generate` does not write these, because it
never emits an operator; add them to the aliases file by hand if the user wants
to write conditions in their own language:

```json
{
  "operatorAliases": {
    "больше или равно": ">=",
    "больше": ">",
    "равно": "===",
    "и": "&&"
  }
}
```

Longest alias wins, so `больше или равно` is not mistaken for `больше`, and
aliases are never substituted inside string literals.

---

## Images, links and HTML

`+++IMAGE expr+++` expects the expression to return
`{ width, height, data, extension }` — width and height in centimetres, `data` a
base64 string or buffer, `extension` one of `.png .gif .jpg .jpeg .svg`. Those
helpers live in the caller's `additionalJsContext`, so to verify a template that
uses them, write a small module and pass it:

```js
// context.mjs
export default {
  logo: () => ({
    width: 4,
    height: 2,
    data: fs.readFileSync('logo.png'),
    extension: '.png',
  }),
};
```

```bash
node agent/dist/verify.mjs templates/contract.docx data.json --context context.mjs
```

Centre an image by centring the `IMAGE` command in the template.

## Floating tables

Word uses absolutely positioned tables (`w:tblpPr`) for side-by-side layouts such
as signature blocks. They drift out of alignment easily and are miserable to
template. Merge them into one inline table:

```json
{ "modifications": [{ "type": "mergeFloatingTables" }] }
```

With no `tableIndices`, every floating table in the document is merged into the
first one. Pass `"tableIndices": [0, 1]` to merge specific tables instead. The
merged table keeps the original column widths, loses its borders so the seams
stay invisible, and drops the empty paragraphs that sat between the originals.

---

## Common mistakes

1. **Markers sharing a paragraph with content.** `FOR`, `END-FOR`, `IF` and
   `END-IF` need a paragraph or row of their own, or the surrounding text is
   deleted along with them.
2. **A missing `$`** on a loop variable.
3. **Curly quotes** in an expression, courtesy of Word's autocorrect.
4. **Mapping a variable that is not in the document.** `generate` fails loudly
   on this — copy the text from `analyze`, do not retype it.
5. **Substring collisions.** Mapping `Acme` when `Acme Corporation` also appears
   leaves ` Corporation` stranded. Map the longer string too; `generate` applies
   the longest first.
6. **Templating the wrong table.** `tableIndex` skips nested tables. Check
   against `analyze`'s output rather than counting in Word.
7. **Nullish results.** `verify` rejects a command that evaluates to `null` or
   `undefined`, because it almost always means a typo in the expression or a
   missing key in the data. Pass `--allow-nullish` only when a blank is genuinely
   intended.
8. **Skipping the render.** A template that lists the right commands can still
   produce nonsense. Always read `verify`'s output.

## Where to put things

- Source documents: `examples/<name>.docx`
- Templates: `templates/<name>.docx`
- Sample data: `templates/<name>_sample_data.json`
- Mappings and modification files: alongside, for the next round of edits
