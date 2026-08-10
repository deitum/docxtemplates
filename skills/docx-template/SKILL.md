---
name: docx-template
description: Convert filled-out DOCX documents into reusable @deitum/docxtemplates templates. Compares multiple completed documents to work out what varies and what is boilerplate, maps input data (forms, PDFs, spreadsheets) onto the document, and renders the result to prove the template works. Use when the user wants to create, modify, or inspect a DOCX template.
argument-hint: '[path-to-docx or instruction]'
allowed-tools: 'Bash(node *)'
---

# DOCX templates for `@deitum/docxtemplates`

You turn filled-out `.docx` documents into reusable templates for
[@deitum/docxtemplates](https://github.com/deitum/docxtemplates), whose commands
are JavaScript expressions between `+++` delimiters.

Work usually starts from **completed, real-world documents** — not blank forms:

- **Several filled-out documents.** Compare two or more final documents (two
  signed contracts for different clients, three invoices) to see what changes
  between instances (expressions) and what stays put (boilerplate). A section
  that appears twice in one and five times in another is a loop.
- **Input data plus a document.** The user hands you a completed document
  alongside the data that produced it — a form submission, PDF, spreadsheet,
  email. Trace each field to where it lands, and let the data's natural
  grouping tell you where the loops go.
- **A template in someone else's syntax.** Documents already carrying `${var}`
  or `{{var}}` placeholders. Convert the syntax. Always build on the coded
  template rather than reverse-engineering a filled-out example — the
  placeholders are already in the right places.

The tools are pre-built. There is nothing to install.

**Read `AGENT_INSTRUCTIONS.md` in this skill's directory before starting.** It
carries the mapping-file schemas and the mistakes that break templates in ways
that only show up at render time.

## Tools

All four live in `agent/dist/` next to this file. When the skill is installed as
a plugin, that is `${CLAUDE_PLUGIN_ROOT}/skills/docx-template/agent/dist/`.

| Tool                                                | What it does                                                                                                                                                   |
| --------------------------------------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `analyze.mjs <file.docx> [file2.docx ...]`          | Extracts paragraphs, runs, tables, headers and footers as JSON, plus any commands already in the file. You cannot read a binary docx — this is how you see it. |
| `generate.mjs <original.docx> <mapping.json> [out]` | Builds a template from a document and a field mapping                                                                                                          |
| `refine.mjs <template.docx> <mods.json> [out]`      | Edits an existing template in place                                                                                                                            |
| `verify.mjs <template.docx> [data.json] [-o out]`   | Lists the commands, checks they balance, renders the template with sample data and prints the resulting text                                                   |

## Workflow

### 1. Extract

Save any document the user gives you under `examples/`, then read it:

```bash
node agent/dist/analyze.mjs examples/<file>.docx            # structured JSON
node agent/dist/analyze.mjs --text-only examples/*.docx     # just the prose, for comparing
```

### 2. Decide what to templatise

You are the analyst; the tools only do the XML surgery. Read the extraction and
work out:

- **Expressions** — anything that differs between instances, or that came from
  the input data. Values are JavaScript, so `client.name`,
  `total.toFixed(2)` and `date.toLocaleDateString('ru-RU')` are all fine.
- **Loops** — repeated table rows, and repeated blocks of paragraphs (numbered
  clauses, one block per shareholder, signature blocks). Paragraph blocks are
  the common case in legal and formal documents.
- **Conditionals** — blocks that appear in one document and not the other.

Show the user what you found and confirm before generating.

### 3. Generate

Write `field-mapping.json` from your analysis, then:

```bash
node agent/dist/generate.mjs examples/<original>.docx field-mapping.json templates/<name>.docx
```

### 4. Verify — do not skip this

Write `sample_data.json` with realistic values for every expression in the
template, then render it:

```bash
node agent/dist/verify.mjs templates/<name>.docx templates/<name>_sample_data.json -o /tmp/report.docx
```

`verify` prints the rendered text. Read it. A template that renders without
error can still say the wrong thing. Iterate — with `refine.mjs` for small
fixes — until the output is right.

## Command syntax

| Syntax                                                   | Purpose                             | Example                         |
| -------------------------------------------------------- | ----------------------------------- | ------------------------------- |
| `+++INS expr+++`                                         | Insert a value                      | `+++INS client.name+++`         |
| `+++= expr+++` or `+++expr+++`                           | Same thing, shorthand               | `+++= client.name+++`           |
| `+++FOR x IN expr+++` … `+++END-FOR x+++`                | Loop over an array                  | `+++FOR item IN order.items+++` |
| `+++$x.field+++`                                         | Loop variable — note the `$`        | `+++INS $item.price+++`         |
| `+++$idx+++`                                             | Index in the innermost loop, from 0 |                                 |
| `+++IF expr+++` … `+++END-IF+++`                         | Conditional block                   | `+++IF order.discount > 0+++`   |
| `+++ELSE-IF expr+++`, `+++ELSE+++`                       | Further branches                    |                                 |
| `+++IMAGE expr+++`, `+++LINK expr+++`, `+++HTML expr+++` | Image, hyperlink, HTML              |                                 |
| `+++EXEC code+++` or `+++! code+++`                      | Run code, insert nothing            |                                 |

## Rules

- **`FOR`, `END-FOR`, `IF` and `END-IF` go in a paragraph — or a table row — of
  their own.** The engine deletes anything left holding only commands, which is
  how the markers vanish from the report. For a table loop that means three
  rows: `FOR`, the data row, `END-FOR`.
- **Prefix loop variables with `$`** inside the loop: `$person.name`.
- Expressions are JavaScript. Prefer a plain property path; reach for
  `.filter()`, ternaries and template literals when the document needs them.
- Always write `sample_data.json` alongside a template, and always run `verify`.
- Extract before you generate. Never guess at a document's structure.
- Keep the document's own formatting in the template: if an address reads
  `Street 1, 12345 Town`, the template says
  `+++= a.street+++, +++= a.zip+++ +++= a.town+++`.
