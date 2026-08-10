# Reading an extraction

`AGENT_INSTRUCTIONS.md` has the rules. This file is the judgement call: given
`analyze.mjs` output, what should become a command?

## Read the prose first

```bash
node agent/dist/analyze.mjs --text-only examples/*.docx
```

Two documents side by side, as text, tell you more in ten seconds than the JSON
does in ten minutes. Only go to the structured output once you know _what_ you
are looking for and need to know _where_ it is.

## A worked comparison

Two signed contracts:

```
--- contract-acme.docx                    --- contract-globex.docx
AGREEMENT                                 AGREEMENT
between Acme Corporation, Berlin          between Globex Inc, Prague
and the Supplier, dated 15 January 2025   and the Supplier, dated 3 March 2025

1. Scope                                  1. Scope
The Supplier shall deliver the goods      The Supplier shall deliver the goods
listed in Annex A.                        listed in Annex A.

2. Payment                                2. Payment
EUR 12,500 within 30 days.                EUR 8,000 within 14 days.

3. Delivery to Berlin                     3. Delivery to Prague
                                          4. Delivery to Brno
                                          5. Penalties
                                          Late delivery incurs 0.5% per day.

Signed for Acme Corporation               Signed for Globex Inc
```

What that yields:

| Reading                                                                 | Becomes                                                       |
| ----------------------------------------------------------------------- | ------------------------------------------------------------- |
| `Acme Corporation` differs, and appears twice                           | one variable, `client.name`, replaced in both places          |
| `Berlin` / `Prague` differ, and also appear inside the delivery clauses | careful — see below                                           |
| `15 January 2025` differs                                               | `formatDate(contract.date)`                                   |
| `EUR 12,500 within 30 days` differs in two independent ways             | two variables, not one                                        |
| "The Supplier shall deliver…" is identical                              | boilerplate, leave it alone                                   |
| one delivery clause versus two                                          | a `sectionLoop` over `deliveries`                             |
| the penalties section exists in only one                                | a `conditional`                                               |
| clause numbers run 1, 2, 3 versus 1, 2, 3, 4, 5                         | computed numbering — `$idx + 3`, then `deliveries.length + 3` |

## Heuristics

**A value that differs is a variable.** A value that is identical in two
documents may still be a variable — two invoices for the same customer prove
nothing. When in doubt, ask the user rather than guessing.

**Beware of a string that appears in two roles.** `Berlin` above is both the
client's city and part of a repeated delivery clause. Mapping it globally would
put `client.city` inside the loop, where the correct expression is
`$delivery.city`. Handle the loop first — `generate` applies structural edits
before variables for exactly this reason — and let the loop's `fields` or the
block's own text carry the per-item value.

**Count repetitions across documents, not within one.** One document with three
identical-looking blocks might be three hardcoded sections; two documents with
three and five is proof of a loop.

**A number next to a heading is rarely a constant.** As soon as one section
comes from a loop, every number after it moves.

**Check the runs, not just the text.** In the structured output, a paragraph
whose runs carry different `bold`/`size` values is telling you which fragment is
the value and which is the label: `runs: [{text: "Total: "}, {text: "12,500",
bold: true}]` means the variable is `12,500`, not the whole line.

**Headers and footers count.** They are extracted separately under
`headersAndFooters`, and `variables` are applied to them too.

## Before you generate

Show the user a table: text, proposed expression, why. Loops and conditionals
get a sentence each. Getting a correction here costs one message; getting it
after generating costs a rebuild.
