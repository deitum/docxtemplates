# Building a mapping

`AGENT_INSTRUCTIONS.md` has the schema. This file is the worked example, using
the two contracts from `analyze.md`.

## The extraction

```
0  AGREEMENT
1  between Acme Corporation, Berlin
2  and the Supplier, dated 15 January 2025
3  1. Scope
4  The Supplier shall deliver the goods listed in Annex A.
5  2. Payment
6  EUR 12,500 within 30 days.
7  3. Delivery to Berlin
8  Signed for Acme Corporation
```

…plus a table of line items at `tableIndex: 0`, header row and three data rows.

## The mapping

```json
{
  "tableLoops": [
    {
      "var": "line",
      "over": "order.lines",
      "tableIndex": 0,
      "startRow": 1,
      "fields": {
        "0": "$line.description",
        "1": "$line.quantity",
        "2": "$line.price.toFixed(2)"
      }
    }
  ],
  "sectionLoops": [
    {
      "var": "delivery",
      "over": "deliveries",
      "startText": "3. Delivery to Berlin"
    }
  ],
  "conditionals": [
    {
      "expr": "contract.penalties != null",
      "paragraphText": "5. Penalties",
      "endParagraphText": "per day"
    }
  ],
  "variables": {
    "Acme Corporation": "client.name",
    "15 January 2025": "formatDate(contract.date)",
    "12,500": "contract.amount.toLocaleString('de-DE')",
    "30": "contract.paymentDays"
  }
}
```

## Why it is shaped that way

**The loop comes before the variables.** `generate` applies table loops, section
loops and conditionals first, then variables — all the structural edits are
located by the document's original prose, and once a phrase has become
`+++INS …+++` nothing will match it again. So `startText` is
`"3. Delivery to Berlin"`, the text as it stands in the filled-out document, and
`Berlin` is deliberately **not** in `variables`: it lives inside the loop, where
the delivery clause's own text becomes `+++= $delivery.city+++` in a follow-up
`refine` pass.

**`"30"` is dangerous and `"12,500"` is not.** A two-character key matches
anywhere — inside `12,500`, inside a date, inside `£30.00`. Before mapping a
short string, search the extracted text for it. If it occurs more than once, map
a longer phrase (`"within 30 days"` → `` `within ${contract.paymentDays} days` ``)
or make the edit with `refine`'s `addCommand` after the surrounding text has
already been templated.

**Expressions do the formatting.** `price.toFixed(2)` and
`toLocaleString('de-DE')` keep the document's own number format without the
caller having to pre-format anything. That is the advantage of expressions over
tag names — use it, but keep it readable: heavy logic belongs in the data, not
in the template.

**Section numbers are computed.** After generating, the numbers around the loop
need fixing up, because clause 4 onwards now depends on how many deliveries
there are. Inside the loop, `+++= $idx + 3+++`; after it,
`+++= deliveries.length + 3+++`.

## Then

```bash
node agent/dist/generate.mjs examples/contract-acme.docx field-mapping.json templates/contract.docx
```

`generate` prints one line per edit and then lists every command it produced.
Read that list: it is the cheapest place to notice that a variable landed
somewhere unintended.

Write the sample data to match — realistic values, and an array long enough to
prove the loop repeats:

```json
{
  "client": { "name": "Umbrella Ltd" },
  "contract": {
    "date": "2025-01-15",
    "amount": 12500,
    "paymentDays": 30,
    "penalties": null
  },
  "deliveries": [{ "city": "Berlin" }, { "city": "Prague" }],
  "order": {
    "lines": [
      { "description": "Widget", "quantity": 10, "price": 99.5 },
      { "description": "Gadget", "quantity": 2, "price": 1250 }
    ]
  }
}
```

Then render it and read the output:

```bash
node agent/dist/verify.mjs templates/contract.docx templates/contract_sample_data.json -o /tmp/report.docx
```

Two deliveries in the sample data, two delivery clauses in the output, numbered
3 and 4, with "5. Penalties" absent because `penalties` is `null`. If any of
that is wrong, the template is wrong — fix it with `refine` and render again.
