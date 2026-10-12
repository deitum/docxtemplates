---
'@deitum/docxtemplates': minor
---

Allow an `IF` nested into another one on the same line or table row, where that
is safe.

Until now, any such `IF` was rejected:

```
Invalid IF command nested into another IF command on the same table row
```

That ruled out a common layout — a table cell whose content depends on one
condition, with a detail inside it that depends on another:

```
+++IF company.type === 'LLC'+++
+++company.name+++
Represented by: +++IF signer.byProxy+++proxy+++ELSE+++director+++END-IF+++
+++END-IF+++
```

The rule guarded against an infinite loop (upstream issue #340): to render the
branch it has selected, an `IF` goes back to where it starts, and if that
takes it over the outer `IF`, it opens the outer one a second time. Whether
that happens depends on where the inner block ends, and the engine now works
that out before rendering. An inner block that ends in its own paragraph, or in
the same cell when it starts in a paragraph of its own, renders; one that would
run into the outer `IF` is still rejected, with the same error as before.

Nothing that rendered before renders differently. An `IF` that is still
rejected is rejected at the same point, with the same message, so a template
with no safe nested `IF` produces exactly the errors it did before, in the same
order, with `failFast: false` too. A part that uses `*shorthands` keeps the old
rule, since a shorthand might stand for an `END-IF`.
