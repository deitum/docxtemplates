---
'@deitum/docxtemplates': patch
---

Narrow the `ctx` a custom `runJs` sandbox receives.

`runJs` was handed the engine's entire internal context — twenty-seven fields of
walk state, output buffers and half-built resources, none of which a sandbox has
any business with, and all of which change as the engine does. It now gets four:

```ts
runJs: ({ sandbox, ctx }) => {
  ctx.options; // the report options, with defaults filled in
  ctx.vars; // loop variables, as `$name` in snippets
  ctx.loops; // the FOR/IF constructs currently open
  ctx.jsSandbox; // the sandbox as the previous snippet left it
  // ...
};
```

Those four keep the names and the shapes they had, and are still live references
to the engine's own state, so a sandbox reading any of them works unchanged. A
sandbox reaching further in will no longer type-check — in TypeScript at the call
site, in JavaScript not at all — which is the point: those fields were never
meant to be part of the contract.
