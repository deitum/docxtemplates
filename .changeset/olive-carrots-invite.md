---
'@deitum/docxtemplates': minor
---

Reject options of the wrong type instead of ignoring them.

The motivating case:

```ts
errorHandler: typeof options.errorHandler === 'function' ? options.errorHandler : null,
```

Pass anything but a function and the handler simply never ran — no error, no
warning, and a report full of the failures it was meant to catch. `template` was
not checked at all, so passing the wrong thing surfaced as an opaque complaint
from JSZip about not finding the end of a central directory.

Every documented option is now checked against its type, and a mismatch throws
the new `InvalidOptionError`, which names the option, what it expected and what
it got:

```
Option 'errorHandler' must be a function, but received a string
```

**This can break code that is passing a bad value today** and getting the
default behaviour by accident — hence the minor. Setting an option to
`undefined` is still the same as leaving it out.

Unknown option names are _not_ rejected: they go to the debug log. TypeScript
already catches a typo for anyone who has types, and in plain JavaScript an
extra key on a shared options object is more often deliberate than not.
