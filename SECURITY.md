# Security policy

## Supported versions

Security fixes are released for the latest published version only.

## Reporting a vulnerability

Please report vulnerabilities privately through
[GitHub's security advisory form](https://github.com/deitum/docxtemplates/security/advisories/new)
rather than opening a public issue. We aim to acknowledge reports within a few
working days.

## Scope: the JavaScript sandbox

This library evaluates the JavaScript snippets found in a `.docx` template. By
default those snippets run through Node's built-in [`vm`](https://nodejs.org/api/vm.html)
module, which is **explicitly not a security boundary** — the Node documentation
says so directly. Setting `noSandbox: true` removes even that, and evaluates
snippets in the current context.

Consequently:

- **Treat templates as code.** Rendering a template from an untrusted source is
  equivalent to executing arbitrary code with the privileges of your process.
- If you must handle untrusted templates, supply your own isolation through the
  `runJs` option (a separate process, a WASM/V8 isolate, a container) and apply
  your own resource limits.

Reports that amount to "`vm` can be escaped" describe documented behaviour of the
underlying platform rather than a vulnerability in this library, but reports of
sandbox escapes that bypass a _user-provided_ `runJs` boundary, or any other
memory-safety or injection issue in the library itself, are very welcome.
