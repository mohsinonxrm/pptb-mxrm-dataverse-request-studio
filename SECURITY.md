# Security Policy

## Reporting a Vulnerability

Please do **not** open a public issue for security concerns. Use GitHub
Security Advisories (private):
<https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/security/advisories/new>

I'll acknowledge receipt within a few business days and coordinate a fix +
disclosure timeline with you.

## Scope

The studio is a client-side composer + executor for Dataverse Web API
requests, running either embedded in Power Platform ToolBox or standalone
in a browser. Vulnerability reports of interest:

- **Authentication / token handling** — any path that could leak a
  Dataverse access token, the PPTB session token, or a Connection's
  credentials from the host bridge.
- **Cross-site / injection** — XSS, HTML / Markdown injection in
  user-supplied request fields, code-tab snippet output, or results-grid
  cell rendering.
- **CSP / sandbox escape** — anything that lets DRS code execute outside
  the iframe / host's allowed origins, or fetch from a non-Dataverse host.
- **Dataverse request integrity** — paths where the studio could emit a
  request that escalates privileges, bypasses documented `MSCRM.*` header
  rules, or alters another tool's state.
- **Saved-request library** — paths where a crafted saved-request payload
  triggers code execution on load, exfiltrates localStorage from another
  org scope, or escapes the per-org scoping bucket.
- **Dependency vulnerabilities** — known CVEs in production dependencies
  (`@fluentui/*`, `monaco-editor`, `@pptb/types`, etc.) that affect the
  shipped bundle.

## Out of scope

- **Mistakes the user makes against their own environment** — DRS lets you
  Delete, Update, and Merge records. Destructive operations require typed
  confirmation; beyond that, the user is responsible for what they send.
- **Standalone-mode network errors** — the standalone build (`npm run dev`)
  has no host bridge; Execute returns 501. This is by design, not a bug.
- **Antipattern advisories** — the studio surfaces guidance per the
  Microsoft query-antipatterns doc, but it's advisory only; the request
  still goes through if the user proceeds.

## Supported versions

Only the latest released minor version on `main` receives security fixes.

Thanks for helping keep the tool safe.
