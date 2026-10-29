# Contributing

Bug reports, accessibility experiences, documentation improvements, and code contributions are welcome. Read the [Code of Conduct](CODE_OF_CONDUCT.md).

Use the [issue chooser](https://github.com/Hart365/PP-MD/issues/new/choose). For security issues use the [security policy](SECURITY.md), not public issues. Never upload customer solution files, sensitive generated documents, credentials, or personal data.

## Development

Use Node.js 22 or later, npm, and Windows for desktop release builds.

```powershell
npm ci
npm run dev
npm run quality:gate
npm run desktop:build:latest
```

Focused checks include `npm run lint`, `npm run test:run`, `npm run smoke:doc`, `npm run check:contrast`, `npm run build`, and `npm run test:a11y`. Use synthetic ZIP fixtures. Validate portable and installed Electron behaviour when changing desktop integration. PP-MD has no Dataverse connection or ToolBox dependency.

## Accessibility

Target WCAG 2.2 AA. Check keyboard operation, logical focus, visible focus, menu dismissal, labels, status announcements, and screen-reader output. Check light/dark themes, Windows forced colours, reduced motion, 200% text resizing, and 320 CSS-pixel reflow or 400% zoom. Tables and diagrams may scroll in two dimensions; surrounding controls must remain usable.

For exports check headings, reading order, table structure, links, diagram alternatives, and PDF tagging separately. A searchable or visually tidy PDF is not automatically an accessible PDF. Report exact checks and untested combinations in your PR; do not infer manual conformance from automated tests.

## Pull Requests

Keep changes focused. Explain the problem and result, list checks and limitations, and update the README, release notes, and relevant design documentation. Do not commit generated output or unrelated changes.