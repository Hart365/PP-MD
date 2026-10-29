# Security Policy

Use the latest PP-MD release. Local solution ZIPs are parsed on the user's device; there is no Dataverse connection or solution upload service. Optional update checks contact GitHub.

Do not post suspected vulnerabilities or sensitive ZIPs in public issues. Use [GitHub private vulnerability reporting](https://github.com/Hart365/PP-MD/security/advisories/new) when available, or contact the maintainer through [Hart of the Midlands](https://HartOfTheMidlands.co.uk). Include the version, impact, and a minimal synthetic reproduction. Do not include credentials, tokens, customer metadata, or personal data. No response-time guarantee is made.

Treat solution archives and imported Markdown as untrusted input. Rendered HTML is sanitised; exports remain under user control. Review generated documents before sharing, since source names, descriptions, defaults, and configuration can be sensitive. Never commit real solution archives or generated customer documentation.