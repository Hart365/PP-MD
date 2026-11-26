# Local-File Parity

The current 1.3.12 source includes the solution-file capabilities available through `pp-md-pptb-edition` commit `db47c99` (2026-11-23), adapted for PP-MD's standalone Electron file workflow. The documented baseline began in version 1.3.0, with later 1.3.x capabilities recorded here.

This port retains PP-MD's Electron packaging, update checking, configuration storage, Markdown imports/search, form layout, canvas insights, component inventory and diagram palette. It intentionally omits PPTB and connected Dataverse-specific code.

## Included

| Capability | PP-MD implementation |
| --- | --- |
| ZIP processing | Manifest prerequisites, table properties, AbortSignal checkpoints and two-job inline-worker concurrency with explicit compatibility fallback |
| Consolidation | Full All Selected Solutions document alongside the existing summary, with an opt-out |
| Peer gap fill | Matching component metadata only; source component membership and prerequisites remain solution-scoped |
| Diagnostics | Separate per-solution and consolidated warning/issue documents |
| Dependency reports | Optional document-wide reports, structured installation analysis and table-scoped Excel sections; shared classification with component IDs/schema names, parent evidence, packages, shell/full ownership, provider-first ordering, blocker propagation, readiness layers, cycle groups and downstream waiters |
| HTML | Standalone sanitised HTML with embedded styles and Mermaid SVGs |
| PDF | Searchable text, pagination, headings, lists, tables, links, rasterised diagrams, source fallback |
| Excel | Native Excel tables, frozen headers, styled inventories, global choices, per-table tabs, metadata blocks, navigation, sorting, wrapped descriptions, dependency identities/evidence, sectioned installation guides and overflow continuations |
| Batch export | Markdown/HTML/PDF/Excel ZIPs; installation guides remain workbooks while diagnostics and dependency reports accompany Excel as Markdown |
| Categories | Self-contained category documents with contents links and solution/consolidated folders |
| Viewer | Sanitised GFM, raw view, search, copy, heading navigation, diagram progress and export menu |
| Diagrams | Low-priority visible-first background scheduler paused during scrolling; 100-entry LRU cache; selectable themes; classic Dagre configuration; validated SVG dimensions and solid background |
| UI | Versioned branding, grouped sidebar, responsive toolbar disclosure, output options, visible cancellation and GitHub link |
| Community | Accessibility/bug/feature issue forms, issue chooser, PR template, contribution/accessibility/security policies and Code of Conduct |

## Excluded

No Dataverse connection browser, authentication, connected collectors, collection policies, environment appendix, tenant/user counts, or Power Platform ToolBox bridge is included. ZIP content stays on the device. Source-derived configuration and defaults can still contain sensitive information; review documentation before sharing.

## Verification

### Option Output Contracts

| Controls | Expected output |
| --- | --- |
| Include diagrams / separate diagrams | Off means no diagrams; included diagrams stay inline unless separation is checked. No document-size or consolidated-output override. |
| Create consolidated document | Produces the combined summary and full document only for two or more solutions when enabled. |
| Missing dependencies | Adds declared per-solution reports and table-scoped workbook dependency sections only when enabled. Table dependency rows expose dependent/required IDs, owners and evidence. |
| Section scope | All ten category toggles govern detail, relationship diagrams, consolidated summaries, and workbook inventories. |
| Column-detail checkboxes | Type, Required, Custom, Audited, Field Security, Advanced Find, Metadata Source, Notes, and Description control their table/workbook columns. |
| Default / virtual columns and selection mode | Default and virtual filters combine with all, custom-only, on-form, not-on-form, option-set-focused, or unmanaged-only selection. Relationship keys remain in relationship documentation. |
| Security filters | Limit role privileges to solution tables and/or custom tables in documentation and workbooks. |
| Summary / detailed | Summary documentation omits detailed-only sections; summary workbooks omit per-table column worksheets. |
| ERD detail and diagram theme/palette | Compact/detailed modes affect diagrams, selected themes are exported, and all four palette colours are used. |
| Presets and live changes | Settings persist, regenerate output consistently, and retain the active document where possible. Imported Markdown is not regenerated or filtered. |

Companions are grouped in a dedicated Diagrams sidebar region; diagnostics and dependency reports have their own region. Installation reports may use the active Excel handler. Browser regressions inspect actual ZIP contents while changing settings and importing Markdown.

Run `npm run quality:gate` and `npm run desktop:build:latest`. Focused unit tests cover parser cancellation, authoritative prerequisite scoping, all-subcomponent/shell matching, cycle membership, peer enrichment, code-fence-safe category splitting, export sanitisation, archive collisions and cancellation, native Excel table packaging, installation-guide worksheets, overflow continuations, PDF content and LRU eviction. Browser tests cover home/settings and generated-document flows. Automated accessibility results are not a complete conformance claim; screen-reader and packaged Electron checks remain separate manual responsibilities.

Browser regression coverage also downloads and inspects HTML, PDF (including an actual embedded diagram image), Excel, category ZIPs, and batch HTML/PDF/Excel archives. It checks generated-viewer contrast, keyboard-accessible scroll regions, 320px toolbar reflow, real dark-theme switching, and forced-colours captures. SVG rasterisation uses embedded data URLs, a timeout and exception-safe fallbacks so canvas errors cannot leave an export pending forever.