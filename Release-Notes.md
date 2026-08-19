# PP-MD Release Notes

---

## Version 1.2.2

This release fixes the documentation viewer toolbar anchoring, tightens search behavior and contrast, and restores the missing app/window icon in packaged builds.

### Highlights

- **Toolbar anchoring fix**: Fixed the root cause of the toolbar (and document title) scrolling off screen — the viewer panel had no bounded height, so its internal scroll region never took effect. The toolbar now stays fixed in place while only the document content scrolls beneath it.
- **Back to top**: Added a **⬆️ Top** button that appears once you've scrolled down, returning you to the top of the document.
- **Search refinement**: Search now requires at least 3 characters before it starts matching, showing a "Type 3+ characters" hint for shorter queries.
- **Search contrast fix**: Replaced the translucent orange/black active-match highlight (which failed WCAG 2.2 AA contrast) with opaque, verified-contrast colors (≥ 4.5:1) for both the general match and current-match highlights.
- **App/window icon fix**: `build/icon.ico` is now bundled into the packaged app so the Electron window icon resolves correctly at runtime, in addition to the executable icon.

---

## Version 1.2.1

This release refines the documentation viewer toolbar: export actions are now merged into a single dropdown, and an in-document search capability has been added.

### Highlights

- **Toolbar refinement**: Consolidated the three separate Export .md/.xlsx/.pdf buttons into a single **Export ▾** dropdown menu, reducing toolbar clutter.
- **Toolbar visibility fix**: The export dropdown is positioned relative to the viewport instead of the toolbar, so it is never clipped by the Markdown viewer's `overflow: hidden` container.
- **Search capability**: Added a **🔍 Search** toggle to the viewer toolbar that finds and highlights matching text in the rendered document, with next/previous navigation and a match counter.

---

## Version 1.2.0

This release expands Document Options with new toggles for previously always-on sections, removes the Manual Attributes selection mode, and updates all dependencies to their latest secure versions.

### Highlights

- **Document Options**: Added dedicated toggles for Web Resources, Desktop Flows & Dataflows, Custom APIs & Offline Profiles, and Copilot Studio Agents & AI Models, so every documented item type can be explicitly included or excluded.
- **Table Options**: Removed the "Manual Attributes" text field and the "Manually Selected" attribute selection mode, simplifying the attribute filtering options.
- **Maintenance**: Updated all dependencies to their latest compatible versions and removed the unused, vulnerable `xlsx` dependency, resolving all `npm audit` findings (0 vulnerabilities).
- **CI**: Hardened the CI workflow's `GITHUB_TOKEN` permissions to read-only, resolving a CodeQL code scanning alert.
- **Security**: Replaced regex-based HTML sanitization in the PDF export path with DOMPurify, and fixed unsafe HTML-entity decoding order in the PDF and Excel export utilities, resolving five CodeQL alerts (bad HTML tag filtering, incomplete sanitization, and double-escaping).

---

## Version 1.1.24

This release fixes accessibility issues with the new collapsible sections and adds missing column type mappings for Choice fields.

### Highlights

- **Accessibility**: Updated the hover contrast color on the new UI collapsible section headers to ensure they pass WCAG 2.2 AA contrast checks on both light and dark themes.
- **Reporting**: Fixed a bug where Dataverse `Choice` and `Choices` columns were mislabeled as "Unknown" by mapping them correctly to OptionSet classifications.

---

## Version 1.1.23

This release introduces UI improvements by adding collapsible sections and fine-grained column metadata configuration.

### Highlights

- Restructured the opening settings screen with discrete, collapsible containers for Document Header Details, Document Options, Table Options, Security Role Options, and Diagram Options.
- Added a new Column Metadata options section (under Table Options) to select exactly which attribute metadata fields should render in the table documentation.

---

## Version 1.1.22

This release adds configurable Mermaid styling for the generated diagrams, including a stronger, more diagrammatic palette and source-based line colors that make it easier to follow connections from individual items.

### Highlights

- Mermaid diagrams now support configurable colors from a collapsed section on the start page, with a stronger contrast palette by default.
- Relationship lines in all Mermaid diagrams use distinct source-based colors so it is easier to identify connections that radiate from a single node.
- The same styling approach is applied consistently across component graphs, ERDs, and process-flow diagrams.

---

## Version 1.1.21

This release consolidates the main improvements delivered since version 1.1.6, covering export reliability, reporting quality, and the desktop experience.

### Highlights

- PDF export handling for Mermaid diagrams is much more reliable, with better text sizing, multi-page diagram splitting, and improved preservation of labels and styling.
- Excel export output is cleaner and more consistent, with better table structure, header handling, and worksheet formatting.
- The desktop app now provides clearer export progress feedback, direct loading of previously generated Markdown files, and stronger accessibility and regression coverage.

### Quality

- Expanded automated regression and accessibility checks to keep PDF, Excel, and viewer workflows stable.

---

## Version 1.1.8

### Fixes

- Removed the **Status** column from Processes & Automation summaries and details because this value cannot be reliably reported from Dataverse-connected exports.
- Improved Power Automate flow display-name normalization for partially condensed names so labels like `Customerandtriageteamwhenanewprocure` render with readable spacing.

---

## Version 1.1.7

### Fixes

- Security Role permission columns now render with color-coded labels and matching circle indicators (⚫ 🔵 🟢 🟡 🟠 🔴) to make access depth easier to scan.
- Power Automate flow display names now consistently prefer JSON display metadata, preventing corrupted names with removed spacing.
- Processes & Automation summary has been split into automation-type summary tables with relevant columns per type.

### Improvements

- Markdown filenames now follow `<Solution Name>.md` and `<Solution Name> - Diagrams.md`.
- The Markdown Viewer title now exposes the active filename as a hover tooltip to help identify truncated sidebar entries.

---

## Versions 1.1.3 to 1.1.5

### Fixes

- Fixed a regression where the Security Roles section could disappear when active security-role filters removed all matching table privileges.
- Added an explicit placeholder row in empty filtered matrices: _No table privileges matched the active filters for this role._
- Added support for matching role privilege table names against both Dataverse logical names and entity-set names to reduce false exclusions.
- Hardened custom-table detection for security role filters so custom tables are retained even when Dataverse exports do not reliably mark `IsCustomEntity`.
- Added fallback custom-table heuristics using object type code and Dataverse naming patterns for logical names and entity-set names.

### New Features

- Added a new document option: **Generate Companion Diagrams Document**.
- When enabled, Mermaid diagrams are moved into a companion markdown file while the main document keeps the same structure and includes a notice that diagrams are available in the companion file.
- Companion and primary outputs both preserve Table of Contents and Back-to-Top navigation links.

### Quality

- Added a regression test to ensure filtered-empty role matrices do not remove the Security Roles section.
- Updated regression tests for security role filtering semantics and added coverage for logical-name/entity-set-name matching.
- Added regression tests for fallback custom-table filtering when `IsCustomEntity` metadata is unavailable.
- Added tests validating diagram extraction into a companion markdown output.

---

## Version 1.1.2

### New Features

#### Security Role Filters

- **Only include tables in current solution** — limits the security role privilege matrix to tables that exist in the loaded solution.
- **Only include custom tables** — further narrows the privilege matrix to custom tables only, excluding out-of-box tables even when they appear in the role.

---

## Version 1.1.0

### New Features

#### New Solution Component Sections

The generated documentation now covers a significantly broader set of solution components. The following sections are new in v1.1.0:

- **Copilot Studio Agents** — lists agents included in the solution, with agent type, language, trigger/channel metadata, and referenced connectors where discoverable.
- **AI Models** — documents AI model artefacts, including model type, provider, version, and runtime endpoint/deployment reference.
- **Desktop Flows** — lists desktop (RPA) flows with folder grouping, enabled/disabled status, estimated step count, and referenced connectors.
- **Dataflows** — documents Dataflow artefacts with connector/data-source hints and refresh mode where available.
- **Custom APIs** — lists custom API definitions, bound table, and whether each API is a function-style endpoint.
- **Offline (Mobile) Profiles** — documents Mobile Offline profiles included in the solution.
- **Solution Dependencies** — a new dedicated section lists all declared solution dependencies (required and missing), including display name, schema name, and version.
- **Solution Component Inventory** — a high-level inventory table categorises every component in the solution by type (tables, flows, apps, agents, plugins, reports, etc.) with component counts.
- **Solution Component Relationship Graph** — a Mermaid graph visualises connections and dependencies between major solution components.

#### Enhanced App Documentation

- **Canvas App insights** — canvas and custom-page apps now include screen count, control count, data source count, variable count, and resource count. Detailed view lists screen names, data sources, variables, and per-screen control lists.
- **Model-Driven App site map** — model-driven apps now include their full site map structure (areas, groups, sub-areas) and site map settings (show Home, show Pinned, show Recents, collapsible groups).

#### Enhanced Table & Column Documentation

Column (attribute) tables now support several additional optional columns, all independently toggleable:

| Option | Description |
|---|---|
| **Required Level** | Displays the field's required level (None, Recommended, ApplicationRequired). |
| **Field Security** | Shows whether column-level security is enabled for each field. |
| **Advanced Find** | Indicates whether the field is visible in Advanced Find. |
| **Metadata Diagnostic Info** | Shows the source metadata key used to derive each flag — useful when troubleshooting parsed output. |

Additional column-level data is captured and rendered where available:

- **Polymorphic lookup targets** — multi-target lookups list all possible target tables.
- **Min/max values** — numeric and date-like fields show configured minimum and maximum values.
- **Format hint** — fields with a format hint (e.g. `Email`, `Url`, `DateOnly`, `Duration`) display it.
- **Default value** — fields with a configured default value show it in the table.
- **Form placement** — for attributes found on forms, the tab name and section name where the field appears is recorded.

#### Attribute Selection Modes

A new **Attribute Selection Mode** drop-down gives precise control over which columns appear in the Tables & Columns section:

| Mode | Behaviour |
|---|---|
| **All** | Every attribute in the solution metadata. |
| **Custom Only** | Only attributes with `IsCustomAttribute = true`. |
| **Attributes On Form** | Only attributes that appear on at least one form in the solution. |
| **Attributes Not On Form** | Attributes present in metadata but absent from all forms. |
| **Option-Set Focused** | Only attributes with choice (option set) or boolean data types. |
| **Manually Selected** | A user-defined comma-separated list of logical attribute names. |
| **Unmanaged Only** | Only unmanaged (customisable) attributes. |

#### Summary Documentation Mode

A new **Documentation Detail Level** control lets you choose between:

- **Detailed** — the full report including all tables, diagrams, and component sections (existing behaviour).
- **Summary** — a condensed report containing the overview, solution metadata, component inventory, and scoped section summaries without per-row detail tables. Useful for quick executive summaries or large solutions.

#### Documentation Scope Controls

A new **Scope** panel lets you independently include or exclude whole categories of content from the generated document:

- Flows
- Apps
- Security (security roles, field security profiles, access teams)
- Integration (connection references, environment variables, email templates, service endpoints)
- Plugins & SDK steps
- Reports

#### Table of Contents

All generated documents now open with an auto-generated **Table of Contents** that lists every section with the count of items it covers, giving an instant summary of what is in the solution before reading the detail.

#### Update Checker Improvements

- The update check dialog is now fully keyboard-accessible: focus moves to the close button when it opens, and pressing `Escape` dismisses it.
- The dialog uses proper ARIA roles (`role="dialog"`, `aria-modal`, `aria-labelledby`, `aria-describedby`) for screen-reader compatibility.
- The Electron main process now performs its own GitHub release check using a built-in HTTPS fetch (no external runtime dependency), with redirect support and a versioned `User-Agent` header.

#### Developer / Quality

- A full **unit test suite** (Vitest) has been added, covering the solution parser, markdown generator, output model, metadata grid, settings controls, and version utilities.
- **Playwright accessibility tests** (`@axe-core/playwright`) are included for the home page.
- A CI workflow runs lint, unit tests, smoke tests, production build, and accessibility tests on every push.
- A **contrast audit** script (`check-contrast.mjs`) validates all theme token pairs against WCAG AA/non-text minimums and is part of every build.

---

## Version 1.0.3

_Adds update checking, multi-platform builds, and documentation updates._

- Added in-app update checker that queries GitHub Releases and shows the latest version, release date, and download size.
- Added platform and architecture detection so update links point to the correct download artefact (Windows x64/arm64, macOS, Linux).
- GitHub Actions release workflow extended with macOS and Linux builds.
- Multi-arch portable builds added for Windows arm64.
- Explicit architecture suffixes added to Linux and macOS artefact filenames.
- Package version bumped to 1.0.3.

---

## Version 1.0.1 – 1.0.2

_Build and release pipeline stabilisation._

- Fixed release workflow to disable electron-builder auto-publish.
- Release output folders removed from Git tracking and added to `.gitignore`.
- GitHub Actions release workflow added to automate Windows executable builds.

---

## Version 1.0.0

_Initial public release._

- Power Platform solution documentation generator for Windows desktop (Electron + React).
- Parses `.zip` solution packages and generates structured Markdown documentation.
- Supports: tables & columns, model-driven and canvas apps, cloud flows, security roles, field security profiles, connection references, environment variables, plugins & SDK steps, reports, web resources, email templates, and access teams.
- Light/dark theme toggle.
- Entity Relationship Diagram (ERD) generation via Mermaid.
- Drag-and-drop and file-picker drop zone for solution files.
- Multi-solution processing with merged consolidated output.
- Sidebar navigation between multiple loaded solutions.
- Markdown preview with copy-to-clipboard.
