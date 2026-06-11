# PP-MD Release Notes

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
