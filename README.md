# PP-MD — Power Platform Solution Documentation

PP-MD is a Windows desktop application that turns Microsoft Power Platform solution ZIP exports into structured, portable Markdown documentation with Mermaid diagrams. No server, no cloud upload, and no licence key required.

Version 1.3.0 brings the local-file documentation and export capabilities of PP-MD-PPTB-Edition to the standalone desktop app. Dataverse-connected collection, environment appendices, and ToolBox host APIs are intentionally excluded. See the [parity guide](docs/local-file-parity.md), [Accessibility statement](ACCESSIBILITY.MD), [contributing guide](CONTRIBUTING.md), [Code of Conduct](CODE_OF_CONDUCT.md), and [security policy](SECURITY.md).

---

## Table of Contents

- [What It Does](#what-it-does)
- [Generated Documentation Sections](#generated-documentation-sections)
- [Getting Started](#getting-started)
- [Using the App](#using-the-app)
	- [1. Load Solution ZIPs](#1-load-solution-zips)
	- [2. Fill in Document Context](#2-fill-in-document-context)
	- [3. Generate and Review](#3-generate-and-review)
	- [4. Export](#4-export)
	- [5. Consolidated Report](#5-consolidated-report)
- [Document Configuration Presets](#document-configuration-presets)
- [ERD Mode](#erd-mode)
- [Tech Stack](#tech-stack)
- [Building from Source](#building-from-source)
- [Security Notes](#security-notes)
- [System Requirements](#system-requirements)

---

## What It Does

Documentation. Usually a task people try to avoid, often difficult to get consistent output, and not fun to include all of the details you need. Then add in updating ADO, Jira etc. with the latest notes and people run for the hills.

This is where PP-MD steps in to the mix. Load the app, configure or load your document information (project, sprint, client etc) and document options, and then select solution ZIP files and click "Generate Documentation" to get a full suite of Markdown documentation automatically and consistently.

PP-MD parses you solution zip files and generates documentation that covers:

- An auto-generated **Entity Relationship Diagram** (ERD) rendered as a Mermaid diagram with the ability to have diagrams in a companion document to the core solution document
- Every **Dataverse table** with its columns, data types, and relationships
- All **Power Automate Flows** and classic workflows, with step-by-step details
- **Model-driven and Canvas apps**, including the tables they reference
- **Security roles** and **field security profiles** with privilege depth matrices and option to only include custom tables and/or tables in the current solution
- **Connection references** and **environment variables**
- **Plugins** and SDK message processing steps
- **Reports**, **dashboards**, **web resources**, and **email templates**
- Optional document metadata header (client, project, contract, sprint, release date)
- Column filtering controls, including **Include/Exclude Default Columns** (system columns), while always retaining relationship key columns in relationship documentation and ERD relationship context
- Document Options toggles to include/exclude Flows & Automation, Apps, Security, Integration, Plugins, Reports & Dashboards, Web Resources, Desktop Flows & Dataflows, Custom APIs & Offline Profiles, and Copilot Studio Agents & AI Models

Multiple solutions can be loaded at once and reviewed across all loaded solutions.

---

## Generated Documentation Sections

| Section | Contents |
|---|---|
| Header | Solution name, version, publisher, generation timestamp, document context |
| Table of Contents | Auto-generated links to every section |
| Entity Relationship Diagram | Mermaid ERD of all Dataverse tables and their relationships |
| Entities | Per-table: columns (name, type, required), relationships, forms, views |
| Option Sets / Choice Columns | All global and local option sets with label/value pairs |
| Forms & Views | Form names and field lists; view names and columns |
| Processes | Power Automate flows, workflows, BPFs, business rules, actions — with trigger, status, and connector details |
| Apps | Model-driven and canvas apps with referenced tables |
| Web Resources | JS, HTML, CSS, and other web resources with type and description |
| Security | Role privilege matrices and field security profiles |
| Integration | Connection references (connector type) and environment variables |
| Reports & Dashboards | Native reports and dashboards |
| Plugins | Plugin assemblies and SDK message processing steps (stage, mode, message, entity) |
| Diagnostics | Separate documents containing parser warnings and incomplete metadata issues |
| Missing Dependencies | Optional installation prerequisites explicitly declared by each ZIP manifest |

---

## Getting Started

**Download the latest release** from the `release_latest_portable/` or `release_latest_installer/` folders in this repository, or build from source (see [Building from Source](#building-from-source)).

| Artifact | Description |
|---|---|
| `PP-MD-<version>-x64-portable.exe` | Self-contained single executable — no installation required |
| `PP-MD-<version>-x64-installer.exe` | Standard Windows installer with Start Menu shortcut |

Both require **Windows 10/11 x64**.

---

## Using the App

### 1. Load Solution ZIPs

Export your solution from Power Platform:

1. Open [make.powerapps.com](https://make.powerapps.com)
2. Go to **Solutions** → select your solution → **Export**
3. Choose **Unmanaged** or **Managed** (both are supported)
4. Save the `.zip` file to your machine

In PP-MD, either:
- **Drag and drop** one or more `.zip` files onto the drop zone, or
- Click **Browse** to open a file picker

Multiple ZIPs can be loaded in a single batch. Each file is parsed independently with its own progress indicator.

### 2. Fill in Document Context

Above the drop zone you can optionally fill in:

| Field | Purpose |
|---|---|
| Client | Client or customer name — appears in the document header |
| Project | Project name |
| Contract | Contract or engagement reference |
| SOW | Statement of Work reference number |
| Sprint | Sprint or iteration name |
| Release Date | Intended release or delivery date |

These fields are included in the generated document header. They can be left blank if not needed.

### 3. Generate and Review

After uploading, the documentation is generated automatically. Use the **sidebar** on the left to switch between loaded solutions. The **Markdown Viewer** renders the output with full Mermaid diagram support, syntax highlighting, and GFM tables.

You can also load previously generated documents directly using **Open .md** in the top header toolbar.

When a document contains Mermaid content, the viewer toolbar shows live diagram render progress as **Diagrams X/Y rendered**.

Use **🔍 Search** in the viewer toolbar to find text within the rendered document (type at least 3 characters) — matches are highlighted, with next/previous navigation and a match counter.

### 4. Export

Export actions and Options rows have full-row, high-contrast hover highlights and distinct keyboard-focus outlines in light, dark, and Windows forced-colours modes.

From version 1.3.1, adjacent toolbar controls use a shared height and corner radius. **Export All** and **Options** have dropdown arrows and open non-modal popovers above the document; press Escape or click outside to dismiss them. Choosing an export closes its menu automatically.

Use the single **⬇️ Export ▾** menu in the Markdown viewer toolbar to export the active document:

- **Export .md**: saves the raw Markdown source.
- **Export .html**: saves standalone sanitised HTML with embedded styles and rendered Mermaid diagrams.
- **Export .pdf**: saves a searchable, paginated text PDF with diagrams, lists, tables, links, and readable source fallbacks.
- **Export .xlsx**: saves a structured solution workbook with inventories, global choices, per-table column worksheets, table properties, navigation links, and optional table-scoped missing dependencies. Imported Markdown and report documents retain the Markdown-table Excel fallback.

The new searchable PDF path does not claim PDF/UA conformance or complete tagging. The existing Electron tagged-PDF utility is retained for compatibility. See [Accessibility](ACCESSIBILITY.MD).

Use **Export All** for Markdown, HTML, PDF, or Excel ZIP archives, or **By Category** for self-contained category documents organised under solution and consolidated folders. Excel batch archives include diagnostics and dependency reports as Markdown companions. Duplicate filenames are disambiguated rather than overwritten.

### 5. Consolidated Report

**Separate diagrams is off by default.** Diagrams remain inline in individual and consolidated documents unless **Generate companion diagrams document** (or **Options > Separate diagrams document**) is checked. Document size does not override this setting. Separate documents appear together under **Diagrams** in the sidebar. Turning **Include diagrams** off removes inline diagrams and companion documents, even if separation remains checked.

When two or more solutions are loaded, the **Consolidated** sidebar section contains a complete **All Selected Solutions** document and the existing consolidated summary. The full document includes:

- A combined ERD across all tables
- Merged component sections de-duplicated by schema name
- A solution inventory table at the top

Clear **Create All Selected Solutions document** to generate individual documents only. **Options** can update output settings after generation. Optional missing-dependency reports retain each source solution's declared prerequisites; they do not infer dependencies from runtime references or test whether a target environment already satisfies them. Matching peer components can fill missing metadata without adding unrelated top-level components to an individual solution.

Choose a diagram colour theme or disable diagrams. Rendering runs lazily in the background, pauses while scrolling, prioritises visible diagrams, and retains at most 100 rendered SVGs. **Cancel** remains accessible outside the narrow-screen toolbar menu and stops at the next processing/export checkpoint. Completed solution documents are retained; unfinished exports are not downloaded.

Section scope, attribute selection, metadata-column visibility, and security-role filters apply to structured Excel exports as well as documentation. Summary-mode workbooks omit detailed per-table worksheets. Diagram theme choices are retained in exported Mermaid; changing a custom palette selects the base theme required for custom theme variables.

---

## Document Configuration Presets

Document context fields (client, project, etc.) can be saved as named presets and recalled from the dropdown in the toolbar.

**Built-in presets** are loaded from `public/doc-configurations.json`. You can pre-populate this file to share standard configurations across a team:

```json
{
	"configurations": [
		{
			"id": "project-home",
			"name": "Home Inventory",
			"client": "Hart of the Midlands",
			"project": "Home Inventory",
			"contract": "HART-001-HOME",
			"sow": "HART-001",
			"sprint": "Sprint 2",
			"releaseDate": "2026-06-30"
		}
	]
}
```

**Local presets** (saved from within the app) are stored in `localStorage` and persist between sessions on the same machine.

---

## ERD Mode

The ERD can be rendered in two modes, toggled in the toolbar:

| Mode | Description |
|---|---|
| **Detailed Relationships** | Shows all foreign-key and many-to-many relationships between tables |
| **Compact** | Shows tables only, without relationship lines — useful for solutions with very large schemas |

---

## Tech Stack

| Layer | Technology |
|---|---|
| Desktop shell | Electron 41 |
| Frontend | React 19 + TypeScript 6 + Vite 8 |
| ZIP parsing | JSZip |
| XML parsing | fast-xml-parser (XXE-safe configuration) |
| Markdown rendering | react-markdown + remark-gfm + rehype-raw |
| Diagram rendering | Mermaid 11 |
| Packaging | electron-builder (portable + NSIS installer) |

All processing runs **entirely locally** — no data leaves your machine.

---

## Building from Source

**Prerequisites:** Node.js 20+, npm, Windows (for packaging)

```powershell
# Install dependencies
npm install

# Run in development mode (hot-reload)
npm run desktop:dev

# Production build only (no packaging)
npm run build

# Build and package both portable + installer into release_latest_* folders
npm run desktop:build:latest

# End-of-phase workflow: bump version, run full gate, and build release executables
npm run phase:complete

# Same but also mirror the unpacked runtime to release_latest/
npm run desktop:build:latest -- -MirrorUnpacked

# Skip the npm build step (re-package from existing dist/)
npm run desktop:build:latest:skip-build
```

Build artifacts are written to:

| Folder | Contents |
|---|---|
| `release_latest/` | Unpacked Electron runtime (PP-MD.exe + supporting files) |
| `release_latest_portable/` | Self-contained portable `.exe` |
| `release_latest_installer/` | NSIS installer `.exe` + `.blockmap` |

Large binary artifacts in these folders are tracked by **Git LFS** (`.exe`, `.dll`, `.asar`, `.pak`, `.dat`, `.bin`, `.blockmap`).

---

## Security Notes

- **No network access** — the app performs no outbound requests during solution processing.
- Solution ZIPs are parsed entirely in the Electron renderer process using a memory-safe XML parser configured to prevent entity-expansion (XXE) attacks.
- Electron is configured with `contextIsolation: true` and `sandbox: true`.
- External hyperlinks open in the system default browser, not inside the app.
- A strict Content Security Policy is applied via `index.html`.
- The contrast of all UI theme tokens is verified at build time (WCAG 2.2 AA minimum).

---

## System Requirements

| Requirement | Minimum |
|---|---|
| OS | Windows 10 or Windows 11 (x64) |
| RAM | 4 GB (8 GB recommended for very large solutions) |
| Disk | 500 MB free for the app and runtime |
| Display | 1280 × 768 or higher |

---

*Built by Mike Hartley — Hart of the Midlands*
