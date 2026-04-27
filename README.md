# PP-MD — Power Platform Solution Documenter

PP-MD is a Windows desktop application that turns Microsoft Power Platform solution ZIP exports into structured, portable Markdown documentation — with Mermaid diagrams included. No server, no cloud upload, no licence key required.

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

When you export a solution from Power Platform / Dataverse, you receive a ZIP archive containing XML definitions of every component — tables, columns, relationships, flows, apps, security roles, plugins, and more. Reading that raw XML is slow and error-prone.

PP-MD parses that ZIP entirely on your local machine and produces a single, readable `.md` file that covers:

- An auto-generated **Entity Relationship Diagram** (ERD) rendered as a Mermaid diagram
- Every **Dataverse table** with its columns, data types, and relationships
- All **Power Automate Flows** and classic workflows, with step-by-step details
- **Model-driven and Canvas apps**, including the tables they reference
- **Security roles** and **field security profiles** with privilege depth matrices
- **Connection references** and **environment variables**
- **Plugins** and SDK message processing steps
- **Reports**, **dashboards**, **web resources**, and **email templates**
- Optional document metadata header (client, project, contract, sprint, release date)

Multiple solutions can be loaded at once and reviewed side-by-side, or merged into a single **consolidated report** across all loaded solutions.

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
| Reports & Dashboards | SSRS reports and dashboards |
| Plugins | Plugin assemblies and SDK message processing steps (stage, mode, message, entity) |
| Warnings | Parser warnings for components that could not be fully parsed |

---

## Getting Started

**Download the latest release** from the `release_latest_portable/` or `release_latest_installer/` folders in this repository, or build from source (see [Building from Source](#building-from-source)).

| Artifact | Description |
|---|---|
| `PP-MD-1.0.0-x64-portable.exe` | Self-contained single executable — no installation required |
| `PP-MD-1.0.0-x64-installer.exe` | Standard Windows installer with Start Menu shortcut |

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

### 4. Export

Click **Export .md** (or the copy button) to save the current document as a `.md` file. The file name is derived from the solution's unique name. The exported file can be opened in any Markdown editor (VS Code, Obsidian, Azure DevOps Wiki, GitHub, etc.).

### 5. Consolidated Report

When two or more solutions are loaded, a **Consolidated Summary** entry appears in the sidebar. This merges all solutions into a single report with:

- A combined ERD across all tables
- Merged component sections de-duplicated by schema name
- A solution inventory table at the top

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
| Desktop shell | Electron 32 |
| Frontend | React 18 + TypeScript + Vite 6 |
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
