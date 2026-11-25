# PP-MD File Catalog and Naming Review

## Current Capability Additions

Portable export modules live under `src/api/`; safe single/batch archive handling is in `src/utils/documentExport.ts`. Local reports, installation analysis, category splitting, table properties, and declared dependency helpers live under `src/generator/`. Background ZIP processing lives under `src/workers/`. `MermaidRenderer.tsx` owns background rendering and uses `utils/lruCache.ts`. Focused tests cover these paths alongside the existing parser, viewer, generator and export suites.

Community files: `.github/ISSUE_TEMPLATE/*.yml`, `.github/PULL_REQUEST_TEMPLATE.md`, `CONTRIBUTING.md`, `CODE_OF_CONDUCT.md`, `SECURITY.md`, and `ACCESSIBILITY.MD`. See `docs/local-file-parity.md` for the local-only capability matrix.

## Purpose
This document lists the application code and resource files, explains their responsibilities, and reviews naming clarity.

## Scope
Included:
- Source code and resources under src/, electron/, public/, scripts/
- Core project configuration files used by the app build/runtime

Excluded:
- node_modules/ (third-party dependencies)
- dist/ and release*/ (generated build artifacts)
- temporary packaging folders

## Naming Clarity Review
Current naming is generally clear and consistent for source and script files.

Findings:
- Component files use PascalCase and match exported component names.
- CSS module files mirror component names (for example: ThemeToggle.tsx and ThemeToggle.module.css).
- Functional utility modules are named by role (solutionParser.ts, markdownGenerator.ts).
- Build/packaging scripts are action-oriented and descriptive (build-latest.ps1, stamp-exe.mjs).

Notes:
- Generated files in dist/ include hash suffixes by design and should not be renamed.
- Release output folder names include timestamps by design for traceability.

## Naming Convention Baseline
- React components: PascalCase.tsx
- Context/providers: FeatureContext.tsx
- Processing modules: verbNoun.ts
- Script files: kebab-case with clear action intent
- Style files: <Component>.module.css for component-scoped styles
- Static resources: lower-case, descriptive names

## File Inventory

### Root Configuration and Entry Files
| File | Function |
|---|---|
| index.html | Browser shell, CSP definition, root mount element. |
| package.json | Dependency manifest, npm scripts, electron-builder config. |
| tsconfig.json | TypeScript compiler rules for strict frontend typing. |
| vite.config.ts | Vite build config for Electron-friendly relative assets. |

### Electron Runtime
| File | Function |
|---|---|
| electron/main.cjs | Electron main process bootstrap, BrowserWindow setup, external link handling, load/crash error dialogs. |

### Public Static Resources
| File | Function |
|---|---|
| public/doc-configurations.json | Default document context presets loaded by the UI. |
| public/favicon.svg | App favicon for browser/Electron renderer window. |
| public/icons.svg | Shared SVG icon sprite/static icon resource. |

### Build and Utility Scripts
| File | Function |
|---|---|
| scripts/build-latest.ps1 | Full build and packaging orchestrator; mirrors latest outputs into release_latest folders. |
| scripts/check-contrast.mjs | WCAG contrast token audit for light/dark theme variables in global.css. |
| scripts/sample-doc-from-zip.ts | CLI tool to parse a solution ZIP and emit sample markdown with relationship stats. |
| scripts/smoke-doc.ts | Synthetic end-to-end smoke test for parser and markdown generator behavior. |
| scripts/stamp-exe.mjs | Stamps Windows executable metadata/version strings on packaged binaries. |
| scripts/phase-complete.ps1 | Runs the release close-out workflow, quality checks and packaging. |

### React Application Core
| File | Function |
|---|---|
| src/main.tsx | React entry point; mounts App inside ThemeProvider and StrictMode. |
| src/App.tsx | Main UI and orchestration: file intake, worker parsing, status/progress, result selection and export. |
| src/App.module.css | Root app layout and visual styling for App.tsx. |
| src/vite-env.d.ts | Vite TypeScript ambient declarations. |

### Styling and Assets
| File | Function |
|---|---|
| src/assets/global.css | Global theme tokens, base styles, WCAG-aware color system. |
| src/assets/hero.png | Hero/branding image asset. |
| src/assets/typescript.svg | TypeScript branding icon resource. |
| src/assets/vite.svg | Vite branding icon resource. |

### UI Components
| File | Function |
|---|---|
| src/components/SolutionSidebar.tsx | Left navigation for parsed solutions and component counts. |
| src/components/SolutionSidebar.module.css | Styles for solution sidebar structure and states. |
| src/components/ui/DropZone.tsx | Accessible drag-and-drop/browse control for ZIP intake and queue management. |
| src/components/ui/DropZone.module.css | Styles for drop zone, file queue list, and state transitions. |
| src/components/ui/MarkdownViewer.tsx | Markdown renderer with Mermaid interception, copy/export, and raw toggle. |
| src/components/ui/MarkdownViewer.module.css | Styles for markdown viewer, toolbar, and rendered content. |
| src/components/ui/MermaidRenderer.tsx | Scheduled Mermaid generation, SVG/background validation, rendering cache and batch-export SVG support. |
| src/components/ui/MermaidDiagram.tsx | Interactive Mermaid wrapper with zoom, fit, fullscreen and source fallback. |
| src/components/ui/ProgressBar.tsx | Accessible progress bar for processing status. |
| src/components/ui/ProgressBar.module.css | Progress bar styles and motion behavior. |
| src/components/ui/ThemeToggle.tsx | Theme switch UI connected to ThemeContext. |
| src/components/ui/ThemeToggle.module.css | Theme toggle styles and interaction states. |

### Domain Logic and Models
| File | Function |
|---|---|
| src/context/ThemeContext.tsx | Theme state, persistence, OS preference sync, document-level theme application. |
| src/parser/solutionParser.ts | Core parser that reads solution ZIP/XML and maps content to typed ParsedSolution. |
| src/generator/markdownGenerator.ts | Markdown and Mermaid generation for single and consolidated solution outputs. |
| src/generator/dependencyAnalysis.ts | Classifies manifest prerequisites using ZIP component identities and selected providers. |
| src/generator/localDocumentation.ts | Produces dependency, diagnostics and provider-first installation guidance. |
| src/generator/missingDependencies.ts | Normalises component labels and scopes dependency records. |
| src/generator/installationAnalysis.ts | Calculates the reusable selected-ZIP installation plan, blockers, evidence and readiness. |
| src/generator/installationGraph.ts | Detects strongly connected cycles and computes graph IDs/readiness layers. |
| src/generator/tableMetadata.ts | Shared table metadata blocks for Markdown and Excel. |
| src/types/solution.ts | Domain enums/interfaces for all parsed and generated solution structures. |
| src/api/excelExporter.ts | Builds solution and installation-guide workbooks. |
| src/api/excelPackage.ts | Adds native table parts, frozen panes and repaired styles to XLSX package XML. |
| src/api/documentRendering.tsx | Renders sanitised standalone HTML and export-ready Mermaid SVGs. |
| src/api/pdfExporter.ts | Draws searchable paginated PDF output. |
| src/utils/documentExport.ts | Handles safe single/batch downloads, collisions, archives, progress and cancellation. |
| src/utils/lruCache.ts | Bounded least-recently-used cache for rendered diagrams. |
| src/workers/parseSolutions.ts | Coordinates two-job concurrent solution parsing and compatibility fallback. |
| src/workers/solution.worker.ts | Inline worker entry point for ZIP parsing. |

### Reserved Folders
| Folder | Function |
|---|---|
| src/hooks/ | Reserved for future custom React hooks. |
| src/utils/ | Shared export, cache and document helpers. |

## Core and Key Functions

### Core Runtime Functions
| Module | Function | Responsibility |
|---|---|---|
| src/parser/solutionParser.ts | parseSolutionZip | Parse solution.zip content into canonical ParsedSolution model. |
| src/generator/markdownGenerator.ts | generateMarkdown | Generate a full markdown document for a single parsed solution. |
| src/generator/markdownGenerator.ts | generateConsolidatedMarkdown | Generate combined markdown across multiple parsed solutions. |
| src/generator/markdownGenerator.ts | consolidateSolutions | Merge multiple ParsedSolution objects into a single aggregate model. |
| src/App.tsx | buildConsolidatedResult | Build synthetic consolidated result tab for UI display/export. |
| src/workers/parseSolutions.ts | parseSolutionsConcurrently | Parse queued ZIPs in a bounded worker pool and fall back to the main thread when required. |
| src/api/excelExporter.ts | createInstallationGuideWorkbook | Convert installation-guide sections and Markdown tables into workbook sheets. |
| src/api/excelExporter.ts | createPackagedSolutionWorkbook | Build a solution workbook and post-process its OOXML parts. |
| src/generator/dependencyAnalysis.ts | analyzeDependencies | Classify declared prerequisites and identify selected/component owners. |
| src/generator/installationAnalysis.ts | analyzeInstallation | Calculate the selected solutions' install order, blockers, cycles, downstream waiters and evidence. |
| src/generator/installationGraph.ts | buildInstallationGraph | Identify cycle groups and prerequisite readiness using a selected-solution graph. |

### Key Supporting Functions
| Module | Function | Responsibility |
|---|---|---|
| src/parser/solutionParser.ts | getLocalizedLabel | Resolve display labels from localized XML structures and direct display tags. |
| src/parser/solutionParser.ts | xmlStr | Safe extraction helper for mixed XML node/attribute value patterns. |
| src/generator/markdownGenerator.ts | buildEntityDisplayMap | Logical-name to display-name lookup map for tables. |
| src/generator/markdownGenerator.ts | buildAttributeDisplayMap | Logical-name to display-name lookup map for columns. |
| src/generator/markdownGenerator.ts | resolveAttributeDisplayName | Robust CLS column display-name resolution for qualified/unqualified names. |
| src/components/ui/MarkdownViewer.tsx | code renderer override | Intercepts mermaid fenced blocks and renders diagram component. |
| src/components/ui/MermaidRenderer.tsx | renderSvg | Validate and background-fill Mermaid SVG for viewer and exports. |
| src/components/ui/MermaidDiagram.tsx | render effect | Interactive Mermaid SVG display, accessibility attributes, zoom/fit/fullscreen behavior. |
| src/context/ThemeContext.tsx | applyTheme | Applies active theme to document root and theme-color metadata. |
| scripts/build-latest.ps1 | build-latest flow | Build, package, mirror latest artifacts, and clean temp outputs. |
