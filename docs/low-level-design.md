# PP-MD Low-Level Design

## Current Local-File Modules

`rebuildResults` respects explicit inclusion/separation settings, and a shared regeneration callback preserves the current filename or returns from a removed companion to its main document. Sidebar groups retain original result indexes. `applyDocumentationScope`, `selectDocumentationAttributes`, and `filterDocumentationPrivileges` align Markdown and workbook option behavior.

- `parser/solutionParser.ts`: optional AbortSignal checkpoints, ZIP table-property metadata and authoritative MissingDependencies parsing.
- `generator/dependencyAnalysis.ts`, `localDocumentation.ts` and `missingDependencies.ts`: shared classification, component identity evidence, installation ordering and cycle/member distinction.
- `generator/installationAnalysis.ts` and `installationGraph.ts`: reusable selected-ZIP planner, dependency layers, blocker propagation, Tarjan strongly connected components, cycle edges and downstream waiters.
- `generator/documentSplitter.ts`: category generation using Markdown AST heading boundaries and self-contained contents links.
- `generator/tableMetadata.ts`: shared table-property blocks for Markdown and Excel, retaining unknown/unavailable states.
- `api/documentRendering.tsx`: sanitised static HTML, Mermaid SVG generation and bounded rasterisation with error fallback.
- `api/pdfExporter.ts`: searchable paginated PDF drawing with browser-yield timer fallback.
- `api/excelExporter.ts`: styled inventories, per-table worksheets, Overview contents links, return-navigation buttons, linked overflow continuations, installation dashboards/charts, and dependency evidence.
- `api/excelPackage.ts`: OOXML post-processing for native Excel tables, valid headers, frozen panes, and repaired style XML.
- `utils/documentExport.ts`: single and batch downloads, collision-safe archive paths and cancellation before download.
- `workers/parseSolutions.ts`, `solution.worker.ts`: bounded concurrent parsing and worker-unavailable fallback.
- `components/ui/MermaidRenderer.tsx`: low-priority diagram scheduling, classic Dagre settings, SVG/background validation and a bounded cache; `MermaidDiagram.tsx` provides the interactive viewer.

Connected environment appendices and collectors are deliberately absent. Existing Electron tagged-PDF support remains available. Imported Markdown can be viewed/exported in non-structured formats, but does not claim a parsed solution workbook.

## Purpose
This document details module internals, key functions, data contracts, and implementation-level behavior.

## Source Layout
- electron/main.cjs: Desktop host runtime and browser window lifecycle
- src/App.tsx: UI orchestration and application state machine
- src/parser/solutionParser.ts: Solution ZIP and XML parsing pipeline
- src/generator/markdownGenerator.ts: Markdown composition and diagram generation
- src/generator/dependencyAnalysis.ts: Component-level prerequisite classification
- src/generator/localDocumentation.ts: Installation/dependency report generation
- src/api/excelExporter.ts and excelPackage.ts: Workbook construction and OOXML packaging
- src/workers/: Background ZIP parsing coordination and worker entry point
- src/types/solution.ts: Canonical domain model contracts
- src/components/ui/*: Reusable UI widgets
- src/context/ThemeContext.tsx: Theme state and persistence

## Module Details

### Electron Main Process
File: electron/main.cjs

Responsibilities:
- configure BrowserWindow dimensions and security-related webPreferences
- choose renderer entry point based on app.isPackaged
- handle load and renderer-failure diagnostics
- deny new in-app windows and route external links to system browser

Control flow:
1. app.whenReady triggers createMainWindow
2. activate recreates a window when no windows remain (macOS convention)
3. window-all-closed quits app for non-darwin platforms

### Application Orchestration
File: src/App.tsx

Primary state:
- results: successful solution parsing and markdown outputs
- activeIdx: currently selected documentation tab
- processing: per-file progress state
- isProcessing/statusMsg: runtime status and announcements
- documentContext/configurations: metadata context and reusable presets
- erdMode and viewer loading flags

Key internal functions:
- readSavedConfigurations and writeSavedConfigurations
  - localStorage persistence for custom doc presets
- buildConsolidatedResult
  - merges all loaded ParsedSolution objects and regenerates combined markdown
- isLargeOrComplexMarkdown
  - computes heuristic for loading feedback on heavy markdown payloads
- file processing callbacks
  - dispatch ZIP reads through `parseSolutionsConcurrently`, update per-file progress, generate markdown and update UI state transitions

State machine summary:
- Idle -> FilesQueued -> Processing -> ResultsAvailable -> Export/Reset

### Parsing Pipeline
File: src/parser/solutionParser.ts

Entry point:
- `parseSolutionZip(input, onProgress, signal?)`

Core parser stages:
1. open ZIP with JSZip
2. parse solution.xml for metadata
3. parse customizations.xml for entities and component inventories
4. parse additional folders/files (workflows, apps, plugins, reports)
5. assemble ParsedSolution with warnings and progress updates

Background coordinator:
- `parseSolutionsConcurrently(files, signal, callbacks)` maintains a two-job bound.
- `solution.worker.ts` transfers an `ArrayBuffer` and returns progress/result/error messages.
- A `WorkerUnavailableError` switches the remaining queue to local `parseSolutionZip` calls.
- Cancellation stops outstanding scheduling; settled parse results remain available to the caller.

Important helpers:
- xmlStr
  - robust value extraction across text nodes, mixed object shapes, and attributes
- getLocalizedLabel
  - resolves display labels from direct tags and localized containers
- parsePermissionFlag and parseBooleanLike
  - normalizes non-uniform XML boolean/permission representations
- getEntriesWithPrefix and readZipEntry
  - resilient ZIP entry resolution and traversal
- readManifestMissingDependencies
  - preserves required/dependent IDs, `id.*` keys, parent identities, component types, schema names and managed package owners
- parseManifestComponents
  - retains root-component behavior used to distinguish a full table from a shell

Notable defensive behavior:
- case-insensitive tag access patterns
- mixed attribute/tag fallback chains for legacy and variant exports
- explicit parser hardening settings (entity processing disabled)

### Dependency and Installation Analysis

Files: `src/generator/dependencyAnalysis.ts`, `src/generator/missingDependencies.ts`, `src/generator/localDocumentation.ts`

- `analyzeDependencies` builds IDs and typed-name indexes for each selected ZIP.
- Matches prefer a full owner over a table shell and understand parent-table include-all-subcomponents behavior.
- Manifest owners, standard platform components, Microsoft/package owners and selected providers use a shared classification result.
- `requiredComponentIdentity` separates table, column, schema name and component ID for reports and workbooks.
- Installation analysis performs version checks, provider-first ordering and actual-cycle membership.
- `buildInstallationGraph` performs Tarjan strongly connected-component analysis and computes dependency readiness layers.
- Solutions depending on a cycle are marked as waiting for a blocked provider, not as members of the cycle.
- The same resolutions feed Markdown reports, installation guidance and table-specific Excel evidence.

### Excel Packaging

Files: `src/api/excelExporter.ts`, `src/api/excelPackage.ts`, `src/utils/documentExport.ts`

- `createSolutionWorkbook` retains the synchronous Blob API for focused workbook tests.
- `createPackagedSolutionWorkbook` post-processes OOXML before download.
- `registerExcelTable` normalises headers and records multiple table ranges per worksheet.
- `packageExcelWorkbook` emits table parts, relationships and frozen-pane XML and repairs style order/colour alpha values expected by Excel.
- Solution cells over 32,767 characters keep the first chunk, preserve the remainder on Text Continuations, and receive a comment plus bidirectional hyperlinks.
- Packaged solution and installation workbooks receive a linked Overview Table of Contents and Back to Contents controls; table worksheets retain Back to Tables.
- `createInstallationGuideWorkbook` parses guide headings and Markdown tables into section-specific worksheets.
- Dependency tables include dependent/required IDs plus status and evidence; installation guides export as workbooks rather than Markdown companions.
- Installation workbooks add a decision dashboard with native OOXML charts and equivalent data tables, plus nodes, readiness, evidence, target prerequisite, blocker and cycle sheets.
- When an `InstallationAnalysis` is supplied, the workbook adds Installation Nodes, Dependency Edges, Blocker Details and Circular Dependencies sheets from the exact model used for Markdown.

### Markdown Generation Pipeline
File: src/generator/markdownGenerator.ts

Public API:
- consolidateSolutions(solutions)
- generateMarkdown(solution, options)
- generateConsolidatedMarkdown(solutions, options)

Generation stages:
1. preprocess lookup maps (entity and attribute display names)
2. emit document header and table of contents
3. emit section generators in deterministic order
4. append back-to-top anchors and return markdown string

Representative section generators:
- entities/tables and columns
- option sets
- forms and views
- processes and automation
- apps
- web resources
- security roles and field security profiles
- integration objects (connection references, env vars)
- reports/dashboards and plugin assemblies

Key helpers:
- labelWithSchema
  - combines display and logical/schema labels consistently
- buildEntityDisplayMap and buildAttributeDisplayMap
  - lookup maps used across sections for user-friendly labels
- resolveAttributeDisplayName
  - resilient lookup for CLS attribute names with qualified-path normalization
- parseRolePrivilege
  - decomposes privilege names into operation and table logical name

### Markdown Rendering
File: src/components/ui/MarkdownViewer.tsx

Responsibilities:
- render markdown with GitHub Flavored Markdown support
- inject heading IDs for table-of-contents linking
- intercept mermaid code fences and render MermaidDiagram component
- provide raw markdown toggle, copy, and export actions

Critical methods:
- extractText and headingId
  - deterministic heading anchor generation
- handleCopy
  - clipboard write and status feedback
- code renderer override
  - route language-mermaid blocks to diagram renderer

### Diagram Renderer
Files: src/components/ui/MermaidRenderer.tsx and MermaidDiagram.tsx

Responsibilities:
- lazily import Mermaid and render SVG asynchronously
- enforce accessible SVG attributes and fallback source disclosure
- support zoom, fit-to-width, and fullscreen interactions

Implementation details:
- useId-based unique diagram IDs to avoid collisions
- strict Mermaid security mode
- runtime font-size normalization for readability
- classic Dagre layout settings shared by viewer and export rendering
- validated SVG viewBox and theme background before background-rect insertion
- visible-first queue that pauses during scrolling and a 100-entry LRU SVG cache

### Theme Management
File: src/context/ThemeContext.tsx

Responsibilities:
- initialize theme from localStorage or OS preference
- apply active theme to html[data-theme]
- sync browser theme-color metadata
- expose toggleTheme and setTheme API to consumers

## Data Contracts
File: src/types/solution.ts

Core models:
- ParsedSolution: root aggregate for all parsed components
- SolutionMetadata: identity/version/publisher metadata
- EntityDefinition and EntityAttribute: table and column schema
- ProcessDefinition and ProcessStep: workflow/process structures
- SecurityRoleDefinition and FieldSecurityProfileDefinition
- ConnectionReferenceDefinition and EnvironmentVariableDefinition
- PluginAssemblyDefinition and PluginStepDefinition

Design intent:
- strict typing for parser-generator contract stability
- explicit enums for categories/types where possible
- optional fields used for partial availability across source variants

## Error Handling Strategy
- Parser warnings collected for non-fatal parse gaps.
- Component-level try/catch around variable XML structures.
- UI status messaging for processing and validation errors.
- Electron modal dialogs for packaged renderer load/crash failures.

## Performance Considerations
- Per-file progress callbacks to avoid opaque long-running operations.
- Heuristic loading state for large markdown rendering workloads.
- Mermaid dynamically imported to reduce initial bundle cost.
- ZIP parsing bounded to two workers to limit CPU and buffer-copy cost.
- Excel table and overflow metadata generated once during packaged export.
- Build warns for large chunks but allows diagram-heavy output needs.

## Build and Packaging Internals
- build-latest.ps1 executes build and both package targets, then mirrors outputs.
- check-contrast.mjs validates token-level WCAG contrast thresholds.
- smoke-doc.ts validates parser/generator behavior against synthetic fixture ZIP.

## Low-Level Improvement Backlog
- add unit tests around parser fallback chains and CLS name normalization
- move parser helper clusters into dedicated modules for smaller compilation units
- introduce dedicated hooks under src/hooks for complex App.tsx orchestration logic
- add schema validation for doc-configurations.json payloads
