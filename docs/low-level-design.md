# PP-MD Low-Level Design

## Purpose
This document details module internals, key functions, data contracts, and implementation-level behavior.

## Source Layout
- electron/main.cjs: Desktop host runtime and browser window lifecycle
- src/App.tsx: UI orchestration and application state machine
- src/parser/solutionParser.ts: Solution ZIP and XML parsing pipeline
- src/generator/markdownGenerator.ts: Markdown composition and diagram generation
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
  - validate, parse, generate markdown, and update UI state transitions

State machine summary:
- Idle -> FilesQueued -> Processing -> ResultsAvailable -> Export/Reset

### Parsing Pipeline
File: src/parser/solutionParser.ts

Entry point:
- parseSolutionZip(input, options)

Core parser stages:
1. open ZIP with JSZip
2. parse solution.xml for metadata
3. parse customizations.xml for entities and component inventories
4. parse additional folders/files (workflows, apps, plugins, reports)
5. assemble ParsedSolution with warnings and progress updates

Important helpers:
- xmlStr
  - robust value extraction across text nodes, mixed object shapes, and attributes
- getLocalizedLabel
  - resolves display labels from direct tags and localized containers
- parsePermissionFlag and parseBooleanLike
  - normalizes non-uniform XML boolean/permission representations
- getEntriesWithPrefix and readZipEntry
  - resilient ZIP entry resolution and traversal

Notable defensive behavior:
- case-insensitive tag access patterns
- mixed attribute/tag fallback chains for legacy and variant exports
- explicit parser hardening settings (entity processing disabled)

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
File: src/components/ui/MermaidDiagram.tsx

Responsibilities:
- lazily import Mermaid and render SVG asynchronously
- enforce accessible SVG attributes and fallback source disclosure
- support zoom, fit-to-width, and fullscreen interactions

Implementation details:
- useId-based unique diagram IDs to avoid collisions
- strict Mermaid security mode
- runtime font-size normalization for readability

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
