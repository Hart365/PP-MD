/**
 * @file App.tsx
 * @description Root application component for PP-MD.
 *
 * Application flow:
 *  1. User lands on the drop zone (welcome screen).
 *  2. User selects/drops one or more solution ZIP files.
 *  3. Each ZIP is parsed asynchronously (individual progress per file).
 *  4. For each parsed solution a Markdown document is generated.
 *  5. The generated documentation is displayed in the MarkdownViewer.
 *  6. A sidebar lists all parsed solutions for navigation.
 *  7. Users can export each document as a .md file.
 *
 * WCAG 2.2 compliance:
 *  - 2.4.1 Bypass Blocks: A "Skip to main content" link is the first
 *    focusable element on the page.
 *  - 1.3.1 Info & Relationships: Landmark roles (header, main, nav) used.
 *  - 4.1.3 Status Messages: Processing status announced via aria-live.
 *  - 3.2.2 On Input: No unexpected context changes on file selection alone;
 *    user must click Generate.
 */

import { useState, useCallback, useEffect, type ChangeEvent } from 'react';
import { saveAs } from 'file-saver';
import JSZip from 'jszip';
import { DropZone }          from './components/ui/DropZone';
import { ProgressBar }       from './components/ui/ProgressBar';
import { MarkdownViewer }    from './components/ui/MarkdownViewer';
import { SolutionSidebar }   from './components/SolutionSidebar';
import { ThemeToggle }       from './components/ui/ThemeToggle';
import { parseSolutionZip }  from './parser/solutionParser';
import {
  generateMarkdown,
  generateConsolidatedMarkdown,
  consolidateSolutions,
  type DocumentContext,
} from './generator/markdownGenerator';
import type { ParsedSolution } from './types/solution';
import appIcon from './assets/app-icon.svg';
import styles from './App.module.css';

// ---------------------------------------------------------------------------
// Types
// ---------------------------------------------------------------------------

/** State for a single file being processed */
interface ProcessingEntry {
  fileName: string;
  progress: number; // 0–100
  error?: string;
}

/** A fully-processed result */
interface SolutionResult {
  solution:  ParsedSolution;
  markdown:  string;
  fileName:  string;
  isConsolidated?: boolean;
}

type ErdMode = 'compact' | 'detailed-relationships';

interface SavedDocumentConfiguration extends DocumentContext {
  id: string;
  name: string;
}

interface ConfigurationFile {
  configurations: SavedDocumentConfiguration[];
}

const LOCAL_CONFIG_STORAGE_KEY = 'pp-md-doc-configurations';
const LOCAL_HIDDEN_CONFIG_IDS_KEY = 'pp-md-hidden-doc-configuration-ids';

const EMPTY_DOCUMENT_CONTEXT: DocumentContext = {
  client: '',
  project: '',
  contract: '',
  sow: '',
  sprint: '',
  releaseDate: '',
};

function toSafeMarkdownBaseName(rawName: string | undefined | null, fallback: string): string {
  const source = (rawName || fallback).trim();
  return source
    .replace(/[^\w\s.-]/g, '')
    .replace(/\s+/g, '_');
}

function solutionDisplayName(solution: ParsedSolution): string {
  return (solution.metadata.displayName || solution.metadata.uniqueName || '').trim();
}

function sortSolutionResults(results: SolutionResult[]): SolutionResult[] {
  return [...results].sort((left, right) => solutionDisplayName(left.solution).localeCompare(
    solutionDisplayName(right.solution),
    undefined,
    { numeric: true, sensitivity: 'base' },
  ));
}

function sortFilesByName(files: File[]): File[] {
  return [...files].sort((left, right) => left.name.localeCompare(
    right.name,
    undefined,
    { numeric: true, sensitivity: 'base' },
  ));
}

function readSavedConfigurations(): SavedDocumentConfiguration[] {
  try {
    const raw = localStorage.getItem(LOCAL_CONFIG_STORAGE_KEY);
    if (!raw) return [];
    const parsed = JSON.parse(raw) as unknown;
    if (!Array.isArray(parsed)) return [];

    return parsed
      .filter((entry) => typeof entry === 'object' && entry !== null)
      .map((entry) => entry as SavedDocumentConfiguration)
      .filter((entry) => !!entry.id && !!entry.name);
  } catch {
    return [];
  }
}

function writeSavedConfigurations(configs: SavedDocumentConfiguration[]): void {
  try {
    localStorage.setItem(LOCAL_CONFIG_STORAGE_KEY, JSON.stringify(configs));
  } catch {
    // Best effort only; continue without blocking UX.
  }
}

function readHiddenConfigurationIds(): string[] {
  try {
    const raw = localStorage.getItem(LOCAL_HIDDEN_CONFIG_IDS_KEY);
    if (!raw) return [];
    const parsed = JSON.parse(raw) as unknown;
    if (!Array.isArray(parsed)) return [];
    return parsed.filter((entry): entry is string => typeof entry === 'string' && entry.trim().length > 0);
  } catch {
    return [];
  }
}

function writeHiddenConfigurationIds(ids: string[]): void {
  try {
    localStorage.setItem(LOCAL_HIDDEN_CONFIG_IDS_KEY, JSON.stringify(ids));
  } catch {
    // Best effort only; continue without blocking UX.
  }
}

function buildConsolidatedResult(
  results: SolutionResult[],
  _erdMode: ErdMode,
  documentContext: DocumentContext,
): SolutionResult {
  const solutions = results.map((r) => r.solution);
  const markdown = generateConsolidatedMarkdown(solutions, { documentContext });
  const aggregated: ParsedSolution = consolidateSolutions(solutions);

  return {
    solution: aggregated,
    markdown,
    fileName: 'consolidated-summary.md',
    isConsolidated: true,
  };
}

function rebuildResults(
  results: SolutionResult[],
  erdMode: ErdMode,
  documentContext: DocumentContext,
): SolutionResult[] {
  const base = sortSolutionResults(results
    .filter((entry) => !entry.isConsolidated)
    .map((entry) => ({
      ...entry,
      markdown: generateMarkdown(entry.solution, { erdMode, documentContext }),
    })));

  if (base.length > 1) {
    return [buildConsolidatedResult(base, erdMode, documentContext), ...base];
  }

  return base;
}

/**
 * Estimates whether Markdown content is expensive to render.
 *
 * This keeps the UI responsive by showing a temporary "Please Wait" state
 * before mounting very large/complex markdown documents.
 */
function isLargeOrComplexMarkdown(markdown: string): boolean {
  const lengthThreshold = 45000;
  const headingThreshold = 80;
  const tableRowThreshold = 250;
  const mermaidThreshold = 6;

  const headingCount = (markdown.match(/^#{2,6}\s+/gm) ?? []).length;
  const tableRowCount = (markdown.match(/^\|.*\|\s*$/gm) ?? []).length;
  const mermaidCount = (markdown.match(/```mermaid/g) ?? []).length;

  return markdown.length >= lengthThreshold
    || headingCount >= headingThreshold
    || tableRowCount >= tableRowThreshold
    || mermaidCount >= mermaidThreshold;
}

// ---------------------------------------------------------------------------
// Component
// ---------------------------------------------------------------------------

/**
 * Main application component.
 */
export default function App() {
  /** List of results (one per successfully processed ZIP) */
  const [results,       setResults]       = useState<SolutionResult[]>([]);
  /** Index of the currently displayed result */
  const [activeIdx,     setActiveIdx]     = useState<number>(0);
  /** Per-file progress tracking (shown during processing) */
  const [processing,    setProcessing]    = useState<ProcessingEntry[]>([]);
  /** Whether any files are currently being processed */
  const [isProcessing,  setIsProcessing]  = useState<boolean>(false);
  /** Top-level status message (announced to screen readers) */
  const [statusMsg,     setStatusMsg]     = useState<string>('');
  /** ERD rendering mode */
  const [erdMode,       setErdMode]       = useState<ErdMode>('detailed-relationships');
  /** Document context for MD header details */
  const [documentContext, setDocumentContext] = useState<DocumentContext>(EMPTY_DOCUMENT_CONTEXT);
  /** Preset configurations loaded from JSON */
  const [configurations, setConfigurations] = useState<SavedDocumentConfiguration[]>([]);
  /** Selected configuration id from dropdown */
  const [selectedConfigId, setSelectedConfigId] = useState<string>('custom');
  /** Configuration load failure text */
  const [configLoadError, setConfigLoadError] = useState<string>('');
  /** New configuration name for save action */
  const [newConfigName, setNewConfigName] = useState<string>('');
  /** True when switching to a heavy markdown document so we can show feedback */
  const [isViewerLoading, setIsViewerLoading] = useState<boolean>(false);

  useEffect(() => {
    let active = true;

    const loadConfigurations = async () => {
      const localConfigs = readSavedConfigurations();
      const hiddenConfigIds = new Set(readHiddenConfigurationIds());

      try {
        const response = await fetch('./doc-configurations.json', { cache: 'no-store' });
        if (!response.ok) {
          throw new Error(`HTTP ${response.status}`);
        }
        const data = (await response.json()) as ConfigurationFile;
        if (!Array.isArray(data.configurations)) {
          throw new Error('Invalid configuration shape');
        }
        if (!active) return;
        setConfigurations(
          [...data.configurations, ...localConfigs].filter((config) => !hiddenConfigIds.has(config.id)),
        );
        setConfigLoadError('');
      } catch (err: unknown) {
        if (!active) return;
        setConfigurations(localConfigs.filter((config) => !hiddenConfigIds.has(config.id)));
        setConfigLoadError(`Could not load doc-configurations.json: ${(err as Error).message}`);
      }
    };

    loadConfigurations();
    return () => {
      active = false;
    };
  }, []);

  const applyConfiguration = useCallback((config: SavedDocumentConfiguration) => {
    const nextDocumentContext: DocumentContext = {
      client: config.client ?? '',
      project: config.project ?? '',
      contract: config.contract ?? '',
      sow: config.sow ?? '',
      sprint: config.sprint ?? '',
      releaseDate: config.releaseDate ?? '',
    };

    setDocumentContext({
      client: nextDocumentContext.client,
      project: nextDocumentContext.project,
      contract: nextDocumentContext.contract,
      sow: nextDocumentContext.sow,
      sprint: nextDocumentContext.sprint,
      releaseDate: nextDocumentContext.releaseDate,
    });
    setResults((prev) => rebuildResults(prev, erdMode, nextDocumentContext));
  }, [erdMode]);

  const handleConfigurationSelect = useCallback((event: ChangeEvent<HTMLSelectElement>) => {
    const value = event.target.value;
    setSelectedConfigId(value);
    if (value === 'custom') return;
    const selected = configurations.find((config) => config.id === value);
    if (!selected) return;
    applyConfiguration(selected);
    setNewConfigName(selected.name);
  }, [configurations, applyConfiguration]);

  const handleContextChange = useCallback((field: keyof DocumentContext, value: string) => {
    const nextDocumentContext = { ...documentContext, [field]: value };
    setSelectedConfigId('custom');
    setDocumentContext(nextDocumentContext);
    setResults((prev) => rebuildResults(prev, erdMode, nextDocumentContext));
  }, [documentContext, erdMode]);

  const handleSaveConfiguration = useCallback(() => {
    const name = newConfigName.trim();
    if (!name) {
      setStatusMsg('Please enter a configuration name before saving.');
      return;
    }

    const configId = `local-${Date.now()}`;
    const configToSave: SavedDocumentConfiguration = {
      id: configId,
      name,
      client: documentContext.client,
      project: documentContext.project,
      contract: documentContext.contract,
      sow: documentContext.sow,
      sprint: documentContext.sprint,
      releaseDate: documentContext.releaseDate,
    };

    const savedConfigs = readSavedConfigurations();
    const withoutExistingName = savedConfigs.filter((cfg) => cfg.name.toLowerCase() !== name.toLowerCase());
    const updatedSavedConfigs = [...withoutExistingName, configToSave];
    writeSavedConfigurations(updatedSavedConfigs);

    setConfigurations((prev) => {
      const withoutExistingNameInDropdown = prev.filter((cfg) => cfg.name.toLowerCase() !== name.toLowerCase());
      return [...withoutExistingNameInDropdown, configToSave];
    });

    setSelectedConfigId(configId);
    setStatusMsg(`Configuration "${name}" saved.`);
  }, [newConfigName, documentContext]);

  const handleDeleteConfiguration = useCallback((configId: string) => {
    if (configId === 'custom') return;

    const configToDelete = configurations.find((config) => config.id === configId);
    if (!configToDelete) return;

    const updatedHiddenIds = Array.from(new Set([...readHiddenConfigurationIds(), configId]));
    writeHiddenConfigurationIds(updatedHiddenIds);

    if (configId.startsWith('local-')) {
      const updatedSaved = readSavedConfigurations().filter((config) => config.id !== configId);
      writeSavedConfigurations(updatedSaved);
    }

    setConfigurations((prev) => prev.filter((config) => config.id !== configId));

    if (selectedConfigId === configId) {
      setSelectedConfigId('custom');
      setNewConfigName('');
    }

    setStatusMsg(`Configuration "${configToDelete.name}" deleted.`);
  }, [configurations, selectedConfigId]);

  const updateProcessingEntry = useCallback(
    (index: number, updater: (entry: ProcessingEntry) => ProcessingEntry) => {
      setProcessing((prev) => {
        if (!prev[index]) return prev;
        const updated = [...prev];
        updated[index] = updater(updated[index]);
        return updated;
      });
    },
    [],
  );

  // ── File processing ───────────────────────────────────────────────────────

  /**
   * Processes an array of ZIP files sequentially, updating progress state
   * as each file is parsed.
   *
   * @param files - Files selected by the user
   */
  const handleFilesSelected = useCallback(async (files: File[]) => {
    const sortedFiles = sortFilesByName(files);

    setIsProcessing(true);
    setStatusMsg(`Processing ${sortedFiles.length} file${sortedFiles.length > 1 ? 's' : ''}…`);

    // Initialise progress tracking for all files
    const initial: ProcessingEntry[] = sortedFiles.map((f) => ({
      fileName: f.name,
      progress: 0,
    }));
    setProcessing(initial);

    const newResults: SolutionResult[] = [];

    for (let i = 0; i < sortedFiles.length; i++) {
      const file = sortedFiles[i];

      try {
        setStatusMsg(`Parsing ${file.name}…`);

        const solution = await parseSolutionZip(file, (pct) => {
          // Update progress for this specific file
          updateProcessingEntry(i, (entry) => ({ ...entry, progress: pct }));
        });

        setStatusMsg(`Generating documentation for ${file.name}…`);
        const markdown = generateMarkdown(solution, { erdMode, documentContext });

        newResults.push({ solution, markdown, fileName: file.name });

        // Mark as complete
        updateProcessingEntry(i, (entry) => ({ ...entry, progress: 100 }));
      } catch (err: unknown) {
        const msg = (err as Error).message ?? 'Unknown error';
        updateProcessingEntry(i, (entry) => ({ ...entry, progress: 0, error: msg }));
        // Continue with remaining files rather than aborting
      }
    }

    if (newResults.length > 0) {
      const existingBase = results.filter((r) => !r.isConsolidated);
      const mergedBase = sortSolutionResults([...existingBase, ...newResults]);
      const nextResults = mergedBase.length > 1
        ? [buildConsolidatedResult(mergedBase, erdMode, documentContext), ...mergedBase]
        : mergedBase;

      setResults(nextResults);
      setActiveIdx(nextResults.length > 1 && nextResults[0].isConsolidated ? 0 : Math.max(0, nextResults.length - 1));
      const count  = newResults.length;
      const failed = sortedFiles.length - count;
      setStatusMsg(
        failed > 0
          ? `Done: ${count} document${count > 1 ? 's' : ''} generated, ${failed} file${failed > 1 ? 's' : ''} failed.`
          : `Done: ${count} document${count > 1 ? 's' : ''} generated successfully.`,
      );
    } else {
      setStatusMsg('No documents generated. Check that the files are valid Power Platform solution ZIPs.');
    }

    setIsProcessing(false);
    // Clear progress indicators after a brief delay
    setTimeout(() => setProcessing([]), 1500);
  }, [results, erdMode, documentContext, updateProcessingEntry]);

  const handleToggleErdMode = useCallback(() => {
    const nextMode: ErdMode = erdMode === 'detailed-relationships' ? 'compact' : 'detailed-relationships';

    setResults((prev) => rebuildResults(prev, nextMode, documentContext));

    setErdMode(nextMode);
    setStatusMsg(`ERD mode switched to ${nextMode === 'compact' ? 'Compact' : 'Detailed-Relationships'}.`);
  }, [erdMode, documentContext]);

  /**
   * Handles sidebar document selection with optional loading feedback for
   * large markdown payloads.
   */
  const handleSelectResult = useCallback((index: number) => {
    const next = results[index];
    if (!next) return;

    if (!isLargeOrComplexMarkdown(next.markdown)) {
      setActiveIdx(index);
      setIsViewerLoading(false);
      return;
    }

    setIsViewerLoading(true);

    // Yield one frame so the loading indicator can paint before heavy render work.
    window.setTimeout(() => {
      setActiveIdx(index);
      window.setTimeout(() => {
        setIsViewerLoading(false);
      }, 300);
    }, 50);
  }, [results]);

  // ── Export ────────────────────────────────────────────────────────────────

  /**
   * Exports the active solution's Markdown as a downloadable .md file.
   */
  const handleExport = useCallback(() => {
    const result = results[activeIdx];
    if (!result) return;
    const blob = new Blob([result.markdown], { type: 'text/markdown;charset=utf-8' });
    const safeName = toSafeMarkdownBaseName(
      result.solution.metadata.displayName || result.solution.metadata.uniqueName,
      'solution',
    );
    saveAs(blob, `${safeName}-documentation.md`);
  }, [results, activeIdx]);

  /**
   * Exports all generated Markdown documents as a single ZIP archive.
   */
  const handleExportAll = useCallback(async () => {
    if (results.length === 0) return;

    const zip = new JSZip();
    results.forEach((result, idx) => {
      const fallback = result.isConsolidated ? 'consolidated' : `solution_${idx + 1}`;
      const safeName = toSafeMarkdownBaseName(
        result.solution.metadata.displayName || result.solution.metadata.uniqueName,
        fallback,
      );
      const suffix = result.isConsolidated ? '-summary.md' : '-documentation.md';
      zip.file(`${safeName}${suffix}`, result.markdown);
    });

    const blob = await zip.generateAsync({ type: 'blob' });
    saveAs(blob, 'pp-md-markdown-documents.zip');
  }, [results]);

  // ── Reset ─────────────────────────────────────────────────────────────────

  /**
   * Clears all results and returns to the welcome/drop zone screen.
   */
  const handleReset = useCallback(() => {
    setResults([]);
    setActiveIdx(0);
    setProcessing([]);
    setStatusMsg('');
  }, []);

  // ── Derived state ─────────────────────────────────────────────────────────

  const hasResults     = results.length > 0;
  const activeResult   = results[activeIdx];
  const isWelcome      = !hasResults && !isProcessing;

  // ── Render ────────────────────────────────────────────────────────────────

  return (
    <div className={styles.appRoot}>
      {/*
       * Skip to main content link — MUST be the first focusable element
       * (WCAG 2.4.1 Bypass Blocks).
       */}
      <a href="#main-content" className={styles.skipLink}>
        Skip to main content
      </a>

      {/* ── Global header ─────────────────────────────────────────────── */}
      <header className={styles.header} role="banner">
        <div className={styles.headerInner}>
          {/* Logo / app name */}
          <div className={styles.brand}>
            <img src={appIcon} className={styles.brandIcon} alt="" aria-hidden="true" />
            <h1 className={styles.brandName}>PP-MD</h1>
            <span className={styles.brandTagline}>Power Platform Documentation Generator</span>
          </div>

          {/* Header actions */}
          <div className={styles.headerActions}>
            {hasResults && !isProcessing && (
              <button
                type="button"
                className={styles.headerBtn}
                onClick={handleToggleErdMode}
                aria-label="Toggle ERD detail level"
              >
                {erdMode === 'compact' ? 'ERD: Compact' : 'ERD: Detailed'}
              </button>
            )}

            {hasResults && !isProcessing && (
              <button
                type="button"
                className={styles.headerBtn}
                onClick={handleExportAll}
                aria-label="Download all generated Markdown files"
              >
                ⬇ All .md
              </button>
            )}

            {hasResults && !isProcessing && (
              <button
                type="button"
                className={styles.headerBtn}
                onClick={handleReset}
                aria-label="Clear all results and add new files"
              >
                ＋ New
              </button>
            )}
            <ThemeToggle />
          </div>
        </div>
      </header>

      {/*
       * Status message region — announced to screen readers via aria-live.
       * Visually hidden when empty.
       */}
      <div
        aria-live="polite"
        aria-atomic="true"
        className={styles.statusRegion}
        role="status"
      >
        {statusMsg}
      </div>

      {/* ── Main content area ─────────────────────────────────────────── */}
      <div className={styles.body}>
        {/* Sidebar (only when we have results) */}
        {hasResults && (
          <SolutionSidebar
            solutions={results.map((r) => r.solution)}
            activeIndex={activeIdx}
            onSelect={handleSelectResult}
            onReset={handleReset}
          />
        )}

        {/* Main content */}
        <main id="main-content" className={styles.main} tabIndex={-1}>

          {/* Welcome / drop zone screen */}
          {isWelcome && (
            <section
              className={styles.welcomeSection}
              aria-labelledby="welcome-heading"
            >
              <h2 id="welcome-heading" className={styles.welcomeHeading}>
                Generate documentation from your solution files
              </h2>
              <p className={styles.welcomeSubtitle}>
                Drop one or more Power Platform solution <code>.zip</code> archives below.
                PP-MD will parse every component and produce comprehensive Markdown
                documentation — including architecture and ERD Mermaid diagrams.
              </p>

              <section className={styles.contextPanel} aria-labelledby="context-heading">
                <h3 id="context-heading" className={styles.contextHeading}>Document Header Details</h3>

                <div className={styles.contextRow}>
                  <label htmlFor="config-select" className={styles.contextLabel}>Configuration</label>
                  <div className={styles.contextSelectRow}>
                    <select
                      id="config-select"
                      className={styles.contextSelect}
                      value={selectedConfigId}
                      onChange={handleConfigurationSelect}
                      aria-label="Select document configuration"
                    >
                      <option value="custom">Custom (manual entry)</option>
                      {configurations.map((config) => (
                        <option key={config.id} value={config.id}>{config.name}</option>
                      ))}
                    </select>
                    <button
                      type="button"
                      className={styles.contextDeleteBtn}
                      onClick={() => handleDeleteConfiguration(selectedConfigId)}
                      disabled={selectedConfigId === 'custom'}
                      aria-label="Delete selected configuration"
                      title="Delete selected configuration"
                    >
                      🗑
                    </button>
                  </div>
                </div>

                <div className={styles.contextSaveRow}>
                  <input
                    type="text"
                    className={styles.contextNameInput}
                    placeholder="Configuration name"
                    value={newConfigName}
                    onChange={(e) => setNewConfigName(e.target.value)}
                    aria-label="Configuration name"
                  />
                  <button
                    type="button"
                    className={styles.contextSaveBtn}
                    onClick={handleSaveConfiguration}
                  >
                    Save Configuration
                  </button>
                </div>

                {configLoadError && (
                  <p className={styles.contextWarning} role="alert">⚠ {configLoadError}</p>
                )}

                <div className={styles.contextGrid}>
                  <label className={styles.contextField}>
                    <span>Client</span>
                    <input
                      type="text"
                      value={documentContext.client}
                      onChange={(e) => handleContextChange('client', e.target.value)}
                    />
                  </label>
                  <label className={styles.contextField}>
                    <span>Contract</span>
                    <input
                      type="text"
                      value={documentContext.contract}
                      onChange={(e) => handleContextChange('contract', e.target.value)}
                    />
                  </label>
                  <label className={styles.contextField}>
                    <span>Contract ID/SoW</span>
                    <input
                      type="text"
                      value={documentContext.sow}
                      onChange={(e) => handleContextChange('sow', e.target.value)}
                    />
                  </label>
                  <label className={styles.contextField}>
                    <span>Project</span>
                    <input
                      type="text"
                      value={documentContext.project}
                      onChange={(e) => handleContextChange('project', e.target.value)}
                    />
                  </label>
                  <label className={styles.contextField}>
                    <span>Sprint</span>
                    <input
                      type="text"
                      value={documentContext.sprint}
                      onChange={(e) => handleContextChange('sprint', e.target.value)}
                    />
                  </label>
                  <label className={styles.contextField}>
                    <span>Release Date</span>
                    <input
                      type="date"
                      value={documentContext.releaseDate}
                      onChange={(e) => handleContextChange('releaseDate', e.target.value)}
                    />
                  </label>
                </div>

              </section>

              <div className={styles.dropZoneWrapper}>
                <DropZone
                  onFilesSelected={handleFilesSelected}
                  disabled={isProcessing}
                />
              </div>

              {/* Features list */}
              <section
                className={styles.featureGrid}
                aria-labelledby="features-heading"
              >
                <h3 id="features-heading" className="sr-only">Supported components</h3>
                {[
                  { icon: '🗃️', label: 'Tables & Columns'         },
                  { icon: '📋', label: 'Forms & Views'             },
                  { icon: '⚙️', label: 'Power Automate Flows'      },
                  { icon: '🔄', label: 'Classic Workflows & BPFs'  },
                  { icon: '🎨', label: 'Canvas & Model Apps'       },
                  { icon: '🌐', label: 'Web Resources (JS/TS/HTML)'},
                  { icon: '🔒', label: 'Security Roles & CLS'      },
                  { icon: '🔌', label: 'Plugins & Plugin Steps'    },
                  { icon: '🔗', label: 'Connection References'      },
                  { icon: '⚙️', label: 'Environment Variables'     },
                  { icon: '📊', label: 'Reports & Dashboards'      },
                  { icon: '📈', label: 'Mermaid Diagrams'          },
                ].map(({ icon, label }) => (
                  <div key={label} className={styles.featureCard}>
                    <span className={styles.featureIcon} aria-hidden="true">{icon}</span>
                    <span className={styles.featureLabel}>{label}</span>
                  </div>
                ))}
              </section>
            </section>
          )}

          {/* Processing state */}
          {isProcessing && (
            <section
              className={styles.processingSection}
              aria-labelledby="processing-heading"
              aria-busy="true"
            >
              <h2 id="processing-heading" className={styles.processingHeading}>
                Processing solutions…
              </h2>
              <div className={styles.progressList}>
                {processing.map((entry) => (
                  <div key={entry.fileName} className={styles.progressItem}>
                    {entry.error ? (
                      <div
                        role="alert"
                        className={styles.progressError}
                      >
                        <span aria-hidden="true">❌</span> {entry.fileName}: {entry.error}
                      </div>
                    ) : (
                      <ProgressBar
                        value={entry.progress}
                        label={entry.fileName}
                        showLabel
                      />
                    )}
                  </div>
                ))}
              </div>

              {/* Show drop zone to allow additional files during processing */}
            </section>
          )}

          {/* Post-processing errors (when processing done but some failed) */}
          {!isProcessing && processing.some((e) => e.error) && (
            <div
              role="alert"
              className={styles.errorSummary}
            >
              {processing
                .filter((e) => e.error)
                .map((e) => (
                  <p key={e.fileName} className={styles.errorLine}>
                    ❌ <strong>{e.fileName}</strong>: {e.error}
                  </p>
                ))}
            </div>
          )}

          {/* Documentation viewer */}
          {hasResults && activeResult && !isProcessing && (
            <div className={styles.viewerWrapper}>
              {isViewerLoading ? (
                <div className={styles.viewerLoadingState} role="status" aria-live="polite" aria-busy="true">
                  <div className={styles.viewerLoadingCard}>
                    <span className={styles.viewerLoadingSpinner} aria-hidden="true" />
                    <p className={styles.viewerLoadingText}>Please Wait - rendering markdown document...</p>
                  </div>
                </div>
              ) : (
                <MarkdownViewer
                  markdown={activeResult.markdown}
                  title={activeResult.solution.metadata.displayName || activeResult.solution.metadata.uniqueName}
                  onExport={handleExport}
                />
              )}
            </div>
          )}

          {/* Additional file drop zone when results exist */}
          {hasResults && !isProcessing && (
            <details className={styles.addMoreDetails}>
              <summary className={styles.addMoreSummary}>
                ＋ Add more solution files
              </summary>
              <div className={styles.addMoreBody}>
                <DropZone
                  onFilesSelected={handleFilesSelected}
                  disabled={isProcessing}
                />
              </div>
            </details>
          )}
        </main>
      </div>

      {/* ── Footer ────────────────────────────────────────────────────── */}
      <footer className={styles.footer} role="contentinfo">
        <p>
          PP-MD: Power Platform Documentation Generator. Created by Mike Hartley - Hart of the Midlands.
          All processing happens locally on your PC; your solution data never leaves your machine.
        </p>
      </footer>
    </div>
  );
}
