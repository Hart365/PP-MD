/**
 * @file DropZone.tsx
 * @description Accessible drag-and-drop / click-to-browse file input component.
 *
 * Accessibility features:
 *  - Role: the outer div is keyboard-focusable and announces as a "region" with
 *    an aria-label.  The hidden <input> receives actual file selections.
 *  - When drag is active, aria-live="polite" announces the state change.
 *  - Each selected file can be removed via a labelled button.
 *  - Supports multiple ZIP files simultaneously.
 *  - File type validation with clear error messaging (aria-live="assertive").
 */

import { useRef, useState, useCallback, type DragEvent, type ChangeEvent } from 'react';
import styles from './DropZone.module.css';

// ---------------------------------------------------------------------------
// Types
// ---------------------------------------------------------------------------

export interface DropZoneProps {
  /** Called when the user has selected/dropped valid ZIP files */
  onFilesSelected: (files: File[]) => void;
  /** Whether the component is currently disabled (e.g. during processing) */
  disabled?: boolean;
}

// ---------------------------------------------------------------------------
// Validation helpers
// ---------------------------------------------------------------------------

/** Only ZIP archives are accepted (MIME and extension check). */
function isValidZip(file: File): boolean {
  return (
    file.type === 'application/zip' ||
    file.type === 'application/x-zip-compressed' ||
    file.name.toLowerCase().endsWith('.zip')
  );
}

/**
 * Formats a file size in bytes to a human-readable string.
 *
 * @param bytes - Raw byte count
 */
function humanSize(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(2)} MB`;
}

function sortFilesByName(files: File[]): File[] {
  return [...files].sort((left, right) => left.name.localeCompare(
    right.name,
    undefined,
    { numeric: true, sensitivity: 'base' },
  ));
}

// ---------------------------------------------------------------------------
// Component
// ---------------------------------------------------------------------------

/**
 * DropZone component — lets users drag-and-drop or browse for Power Platform
 * solution ZIP files.
 */
export function DropZone({ onFilesSelected, disabled = false }: DropZoneProps) {
  /** Files currently queued in the list (not yet processed) */
  const [queuedFiles, setQueuedFiles] = useState<File[]>([]);
  /** Whether a drag operation is currently over the zone */
  const [isDragActive, setIsDragActive] = useState(false);
  /** Validation error message */
  const [error, setError] = useState<string>('');
  /** The hidden file <input> element reference */
  const inputRef = useRef<HTMLInputElement>(null);

  /**
   * Filters a FileList/array for valid ZIPs and updates state.
   *
   * @param rawFiles - Array of files from drag event or input change
   */
  const handleFiles = useCallback(
    (rawFiles: File[]) => {
      const valid   = rawFiles.filter(isValidZip);
      const invalid = rawFiles.filter((f) => !isValidZip(f));

      if (invalid.length > 0) {
        setError(`${invalid.length} file(s) ignored — only .zip files are accepted.`);
      } else {
        setError('');
      }

      if (valid.length === 0) return;

      // Deduplicate by name
      setQueuedFiles((prev) => {
        const merged = [...prev];
        valid.forEach((f) => {
          if (!merged.some((x) => x.name.toLowerCase() === f.name.toLowerCase())) merged.push(f);
        });
        return sortFilesByName(merged);
      });
    },
    [],
  );

  // ── Drag events ────────────────────────────────────────────────────────────

  const handleDragOver = useCallback((e: DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    e.stopPropagation();
    if (!disabled) setIsDragActive(true);
  }, [disabled]);

  const handleDragLeave = useCallback((e: DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragActive(false);
  }, []);

  const handleDrop = useCallback(
    (e: DragEvent<HTMLDivElement>) => {
      e.preventDefault();
      e.stopPropagation();
      setIsDragActive(false);
      if (disabled) return;
      const files = Array.from(e.dataTransfer.files);
      handleFiles(files);
    },
    [disabled, handleFiles],
  );

  // ── Input change ───────────────────────────────────────────────────────────

  const handleInputChange = useCallback(
    (e: ChangeEvent<HTMLInputElement>) => {
      const files = Array.from(e.target.files ?? []);
      handleFiles(files);
      // Reset input so the same file can be re-selected after removal
      e.target.value = '';
    },
    [handleFiles],
  );

  // ── Keyboard activation of the browse button ────────────────────────────

  const handleZoneKeyDown = useCallback(
    (e: React.KeyboardEvent<HTMLDivElement>) => {
      if ((e.key === 'Enter' || e.key === ' ') && !disabled) {
        e.preventDefault();
        inputRef.current?.click();
      }
    },
    [disabled],
  );

  // ── Remove a single file from the queue ────────────────────────────────

  const removeFile = useCallback((name: string) => {
    setQueuedFiles((prev) => prev.filter((f) => f.name !== name));
  }, []);

  // ── Process queued files ──────────────────────────────────────────────────

  const handleProcess = useCallback(() => {
    if (queuedFiles.length === 0) return;
    onFilesSelected(sortFilesByName(queuedFiles));
  }, [queuedFiles, onFilesSelected]);

  // ── Render ─────────────────────────────────────────────────────────────────

  const zoneClass = [
    styles.dropzone,
    isDragActive ? styles.active : '',
    disabled      ? styles.disabled : '',
  ].filter(Boolean).join(' ');

  return (
    <div>
      {/*
       * The drop zone container.
       * role="button" + tabIndex=0 makes it keyboard-accessible
       * aria-describedby points to the subtitle for richer context
       */}
      <div
        className={zoneClass}
        role="button"
        tabIndex={disabled ? -1 : 0}
        aria-label="Drop zone: drag and drop Power Platform solution ZIP files here, or press Enter to browse"
        aria-disabled={disabled}
        aria-describedby="dz-subtitle"
        onDragOver={handleDragOver}
        onDragLeave={handleDragLeave}
        onDrop={handleDrop}
        onKeyDown={handleZoneKeyDown}
        onClick={() => !disabled && inputRef.current?.click()}
      >
        {/* aria-live region announces drag state to screen readers */}
        <div aria-live="polite" className="sr-only">
          {isDragActive ? 'Drop files to add them' : ''}
        </div>

        {/* Icon — presentational */}
        <div className={styles.icon} aria-hidden="true">
          {isDragActive ? '📂' : '📁'}
        </div>

        <p className={styles.title}>
          {isDragActive ? 'Release to add files' : 'Drag & drop solution ZIP files here'}
        </p>
        <p id="dz-subtitle" className={styles.subtitle}>
          Accepts one or more Power Platform solution <code>.zip</code> archives.
          You can also click or press Enter to browse.
        </p>

        <button
          type="button"
          className={styles.browseBtn}
          disabled={disabled}
          aria-label="Browse for solution ZIP files"
          tabIndex={-1} /* Parent div handles keyboard; avoid double tab stop */
          onClick={(e) => {
            e.stopPropagation();
            inputRef.current?.click();
          }}
        >
          Browse files
        </button>

        {/* Hidden file input */}
        <input
          ref={inputRef}
          type="file"
          accept=".zip,application/zip,application/x-zip-compressed"
          multiple
          disabled={disabled}
          aria-hidden="true"
          tabIndex={-1}
          style={{ display: 'none' }}
          onChange={handleInputChange}
        />
      </div>

      {/* Error message — assertive for immediate announcement */}
      {error && (
        <p
          role="alert"
          aria-live="assertive"
          style={{
            marginTop: '0.5rem',
            color: 'var(--color-error)',
            fontSize: '0.875rem',
          }}
        >
          ⚠️ {error}
        </p>
      )}

      {/* Queued files list */}
      {queuedFiles.length > 0 && (
        <div style={{ marginTop: '1rem' }}>
          <p
            style={{ fontSize: '0.875rem', color: 'var(--color-text-secondary)', marginBottom: '0.5rem' }}
          >
            {queuedFiles.length} file{queuedFiles.length > 1 ? 's' : ''} ready to process:
          </p>
          <ul className={styles.fileList} aria-label="Files queued for processing">
            {queuedFiles.map((file) => (
              <li key={file.name} className={styles.fileItem}>
                <span className={styles.fileIcon} aria-hidden="true">🗜️</span>
                <span className={styles.fileName} title={file.name}>
                  {file.name}
                </span>
                <span
                  style={{ color: 'var(--color-text-muted)', fontSize: '0.8125rem' }}
                  aria-label={`Size: ${humanSize(file.size)}`}
                >
                  {humanSize(file.size)}
                </span>
                <button
                  type="button"
                  className={styles.removeBtn}
                  aria-label={`Remove ${file.name} from the queue`}
                  onClick={() => removeFile(file.name)}
                >
                  ✕
                </button>
              </li>
            ))}
          </ul>

          <button
            type="button"
            onClick={handleProcess}
            disabled={disabled}
            style={{
              marginTop: '1rem',
              padding: '0.6rem 1.5rem',
              background: 'var(--color-accent)',
              color: 'var(--color-text-inverse)',
              border: 'none',
              borderRadius: 'var(--border-radius-md)',
              fontSize: '1rem',
              fontWeight: 600,
              cursor: disabled ? 'not-allowed' : 'pointer',
              opacity: disabled ? 0.6 : 1,
              transition: 'background var(--transition-fast)',
            }}
            aria-label={`Generate documentation for ${queuedFiles.length} file${queuedFiles.length > 1 ? 's' : ''}`}
          >
            Generate Documentation
          </button>
        </div>
      )}
    </div>
  );
}
