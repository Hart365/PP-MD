/**
 * @file ProgressBar.tsx
 * @description Accessible progress bar with ARIA progressbar role.
 *
 * WCAG compliance:
 *  - 4.1.2 Name, Role, Value: Uses role="progressbar", aria-valuenow,
 *    aria-valuemin, aria-valuemax, and aria-label.
 *  - 1.4.3 Contrast: The filled track colour meets 4.5:1 against the
 *    progress track background.
 *  - 2.3.3 Animation: Respects prefers-reduced-motion via CSS.
 */

import styles from './ProgressBar.module.css';

export interface ProgressBarProps {
  /** Progress value from 0 to 100 */
  value: number;
  /** Descriptive label for screen readers */
  label?: string;
  /** Whether to show the percentage text visually */
  showLabel?: boolean;
}

/**
 * An accessible progress bar component.
 *
 * @param value     - Current progress (0–100)
 * @param label     - Accessible label (used as aria-label and visible caption)
 * @param showLabel - Whether to show the percentage number
 */
export function ProgressBar({ value, label = 'Processing', showLabel = true }: ProgressBarProps) {
  const clamped = Math.min(100, Math.max(0, value));

  return (
    <div className={styles.wrapper}>
      {/* Visible label above the bar */}
      {showLabel && (
        <div className={styles.labelRow}>
          <span className={styles.labelText}>{label}</span>
          <span className={styles.percent} aria-hidden="true">{Math.round(clamped)}%</span>
        </div>
      )}

      {/*
       * The progressbar role element.
       * aria-label duplicates the visible caption for AT-only contexts.
       * aria-valuetext provides a richer textual description.
       */}
      <div
        role="progressbar"
        aria-label={label}
        aria-valuenow={Math.round(clamped)}
        aria-valuemin={0}
        aria-valuemax={100}
        aria-valuetext={`${Math.round(clamped)} percent complete`}
        className={styles.track}
      >
        <div
          className={styles.fill}
          style={{ width: `${clamped}%` }}
        />
      </div>
    </div>
  );
}
