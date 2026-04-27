/**
 * @file SolutionSidebar.tsx
 * @description Navigation sidebar listing parsed solutions.  When multiple
 * ZIP files have been processed, this sidebar lets the user switch between
 * their generated documentation tabs.
 *
 * WCAG compliance:
 *  - 1.3.1 Info & Relationships: Uses <nav> with aria-label.
 *  - 2.1.1 Keyboard: All items keyboard-navigable.
 *  - 4.1.2 Name, Role, Value: aria-current="page" on the active item.
 *  - 1.4.3 Contrast: Active/hover states use tokens meeting 4.5:1.
 */

import styles from './SolutionSidebar.module.css';
import type { ParsedSolution } from '../types/solution';

export interface SolutionSidebarProps {
  /** All successfully parsed solutions */
  solutions: ParsedSolution[];
  /** Index of the currently visible solution */
  activeIndex: number;
  /** Callback when the user selects a solution */
  onSelect: (index: number) => void;
  /** Callback to clear all solutions and return to the drop zone */
  onReset: () => void;
}

/**
 * Sidebar listing each parsed solution for navigation.
 */
export function SolutionSidebar({
  solutions,
  activeIndex,
  onSelect,
  onReset,
}: SolutionSidebarProps) {
  if (solutions.length === 0) return null;

  return (
    <nav
      className={styles.sidebar}
      aria-label="Solution navigation"
    >
      {/* Header */}
      <div className={styles.sidebarHeader}>
        <span className={styles.sidebarTitle} aria-hidden="true">📦</span>
        <span className={styles.sidebarTitleText}>Solutions</span>
        <button
          type="button"
          className={styles.resetBtn}
          onClick={onReset}
          aria-label="Clear all solutions and start over"
          title="Start over"
        >
          ✕
        </button>
      </div>

      {/* Solution list */}
      <ul className={styles.list} role="list">
        {solutions.map((sol, idx) => {
          const isActive  = idx === activeIndex;
          const itemClass = [styles.item, isActive ? styles.active : ''].filter(Boolean).join(' ');

          return (
            <li key={`${sol.metadata.uniqueName}-${idx}`} className={styles.listItem}>
              <button
                type="button"
                className={itemClass}
                onClick={() => onSelect(idx)}
                aria-current={isActive ? 'page' : undefined}
                aria-label={`View documentation for ${sol.metadata.displayName}`}
              >
                <span className={styles.solIcon} aria-hidden="true">
                  {sol.metadata.isManaged ? '🔒' : '📄'}
                </span>
                <span className={styles.solName}>
                  {sol.metadata.displayName || sol.metadata.uniqueName}
                </span>
                <span className={styles.solVersion} aria-label={`Version ${sol.metadata.version}`}>
                  v{sol.metadata.version}
                </span>
              </button>
            </li>
          );
        })}
      </ul>

      {/* Summary counts */}
      {solutions[activeIndex] && (
        <div className={styles.summaryPanel} aria-label="Active solution component counts">
          <p className={styles.summaryTitle}>Components</p>
          <dl className={styles.summaryList}>
            {[
              { label: 'Tables',      count: solutions[activeIndex].entities.length },
              { label: 'Processes',   count: solutions[activeIndex].processes.length },
              { label: 'Apps',        count: solutions[activeIndex].apps.length },
              { label: 'Web Res.',    count: solutions[activeIndex].webResources.length },
              { label: 'Plugins',     count: solutions[activeIndex].pluginAssemblies.length },
              { label: 'Roles',       count: solutions[activeIndex].securityRoles.length },
              { label: 'Env. Vars.',  count: solutions[activeIndex].environmentVariables.length },
            ]
              .filter((s) => s.count > 0)
              .map(({ label, count }) => (
                <div key={label} className={styles.summaryRow}>
                  <dt className={styles.summaryLabel}>{label}</dt>
                  <dd className={styles.summaryCount}>{count}</dd>
                </div>
              ))}
          </dl>
        </div>
      )}
    </nav>
  );
}
