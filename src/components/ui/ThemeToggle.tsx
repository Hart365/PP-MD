/**
 * @file ThemeToggle.tsx
 * @description Accessible dark/light mode toggle button.
 *
 * WCAG compliance:
 *  - 1.4.11 Non-text Contrast: The toggle button border meets 3:1 against bg.
 *  - 4.1.2 Name, Role, Value: Uses role="switch" to convey on/off state.
 *  - 2.4.7 Focus Visible: Inherits global :focus-visible ring.
 *  - 1.3.3 Sensory Characteristics: State described in text, not icon alone.
 */

import { useTheme } from '../../context/ThemeContext';
import styles from './ThemeToggle.module.css';

/**
 * A toggle button that switches between light and dark themes.
 * Intended for placement in the application header.
 */
export function ThemeToggle() {
  const { theme, toggleTheme } = useTheme();
  const isDark = theme === 'dark';

  return (
    <button
      type="button"
      role="switch"
      aria-checked={isDark}
      aria-label={`Switch to ${isDark ? 'light' : 'dark'} theme`}
      title={`Switch to ${isDark ? 'light' : 'dark'} theme (currently ${theme})`}
      className={styles.toggle}
      onClick={toggleTheme}
    >
      {/* Icon area — aria-hidden because the label carries the meaning */}
      <span className={styles.icon} aria-hidden="true">
        {isDark ? '☀️' : '🌙'}
      </span>
      {/* Visible text so intent is communicated without icon knowledge */}
      <span className={styles.label}>
        {isDark ? 'Light mode' : 'Dark mode'}
      </span>
    </button>
  );
}
