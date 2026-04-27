/**
 * @file ThemeContext.tsx
 * @description React context for managing light/dark theme across the app.
 * Reads the user's OS preference via the `prefers-color-scheme` media query
 * and persists the choice to localStorage.  Applies the chosen theme as a
 * `data-theme` attribute on <html> so that CSS custom properties can switch
 * all colours in one place.
 *
 * WCAG 2.2 compliance notes:
 *  - 1.4.3 Contrast (Minimum): all colour tokens defined in global.css meet
 *    or exceed the 4.5:1 ratio for normal text and 3:1 for large text.
 *  - 1.4.11 Non-text Contrast: interactive component borders use tokens that
 *    meet 3:1 against their backgrounds.
 */

import {
  createContext,
  useContext,
  useEffect,
  useState,
  useCallback,
  type ReactNode,
} from 'react';

// ---------------------------------------------------------------------------
// Types
// ---------------------------------------------------------------------------

/** The two supported colour schemes. */
export type Theme = 'light' | 'dark';

/** Shape of the context value exposed to consumers. */
export interface ThemeContextValue {
  /** Currently active theme. */
  theme: Theme;
  /**
   * Toggle between light and dark.
   * Safe to call from ARIA button handlers.
   */
  toggleTheme: () => void;
  /** Explicitly set the theme. */
  setTheme: (theme: Theme) => void;
}

// ---------------------------------------------------------------------------
// Context
// ---------------------------------------------------------------------------

const ThemeContext = createContext<ThemeContextValue | undefined>(undefined);

// ---------------------------------------------------------------------------
// Storage key & helpers
// ---------------------------------------------------------------------------

/** localStorage key under which the user's theme preference is stored. */
const STORAGE_KEY = 'pp-md-theme';

/**
 * Reads the user's stored preference or falls back to the OS setting.
 *
 * @returns The resolved {@link Theme}
 */
function getInitialTheme(): Theme {
  try {
    const stored = localStorage.getItem(STORAGE_KEY);
    if (stored === 'light' || stored === 'dark') return stored;
  } catch {
    // localStorage may be blocked (e.g. private browsing with strict settings)
  }
  // Fall back to OS preference
  return window.matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light';
}

/**
 * Applies the theme to the document root and updates the browser chrome
 * `theme-color` meta tag so that device taskbars / address bars reflect
 * the current colour mode.
 *
 * @param theme - The new theme to apply
 */
function applyTheme(theme: Theme): void {
  document.documentElement.setAttribute('data-theme', theme);
  // Update <meta name="theme-color"> to match the background token
  const themeColor = theme === 'dark' ? '#1e1e2e' : '#f8f9fa';
  const metaTag = document.querySelector('meta[name="theme-color"]');
  if (metaTag) metaTag.setAttribute('content', themeColor);
}

// ---------------------------------------------------------------------------
// Provider
// ---------------------------------------------------------------------------

/**
 * ThemeProvider wraps the application and makes the theme context available
 * to all descendant components.
 *
 * @param children - React tree to render inside the provider
 */
export function ThemeProvider({ children }: { children: ReactNode }) {
  const [theme, setThemeState] = useState<Theme>(getInitialTheme);

  // Apply theme to DOM on every change
  useEffect(() => {
    applyTheme(theme);
    try {
      localStorage.setItem(STORAGE_KEY, theme);
    } catch {
      // Ignore storage errors
    }
  }, [theme]);

  // Also listen for OS-level theme changes (e.g. user changes system dark mode)
  useEffect(() => {
    const mediaQuery = window.matchMedia('(prefers-color-scheme: dark)');
    const handleChange = (e: MediaQueryListEvent) => {
      // Only react to OS change if the user has NOT manually set a preference
      try {
        const stored = localStorage.getItem(STORAGE_KEY);
        if (!stored) {
          setThemeState(e.matches ? 'dark' : 'light');
        }
      } catch {
        setThemeState(e.matches ? 'dark' : 'light');
      }
    };
    mediaQuery.addEventListener('change', handleChange);
    return () => mediaQuery.removeEventListener('change', handleChange);
  }, []);

  const setTheme = useCallback((newTheme: Theme) => {
    setThemeState(newTheme);
  }, []);

  const toggleTheme = useCallback(() => {
    setThemeState((prev) => (prev === 'dark' ? 'light' : 'dark'));
  }, []);

  return (
    <ThemeContext.Provider value={{ theme, toggleTheme, setTheme }}>
      {children}
    </ThemeContext.Provider>
  );
}

// ---------------------------------------------------------------------------
// Hook
// ---------------------------------------------------------------------------

/**
 * Custom hook for consuming the theme context.
 * Must be used inside a {@link ThemeProvider}.
 *
 * @returns The current {@link ThemeContextValue}
 * @throws If used outside a ThemeProvider
 */
export function useTheme(): ThemeContextValue {
  const ctx = useContext(ThemeContext);
  if (!ctx) {
    throw new Error('useTheme must be used within a ThemeProvider');
  }
  return ctx;
}
