/**
 * @file main.tsx
 * @description Application entry point.  Mounts the React tree into the
 * #root element defined in index.html.
 *
 * Wraps the application in:
 *  - React.StrictMode for development-time checks
 *  - ThemeProvider for global colour-scheme management
 */

import { StrictMode } from 'react';
import { createRoot } from 'react-dom/client';

import { ThemeProvider } from './context/ThemeContext';
import App              from './App';

// Import global styles first so they form the cascade base
import './assets/global.css';

// ---------------------------------------------------------------------------
// Mount
// ---------------------------------------------------------------------------

const rootElement = document.getElementById('root');

if (!rootElement) {
  /*
   * This should never happen in production because index.html always has #root.
   * Throwing gives a clear developer error rather than a silent blank screen.
   */
  throw new Error(
    'PP-MD: Could not find #root element. Ensure index.html contains <div id="root"></div>.',
  );
}

createRoot(rootElement).render(
  <StrictMode>
    <ThemeProvider>
      <App />
    </ThemeProvider>
  </StrictMode>,
);
