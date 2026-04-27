/**
 * Vite configuration for PP-MD (Power Platform Markdown Documentation Generator).
 * Uses the official Vite React plugin for JSX/TSX transform support.
 */
import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';

export default defineConfig(({ mode }) => {
  const enableSourceMaps = mode !== 'production';

  return {
    // Relative asset paths are required when loading index.html from file:// in Electron.
    base: './',
    plugins: [react()],
    build: {
      // Output to the standard dist folder
      outDir: 'dist',
      // Keep production bundles compact; source maps stay available in non-prod builds.
      sourcemap: enableSourceMaps,
      // Mermaid and diagram libraries produce large chunks; keep build noise manageable.
      chunkSizeWarningLimit: 2000,
    },
  };
});
