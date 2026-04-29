/**
 * @file vite-env.d.ts
 * @description Vite client type declarations — exposes import.meta.env and
 * other Vite-specific globals to TypeScript.
 */

/// <reference types="vite/client" />

/**
 * Electron IPC API exposed to renderer process
 */
interface ElectronAPI {
  invoke: (channel: string, ...args: any[]) => Promise<any>;
}

/**
 * Extend the global window object to include Electron and app-specific properties
 */
declare global {
  interface Window {
    electron?: ElectronAPI;
    __PPMD_VERSION__?: string;
  }
}
