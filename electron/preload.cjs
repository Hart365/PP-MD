/**
 * @file preload.cjs
 * @description Preload script that exposes safe IPC channels to the renderer process.
 * 
 * This runs in a context that has access to both Node.js APIs and the renderer process,
 * allowing us to safely expose IPC methods while maintaining security through context isolation.
 */

const { contextBridge, ipcRenderer } = require('electron');

// Expose only the IPC methods we need for the update checker
contextBridge.exposeInMainWorld('electron', {
  invoke: (channel, ...args) => ipcRenderer.invoke(channel, ...args),
});
