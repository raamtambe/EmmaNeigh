const { contextBridge, ipcRenderer } = require('electron');

// Expose protected methods that allow the renderer process to use
// the ipcRenderer without exposing the entire object
contextBridge.exposeInMainWorld('api', {
  // File selection dialogs
  selectFiles: () => ipcRenderer.invoke('select-files'),
  selectFolder: () => ipcRenderer.invoke('select-folder'),
  saveFile: (defaultName) => ipcRenderer.invoke('save-file', defaultName),

  // Processing functions
  processSignaturePackets: (filePaths) =>
    ipcRenderer.invoke('process-signature-packets', filePaths),
  createExecutionVersion: (originalPath, signedPath, insertAfter) =>
    ipcRenderer.invoke('create-execution-version', originalPath, signedPath, insertAfter),

  // File operations
  copyFile: (sourcePath, destPath) =>
    ipcRenderer.invoke('copy-file', sourcePath, destPath),

  // Progress event listener
  onProgress: (callback) => {
    ipcRenderer.on('progress', (event, data) => callback(data));
    // Return cleanup function
    return () => ipcRenderer.removeAllListeners('progress');
  },

  // Platform info
  platform: process.platform,
});
