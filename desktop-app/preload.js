const { contextBridge, ipcRenderer } = require('electron');

const invokeChannels = new Set([
  'agent-general-ask',
  'agent-llm-diagnostics',
  'agent-plan',
  'bootstrap-local-ai',
  'can-access-feedback-log',
  'cancel-active-operations',
  'check-access-policy',
  'check-for-updates',
  'check-litera-installed',
  'check-version-enforcement',
  'clear-history',
  'clear-session-context',
  'collate-documents',
  'complete-2fa-login',
  'create-zip',
  'create-zip-files',
  'delete-api-key',
  'detect-tasks',
  'dismiss-update-warning',
  'download-update',
  'email-login',
  'execute-command',
  'export-history-csv',
  'generate-packet-shell',
  'generate-punchlist',
  'get-2fa-status',
  'get-access-policy-config',
  'get-api-key',
  'get-api-key-value',
  'get-ai-provider',
  'get-app-version',
  'get-deployment-security-status',
  'get-feedback',
  'get-firebase-config',
  'get-firebase-telemetry-status',
  'get-history',
  'get-local-ai-profiles',
  'get-local-ai-status',
  'get-managed-ai-status',
  'get-security-question',
  'get-setting',
  'get-telemetry-ingest-config',
  'get-usage-stats',
  'get-user-api-key',
  'get-user-by-id',
  'get-user-stats',
  'install-update',
  'list-outlook-folders',
  'load-outlook-emails',
  'log-usage',
  'nl-search-emails',
  'open-analytics-storage',
  'open-external-url',
  'open-feedback-log',
  'open-file',
  'open-folder',
  'parse-checklist',
  'parse-email-csv',
  'parse-incumbency',
  'preflight-signature-packets',
  'process-execution-version',
  'process-folder',
  'process-sigblocks',
  'quit-app',
  'redline-documents',
  'request-verification-code',
  'reset-password',
  'save-access-policy-config',
  'save-file',
  'save-firebase-config',
  'save-setting',
  'save-telemetry-ingest-config',
  'save-zip',
  'select-documents-multiple',
  'select-docx',
  'select-docx-multiple',
  'select-file',
  'select-files',
  'select-folder',
  'select-pdf',
  'select-pdfs-multiple',
  'select-spreadsheet',
  'set-api-key',
  'submit-feedback',
  'test-access-policy-config',
  'test-api-key',
  'test-telemetry-ingest-config',
  'update-checklist',
  'update-user-email',
  'verify-code'
]);

const receiveChannels = new Set([
  'agent-execution-progress',
  'checklist-progress',
  'collate-progress',
  'email-progress',
  'local-ai-progress',
  'navigate-tab',
  'nl-search-progress',
  'progress',
  'punchlist-progress',
  'redline-progress',
  'shell-progress',
  'sigblocks-progress',
  'task-detect-progress',
  'trigger-email-search',
  'trigger-redline',
  'update-available',
  'update-downloaded',
  'update-error',
  'update-not-available',
  'update-progress'
]);

contextBridge.exposeInMainWorld('emmaElectron', {
  invoke(channel, ...args) {
    if (!invokeChannels.has(channel)) {
      return Promise.reject(new Error(`Blocked IPC invoke for channel '${channel}'.`));
    }
    return ipcRenderer.invoke(channel, ...args);
  },
  on(channel, listener) {
    if (!receiveChannels.has(channel)) {
      throw new Error(`Blocked IPC listener for channel '${channel}'.`);
    }
    if (typeof listener !== 'function') {
      throw new Error('IPC listener must be a function.');
    }
    const wrapped = (_event, ...args) => listener(undefined, ...args);
    ipcRenderer.on(channel, wrapped);
    return () => ipcRenderer.removeListener(channel, wrapped);
  },
  openExternal(url) {
    return ipcRenderer.invoke('open-external-url', url);
  }
});
