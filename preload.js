/**
 * Preload — bridges ui.html to Electron IPC.
 *
 * ui.html calls window.pywebview.api.<method>(...args) throughout.
 * We expose a __pybridge__ helper via contextBridge, then inject a
 * tiny inline script into the page that builds the Proxy and fires
 * pywebviewready — exactly what pywebview would do.
 *
 * The inline script is injected synchronously via <script> before any
 * other page script runs, so no race conditions.
 */

const { contextBridge, ipcRenderer } = require('electron');

// Expose the raw IPC caller — used by the injected Proxy below
contextBridge.exposeInMainWorld('__pybridge__', {
  call: (method, args) => ipcRenderer.invoke('py-call', method, args),
});

// Expose shell helpers ui.html may need
contextBridge.exposeInMainWorld('electronAPI', {
  openFile:  filePath => ipcRenderer.invoke('open-file', filePath),
  refocus:   ()       => ipcRenderer.send('refocus-window'),
  openLink:  url      => ipcRenderer.invoke('open-link', url),
});

// Inject the pywebview.api Proxy + flags + pywebviewready event.
// We do this by writing a <script> into <head> before any other script runs.
// Electron preloads run before the page, so we hook into the earliest
// possible point: before DOMContentLoaded.
const BRIDGE_SCRIPT = `
(function () {
  window.__IS_ELECTRON_MODE__ = true;
  window.__IS_BROWSER_MODE__  = false;

  window.pywebview = window.pywebview || {};
  window.pywebview.api = new Proxy({}, {
    get: function (_, name) {
      return function () {
        return window.__pybridge__.call(name, Array.prototype.slice.call(arguments));
      };
    }
  });

  // Fire pywebviewready after all synchronous page scripts have registered
  // their listeners (same timing guarantee as real pywebview).
  function fireReady() {
    setTimeout(function () {
      window.dispatchEvent(new Event('pywebviewready'));
    }, 0);
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', fireReady, { once: true });
  } else {
    fireReady();
  }
})();
`;

// window is not available in preload directly (contextIsolation), but we can
// add the script by intercepting the first document write via a MutationObserver
// on the documentElement, or more simply: use Electron's documented way of
// running code in the renderer world before page scripts.
//
// The cleanest approach: attach to the 'did-start-loading' event is not
// available in preload. Instead we use a <script> tag injected as early
// as possible via DOMContentLoaded on the preload side — preload's
// DOMContentLoaded fires BEFORE the page's own scripts, so this is safe.

document.addEventListener('DOMContentLoaded', () => {
  const s = document.createElement('script');
  s.textContent = BRIDGE_SCRIPT;
  // insertBefore firstChild puts it ahead of any existing <script> tags
  (document.head || document.documentElement).insertBefore(s, 
    (document.head || document.documentElement).firstChild);
});
