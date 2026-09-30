import React from 'react'
import ReactDOM from 'react-dom/client'
import App from './App.jsx'
import { registerSW } from 'virtual:pwa-register'

// Without this, an already-installed app/tab can keep running the JS from
// whenever it was first loaded, even after a new version is deployed —
// silently missing new features (like live sync) until a manual hard
// refresh. This checks for updates and lets the person know when one's
// ready.
//
// It used to call window.location.reload() the moment an update was
// found — but browsers commonly re-check a service worker for updates
// whenever a tab regains focus, so that meant switching back to this
// tab could force a reload (and silently discard whatever was open,
// like an in-progress Add Asset form) with zero warning, every single
// time. Now it just shows a small dismissible banner and lets the
// person choose when to refresh, instead of doing it to them.
registerSW({
  immediate: true,
  onNeedRefresh() {
    showUpdateBanner()
  },
})

function showUpdateBanner() {
  if (document.getElementById('pwa-update-banner')) return
  const bar = document.createElement('div')
  bar.id = 'pwa-update-banner'
  bar.style.cssText =
    'position:fixed;left:0;right:0;bottom:0;z-index:99999;background:#3B82F6;' +
    'color:#fff;padding:10px 16px;display:flex;align-items:center;justify-content:center;' +
    'gap:12px;font:600 13px system-ui,-apple-system,sans-serif;box-shadow:0 -2px 10px rgba(0,0,0,0.15);'
  bar.innerHTML =
    '<span>A new version of AssetHub is available.</span>' +
    '<button id="pwa-update-btn" style="background:#fff;color:#3B82F6;border:none;border-radius:6px;' +
    'padding:6px 14px;font-weight:700;font-size:13px;cursor:pointer;">Refresh now</button>' +
    '<button id="pwa-update-dismiss" style="background:none;border:none;color:#fff;opacity:0.85;' +
    'cursor:pointer;font-size:18px;line-height:1;padding:0 4px;">&times;</button>'
  document.body.appendChild(bar)
  document.getElementById('pwa-update-btn').onclick = () => window.location.reload()
  document.getElementById('pwa-update-dismiss').onclick = () => bar.remove()
}

ReactDOM.createRoot(document.getElementById('root')).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>,
)
