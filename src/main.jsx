import React from 'react'
import ReactDOM from 'react-dom/client'
import App from './App.jsx'
import { registerSW } from 'virtual:pwa-register'

// Without this, an already-installed app/tab can keep running the JS from
// whenever it was first loaded, even after a new version is deployed —
// silently missing new features (like live sync) until a manual hard
// refresh. This checks for updates and reloads automatically when found.
registerSW({
  immediate: true,
  onNeedRefresh() {
    window.location.reload()
  },
})

ReactDOM.createRoot(document.getElementById('root')).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>,
)
