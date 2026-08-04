import React, { useState, useEffect, useCallback, useMemo, useRef } from "react";
import {
  LayoutDashboard, Package, Wrench, MapPin, Tags, Users, User, LogOut,
  Menu, Sun, Moon, Plus, Pencil, Trash2, Download, Upload, X, Search,
  KeyRound, ShieldCheck, AlertTriangle, Info, Clock,
  Bell, Copy, Truck, CheckSquare, Archive, ExternalLink,
  ChevronUp, ChevronDown, ChevronsUpDown, MessageCircle,
} from "lucide-react";
import { PieChart, Pie, Cell, ResponsiveContainer, Legend, Tooltip } from "recharts";
import * as XLSX from "xlsx";
import { fetchOrgData, saveOrgData, subscribeToOrgData } from "./lib/supabase.js";

/* ---------------------------------------------------------
   Helpers
--------------------------------------------------------- */
// Synchronous string hash (cyrb53) — used instead of the Web Crypto API,
// which is unreliable inside sandboxed artifact preview environments.
function sha256(text) {
  const str = String(text ?? "");
  let h1 = 0xdeadbeef ^ 0, h2 = 0x41c6ce57 ^ 0;
  for (let i = 0; i < str.length; i++) {
    const ch = str.charCodeAt(i);
    h1 = Math.imul(h1 ^ ch, 2654435761);
    h2 = Math.imul(h2 ^ ch, 1597334677);
  }
  h1 = Math.imul(h1 ^ (h1 >>> 16), 2246822507) ^ Math.imul(h2 ^ (h2 >>> 13), 3266489909);
  h2 = Math.imul(h2 ^ (h2 >>> 16), 2246822507) ^ Math.imul(h1 ^ (h1 >>> 13), 3266489909);
  return (4294967296 * (2097151 & h2) + (h1 >>> 0)).toString(16);
}
const uid = (p) => `${p}-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`;
const todayISO = () => new Date().toISOString().split("T")[0];

// Generates the next sequential ASTUTE### tag when the Asset Tag field is left
// blank — looks at existing tags matching the ASTUTE prefix and picks max+1.
function nextAutoTag(assets) {
  let max = 0;
  (assets || []).forEach((a) => {
    const m = /^ASTUTE(\d+)$/i.exec(String(a.tag || "").trim());
    if (m) max = Math.max(max, parseInt(m[1], 10));
  });
  return `ASTUTE${String(max + 1).padStart(3, "0")}`;
}

// Appends an entry to the audit log, tagged with who did it and when.
// Keeps only the most recent 300 entries so it doesn't grow forever.
// locationId (when the action relates to a specific asset) lets Regional
// Staff's Recent Activity / Activity Log show only their own country's
// activity — entries with no locationId (category/location/user/backup
// admin actions) are only ever visible to Admins anyway.
function withLog(data, currentUser, message, locationId = null) {
  const entry = {
    id: uid("log"),
    at: new Date().toISOString(),
    userId: currentUser?.id || null,
    userName: currentUser?.name || "Unknown",
    message,
    locationId: locationId || null,
  };
  const auditLog = [entry, ...(data.auditLog || [])].slice(0, 300);
  return { ...data, auditLog };
}

// Marks an asset "Under Repair", remembering its prior status so it can be
// restored once all its open maintenance work is done. No-op if it's
// already Under Repair.
function markUnderRepair(assets, assetId) {
  return assets.map((a) => {
    if (a.id !== assetId || a.status === "Under Repair") return a;
    return { ...a, preRepairStatus: a.status, status: "Under Repair" };
  });
}

// Certain status changes imply the asset is no longer with anyone, so we
// clear the fields that only make sense for something actively assigned —
// Retired/Disposed also drop condition to Poor since it's no longer in
// service at all, while In Stock just means it's back on the shelf.
function applyStatusSideEffects(asset) {
  if (asset.status === "Retired" || asset.status === "Disposed") {
    return { ...asset, assignedTo: "", condition: "Poor", department: "" };
  }
  if (asset.status === "In Stock") {
    return { ...asset, assignedTo: "", department: "" };
  }
  return asset;
}

// After a maintenance entry is closed/removed, checks whether the asset
// still has other open (non-"Done") maintenance entries. If not, restores
// its pre-repair status.
function maybeRestoreStatus(assets, maintenance, assetId) {
  const stillOpen = maintenance.some((m) => m.assetId === assetId && m.status !== "Done");
  if (stillOpen) return assets;
  return assets.map((a) => (a.id === assetId && a.status === "Under Repair"
    ? { ...a, status: a.preRepairStatus || "In Use", preRepairStatus: null }
    : a));
}

// Builds the list of "upcoming attention needed" items shown in the
// notification bell — warranty expiring (IT) and calibration due (Non-IT),
// within the next 30 days or already overdue.
function computeAlerts(assets, scopedLocationId) {
  const list = scopedLocationId ? assets.filter((a) => a.locationId === scopedLocationId) : assets;
  const now = new Date();
  const in30 = new Date(now.getTime() + 30 * 24 * 60 * 60 * 1000);
  const alerts = [];
  list.forEach((a) => {
    if (a.assetType === "IT" && a.warrantyExpiry) {
      const d = new Date(a.warrantyExpiry);
      if (!isNaN(d) && d <= in30) {
        alerts.push({ id: `${a.id}-w`, assetId: a.id, urgent: d < now, label: `${a.tag} — warranty ${d < now ? "expired" : "expiring"} ${a.warrantyExpiry}` });
      }
    }
    if (a.assetType === "Non-IT" && a.requiresCalibration) {
      const checkDate = a.nextCalibrationDate || a.calibrationDate;
      if (checkDate) {
        const d = new Date(checkDate);
        if (!isNaN(d) && d <= in30) {
          alerts.push({ id: `${a.id}-c`, assetId: a.id, urgent: d < now, label: `${a.tag} — calibration ${d < now ? "overdue" : "due"} ${checkDate}` });
        }
      }
    }
  });
  return alerts.sort((a, b) => (b.urgent === a.urgent ? 0 : b.urgent ? 1 : -1));
}


// Triggers a browser download for a built XLSX workbook.
function downloadWorkbook(workbook, filename) {
  const wbout = XLSX.write(workbook, { bookType: "xlsx", type: "array" });
  const blob = new Blob([wbout], { type: "application/octet-stream" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

// Reads a File as an XLSX workbook (async).
function readWorkbookFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      try {
        const wb = XLSX.read(reader.result, { type: "array" });
        resolve(wb);
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = () => reject(new Error("Could not read file."));
    reader.readAsArrayBuffer(file);
  });
}

// Reads every row of a sheet as an array of plain objects (empty cells -> "").
function sheetRows(workbook, sheetName) {
  const sheet = workbook.Sheets[sheetName];
  if (!sheet) return [];
  return XLSX.utils.sheet_to_json(sheet, { defval: "" });
}

// Serializes a JS value that might be an array/object into a JSON string for
// a spreadsheet cell (used for the handful of nested fields like transfer
// history), leaving plain scalars untouched.
function cellify(value) {
  if (value === null || value === undefined) return "";
  if (Array.isArray(value) || typeof value === "object") return JSON.stringify(value);
  return value;
}

// Reverses cellify() — parses a cell back into an array/object if it looks
// like JSON, otherwise returns it as-is (or a fallback default).
function parseCell(value, fallback) {
  if (value === "" || value === null || value === undefined) return fallback;
  if (typeof value === "string" && (value.startsWith("[") || value.startsWith("{"))) {
    try { return JSON.parse(value); } catch { return fallback; }
  }
  return value;
}

const CONDITION_OPTIONS = ["New", "Good", "Fair", "Poor"];
const MAINT_STATUS = ["Not Started", "In Progress", "Done"];

const STATUS_COLORS = {
  "In Stock": "#3B82F6",
  "In Use": "#10B981",
  "Under Repair": "#F59E0B",
  "Retired": "#9CA3AF",
  "Disposed": "#EF4444",
};
const STATUS_OPTIONS = Object.keys(STATUS_COLORS);
const CAT_PALETTE = ["#6366F1", "#10B981", "#F59E0B", "#EC4899", "#06B6D4", "#8B5CF6", "#84CC16", "#F97316"];

// Gives each category a stable color (by its position in the categories list)
// so the same category always shows the same dot color across views.
function categoryColor(categories, categoryId) {
  const idx = categories.findIndex((c) => c.id === categoryId);
  return idx >= 0 ? CAT_PALETTE[idx % CAT_PALETTE.length] : "#9CA3AF";
}

// Known vendor warranty/spec lookup pages, matched against the asset's Brand
// field. Falls back to a general search when the brand isn't recognized.
const WARRANTY_LOOKUP_URLS = {
  lenovo: "https://pcsupport.lenovo.com/us/en/warranty-lookup",
  dell: "https://www.dell.com/support/home/en-us/product-support/servicetag/",
  hp: "https://support.hp.com/us-en/checkwarranty",
  apple: "https://checkcoverage.apple.com/",
};
function warrantyLookupUrl(brand, serial) {
  const key = Object.keys(WARRANTY_LOOKUP_URLS).find((k) => (brand || "").toLowerCase().includes(k));
  if (key) return WARRANTY_LOOKUP_URLS[key];
  const q = [brand, serial, "warranty lookup"].filter(Boolean).join(" ");
  return `https://www.google.com/search?q=${encodeURIComponent(q)}`;
}

function seedData() {
  const locations = [
    { id: "loc-main", name: "Main Office" },
  ];
  const categories = [
    { id: "cat-laptop", name: "Laptops & Desktops", type: "IT", usefulLife: 3 },
    { id: "cat-server", name: "Servers & Storage", type: "IT", usefulLife: 5 },
    { id: "cat-network", name: "Network Equipment", type: "IT", usefulLife: 4 },
    { id: "cat-software", name: "Software & Licenses", type: "IT", usefulLife: 1 },
    { id: "cat-monitor", name: "Monitors & Displays", type: "IT", usefulLife: 4 },
    { id: "cat-furniture", name: "Office Furniture", type: "Non-IT", usefulLife: 7 },
    { id: "cat-tools", name: "Tools & Testing Equipment", type: "Non-IT", usefulLife: 5 },
  ];
  const assets = generateMockAssets(locations, categories);
  return { locations, categories, assets, maintenance: generateMockMaintenance(assets), users: [], auditLog: [] };
}

const SAMPLE_STAFF = [
  "Maria Santos", "Wei Chen", "Li Na", "Tan Wei Ling", "Chan Ka Wai", "Jose Reyes",
  "Grace Tan", "Arman Cruz", "Huang Yi", "Sofia Lim", "Marco Villanueva", "Priya Rao",
  "Daniel Ong", "Faith Aquino", "Kenji Sato", "Emily Wong", "Ryan Dela Cruz", "Zhang Wei",
  "Hannah Goh", "Carlos Mendoza",
];

const IT_MODELS = {
  "cat-laptop": [["Lenovo", "ThinkPad X1 Carbon"], ["Dell", "Latitude 5440"], ["Apple", "MacBook Pro 14"], ["HP", "EliteBook 840"], ["Lenovo", "ThinkPad T14"]],
  "cat-server": [["Dell", "PowerEdge R450"], ["HPE", "ProLiant DL360"], ["Lenovo", "ThinkSystem SR630"]],
  "cat-network": [["Cisco", "Catalyst 1200"], ["Ubiquiti", "UniFi Switch 24"], ["TP-Link", "T2600G-28TS"]],
  "cat-software": [["Microsoft", "Office 365 E3"], ["Adobe", "Creative Cloud"], ["Autodesk", "AutoCAD LT"]],
  "cat-monitor": [["Dell", "P2422H"], ["LG", "27UL850"], ["Samsung", "S24C450"]],
};
const NONIT_MODELS = {
  "cat-furniture": [["Herman Miller", "Aeron Chair"], ["IKEA", "Bekant Desk"], ["Steelcase", "Series 2"]],
  "cat-tools": [["Mitutoyo", "Digital Caliper 500-196"], ["Fluke", "179 Multimeter"], ["Bosch", "GLM 50C Laser Meter"]],
};

function generateMockAssets(locations, categories) {
  const assets = [];
  let counter = 1;
  const totalTarget = 130;
  const itCats = categories.filter((c) => c.type === "IT").map((c) => c.id);
  const nonItCats = categories.filter((c) => c.type === "Non-IT").map((c) => c.id);

  for (let i = 0; i < totalTarget; i++) {
    const loc = locations[i % locations.length];
    const isIT = i % 5 !== 0; // ~80% IT, 20% Non-IT
    const catId = isIT ? itCats[i % itCats.length] : nonItCats[i % nonItCats.length];
    const models = isIT ? IT_MODELS[catId] : NONIT_MODELS[catId];
    const [brand, model] = models[i % models.length];
    const status = STATUS_OPTIONS[i % STATUS_OPTIONS.length];
    const condition = CONDITION_OPTIONS[(i + 1) % CONDITION_OPTIONS.length];
    const assignedTo = status === "In Use" ? SAMPLE_STAFF[i % SAMPLE_STAFF.length] : "";
    const year = 2022 + (i % 4);
    const month = String(1 + (i % 12)).padStart(2, "0");
    const day = String(1 + (i % 27)).padStart(2, "0");
    const purchaseDate = `${year}-${month}-${day}`;
    const cost = isIT ? 400 + (i % 12) * 350 : 80 + (i % 10) * 120;
    const locCode = loc.name.slice(0, 2).toUpperCase();

    const requiresCalibration = !isIT && catId === "cat-tools";
    const calDate = requiresCalibration ? `2026-${String(1 + ((i + 3) % 12)).padStart(2, "0")}-15` : "";
    const nextCalDate = requiresCalibration ? `2027-${String(1 + ((i + 3) % 12)).padStart(2, "0")}-15` : "";

    assets.push({
      id: uid("ast"),
      tag: `AST-${locCode}-${String(counter++).padStart(3, "0")}`,
      name: `${model}`,
      categoryId: catId,
      assetType: isIT ? "IT" : "Non-IT",
      brand, model,
      serial: `SN-${100000 + i * 37}`,
      status, condition,
      locationId: loc.id,
      assignedTo,
      purchaseDate,
      purchaseCost: cost,
      warrantyExpiry: isIT ? `${year + 3}-${month}-${day}` : "",
      requiresCalibration,
      calibrationDate: calDate,
      nextCalibrationDate: nextCalDate,
      notes: status === "Under Repair" ? "Reported issue — pending technician review" : "",
      transferHistory: [],
    });
  }
  return assets;
}

function generateMockMaintenance(assets) {
  const logs = [];
  const descriptions = [
    "Routine inspection and cleaning",
    "Battery replacement",
    "Firmware/software update",
    "Hardware fault diagnosis",
    "Annual calibration check",
    "Screen/display repair",
    "Network connectivity issue",
  ];
  const sample = assets.filter((_, i) => i % 9 === 0).slice(0, 14);
  sample.forEach((a, i) => {
    logs.push({
      id: uid("maint"),
      assetId: a.id,
      description: descriptions[i % descriptions.length],
      cost: 20 + (i % 8) * 35,
      date: a.purchaseDate,
      status: MAINT_STATUS[i % MAINT_STATUS.length],
    });
  });
  return logs;
}

/* ---------------------------------------------------------
   Small UI atoms
--------------------------------------------------------- */
function Badge({ children, color }) {
  return (
    <span
      className="badge"
      style={{
        background: `${color}1a`,
        color,
        border: `1px solid ${color}40`,
      }}
    >
      {children}
    </span>
  );
}

function IconBtn({ icon: Icon, onClick, title, danger }) {
  return (
    <button className={`icon-btn ${danger ? "danger" : ""}`} onClick={onClick} title={title} type="button">
      <Icon size={15} />
    </button>
  );
}

function Modal({ title, onClose, children, width = 480 }) {
  return (
    <div className="modal-overlay" onClick={onClose}>
      <div className="modal" style={{ maxWidth: width }} onClick={(e) => e.stopPropagation()}>
        <div className="modal-head">
          <h3>{title}</h3>
          <button className="icon-btn" onClick={onClose}><X size={16} /></button>
        </div>
        <div className="modal-body">{children}</div>
      </div>
    </div>
  );
}

function ConfirmDialog({ message, onCancel, onConfirm, confirmLabel = "Delete", maxWidth }) {
  return (
    <div className="modal-overlay" onClick={onCancel}>
      <div className="modal confirm" style={maxWidth ? { maxWidth } : undefined} onClick={(e) => e.stopPropagation()}>
        <div className="confirm-icon"><AlertTriangle size={20} /></div>
        <p>{message}</p>
        <div className="confirm-actions">
          <button className="btn ghost" onClick={onCancel}>Cancel</button>
          <button className="btn danger" onClick={onConfirm}>{confirmLabel}</button>
        </div>
      </div>
    </div>
  );
}

function TypeToConfirmDialog({ title, message, confirmWord, value, onChange, onCancel, onConfirm }) {
  const matches = value.trim() === confirmWord;
  return (
    <div className="modal-overlay" onClick={onCancel}>
      <div className="modal confirm" onClick={(e) => e.stopPropagation()}>
        <div className="confirm-icon"><AlertTriangle size={20} /></div>
        {title && <h3 style={{ marginBottom: 6 }}>{title}</h3>}
        <p>{message}</p>
        <p style={{ fontSize: 12.5, color: "var(--text-soft)", marginTop: 10 }}>
          Type <strong>{confirmWord}</strong> below to confirm.
        </p>
        <input
          type="text"
          value={value}
          onChange={(e) => onChange(e.target.value)}
          onKeyDown={(e) => { if (e.key === "Enter" && matches) onConfirm(); }}
          placeholder={confirmWord}
          autoFocus
          style={{
            width: "100%", marginTop: 10, padding: "8px 10px", borderRadius: 8,
            border: "1px solid var(--border)", background: "var(--bg)", color: "var(--text)",
            fontFamily: "inherit", fontSize: 13.5, textAlign: "center",
          }}
        />
        <div className="confirm-actions">
          <button className="btn ghost" onClick={onCancel}>Cancel</button>
          <button className="btn danger" onClick={onConfirm} disabled={!matches}>Delete</button>
        </div>
      </div>
    </div>
  );
}

function Field({ label, children, required, className, action }) {
  return (
    <label className={`field ${className || ""}`}>
      <span className="field-label-row">
        <span>{label}{required && <span className="required-mark">*</span>}</span>
        {action}
      </span>
      {children}
    </label>
  );
}

/* ---------------------------------------------------------
   Login Screen
--------------------------------------------------------- */
function LoginScreen({ users, onLogin }) {
  const [username, setUsername] = useState("");
  const [password, setPassword] = useState("");
  const [error, setError] = useState("");
  const [busy, setBusy] = useState(false);

  const submit = async (e) => {
    e.preventDefault();
    setError("");
    setBusy(true);
    try {
      if (!username.trim() || !password) {
        setError("Enter both username and password.");
        return;
      }
      const hash = sha256(password);
      const list = Array.isArray(users) ? users : [];
      const found = list.find(
        (u) => (u.username || "").toLowerCase() === username.trim().toLowerCase() && u.passwordHash === hash
      );
      if (!found) {
        setError("Incorrect username or password.");
        return;
      }
      onLogin(found);
    } catch (err) {
      setError("Something went wrong signing in. Please try again.");
    } finally {
      setBusy(false);
    }
  };

  return (
    <div className="login-wrap">
      <div className="login-card">
        <div className="login-logo">
          <span className="brand-badge"><ShieldCheck size={18} /></span>
          <span>AssetHub</span>
        </div>
        <p className="login-sub">Sign in to manage IT &amp; facility assets.</p>
        <div onKeyDown={(e) => { if (e.key === "Enter") submit(e); }}>
          <Field label="Username">
            <input value={username} onChange={(e) => setUsername(e.target.value)} autoFocus placeholder="e.g. admin" />
          </Field>
          <Field label="Password">
            <input type="password" value={password} onChange={(e) => setPassword(e.target.value)} placeholder="••••••••" />
          </Field>
          {error && <div className="form-error">{error}</div>}
          <button className="btn primary full" disabled={busy} type="button" onClick={submit}>
            {busy ? "Signing in…" : "Sign In"}
          </button>
        </div>
        <div className="login-hint">
          Demo admin login — <strong>admin</strong> / <strong>admin123</strong>
        </div>
      </div>
    </div>
  );
}

/* ---------------------------------------------------------
   Main App
--------------------------------------------------------- */
export default function App() {
  const [loaded, setLoaded] = useState(false);
  const [theme, setTheme] = useState("light");
  const [data, setData] = useState(null);
  const [currentUser, setCurrentUser] = useState(null);
  const [view, setView] = useState("dashboard");
  const [sidebarOpen, setSidebarOpen] = useState(true);
  const [toast, setToast] = useState(null); // { message, phase: 'in' | 'out' }
  const [connectionError, setConnectionError] = useState(null);
  const [lastSynced, setLastSynced] = useState(null);
  const [focusAssetId, setFocusAssetId] = useState(null);
  const dataRef = React.useRef(null);
  useEffect(() => { dataRef.current = data; }, [data]);

  // Clicking a comment notification in the bell jumps straight to that
  // asset's detail view, wherever the user currently is in the app.
  const openAssetFromNotif = useCallback((assetId) => {
    setView("assets");
    setFocusAssetId(assetId);
  }, []);

  const showToast = useCallback((msg) => {
    setToast({ message: msg, phase: "in" });
    window.clearTimeout(showToast._hideTimer);
    window.clearTimeout(showToast._clearTimer);
    showToast._hideTimer = window.setTimeout(() => {
      setToast((t) => (t ? { ...t, phase: "out" } : t));
      showToast._clearTimer = window.setTimeout(() => setToast(null), 220);
    }, 1800);
  }, []);

  const loadFromCloud = useCallback(async ({ seedIfEmpty }) => {
    let orgData = null;
    orgData = await fetchOrgData(); // throws if offline / misconfigured

    if (!orgData) {
      if (!seedIfEmpty) return null;
      orgData = seedData();
      const adminHash = sha256("admin123");
      const staffHash = sha256("hr12345");
      orgData.users = [
        { id: "usr-admin", name: "System Administrator", username: "admin", email: "admin@company.com", position: "IT Systems Admin", passwordHash: adminHash, role: "Admin", locationId: null },
        { id: "usr-staff", name: "Staff User", username: "staff", email: "staff@company.com", position: "Staff", passwordHash: staffHash, role: "Regional Staff", locationId: "loc-main" },
      ];
      await saveOrgData(orgData);
    }
    return orgData;
  }, []);

  // Load persisted state on mount
  useEffect(() => {
    try {
      const t = localStorage.getItem("theme-pref");
      if (t) setTheme(t);
      const s = localStorage.getItem("sidebar-pref");
      if (s) setSidebarOpen(s === "open");
    } catch {}

    (async () => {
      try {
        const orgData = await loadFromCloud({ seedIfEmpty: true });
        setData(orgData);
        setConnectionError(null);
        setLastSynced(new Date());
      } catch (err) {
        setConnectionError(err?.message || "Could not connect to the database.");
      } finally {
        setLoaded(true);
      }
    })();
  }, [loadFromCloud]);

  // Live feed: any change anyone saves is pushed here automatically —
  // no manual sync needed. The Sync button still works as a manual
  // fallback (e.g. right after reconnecting).
  useEffect(() => {
    const unsubscribe = subscribeToOrgData((next) => {
      setData(next);
      setLastSynced(new Date());
    });
    return unsubscribe;
  }, []);

  // Save org data to the shared database whenever it changes locally
  const persist = useCallback(async (next) => {
    const prev = dataRef.current;
    setData(next);
    try {
      await saveOrgData(next, prev);
      setLastSynced(new Date());
    } catch {
      showToast("Could not save — check your connection and try again.");
    }
  }, [showToast]);

  const toggleTheme = () => {
    const next = theme === "light" ? "dark" : "light";
    setTheme(next);
    try { localStorage.setItem("theme-pref", next); } catch {}
  };

  const toggleSidebar = () => {
    const next = !sidebarOpen;
    setSidebarOpen(next);
    try { localStorage.setItem("sidebar-pref", next ? "open" : "closed"); } catch {}
  };

  // Keep the page background (outside the app's own div) in sync with the
  // theme, so overscroll/bounce edges never flash white in dark mode.
  useEffect(() => {
    const bg = theme === "dark" ? "#12141A" : "#F7F8FA";
    document.documentElement.style.backgroundColor = bg;
    document.body.style.backgroundColor = bg;
    const meta = document.querySelector('meta[name="theme-color"]');
    if (meta) meta.setAttribute("content", theme === "dark" ? "#12141A" : "#3B82F6");
  }, [theme]);

  if (!loaded) {
    return <div className="boot"><div className="spinner" /></div>;
  }

  if (connectionError) {
    return (
      <div className={theme === "dark" ? "theme-dark" : "theme-light"}>
        <GlobalStyles />
        <div className="login-wrap">
          <div className="login-card" style={{ textAlign: "center" }}>
            <div className="confirm-icon"><AlertTriangle size={20} /></div>
            <h3 style={{ marginBottom: 8 }}>You're offline</h3>
            <p className="login-sub">
              This app needs an internet connection to load the shared data.
              Check your connection and try again.
            </p>
            <button
              className="btn primary full"
              onClick={() => { setLoaded(false); setConnectionError(null); loadFromCloud({ seedIfEmpty: true }).then((d) => { setData(d); setLastSynced(new Date()); }).catch((e) => setConnectionError(e?.message || "Could not connect.")).finally(() => setLoaded(true)); }}
            >
              Retry
            </button>
          </div>
        </div>
      </div>
    );
  }

  if (!currentUser) {
    return (
      <div className={theme === "dark" ? "theme-dark" : "theme-light"}>
        <GlobalStyles />
        <LoginScreen users={data.users} onLogin={setCurrentUser} />
      </div>
    );
  }

  const isAdmin = currentUser.role === "Admin";
  const scopedLocationId = isAdmin ? null : currentUser.locationId;
  const pendingCount = data.assets.filter((a) => a.pendingDeletion).length;

  return (
    <div className={theme === "dark" ? "theme-dark" : "theme-light"}>
      <GlobalStyles />
      <div className="shell">
        <Sidebar
          open={sidebarOpen}
          onToggle={toggleSidebar}
          view={view}
          setView={setView}
          isAdmin={isAdmin}
          pendingCount={pendingCount}
        />
        <div className="main">
          <TopBar
            theme={theme}
            toggleTheme={toggleTheme}
            currentUser={currentUser}
            onLogout={() => setCurrentUser(null)}
            onToggleSidebar={toggleSidebar}
            locations={data.locations}
            scopedLocationId={scopedLocationId}
            data={data}
            persist={persist}
            onOpenAsset={openAssetFromNotif}
          />
          <div className="content">
            {view === "dashboard" && (
              <Dashboard data={data} scopedLocationId={scopedLocationId} currentUser={currentUser} setView={setView} />
            )}
            {view === "assets" && (
              <AssetsView
                data={data}
                persist={persist}
                isAdmin={isAdmin}
                scopedLocationId={scopedLocationId}
                showToast={showToast}
                currentUser={currentUser}
                focusAssetId={focusAssetId}
                onFocusHandled={() => setFocusAssetId(null)}
              />
            )}
            {view === "maintenance" && (
              <MaintenanceView data={data} persist={persist} showToast={showToast} scopedLocationId={scopedLocationId} currentUser={currentUser} />
            )}
            {view === "categories" && isAdmin && (
              <CategoriesView data={data} persist={persist} showToast={showToast} currentUser={currentUser} />
            )}
            {view === "locations" && isAdmin && (
              <LocationsView data={data} persist={persist} showToast={showToast} currentUser={currentUser} />
            )}
            {view === "users" && isAdmin && (
              <UsersView data={data} persist={persist} showToast={showToast} currentUser={currentUser} />
            )}
            {view === "backup" && isAdmin && (
              <BackupView data={data} persist={persist} showToast={showToast} currentUser={currentUser} />
            )}
            {view === "activity" && (
              <ActivityLogView data={data} isAdmin={isAdmin} scopedLocationId={scopedLocationId} />
            )}
            {view === "approvals" && isAdmin && (
              <ApprovalsView data={data} persist={persist} showToast={showToast} currentUser={currentUser} />
            )}
          </div>
        </div>
      </div>
      {toast && <div className={`toast ${toast.phase === "out" ? "toast-out" : "toast-in"}`}>{toast.message}</div>}
    </div>
  );
}

/* ---------------------------------------------------------
   Sidebar
--------------------------------------------------------- */
function Sidebar({ open, onToggle, view, setView, isAdmin, pendingCount }) {
  const items = [
    { id: "dashboard", label: "Overview", icon: LayoutDashboard },
    { id: "assets", label: "Assets", icon: Package },
    { id: "maintenance", label: "Maintenance", icon: Wrench },
    ...(isAdmin ? [
      { id: "categories", label: "Categories", icon: Tags },
      { id: "locations", label: "Locations", icon: MapPin },
      { id: "users", label: "User Accounts", icon: Users },
      { id: "backup", label: "Backup & Restore", icon: Download },
    ] : []),
    { id: "activity", label: "Activity Log", icon: ShieldCheck },
    ...(isAdmin ? [
      { id: "approvals", label: "Approvals", icon: AlertTriangle, badge: pendingCount },
    ] : []),
  ];
  return (
    <>
      {open && <div className="sidebar-backdrop" onClick={onToggle} />}
      <div className={`sidebar ${open ? "" : "collapsed"}`}>
        <div className="sidebar-top">
          {open && <div className="brand"><span className="brand-badge"><ShieldCheck size={16} /></span><span>AssetHub</span></div>}
          {!open && <div className="brand-mini"><span className="brand-badge"><ShieldCheck size={16} /></span></div>}
        </div>
        <nav>
          {items.map((it) => (
            <button
              key={it.id}
              className={`nav-item ${view === it.id ? "active" : ""}`}
              onClick={() => { setView(it.id); if (window.innerWidth <= 860) onToggle(); }}
              title={it.label}
            >
              <it.icon size={17} />
              {open && <span>{it.label}</span>}
              {!!it.badge && (
                <span style={{ marginLeft: "auto", background: "#EF4444", color: "white", borderRadius: 999, fontSize: 11, fontWeight: 700, padding: "1px 7px" }}>
                  {it.badge}
                </span>
              )}
            </button>
          ))}
        </nav>
      </div>
    </>
  );
}

/* ---------------------------------------------------------
   Top Bar
--------------------------------------------------------- */
function TopBar({ theme, toggleTheme, currentUser, onLogout, locations, scopedLocationId, data, onToggleSidebar, persist, onOpenAsset }) {
  const [notifOpen, setNotifOpen] = useState(false);
  const [confirmLogoutOpen, setConfirmLogoutOpen] = useState(false);
  const isAdmin = currentUser.role === "Admin";
  const locName = scopedLocationId
    ? locations.find((l) => l.id === scopedLocationId)?.name
    : "All Locations (HQ)";

  // Warranty/calibration alerts are derived from live asset data rather than
  // stored rows, so "read" state for them is tracked per-user in
  // localStorage instead of the database — dismissing one just hides it
  // for this user on this browser until it's clicked again.
  const readAlertsKey = `read-alerts-${currentUser.id}`;
  const [readAlertIds, setReadAlertIds] = useState(() => {
    try { return JSON.parse(localStorage.getItem(readAlertsKey)) || []; } catch { return []; }
  });
  useEffect(() => {
    try { setReadAlertIds(JSON.parse(localStorage.getItem(readAlertsKey)) || []); } catch { setReadAlertIds([]); }
  }, [readAlertsKey]);
  const markAlertRead = (alertId) => {
    setReadAlertIds((prev) => {
      if (prev.includes(alertId)) return prev;
      const next = [...prev, alertId];
      try { localStorage.setItem(readAlertsKey, JSON.stringify(next)); } catch {}
      return next;
    });
  };
  const openAlert = (alert) => {
    markAlertRead(alert.id);
    setNotifOpen(false);
    if (onOpenAsset) onOpenAsset(alert.assetId);
  };

  const allAlerts = useMemo(() => computeAlerts(data.assets, scopedLocationId), [data.assets, scopedLocationId]);
  const alerts = useMemo(() => allAlerts.filter((a) => !readAlertIds.includes(a.id)), [allAlerts, readAlertIds]);
  const recentActivity = useMemo(() => (data.auditLog || []).slice(0, 6), [data.auditLog]);
  const myComments = useMemo(
    () => (data.comments || []).filter((c) => {
      if (!(c.targetUserIds || []).includes(currentUser.id) || (c.readBy || []).includes(currentUser.id)) return false;
      // Live access check: even if targetUserIds was set at comment time,
      // an asset that has since been transferred away from this user's
      // location should no longer surface a notification for it (admins
      // always retain access).
      if (isAdmin) return true;
      const asset = data.assets.find((a) => a.id === c.assetId);
      return !!asset && asset.locationId === scopedLocationId;
    }).sort((a, b) => new Date(b.at) - new Date(a.at)),
    [data.comments, data.assets, currentUser.id, isAdmin, scopedLocationId]
  );
  const notifCount = alerts.length + myComments.length;

  const markCommentRead = (commentId) => {
    if (!persist) return;
    persist({
      ...data,
      comments: (data.comments || []).map((c) =>
        c.id === commentId ? { ...c, readBy: [...(c.readBy || []), currentUser.id] } : c
      ),
    });
  };

  const openComment = (comment) => {
    markCommentRead(comment.id);
    setNotifOpen(false);
    if (onOpenAsset) onOpenAsset(comment.assetId);
  };

  const initials = currentUser.name.split(" ").map((s) => s[0]).slice(0, 2).join("");

  return (
    <div className="topbar">
      <div className="topbar-left">
        <button className="icon-btn mobile-menu-btn" onClick={onToggleSidebar} title="Menu">
          <Menu size={16} />
        </button>
        <span className="topbar-region"><MapPin size={13} /> {locName}</span>
      </div>
      <div className="topbar-right">
        <div className="notif-wrap">
          <button className="icon-btn" onClick={() => setNotifOpen((o) => !o)} title="Notifications">
            <Bell size={16} />
            {notifCount > 0 && <span className="notif-dot">{notifCount > 9 ? "9+" : notifCount}</span>}
          </button>
          {notifOpen && (
            <>
              <div className="notif-backdrop" onClick={() => setNotifOpen(false)} />
              <div className="notif-panel">
                <div className="notif-section-title">Needs Attention</div>
                {alerts.length === 0 && myComments.length === 0 && (
                  <div className="notif-empty">You're all caught up.</div>
                )}
                {alerts.slice(0, 8).map((a) => (
                  <button
                    key={a.id}
                    type="button"
                    className={`notif-item notif-item-btn ${a.urgent ? "urgent" : ""}`}
                    onClick={() => openAlert(a)}
                    title="View asset"
                  >
                    {a.label}
                  </button>
                ))}
                {myComments.map((c) => {
                  const asset = data.assets.find((a) => a.id === c.assetId);
                  // If I already have an earlier comment in this same thread,
                  // this new one is a reply to me specifically.
                  const isReplyToMe = (data.comments || []).some(
                    (other) => other.assetId === c.assetId && other.authorId === currentUser.id && new Date(other.at) < new Date(c.at)
                  );
                  return (
                    <button
                      key={c.id}
                      type="button"
                      className="notif-item notif-item-btn"
                      onClick={() => openComment(c)}
                      title="View asset"
                    >
                      {isReplyToMe
                        ? <>{c.authorName} replied to your comment on Asset #{asset?.tag || "—"}</>
                        : <>{c.authorName} commented on Asset #{asset?.tag || "—"}</>}
                    </button>
                  );
                })}
                {isAdmin && (
                  <>
                    <div className="notif-section-title" style={{ marginTop: 10 }}>Recent Activity</div>
                    {recentActivity.length === 0 && <div className="notif-empty">No activity yet.</div>}
                    {recentActivity.map((l) => (
                      <div key={l.id} className="notif-item">{l.userName}: {l.message}</div>
                    ))}
                  </>
                )}
              </div>
            </>
          )}
        </div>
        <button className="icon-btn" onClick={toggleTheme} title="Toggle theme">
          {theme === "light" ? <Moon size={16} /> : <Sun size={16} />}
        </button>
        <div className="user-chip">
          <div className="avatar">{initials}</div>
          <div className="user-meta">
            <div className="user-name">{currentUser.name}</div>
            <div className="user-role">{currentUser.position || currentUser.role}</div>
          </div>
        </div>
        <button className="icon-btn" onClick={() => setConfirmLogoutOpen(true)} title="Log out"><LogOut size={16} /></button>
      </div>
      {confirmLogoutOpen && (
        <ConfirmDialog
          message="Sign out of AssetHub?"
          confirmLabel="Yes"
          maxWidth={260}
          onCancel={() => setConfirmLogoutOpen(false)}
          onConfirm={() => { setConfirmLogoutOpen(false); onLogout(); }}
        />
      )}
    </div>
  );
}

/* ---------------------------------------------------------
   Dashboard
--------------------------------------------------------- */
function Dashboard({ data, scopedLocationId, currentUser, setView }) {
  const isAdmin = currentUser?.role === "Admin";
  const assets = scopedLocationId
    ? data.assets.filter((a) => a.locationId === scopedLocationId)
    : data.assets;

  const statusData = useMemo(() => {
    const counts = {};
    assets.forEach((a) => { counts[a.status] = (counts[a.status] || 0) + 1; });
    return Object.entries(counts).map(([name, value]) => ({ name, value }));
  }, [assets]);

  const categoryData = useMemo(() => {
    const counts = {};
    assets.forEach((a) => {
      const cat = data.categories.find((c) => c.id === a.categoryId);
      const name = cat ? cat.name : "Uncategorized";
      counts[name] = (counts[name] || 0) + 1;
    });
    return Object.entries(counts).map(([name, value]) => ({ name, value })).sort((a, b) => b.value - a.value);
  }, [assets, data.categories]);

  // Admin-only: how assets break down across every location/country. Not
  // scoped by scopedLocationId (that's always null for an admin anyway),
  // so this always reflects the whole company.
  const locationData = useMemo(() => {
    if (!isAdmin) return [];
    const counts = {};
    data.assets.forEach((a) => {
      const loc = data.locations.find((l) => l.id === a.locationId);
      const name = loc ? loc.name : "Unassigned";
      counts[name] = (counts[name] || 0) + 1;
    });
    return Object.entries(counts).map(([name, value]) => ({ name, value })).sort((a, b) => b.value - a.value);
  }, [isAdmin, data.assets, data.locations]);

  // Same index-based assignment as categoryColor() so the pie chart and the
  // category dots shown in the tables always agree on a category's color.
  const categoryPalette = useMemo(() => {
    const map = {};
    data.categories.forEach((c, i) => { map[c.name] = CAT_PALETTE[i % CAT_PALETTE.length]; });
    return map;
  }, [data.categories]);

  const locationPalette = useMemo(() => {
    const map = {};
    data.locations.forEach((l, i) => { map[l.name] = CAT_PALETTE[(i + 3) % CAT_PALETTE.length]; });
    return map;
  }, [data.locations]);

  const totals = {
    total: assets.length,
    inUse: assets.filter((a) => a.status === "In Use").length,
    underRepair: assets.filter((a) => a.status === "Under Repair").length,
    inStock: assets.filter((a) => a.status === "In Stock").length,
  };
  const pct = (n) => (assets.length ? Math.round((n / assets.length) * 100) : 0);

  // Asset Health: a derived read on the fleet's condition, built from data
  // that's already tracked (Condition, status, and the same warranty/
  // calibration alerts the notification bell uses) rather than inventing a
  // new field — Critical is anything actually down or in poor shape, Needs
  // Attention is anything with an upcoming warranty/calibration issue or
  // fair condition, everything else counts as Healthy.
  const alerts = useMemo(() => computeAlerts(assets, null), [assets]);
  const alertAssetIds = useMemo(() => new Set(alerts.map((a) => a.id.replace(/-[wc]$/, ""))), [alerts]);
  const health = useMemo(() => {
    let healthy = 0, attention = 0, critical = 0;
    assets.forEach((a) => {
      if (a.status === "Under Repair" || a.condition === "Poor") critical += 1;
      else if (alertAssetIds.has(a.id) || a.condition === "Fair") attention += 1;
      else healthy += 1;
    });
    return { healthy, attention, critical };
  }, [assets, alertAssetIds]);
  const healthyPct = assets.length ? Math.round((health.healthy / assets.length) * 100) : 0;

  // Warranty Overview: real counts from each IT asset's warrantyExpiry date
  // (assets without a tracked warranty — "N/A" or blank — are excluded).
  const warrantyStats = useMemo(() => {
    const now = new Date();
    const monthEnd = new Date(now.getFullYear(), now.getMonth() + 1, 0, 23, 59, 59);
    const in30 = new Date(now.getTime() + 30 * 24 * 60 * 60 * 1000);
    let thisMonth = 0, next30 = 0, expired = 0;
    assets.forEach((a) => {
      if (a.assetType !== "IT" || !a.warrantyExpiry || a.warrantyExpiry === "N/A") return;
      const d = new Date(a.warrantyExpiry);
      if (isNaN(d)) return;
      if (d < now) expired += 1;
      else {
        if (d <= monthEnd) thisMonth += 1;
        if (d <= in30) next30 += 1;
      }
    });
    return { thisMonth, next30, expired };
  }, [assets]);

  // Calibration Overview: same shape as Warranty Overview, but for Non-IT
  // assets that require calibration, checked against nextCalibrationDate
  // (falling back to calibrationDate if that's the only date on record).
  const calibrationStats = useMemo(() => {
    const now = new Date();
    const monthEnd = new Date(now.getFullYear(), now.getMonth() + 1, 0, 23, 59, 59);
    const in30 = new Date(now.getTime() + 30 * 24 * 60 * 60 * 1000);
    let thisMonth = 0, next30 = 0, overdue = 0;
    assets.forEach((a) => {
      if (a.assetType !== "Non-IT" || !a.requiresCalibration) return;
      const checkDate = a.nextCalibrationDate || a.calibrationDate;
      if (!checkDate) return;
      const d = new Date(checkDate);
      if (isNaN(d)) return;
      if (d < now) overdue += 1;
      else {
        if (d <= monthEnd) thisMonth += 1;
        if (d <= in30) next30 += 1;
      }
    });
    return { thisMonth, next30, overdue };
  }, [assets]);

  // Regional Staff should only see activity for assets in their own
  // location/country — entries get tagged with a locationId when they're
  // logged (see withLog call sites); anything untagged (category/location/
  // user/backup admin actions) only ever shows for Admins, who see
  // everything unfiltered.
  const recentActivity = useMemo(() => {
    const list = scopedLocationId
      ? (data.auditLog || []).filter((l) => l.locationId === scopedLocationId)
      : (data.auditLog || []);
    return list.slice(0, 5);
  }, [data.auditLog, scopedLocationId]);

  const hour = new Date().getHours();
  const greeting = hour < 12 ? "Good morning" : hour < 18 ? "Good afternoon" : "Good evening";
  const firstName = (currentUser?.name || "").split(" ")[0] || "there";

  return (
    <div>
      <div className="welcome-banner">
        <div>
          <h2>{greeting}, {firstName}</h2>
          <p className="welcome-sub">Here's what's happening with your assets today.</p>
        </div>
        <div className="welcome-icon"><ShieldCheck size={26} /></div>
      </div>

      <div className="metrics-row">
        <Metric label="Total Assets" value={totals.total} sub={isAdmin ? `Across ${data.locations.length} location${data.locations.length === 1 ? "" : "s"}` : "In this location"} icon={Package} color="#6366F1" />
        <Metric label="In Use" value={totals.inUse} sub={`${pct(totals.inUse)}% of all assets`} icon={CheckSquare} color={STATUS_COLORS["In Use"]} />
        <Metric label="In Stock" value={totals.inStock} sub={`${pct(totals.inStock)}% of all assets`} icon={Archive} color={STATUS_COLORS["In Stock"]} />
        <Metric label="Under Repair" value={totals.underRepair} sub={`${pct(totals.underRepair)}% of all assets`} icon={Wrench} color={STATUS_COLORS["Under Repair"]} />
      </div>

      <div className={`charts-row ${isAdmin ? "charts-row-4" : "charts-row-3"}`}>
        <DonutCard title="Assets by Status" data={statusData} palette={STATUS_COLORS} total={assets.length} />
        <DonutCard title="Assets by Category" data={categoryData} palette={categoryPalette} total={assets.length} />
        <div className="panel chart-card health-panel">
          <div className="panel-head"><h3>Asset Health</h3></div>
          <div className="health-body">
            <div className="health-ring-wrap">
              <ResponsiveContainer width={110} height={110}>
                <PieChart>
                  <Pie
                    data={[
                      { name: "Healthy", value: health.healthy },
                      { name: "Needs Attention", value: health.attention },
                      { name: "Critical", value: health.critical },
                    ].filter((d) => d.value > 0)}
                    dataKey="value" nameKey="name" innerRadius={38} outerRadius={52} startAngle={90} endAngle={-270} paddingAngle={health.healthy && (health.attention || health.critical) ? 3 : 0}
                  >
                    <Cell fill="#10B981" />
                    <Cell fill="#F59E0B" />
                    <Cell fill="#EF4444" />
                  </Pie>
                </PieChart>
              </ResponsiveContainer>
              <div className="health-ring-center">
                <div className="health-ring-pct">{assets.length ? `${healthyPct}%` : "—"}</div>
                <div className="health-ring-label">Healthy</div>
              </div>
            </div>
            <ul className="legend-list">
              <li><span className="legend-dot" style={{ background: "#10B981" }} /><span className="legend-name">Healthy</span><span className="legend-count">{health.healthy}</span></li>
              <li><span className="legend-dot" style={{ background: "#F59E0B" }} /><span className="legend-name">Needs Attention</span><span className="legend-count">{health.attention}</span></li>
              <li><span className="legend-dot" style={{ background: "#EF4444" }} /><span className="legend-name">Critical</span><span className="legend-count">{health.critical}</span></li>
            </ul>
          </div>
        </div>
        {isAdmin && (
          <DonutCard title="Assets by Location" data={locationData} palette={locationPalette} total={data.assets.length} />
        )}
      </div>

      <div className="bottom-row bottom-row-2">
        <div className="panel">
          <div className="panel-head"><h3>Warranty Overview</h3></div>
          <div className="warranty-stats">
            <div className="warranty-stat">
              <div className="warranty-stat-icon" style={{ background: "#D1FAE522", color: "#10B981" }}><ShieldCheck size={16} /></div>
              <div className="warranty-stat-value">{warrantyStats.thisMonth}</div>
              <div className="warranty-stat-label">This Month<br />Expiring</div>
            </div>
            <div className="warranty-stat">
              <div className="warranty-stat-icon" style={{ background: "#FEF3C722", color: "#F59E0B" }}><Bell size={16} /></div>
              <div className="warranty-stat-value">{warrantyStats.next30}</div>
              <div className="warranty-stat-label">Next 30 Days<br />Expiring</div>
            </div>
            <div className="warranty-stat">
              <div className="warranty-stat-icon" style={{ background: "#FEE2E222", color: "#EF4444" }}><AlertTriangle size={16} /></div>
              <div className="warranty-stat-value">{warrantyStats.expired}</div>
              <div className="warranty-stat-label">Expired<br />Assets</div>
            </div>
          </div>
          <div className="panel-head panel-head-sub"><h3>Calibration Overview</h3></div>
          <div className="warranty-stats">
            <div className="warranty-stat">
              <div className="warranty-stat-icon" style={{ background: "#D1FAE522", color: "#10B981" }}><ShieldCheck size={16} /></div>
              <div className="warranty-stat-value">{calibrationStats.thisMonth}</div>
              <div className="warranty-stat-label">This Month<br />Due</div>
            </div>
            <div className="warranty-stat">
              <div className="warranty-stat-icon" style={{ background: "#FEF3C722", color: "#F59E0B" }}><Bell size={16} /></div>
              <div className="warranty-stat-value">{calibrationStats.next30}</div>
              <div className="warranty-stat-label">Next 30 Days<br />Due</div>
            </div>
            <div className="warranty-stat">
              <div className="warranty-stat-icon" style={{ background: "#FEE2E222", color: "#EF4444" }}><AlertTriangle size={16} /></div>
              <div className="warranty-stat-value">{calibrationStats.overdue}</div>
              <div className="warranty-stat-label">Overdue<br />Assets</div>
            </div>
          </div>
        </div>

        <div className="panel">
          <div className="panel-head"><h3>Recent Activity</h3><button type="button" className="panel-link" onClick={() => setView?.("activity")}>View all</button></div>
          <ul className="activity-list">
            {recentActivity.length === 0 && <li className="activity-empty">No activity yet.</li>}
            {recentActivity.map((l) => {
              const act = activityStyle(l.message);
              return (
                <li key={l.id} className="activity-item">
                  <span className="activity-icon" style={{ background: `${act.color}1a`, color: act.color }}><act.Icon size={13} /></span>
                  <span className="activity-text">{l.message}</span>
                  <span className="activity-time">{formatLogTime(l.at)}</span>
                </li>
              );
            })}
          </ul>
        </div>
      </div>
    </div>
  );
}

// Picks an icon + color for a Recent Activity row based on the audit-log
// message's wording, so the list reads at a glance like the rest of the app
// (green = added, blue = assignment, purple = transfer, orange = repair).
function activityStyle(message) {
  const m = (message || "").toLowerCase();
  if (m.includes("added")) return { Icon: Plus, color: "#10B981" };
  if (m.includes("transferred")) return { Icon: Truck, color: "#8B5CF6" };
  if (m.includes("repair") || m.includes("maintenance")) return { Icon: Wrench, color: "#F59E0B" };
  if (m.includes("deleted")) return { Icon: Trash2, color: "#EF4444" };
  if (m.includes("assigned") || m.includes("checked out")) return { Icon: User, color: "#3B82F6" };
  return { Icon: Info, color: "#6B7280" };
}

// Formats an ISO timestamp the way the reference design does — "Today,
// 11:32 AM" / "Yesterday, 9:15 AM" / a plain date further back.
function formatLogTime(iso) {
  const d = new Date(iso);
  if (isNaN(d)) return "";
  const now = new Date();
  const isSameDay = (a, b) => a.getFullYear() === b.getFullYear() && a.getMonth() === b.getMonth() && a.getDate() === b.getDate();
  const yesterday = new Date(now); yesterday.setDate(now.getDate() - 1);
  const time = d.toLocaleTimeString([], { hour: "numeric", minute: "2-digit" });
  if (isSameDay(d, now)) return `Today, ${time}`;
  if (isSameDay(d, yesterday)) return `Yesterday, ${time}`;
  return d.toLocaleDateString([], { month: "short", day: "numeric" });
}

function Metric({ label, value, sub, icon: Icon, color }) {
  return (
    <div className="metric">
      <div className="metric-top">
        {Icon && (
          <div className="metric-icon" style={{ background: `${color}22`, color }}>
            <Icon size={15} />
          </div>
        )}
        <div className="metric-label">{label}</div>
      </div>
      <div className="metric-value">{value}</div>
      {sub && <div className="metric-sub">{sub}</div>}
    </div>
  );
}

// A legend list that scrolls instead of cramping when there are many
// entries (lots of categories/locations), with the scrollbar itself hidden
// and a small bobbing chevron shown only when there's actually more
// content in that direction — so it reads as "scrollable" without a
// visible scrollbar cluttering the card.
function ScrollableLegend({ children }) {
  const ref = React.useRef(null);
  const [canUp, setCanUp] = useState(false);
  const [canDown, setCanDown] = useState(false);

  const check = () => {
    const el = ref.current;
    if (!el) return;
    setCanUp(el.scrollTop > 2);
    setCanDown(el.scrollHeight - el.scrollTop - el.clientHeight > 2);
  };

  useEffect(() => {
    check();
    const el = ref.current;
    if (!el || typeof ResizeObserver === "undefined") return;
    const ro = new ResizeObserver(check);
    ro.observe(el);
    return () => ro.disconnect();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [children]);

  return (
    <div className="legend-scroll-wrap">
      <ul className="legend-list" ref={ref} onScroll={check}>
        {children}
      </ul>
      {canUp && <ChevronUp size={13} className="legend-scroll-hint legend-scroll-hint-up" />}
      {canDown && <ChevronDown size={13} className="legend-scroll-hint legend-scroll-hint-down" />}
    </div>
  );
}

function DonutCard({ title, data, palette, total }) {
  const colors = (name, i) => (palette && palette[name]) || CAT_PALETTE[i % CAT_PALETTE.length];
  return (
    <div className="panel chart-card">
      <div className="panel-head">
        <h3>{title}</h3>
      </div>
      {data.length === 0 ? (
        <div className="empty-chart">No data to display</div>
      ) : (
        <div className="donut-body">
          <ResponsiveContainer width={120} height={120}>
            <PieChart>
              <Pie data={data} dataKey="value" nameKey="name" innerRadius={36} outerRadius={52} paddingAngle={2}>
                {data.map((entry, i) => (
                  <Cell key={entry.name} fill={colors(entry.name, i)} />
                ))}
              </Pie>
              <Tooltip contentStyle={{ fontSize: 12, padding: "6px 10px", borderRadius: 8 }} itemStyle={{ fontSize: 12 }} labelStyle={{ fontSize: 12 }} />
            </PieChart>
          </ResponsiveContainer>
          <ScrollableLegend>
            {data.map((entry, i) => (
              <li key={entry.name}>
                <span className="legend-dot" style={{ background: colors(entry.name, i) }} />
                <span className="legend-name">{entry.name}</span>
                <span className="legend-count">{entry.value} <span className="legend-pct">({total ? Math.round((entry.value / total) * 100) : 0}%)</span></span>
              </li>
            ))}
          </ScrollableLegend>
        </div>
      )}
    </div>
  );
}

/* ---------------------------------------------------------
   Assets View
--------------------------------------------------------- */
function emptyAsset(defaultLocationId) {
  return {
    id: null, tag: "", name: "", categoryId: "", assetType: "IT", brand: "", model: "",
    yearModel: "", serial: "", status: "", condition: "New", locationId: defaultLocationId || "",
    assignedTo: "", purchaseDate: "", purchaseCost: "", warrantyExpiry: "",
    requiresCalibration: false, calibrationDate: "", nextCalibrationDate: "",
    notes: "", transferHistory: [], department: "",
  };
}

function AssetsView({ data, persist, isAdmin, scopedLocationId, showToast, currentUser, focusAssetId, onFocusHandled }) {
  const [search, setSearch] = useState("");
  const [locationFilter, setLocationFilter] = useState("all");
  const [categoryFilter, setCategoryFilter] = useState("all");
  const [userFilter, setUserFilter] = useState("all");
  const [editing, setEditing] = useState(null);
  const [viewing, setViewing] = useState(null);

  // A comment notification was clicked elsewhere in the app — jump straight
  // to that asset's detail view, then clear the request so it doesn't refire.
  useEffect(() => {
    if (!focusAssetId) return;
    const a = data.assets.find((x) => x.id === focusAssetId);
    // Only open the asset if the current user still has access to it —
    // an asset transferred to a different location shouldn't be openable
    // via a stale notification link either.
    if (a && (isAdmin || a.locationId === scopedLocationId)) setViewing(a);
    if (onFocusHandled) onFocusHandled();
  }, [focusAssetId, data.assets, onFocusHandled, isAdmin, scopedLocationId]);

  // Assets with an unread comment for the current user — shown as a small
  // indicator in the table so the person who added an asset notices right
  // away that there's an update on it.
  const unreadCommentAssetIds = useMemo(() => {
    const set = new Set();
    (data.comments || []).forEach((c) => {
      if ((c.targetUserIds || []).includes(currentUser.id) && !(c.readBy || []).includes(currentUser.id)) {
        set.add(c.assetId);
      }
    });
    return set;
  }, [data.comments, currentUser.id]);

  // Assets with an upcoming/overdue warranty or calibration issue — same
  // source data as the notification bell, mapped by asset id so the table
  // can flag it with a small icon (see the legend above the table).
  const assetAlertMap = useMemo(() => {
    const map = new Map();
    computeAlerts(data.assets, scopedLocationId).forEach((al) => {
      const kind = al.id.endsWith("-w") ? "warranty" : "calibration";
      const existing = map.get(al.assetId);
      if (!existing || (al.urgent && !existing.urgent)) {
        map.set(al.assetId, { urgent: al.urgent, kind });
      }
    });
    return map;
  }, [data.assets, scopedLocationId]);

  // Opening an asset's detail view (from the table, or from a notification)
  // also clears its unread-comment flag for the current user.
  useEffect(() => {
    if (!viewing) return;
    const hasUnread = (data.comments || []).some(
      (c) => c.assetId === viewing.id && (c.targetUserIds || []).includes(currentUser.id) && !(c.readBy || []).includes(currentUser.id)
    );
    if (!hasUnread) return;
    persist({
      ...data,
      comments: (data.comments || []).map((c) =>
        c.assetId === viewing.id && (c.targetUserIds || []).includes(currentUser.id) && !(c.readBy || []).includes(currentUser.id)
          ? { ...c, readBy: [...(c.readBy || []), currentUser.id] }
          : c
      ),
    });
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [viewing?.id]);
  const [confirmDelete, setConfirmDelete] = useState(null);
  const [requestDeleteTarget, setRequestDeleteTarget] = useState(null); // asset id awaiting a reason
  const [deleteReason, setDeleteReason] = useState("");
  const [selected, setSelected] = useState([]);
  const [bulkDeleteConfirmOpen, setBulkDeleteConfirmOpen] = useState(false);
  const [maintPromptAsset, setMaintPromptAsset] = useState(null); // asset awaiting maintenance details before going Under Repair
  const [inUsePromptAsset, setInUsePromptAsset] = useState(null); // asset awaiting department/assigned user before going In Use
  const [inUseDepartment, setInUseDepartment] = useState("");
  const [inUseAssignedTo, setInUseAssignedTo] = useState("");
  const [inUseLocationId, setInUseLocationId] = useState("");
  const [bulkDeleteText, setBulkDeleteText] = useState("");
  const [transferTarget, setTransferTarget] = useState(null); // asset id
  const [transferLocationId, setTransferLocationId] = useState("");
  const [transferNewLocationName, setTransferNewLocationName] = useState("");
  const [transferAssignedTo, setTransferAssignedTo] = useState("");
  const [transferReason, setTransferReason] = useState("");
  const [sort, setSort] = useState({ key: null, dir: "asc" });
  const fileInputRef = React.useRef(null);

  // The base pool of assets this user can see at all — a Regional Staff
  // account only ever sees their own location's assets.
  const scopedAssetsBase = useMemo(
    () => (scopedLocationId ? data.assets.filter((a) => a.locationId === scopedLocationId) : data.assets),
    [data.assets, scopedLocationId]
  );

  // Regional Staff should only see assigned-user names from their own
  // location's assets, not the whole company's.
  const assignedUserOptions = useMemo(() => {
    const set = new Set(scopedAssetsBase.map((a) => a.assignedTo).filter(Boolean));
    return Array.from(set).sort();
  }, [scopedAssetsBase]);

  const departmentOptions = useMemo(() => {
    const set = new Set(data.assets.map((a) => a.department).filter(Boolean));
    return Array.from(set).sort();
  }, [data.assets]);

  // The Location filter should only list locations that actually have an
  // asset somewhere in this user's scope — e.g. an admin shouldn't see an
  // empty "Philippines" location in the dropdown if nothing's there yet.
  const locationOptions = useMemo(() => {
    const ids = new Set(scopedAssetsBase.map((a) => a.locationId));
    return data.locations.filter((l) => ids.has(l.id)).sort((a, b) => a.name.localeCompare(b.name));
  }, [data.locations, scopedAssetsBase]);

  // The Category filter should only list categories that have an asset in
  // the currently selected location (or across the whole scope if no
  // location is picked) — e.g. don't show "Printers" for a country that
  // doesn't have one, unless someone's actually added one there.
  const categoryOptions = useMemo(() => {
    const base = locationFilter !== "all" ? scopedAssetsBase.filter((a) => a.locationId === locationFilter) : scopedAssetsBase;
    const ids = new Set(base.map((a) => a.categoryId));
    return data.categories.filter((c) => ids.has(c.id)).sort((a, b) => a.name.localeCompare(b.name));
  }, [data.categories, scopedAssetsBase, locationFilter]);

  // If switching the Location filter makes the currently-picked Category no
  // longer relevant (e.g. that country has no printers), fall back to "all"
  // instead of silently showing an empty table.
  useEffect(() => {
    if (categoryFilter !== "all" && !categoryOptions.some((c) => c.id === categoryFilter)) {
      setCategoryFilter("all");
    }
  }, [categoryOptions, categoryFilter]);

  const visibleAssets = useMemo(() => {
    let list = scopedAssetsBase;
    if (locationFilter !== "all") list = list.filter((a) => a.locationId === locationFilter);
    if (categoryFilter !== "all") list = list.filter((a) => a.categoryId === categoryFilter);
    if (userFilter !== "all") list = list.filter((a) => a.assignedTo === userFilter);
    if (search.trim()) {
      const q = search.toLowerCase();
      list = list.filter((a) => {
        const loc = data.locations.find((l) => l.id === a.locationId)?.name || "";
        const cat = data.categories.find((c) => c.id === a.categoryId)?.name || "";
        const haystack = [
          a.tag, a.name, a.serial, a.brand, a.model, a.status, a.condition,
          a.assignedTo, a.notes, a.department, loc, cat,
        ].join(" ").toLowerCase();
        return haystack.includes(q);
      });
    }
    return list;
  }, [scopedAssetsBase, data.locations, data.categories, search, locationFilter, categoryFilter, userFilter]);

  const sortedAssets = useMemo(() => {
    if (!sort.key) return visibleAssets;
    const getVal = (a) => {
      switch (sort.key) {
        case "category": return data.categories.find((c) => c.id === a.categoryId)?.name || "";
        case "tag": return a.tag || "";
        case "name": return a.name || "";
        case "location": return data.locations.find((l) => l.id === a.locationId)?.name || "";
        case "assignedTo": return a.assignedTo || "";
        case "status": return a.status || "";
        case "condition": return a.condition || "";
        default: return "";
      }
    };
    const list = [...visibleAssets];
    list.sort((a, b) => {
      const cmp = String(getVal(a)).localeCompare(String(getVal(b)), undefined, { sensitivity: "base" });
      return sort.dir === "asc" ? cmp : -cmp;
    });
    return list;
  }, [visibleAssets, sort, data.categories, data.locations]);

  const toggleSort = (key) => {
    setSort((s) => (s.key === key ? { key, dir: s.dir === "asc" ? "desc" : "asc" } : { key, dir: "asc" }));
  };

  const SortTh = ({ label, sortKey, style }) => (
    <th style={style}>
      <button type="button" className="sort-th-btn" onClick={() => toggleSort(sortKey)}>
        {label}
        {sort.key === sortKey
          ? (sort.dir === "asc" ? <ChevronUp size={13} /> : <ChevronDown size={13} />)
          : <ChevronsUpDown size={13} className="sort-th-idle" />}
      </button>
    </th>
  );

  const save = async (asset) => {
    const { repairReason, _lockLocationField, ...assetFields } = asset;
    let next;
    let autoMaint = null;
    if (assetFields.id) {
      const prev = data.assets.find((a) => a.id === assetFields.id);
      let finalAsset = assetFields;
      if (assetFields.status === "Under Repair" && prev?.status !== "Under Repair") {
        finalAsset = { ...assetFields, preRepairStatus: prev?.status || "In Use" };
        autoMaint = { id: uid("maint"), assetId: assetFields.id, description: (repairReason || "").trim() || "Marked Under Repair from Assets", status: "Not Started", date: todayISO(), cost: "" };
      } else if (assetFields.status !== "Under Repair" && prev?.status === "Under Repair") {
        finalAsset = { ...assetFields, preRepairStatus: null };
      }
      finalAsset = applyStatusSideEffects(finalAsset);
      next = withLog({
        ...data,
        assets: data.assets.map((a) => (a.id === assetFields.id ? finalAsset : a)),
        maintenance: autoMaint ? [autoMaint, ...data.maintenance] : data.maintenance,
      }, currentUser, `Edited asset "${assetFields.name || assetFields.tag}"${autoMaint ? " — added maintenance entry (status: Under Repair)" : ""}`, finalAsset.locationId);
    } else {
      let newAsset = {
        ...assetFields,
        id: uid("ast"),
        tag: String(assetFields.tag || "").trim() || nextAutoTag(data.assets),
        createdById: currentUser.id,
        createdByName: currentUser.name,
      };
      if (newAsset.status === "Under Repair") {
        autoMaint = { id: uid("maint"), assetId: newAsset.id, description: (repairReason || "").trim() || "Marked Under Repair from Assets", status: "Not Started", date: todayISO(), cost: "" };
      }
      newAsset = applyStatusSideEffects(newAsset);
      next = withLog({
        ...data,
        assets: [newAsset, ...data.assets],
        maintenance: autoMaint ? [autoMaint, ...data.maintenance] : data.maintenance,
      }, currentUser, `Added asset "${newAsset.name || newAsset.tag}"`, newAsset.locationId);
    }
    persist(next);
    setEditing(null);
    showToast(autoMaint ? "Asset saved — added to Maintenance." : "Asset saved.");
  };

  // Quick status change straight from the table row, without opening the
  // full Edit form. Under Repair is a special case — instead of silently
  // logging a generic maintenance entry, it pops up the maintenance form
  // so whoever's doing this can say *why*; the status only actually
  // changes once that form is saved (see submitRepairMaintenance below).
  // In Use is also a special case — it pops up a small form to capture
  // who it's now with, since that's normally the whole point of marking
  // something In Use; the status only actually changes once that's saved
  // (see submitInUseAssignment below).
  const quickStatusChange = async (asset, newStatus) => {
    if (newStatus === asset.status) return;
    if (newStatus === "Under Repair") {
      setMaintPromptAsset(asset);
      return;
    }
    if (newStatus === "In Use") {
      setInUseDepartment(asset.department || "");
      setInUseAssignedTo(asset.assignedTo || "");
      setInUseLocationId(asset.locationId || "");
      setInUsePromptAsset(asset);
      return;
    }
    let finalAsset = { ...asset, status: newStatus };
    if (asset.status === "Under Repair") finalAsset.preRepairStatus = null;
    finalAsset = applyStatusSideEffects(finalAsset);
    const next = withLog({
      ...data,
      assets: data.assets.map((a) => (a.id === asset.id ? finalAsset : a)),
    }, currentUser, `Changed status of asset "${asset.name || asset.tag}" to "${newStatus}"`, finalAsset.locationId);
    persist(next);
    showToast("Status updated.");
  };

  // Saves the department/assigned-user details collected from the In Use
  // popup, and only now actually flips the asset's status — cancelling
  // that popup leaves the asset untouched.
  const submitInUseAssignment = () => {
    const asset = inUsePromptAsset;
    if (!asset) return;
    let finalAsset = { ...asset, status: "In Use", department: inUseDepartment, assignedTo: inUseAssignedTo, locationId: inUseLocationId };
    if (asset.status === "Under Repair") finalAsset.preRepairStatus = null;
    finalAsset = applyStatusSideEffects(finalAsset);
    const next = withLog({
      ...data,
      assets: data.assets.map((a) => (a.id === asset.id ? finalAsset : a)),
    }, currentUser, `Changed status of asset "${asset.name || asset.tag}" to "In Use"`, finalAsset.locationId);
    persist(next);
    setInUsePromptAsset(null);
    showToast("Status updated.");
  };

  // Saves the maintenance entry collected from the Under Repair popup, and
  // only now actually flips the asset's status — cancelling that popup
  // leaves the asset untouched.
  const submitRepairMaintenance = (entry) => {
    const asset = maintPromptAsset;
    if (!asset) return;
    const newEntry = { ...entry, id: uid("maint"), assetId: asset.id };
    const finalAsset = { ...asset, status: "Under Repair", preRepairStatus: asset.status };
    const next = withLog({
      ...data,
      assets: data.assets.map((a) => (a.id === asset.id ? finalAsset : a)),
      maintenance: [newEntry, ...data.maintenance],
    }, currentUser, `Changed status of asset "${asset.name || asset.tag}" to "Under Repair" — added maintenance entry`, finalAsset.locationId);
    persist(next);
    setMaintPromptAsset(null);
    showToast("Status updated — added to Maintenance.");
  };

  const remove = async (id) => {
    const asset = data.assets.find((a) => a.id === id);
    const removedLogs = data.maintenance.filter((m) => m.assetId === id).length;
    const suffix = removedLogs > 0 ? ` (and ${removedLogs} maintenance record${removedLogs > 1 ? "s" : ""})` : "";
    const next = withLog({
      ...data,
      assets: data.assets.filter((a) => a.id !== id),
      maintenance: data.maintenance.filter((m) => m.assetId !== id),
    }, currentUser, `Deleted asset "${asset?.name || asset?.tag || id}"${suffix}`, asset?.locationId);
    persist(next);
    setConfirmDelete(null);
    setSelected((s) => s.filter((x) => x !== id));
    showToast("Asset deleted.");
  };

  // The actual bulk-delete operation. Only ever invoked after the admin has
  // typed DELETE into the confirmation modal below — never wired directly
  // to a button, since this is destructive and irreversible.
  const performBulkDelete = async () => {
    const removedLogs = data.maintenance.filter((m) => selected.includes(m.assetId)).length;
    const suffix = removedLogs > 0 ? ` (and ${removedLogs} maintenance record${removedLogs > 1 ? "s" : ""})` : "";
    const next = withLog({
      ...data,
      assets: data.assets.filter((a) => !selected.includes(a.id)),
      maintenance: data.maintenance.filter((m) => !selected.includes(m.assetId)),
    }, currentUser, `Deleted ${selected.length} asset(s) in bulk${suffix}`);
    persist(next);
    setSelected([]);
    setBulkDeleteConfirmOpen(false);
    setBulkDeleteText("");
    showToast(`${selected.length} asset(s) deleted.`);
  };

  // Clicking "Delete (N)" never deletes directly — it opens a second
  // confirmation modal that requires typing DELETE, since bulk-deleting
  // many assets at once is easy to trigger by accident.
  const bulkDelete = () => setBulkDeleteConfirmOpen(true);

  // Posts a comment on an asset. Notifies every user account currently
  // assigned to manage that asset's location/country (e.g. Eve Yew for
  // Singapore) plus admins, and anyone who has already commented in this
  // asset's thread AND still has access to the asset's current location
  // (so a reply reaches whoever it's replying to) — since those are the
  // only people with app access who should hear about it. The "Assigned
  // To" field is just free text for an employee with no app account, so
  // it's never notified.
  //
  // Deliberately does NOT unconditionally include the asset's original
  // creator or every past commenter: once an asset is transferred to a
  // different location, whoever isn't an admin or currently assigned to
  // the new location should stop hearing about it.
  const addComment = (assetId, message) => {
    const text = (message || "").trim();
    if (!text) return;
    const asset = data.assets.find((a) => a.id === assetId);
    const hasCurrentAccess = (userId) => {
      const u = data.users.find((x) => x.id === userId);
      return !!u && (u.role === "Admin" || (u.locationId && u.locationId === asset?.locationId));
    };
    const locationUserIds = data.users
      .filter((u) => u.locationId && u.locationId === asset?.locationId)
      .map((u) => u.id);
    const priorParticipantIds = (data.comments || [])
      .filter((c) => c.assetId === assetId)
      .map((c) => c.authorId)
      .filter(hasCurrentAccess);
    const createdById = hasCurrentAccess(asset?.createdById) ? asset?.createdById : null;
    const targetUserIds = Array.from(new Set(
      [createdById, ...locationUserIds, ...priorParticipantIds].filter((id) => id && id !== currentUser.id)
    ));
    const comment = {
      id: uid("cmt"),
      assetId,
      at: new Date().toISOString(),
      authorId: currentUser.id,
      authorName: currentUser.name,
      message: text,
      targetUserIds,
      readBy: [],
    };
    const next = withLog({
      ...data,
      comments: [comment, ...(data.comments || [])],
    }, currentUser, `Commented on asset "${asset?.name || asset?.tag}"`, asset?.locationId);
    persist(next);
    showToast("Comment sent.");
  };

  // Starts a "New Asset" draft. Non-admins are fixed to their own assigned
  // location, same as when duplicating — the field is locked in the modal.
  const newAssetDraft = () => ({
    ...emptyAsset(scopedLocationId),
    _lockLocationField: !isAdmin,
  });

  const duplicateAsset = (asset) => {
    setEditing({
      ...asset,
      id: null,
      tag: "",
      serial: "",
      assignedTo: "",
      status: "In Stock",
      transferHistory: [],
      // Non-admins can't relocate a duplicated asset — it always starts
      // in their own assigned location, and the field is locked in the
      // modal below. Admins keep full control over the location.
      locationId: !isAdmin && scopedLocationId ? scopedLocationId : asset.locationId,
      _lockLocationField: !isAdmin,
    });
  };

  const startTransfer = (asset) => {
    setTransferTarget(asset.id);
    setTransferLocationId("");
    setTransferNewLocationName("");
    setTransferAssignedTo(asset.assignedTo || "");
    setTransferReason("");
  };

  const submitTransfer = () => {
    const creatingLocation = transferLocationId === "__new__";
    if (creatingLocation && !isAdmin) { alert("Only an Administrator can add a new location."); return; }
    if (!transferLocationId) { alert("Please select a destination location."); return; }
    if (creatingLocation && !transferNewLocationName.trim()) { alert("Please enter a name for the new location."); return; }
    if (!transferReason.trim()) { alert("Please enter a reason for this transfer."); return; }
    const asset = data.assets.find((a) => a.id === transferTarget);
    if (!asset) return;

    let locations = data.locations;
    let destLocationId = transferLocationId;
    if (creatingLocation) {
      const newLoc = { id: uid("loc"), name: transferNewLocationName.trim() };
      locations = [...data.locations, newLoc];
      destLocationId = newLoc.id;
    }

    const fromLoc = data.locations.find((l) => l.id === asset.locationId)?.name || "Unknown";
    const toLoc = locations.find((l) => l.id === destLocationId)?.name || "Unknown";
    const newAssignedTo = transferAssignedTo.trim();
    const assignedChanged = newAssignedTo !== (asset.assignedTo || "");

    const transferEntry = {
      id: uid("xfer"),
      fromLocationId: asset.locationId,
      fromLocationName: fromLoc,
      toLocationId: destLocationId,
      toLocationName: toLoc,
      reason: transferReason.trim(),
      by: currentUser.name,
      at: new Date().toISOString(),
    };
    const logBits = [`Transferred asset "${asset.name || asset.tag}" from ${fromLoc} to ${toLoc}`];
    if (assignedChanged) logBits.push(`reassigned to ${newAssignedTo || "Unassigned"}`);
    logBits.push(`reason: ${transferReason.trim()}`);

    const next = withLog({
      ...data,
      locations,
      assets: data.assets.map((a) => (a.id === transferTarget
        ? { ...a, locationId: destLocationId, assignedTo: newAssignedTo, transferHistory: [transferEntry, ...(a.transferHistory || [])] }
        : a)),
    }, currentUser, logBits.join(" — "), destLocationId);
    persist(next);
    setTransferTarget(null);
    setTransferReason("");
    setTransferNewLocationName("");
    showToast("Asset transferred.");
  };

  // Non-admins can't delete outright — they submit a reason, and the asset
  // is flagged for the Admin to approve or reject under "Approvals".
  const submitDeleteRequest = async () => {
    if (!deleteReason.trim()) { alert("Please enter a reason for this deletion request."); return; }
    const asset = data.assets.find((a) => a.id === requestDeleteTarget);
    const next = withLog({
      ...data,
      assets: data.assets.map((a) => (a.id === requestDeleteTarget
        ? { ...a, pendingDeletion: { requestedBy: currentUser.id, requestedByName: currentUser.name, reason: deleteReason.trim(), requestedAt: new Date().toISOString() } }
        : a)),
    }, currentUser, `Requested deletion of asset "${asset?.name || asset?.tag}" — reason: ${deleteReason.trim()}`, asset?.locationId);
    persist(next);
    setRequestDeleteTarget(null);
    setDeleteReason("");
    showToast("Deletion request sent for Admin approval.");
  };

  const startDelete = (asset) => {
    if (isAdmin) {
      setConfirmDelete(asset.id);
    } else if (!asset.pendingDeletion) {
      setRequestDeleteTarget(asset.id);
    }
  };

  const EXPORT_COLS = ["tag", "name", "department", "assetType", "brand", "model", "yearModel", "serial", "status", "condition", "location", "assignedTo", "purchaseDate", "purchaseCost", "warrantyExpiry", "requiresCalibration", "calibrationDate", "nextCalibrationDate"];

  const exportExcel = () => {
    const rows = visibleAssets.map((a) => {
      const loc = data.locations.find((l) => l.id === a.locationId)?.name || "";
      return {
        tag: a.tag, name: a.name, department: a.department || "", assetType: a.assetType,
        brand: a.brand || "", model: a.model || "", yearModel: a.yearModel || "", serial: a.serial || "",
        status: a.status, condition: a.condition, location: loc, assignedTo: a.assignedTo || "",
        purchaseDate: a.purchaseDate || "", purchaseCost: a.purchaseCost || "", warrantyExpiry: a.warrantyExpiry || "",
        requiresCalibration: a.requiresCalibration ? "Yes" : "No",
        calibrationDate: a.calibrationDate || "", nextCalibrationDate: a.nextCalibrationDate || "",
      };
    });
    const wb = XLSX.utils.book_new();
    const sheet = XLSX.utils.json_to_sheet(rows, { header: EXPORT_COLS });
    XLSX.utils.book_append_sheet(wb, sheet, "Assets");
    downloadWorkbook(wb, "assets-export.xlsx");
    showToast("Assets exported.");
  };

  const triggerImport = () => fileInputRef.current?.click();

  const handleImportFile = async (e) => {
    const file = e.target.files?.[0];
    e.target.value = "";
    if (!file) return;
    try {
      const wb = await readWorkbookFile(file);
      const sheetName = wb.SheetNames.includes("Assets") ? "Assets" : wb.SheetNames[0];
      const rows = sheetRows(wb, sheetName);
      if (rows.length === 0) { showToast("Excel file has no data rows."); return; }
      const newAssets = rows.map((row, i) => {
        const loc = data.locations.find((l) => l.name.toLowerCase() === String(row.location || "").toLowerCase());
        const cat = data.categories.find((c) => c.name.toLowerCase() === String(row.name || "").toLowerCase()) ||
          data.categories.find((c) => (row.assetType || "IT") === c.type);
        return {
          id: uid("ast"),
          tag: row.tag || `AST-IMP-${String(i + 1).padStart(3, "0")}`,
          name: row.name || "Imported Asset",
          department: row.department || "",
          categoryId: cat?.id || data.categories[0]?.id || "",
          assetType: row.assetType === "Non-IT" ? "Non-IT" : "IT",
          brand: row.brand || "",
          model: row.model || "",
          yearModel: row.yearModel || "",
          serial: row.serial || "",
          status: STATUS_OPTIONS.includes(row.status) ? row.status : "In Stock",
          condition: CONDITION_OPTIONS.includes(row.condition) ? row.condition : "Good",
          locationId: loc?.id || scopedLocationId || data.locations[0]?.id || "",
          assignedTo: row.assignedTo || "",
          purchaseDate: row.purchaseDate || todayISO(),
          purchaseCost: Number(row.purchaseCost) || 0,
          warrantyExpiry: (row.warrantyExpiry && row.warrantyExpiry !== "N/A") ? row.warrantyExpiry : "",
          requiresCalibration: String(row.requiresCalibration || "").toLowerCase() === "yes",
          calibrationDate: row.calibrationDate || "",
          nextCalibrationDate: row.nextCalibrationDate || "",
          notes: "",
          transferHistory: [],
        };
      });
      persist(withLog({ ...data, assets: [...newAssets, ...data.assets] }, currentUser, `Imported ${newAssets.length} asset(s) via Excel`));
      showToast(`Imported ${newAssets.length} asset(s).`);
    } catch {
      showToast("Could not parse this Excel file.");
    }
  };

  return (
    <div>
      <div className="view-head">
        <div className="search-box">
          <Search size={15} />
          <input placeholder="Search anything — tag, name, brand, location, category, user…" value={search} onChange={(e) => setSearch(e.target.value)} />
        </div>
        <div className="filter-group">
          <select className="sort-select" value={locationFilter} onChange={(e) => setLocationFilter(e.target.value)}>
            <option value="all">All Locations</option>
            {locationOptions.map((l) => <option key={l.id} value={l.id}>{l.name}</option>)}
          </select>
          <select className="sort-select" value={categoryFilter} onChange={(e) => setCategoryFilter(e.target.value)}>
            <option value="all">All Categories</option>
            {categoryOptions.map((c) => <option key={c.id} value={c.id}>{c.name}</option>)}
          </select>
          <select className="sort-select" value={userFilter} onChange={(e) => setUserFilter(e.target.value)}>
            <option value="all">All Users</option>
            {assignedUserOptions.map((u) => <option key={u} value={u}>{u}</option>)}
          </select>
        </div>
        <div className="view-actions">
          {isAdmin && selected.length > 0 && (
            <button className="btn danger" onClick={bulkDelete}>
              <Trash2 size={14} /> Delete ({selected.length})
            </button>
          )}
          {isAdmin && (
            <>
              <input ref={fileInputRef} type="file" accept=".xlsx" style={{ display: "none" }} onChange={handleImportFile} />
              <button className="btn ghost" onClick={triggerImport}><Upload size={14} /> Import Excel</button>
              <button className="btn ghost" onClick={exportExcel}><Download size={14} /> Export Excel</button>
            </>
          )}
          <button className="btn primary new-asset-btn" onClick={() => setEditing(newAssetDraft())}>
            <Plus size={14} /> New Asset
          </button>
        </div>
      </div>

      <button className="fab-add" onClick={() => setEditing(newAssetDraft())} title="New Asset">
        <Plus size={22} />
      </button>

      {(assetAlertMap.size > 0 || unreadCommentAssetIds.size > 0) && (
        <div className="table-legend">
          {unreadCommentAssetIds.size > 0 && (
            <span className="table-legend-item"><MessageCircle size={13} className="comment-flag" /> New comment</span>
          )}
          {assetAlertMap.size > 0 && (
            <>
              <span className="table-legend-item"><Clock size={13} className="alert-flag" /> Warranty/calibration due soon</span>
              <span className="table-legend-item"><AlertTriangle size={13} className="alert-flag alert-flag-urgent" /> Expired/overdue</span>
            </>
          )}
        </div>
      )}

      <div className="panel">
        <div className="table-wrap">
          <table>
            <thead>
              <tr>
                {isAdmin && (
                  <th style={{ width: 32 }}>
                    <input
                      type="checkbox"
                      checked={visibleAssets.length > 0 && selected.length === visibleAssets.length}
                      onChange={(e) => setSelected(e.target.checked ? visibleAssets.map((a) => a.id) : [])}
                    />
                  </th>
                )}
                <SortTh label="Category" sortKey="category" />
                <SortTh label="Asset Tag" sortKey="tag" />
                <SortTh label="Name" sortKey="name" />
                <SortTh label="Location" sortKey="location" />
                <SortTh label="Assigned User" sortKey="assignedTo" />
                <SortTh label="Status" sortKey="status" />
                <SortTh label="Condition" sortKey="condition" />
                <th className="actions-th" style={{ width: 150 }}></th>
              </tr>
            </thead>
            <tbody>
              {visibleAssets.length === 0 && (
                <tr><td colSpan={isAdmin ? 9 : 8} className="empty-cell">No assets yet — click "New Asset" to add one.</td></tr>
              )}
              {sortedAssets.map((a) => {
                const cat = data.categories.find((c) => c.id === a.categoryId);
                const loc = data.locations.find((l) => l.id === a.locationId);
                const brandModel = [a.brand, a.model].filter(Boolean).join(" / ");
                return (
                  <tr key={a.id} className="asset-row" onClick={() => setViewing(a)}>
                    {isAdmin && (
                      <td className="checkbox-cell" onClick={(e) => e.stopPropagation()}>
                        <input
                          type="checkbox"
                          checked={selected.includes(a.id)}
                          onChange={(e) =>
                            setSelected((s) => (e.target.checked ? [...s, a.id] : s.filter((x) => x !== a.id)))
                          }
                        />
                      </td>
                    )}
                    <td data-label="Category">
                      {cat && <span className="cat-dot" style={{ background: categoryColor(data.categories, a.categoryId) }} />}
                      {cat?.name || "—"}
                    </td>
                    <td data-label="Asset Tag">
                      <span className="link-tag" title="View details, history & comments">{a.tag}</span>
                      {unreadCommentAssetIds.has(a.id) && (
                        <MessageCircle size={13} className="comment-flag" title="New comment on this asset" />
                      )}
                      {assetAlertMap.has(a.id) && (() => {
                        const info = assetAlertMap.get(a.id);
                        const label = `${info.kind === "warranty" ? "Warranty" : "Calibration"} ${info.urgent ? (info.kind === "warranty" ? "expired" : "overdue") : (info.kind === "warranty" ? "expiring soon" : "due soon")}`;
                        return info.urgent
                          ? <AlertTriangle size={13} className="alert-flag alert-flag-urgent" title={label} />
                          : <Clock size={13} className="alert-flag" title={label} />;
                      })()}
                    </td>
                    <td data-label="Name">
                      {a.name}
                      {brandModel && <div className="name-subtext">{brandModel}</div>}
                    </td>
                    <td data-label="Location">{loc?.name || "—"}</td>
                    <td data-label="Assigned User">{a.assignedTo || "—"}</td>
                    <td data-label="Status" onClick={(e) => e.stopPropagation()}>
                      <select
                        className="status-select"
                        value={a.status}
                        onChange={(e) => quickStatusChange(a, e.target.value)}
                        style={{ color: STATUS_COLORS[a.status], borderColor: `${STATUS_COLORS[a.status]}55` }}
                        title="Quick-change status"
                      >
                        {STATUS_OPTIONS.map((s) => <option key={s} value={s}>{s}</option>)}
                      </select>
                      {a.pendingDeletion && (
                        <div style={{ marginTop: 4 }}>
                          <Badge color="#EF4444">Pending Deletion</Badge>
                        </div>
                      )}
                    </td>
                    <td data-label="Condition">{a.condition}</td>
                    <td className="actions-cell" onClick={(e) => e.stopPropagation()}>
                      <div className="row-actions">
                        <IconBtn icon={Pencil} title="Edit" onClick={() => setEditing(a)} />
                        <IconBtn icon={Copy} title="Duplicate" onClick={() => duplicateAsset(a)} />
                        <IconBtn icon={Truck} title="Transfer" onClick={() => startTransfer(a)} />
                      </div>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </div>

      {editing && (
        <AssetModal
          asset={editing}
          categories={data.categories}
          locations={data.locations}
          isAdmin={isAdmin}
          scopedLocationId={scopedLocationId}
          existingAssets={data.assets}
          departmentOptions={departmentOptions}
          onClose={() => setEditing(null)}
          onSave={save}
        />
      )}
      {confirmDelete && (
        <ConfirmDialog
          message="Delete this asset? This cannot be undone."
          onCancel={() => setConfirmDelete(null)}
          onConfirm={() => remove(confirmDelete)}
        />
      )}
      {maintPromptAsset && (
        <Modal title={`Log Repair — ${maintPromptAsset.tag}`} onClose={() => setMaintPromptAsset(null)} width={760}>
          <div className="form-full hint-box" style={{ marginBottom: 14 }}>
            This asset will be marked Under Repair once you save these details — they're added to Service &amp; Maintenance.
          </div>
          <MaintForm
            entry={emptyMaint(maintPromptAsset.id)}
            assets={[maintPromptAsset]}
            onSave={submitRepairMaintenance}
            onClose={() => setMaintPromptAsset(null)}
          />
        </Modal>
      )}
      {inUsePromptAsset && (
        <Modal title={`Mark In Use — ${inUsePromptAsset.tag}`} onClose={() => setInUsePromptAsset(null)} width={480}>
          <div className="form-full hint-box" style={{ marginBottom: 14 }}>
            This asset will be marked In Use once you save who it's with.
          </div>
          <div className="form-grid">
            <Field label="Department">
              <SearchableSelect
                value={inUseDepartment}
                onChange={setInUseDepartment}
                options={departmentOptions.map((v) => ({ value: v, label: v }))}
                placeholder="Search or type a department"
              />
            </Field>
            <Field label="Assigned To">
              <SearchableSelect
                value={inUseAssignedTo}
                onChange={setInUseAssignedTo}
                options={Array.from(new Set(data.assets.map((a) => a.assignedTo).filter(Boolean))).sort().map((v) => ({ value: v, label: v }))}
                placeholder="Search or type a name"
              />
            </Field>
            <Field label="Location">
              <SearchableSelect
                value={inUseLocationId}
                onChange={setInUseLocationId}
                options={data.locations.map((l) => ({ value: l.id, label: l.name }))}
                placeholder="Search locations"
                allowCustom={false}
              />
            </Field>
          </div>
          <div className="form-full modal-actions">
            <button type="button" className="btn ghost" onClick={() => setInUsePromptAsset(null)}>Cancel</button>
            <button type="button" className="btn primary" onClick={submitInUseAssignment}>Save</button>
          </div>
        </Modal>
      )}
      {bulkDeleteConfirmOpen && (
        <TypeToConfirmDialog
          title="Delete Multiple Assets"
          message={`You're about to permanently delete ${selected.length} asset(s)${
            data.maintenance.filter((m) => selected.includes(m.assetId)).length > 0
              ? " along with their maintenance records"
              : ""
          }. This cannot be undone.`}
          confirmWord="DELETE"
          value={bulkDeleteText}
          onChange={setBulkDeleteText}
          onCancel={() => { setBulkDeleteConfirmOpen(false); setBulkDeleteText(""); }}
          onConfirm={performBulkDelete}
        />
      )}
      {requestDeleteTarget && (
        <Modal title="Request Deletion" onClose={() => { setRequestDeleteTarget(null); setDeleteReason(""); }} width={760}>
          <div className="form-grid">
            <div className="form-full hint-box">
              This asset won't be deleted right away — your request and reason go to an
              Admin for approval first.
            </div>
            <div className="form-full">
              <Field label="Reason for deletion">
                <textarea value={deleteReason} onChange={(e) => setDeleteReason(e.target.value)} rows={3} placeholder="e.g. Disposed, duplicate entry, no longer in use..." autoFocus />
              </Field>
            </div>
            <div className="form-full modal-actions">
              <button type="button" className="btn ghost" onClick={() => { setRequestDeleteTarget(null); setDeleteReason(""); }}>Cancel</button>
              <button type="button" className="btn danger" onClick={submitDeleteRequest}>Submit Request</button>
            </div>
          </div>
        </Modal>
      )}
      {viewing && (
        <AssetDetailModal
          asset={viewing}
          categories={data.categories}
          locations={data.locations}
          isAdmin={isAdmin}
          comments={(data.comments || []).filter((c) => c.assetId === viewing.id)}
          maintenance={data.maintenance.filter((m) => m.assetId === viewing.id)}
          currentUser={currentUser}
          onAddComment={(message) => addComment(viewing.id, message)}
          onClose={() => setViewing(null)}
          onDelete={() => { setViewing(null); startDelete(viewing); }}
        />
      )}
      {transferTarget && (
        <Modal title="Transfer Asset" onClose={() => setTransferTarget(null)} width={760}>
          <div className="form-grid">
            <div className="form-full hint-box">
              Moving this asset to a different location will be recorded in its history along with your reason.
            </div>
            <div className="form-full">
              <Field label="Destination Location">
                <select value={transferLocationId} onChange={(e) => setTransferLocationId(e.target.value)}>
                  <option value="">Select location</option>
                  {data.locations.filter((l) => l.id !== data.assets.find((a) => a.id === transferTarget)?.locationId).map((l) => (
                    <option key={l.id} value={l.id}>{l.name}</option>
                  ))}
                  {isAdmin && <option value="__new__">+ Add new location…</option>}
                </select>
              </Field>
            </div>
            {transferLocationId === "__new__" && (
              <div className="form-full">
                <Field label="New Location Name">
                  <input value={transferNewLocationName} onChange={(e) => setTransferNewLocationName(e.target.value)} placeholder="e.g. Cebu Branch" autoFocus />
                </Field>
              </div>
            )}
            <div className="form-full">
              <Field label="Assign To (optional)">
                <input
                  list="transfer-assigned-user-list"
                  value={transferAssignedTo}
                  onChange={(e) => setTransferAssignedTo(e.target.value)}
                  placeholder="Keep current user, or type a new one"
                />
                <datalist id="transfer-assigned-user-list">
                  {assignedUserOptions.map((u) => <option key={u} value={u} />)}
                </datalist>
              </Field>
            </div>
            <div className="form-full">
              <Field label="Reason for transfer">
                <textarea value={transferReason} onChange={(e) => setTransferReason(e.target.value)} rows={3} placeholder='e.g. "Shipped to Philippines for new hire"' />
              </Field>
            </div>
            <div className="form-full modal-actions">
              <button type="button" className="btn ghost" onClick={() => setTransferTarget(null)}>Cancel</button>
              <button type="button" className="btn primary" onClick={submitTransfer}>Confirm Transfer</button>
            </div>
          </div>
        </Modal>
      )}
    </div>
  );
}

/* ---------------------------------------------------------
   New Asset form building blocks — section cards, toggle switch,
   a colored-badge status picker, and a lightweight searchable
   combobox (no external deps). Used by AssetModal below.
--------------------------------------------------------- */
function FormSection({ icon: Icon, title, children, action }) {
  return (
    <div className="form-section">
      <div className="form-section-head">
        <span className="form-section-head-icon"><Icon size={14} /></span>
        <h4>{title}</h4>
        {action && <div className="form-section-head-action">{action}</div>}
      </div>
      <div className="section-grid">{children}</div>
    </div>
  );
}

function ToggleSwitch({ checked, onChange, label }) {
  return (
    <label className="toggle-switch-wrap">
      <span className="toggle-switch">
        <input type="checkbox" checked={checked} onChange={(e) => onChange(e.target.checked)} />
        <span className="toggle-slider" />
      </span>
      {label && <span>{label}</span>}
    </label>
  );
}

// Custom colored-badge dropdown for Status, so each option reads as a
// status pill (matching the colors already used elsewhere in the app)
// instead of a plain text list.
function StatusPicker({ value, onChange, placeholder = "Select status" }) {
  const [open, setOpen] = useState(false);
  const ref = useRef(null);

  useEffect(() => {
    function onDocClick(e) {
      if (ref.current && !ref.current.contains(e.target)) setOpen(false);
    }
    document.addEventListener("mousedown", onDocClick);
    return () => document.removeEventListener("mousedown", onDocClick);
  }, []);

  const color = STATUS_COLORS[value] || "#9CA3AF";

  return (
    <div className="status-picker" ref={ref}>
      <button type="button" className="status-picker-trigger" onClick={() => setOpen((o) => !o)}>
        {value ? (
          <span className="status-dot-badge" style={{ background: `${color}1a`, color, border: `1px solid ${color}40` }}>
            <span className="status-dot" style={{ background: color }} />
            {value}
          </span>
        ) : (
          <span className="status-picker-placeholder">{placeholder}</span>
        )}
        <ChevronDown size={14} />
      </button>
      {open && (
        <div className="status-picker-menu">
          {STATUS_OPTIONS.map((s) => {
            const c = STATUS_COLORS[s];
            return (
              <button
                type="button"
                key={s}
                className="status-picker-option"
                onMouseDown={() => { onChange(s); setOpen(false); }}
              >
                <span className="status-dot" style={{ background: c }} />
                {s}
              </button>
            );
          })}
        </div>
      )}
    </div>
  );
}

// Lightweight searchable combobox — filters a list of {value, label}
// options as you type. With allowCustom (Department / Assigned To / Brand)
// the typed text itself is the value, and picking a suggestion just fills
// it in. Without it (Location) the value must be one of the option ids.
function SearchableSelect({ value, onChange, options, placeholder = "Search…", allowCustom = true, disabled = false }) {
  const [open, setOpen] = useState(false);
  const [query, setQuery] = useState("");
  const ref = useRef(null);

  const selectedLabel = useMemo(() => {
    if (allowCustom) return value || "";
    const opt = options.find((o) => o.value === value);
    return opt ? opt.label : "";
  }, [value, options, allowCustom]);

  useEffect(() => {
    if (!open) setQuery(selectedLabel);
  }, [selectedLabel, open]);

  useEffect(() => {
    function onDocClick(e) {
      if (ref.current && !ref.current.contains(e.target)) setOpen(false);
    }
    document.addEventListener("mousedown", onDocClick);
    return () => document.removeEventListener("mousedown", onDocClick);
  }, []);

  const filtered = useMemo(() => {
    const q = query.trim().toLowerCase();
    if (!q) return options;
    return options.filter((o) => o.label.toLowerCase().includes(q));
  }, [query, options]);

  const pick = (opt) => {
    onChange(allowCustom ? opt.label : opt.value);
    setQuery(opt.label);
    setOpen(false);
  };

  return (
    <div className="searchable-select" ref={ref}>
      <input
        value={open ? query : selectedLabel}
        onChange={(e) => {
          setQuery(e.target.value);
          setOpen(true);
          if (allowCustom) onChange(e.target.value);
        }}
        onFocus={() => { if (disabled) return; setOpen(true); setQuery(selectedLabel); }}
        placeholder={placeholder}
        disabled={disabled}
      />
      {!allowCustom && value && !disabled && (
        <button type="button" className="searchable-select-clear" title="Clear" onMouseDown={(e) => { e.preventDefault(); onChange(""); setQuery(""); }}>
          <X size={12} />
        </button>
      )}
      {!disabled && open && filtered.length > 0 && (
        <div className="searchable-select-menu">
          {filtered.map((o) => (
            <div key={o.value} className="searchable-select-option" onMouseDown={() => pick(o)}>
              {o.label}
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

function AssetModal({ asset, categories, locations, isAdmin, scopedLocationId, existingAssets, departmentOptions, onClose, onSave }) {
  const [form, setForm] = useState(asset);
  const [hasPurchaseInfo, setHasPurchaseInfo] = useState(!!(asset.purchaseDate || asset.purchaseCost));
  const [hasWarrantyExpiry, setHasWarrantyExpiry] = useState(!!(asset.warrantyExpiry && asset.warrantyExpiry !== "N/A"));
  // ^ still checks for legacy "N/A" values from before this fix, so old assets open with the toggle correctly off
  // "New Asset" starts with the tag auto-generated (blank tag → server
  // assigns the next ASTUTE### on save). Editing an asset always shows its
  // real tag — auto-generate only applies at creation time.
  const [autoTag, setAutoTag] = useState(!asset.id && !asset.tag);
  // Optional Information starts open if there's already something in it
  // (editing an asset that has purchase/warranty/notes data), otherwise
  // collapsed so the main form stays short.
  const [optionalOpen, setOptionalOpen] = useState(
    !!(asset.purchaseDate || asset.purchaseCost || (asset.warrantyExpiry && asset.warrantyExpiry !== "N/A") || asset.notes)
  );
  const set = (k, v) => setForm((f) => ({ ...f, [k]: v }));

  const wasUnderRepair = asset.status === "Under Repair";
  const needsReason = form.status === "Under Repair" && !wasUnderRepair;

  const brandOptions = useMemo(() => {
    const brandSet = new Set((existingAssets || []).map((a) => a.brand).filter(Boolean));
    return Array.from(brandSet).sort().map((v) => ({ value: v, label: v }));
  }, [existingAssets]);
  const assignedToOptions = useMemo(() => {
    const nameSet = new Set((existingAssets || []).map((a) => a.assignedTo).filter(Boolean));
    return Array.from(nameSet).sort().map((v) => ({ value: v, label: v }));
  }, [existingAssets]);
  const departmentSelectOptions = useMemo(
    () => (departmentOptions || []).map((v) => ({ value: v, label: v })),
    [departmentOptions]
  );
  const locationSelectOptions = useMemo(
    () => locations.map((l) => ({ value: l.id, label: l.name })),
    [locations]
  );

  // Typing an Assigned To name is a strong signal the asset is in use — set
  // it automatically, but only if the user hasn't already picked a status
  // themselves, so we never override a deliberate choice.
  const onAssignedToChange = (v) => {
    setForm((f) => ({
      ...f,
      assignedTo: v,
      status: !f.status && v.trim() ? "In Use" : f.status,
    }));
  };

  // Switching Asset Type narrows which categories are valid — clear the
  // category if it no longer belongs to the newly selected type.
  const onAssetTypeChange = (type) => {
    setForm((f) => {
      const stillValid = categories.some((c) => c.id === f.categoryId && c.type === type);
      return { ...f, assetType: type, categoryId: stillValid ? f.categoryId : "" };
    });
  };

  // Smart default: picking a category on a brand-new asset suggests a
  // sensible starting status (Condition already defaults to "New"), so the
  // common case — adding a fresh, in-stock item — needs no extra clicks.
  // Never overrides a status the user already chose, and never applies to
  // edits of an existing asset.
  const onCategoryChange = (catId) => {
    setForm((f) => {
      const next = { ...f, categoryId: catId };
      if (!asset.id && catId && !f.status) next.status = "In Stock";
      return next;
    });
  };

  const onAutoTagToggle = (checked) => {
    setAutoTag(checked);
    if (checked) set("tag", "");
  };

  const submit = (e) => {
    if (e && e.preventDefault) e.preventDefault();
    if (!form.name.trim()) {
      alert("Please enter an asset name.");
      return;
    }
    if (!form.categoryId) {
      alert("Please select a category.");
      return;
    }
    if (!form.status) {
      alert("Please select a status.");
      return;
    }
    const tag = String(form.tag || "").trim();
    if (tag) {
      const dup = (existingAssets || []).find(
        (a) => a.id !== form.id && String(a.tag || "").trim().toLowerCase() === tag.toLowerCase()
      );
      if (dup) {
        alert(`Asset Tag "${tag}" is already used by "${dup.name || "another asset"}". Please use a unique tag.`);
        return;
      }
    }
    if (needsReason && !(form.repairReason || "").trim()) {
      alert("Please enter a reason — this creates the matching Maintenance entry.");
      return;
    }
    onSave({
      ...form,
      purchaseDate: hasPurchaseInfo ? form.purchaseDate : "",
      purchaseCost: hasPurchaseInfo ? form.purchaseCost : "",
      warrantyExpiry: form.assetType === "IT" ? (hasWarrantyExpiry ? form.warrantyExpiry : "") : form.warrantyExpiry,
    });
  };

  return (
    <Modal title={asset.id ? "Edit Asset" : "New Asset"} onClose={onClose} width={820}>
      <form onSubmit={submit}>
        <FormSection icon={Package} title="Asset Information">
          <Field label="Asset Name" required>
            <input value={form.name} onChange={(e) => set("name", e.target.value)} autoFocus required />
          </Field>
          <Field
            label="Asset Tag"
            className="field-with-action"
            action={!asset.id && <ToggleSwitch checked={autoTag} onChange={onAutoTagToggle} label="Auto Generate" />}
          >
            <input
              value={form.tag}
              onChange={(e) => set("tag", e.target.value)}
              disabled={autoTag}
              placeholder={autoTag ? "Will be generated automatically (e.g. ASTUTE004)" : "e.g. ASTUTE004"}
            />
          </Field>
          <Field label="Asset Type" required>
            <select value={form.assetType} onChange={(e) => onAssetTypeChange(e.target.value)}>
              <option value="IT">IT Asset</option>
              <option value="Non-IT">Non-IT Asset</option>
            </select>
          </Field>
          <Field label="Category" required>
            <select value={form.categoryId} onChange={(e) => onCategoryChange(e.target.value)} required>
              <option value="">Select category</option>
              {categories.filter((c) => c.type === form.assetType).map((c) => (
                <option key={c.id} value={c.id}>{c.name}</option>
              ))}
            </select>
          </Field>
        </FormSection>

        <FormSection icon={Tags} title="Device Details">
          <Field label="Brand">
            <SearchableSelect value={form.brand} onChange={(v) => set("brand", v)} options={brandOptions} placeholder="Search or type a brand" />
          </Field>
          <Field label="Model">
            <input value={form.model} onChange={(e) => set("model", e.target.value)} />
          </Field>
          <Field label="Serial Number">
            <div className="field-inline">
              <input value={form.serial} onChange={(e) => set("serial", e.target.value)} />
              <button
                type="button"
                className="btn ghost sn-check-btn"
                title="Copy the serial number and open the vendor's warranty lookup in a side window"
                onClick={() => {
                  if (form.serial && navigator.clipboard?.writeText) {
                    navigator.clipboard.writeText(form.serial).catch(() => {});
                  }
                  const w = 480, h = 760;
                  const left = (window.screenX || 0) + (window.outerWidth || w);
                  const top = window.screenY || 0;
                  window.open(warrantyLookupUrl(form.brand, form.serial), "snChecker", `width=${w},height=${h},left=${left},top=${top}`);
                }}
              >
                <ExternalLink size={13} /> Verify
              </button>
            </div>
          </Field>
          <Field label="Manufactured Year">
            <input
              type="number"
              value={form.yearModel}
              onChange={(e) => set("yearModel", e.target.value)}
              placeholder="e.g. 2024"
              min="1990"
              max="2100"
            />
          </Field>
          <Field label="Status" required>
            <StatusPicker value={form.status} onChange={(v) => set("status", v)} />
          </Field>
          <Field label="Condition">
            <select value={form.condition} onChange={(e) => set("condition", e.target.value)}>
              {CONDITION_OPTIONS.map((s) => <option key={s} value={s}>{s}</option>)}
            </select>
          </Field>

          {needsReason && (
            <div className="form-full repair-reason-box">
              <Field label="Reason for Repair">
                <textarea
                  value={form.repairReason || ""}
                  onChange={(e) => set("repairReason", e.target.value)}
                  rows={2}
                  placeholder="What's wrong with it? This becomes the Maintenance entry."
                />
              </Field>
            </div>
          )}

          {form.assetType === "Non-IT" && (
            <div className="form-full">
              <ToggleSwitch
                checked={!!form.requiresCalibration}
                onChange={(checked) => set("requiresCalibration", checked)}
                label="Requires Calibration"
              />
            </div>
          )}
          {form.assetType === "Non-IT" && form.requiresCalibration && (
            <>
              <Field label="Calibration Date">
                <input type="date" value={form.calibrationDate} onChange={(e) => set("calibrationDate", e.target.value)} />
              </Field>
              <Field label="Next Recalibration Date">
                <input type="date" value={form.nextCalibrationDate} onChange={(e) => set("nextCalibrationDate", e.target.value)} />
              </Field>
            </>
          )}
        </FormSection>

        <FormSection icon={User} title="Assignment">
          <Field label="Department">
            <SearchableSelect value={form.department || ""} onChange={(v) => set("department", v)} options={departmentSelectOptions} placeholder="Search or type a department" />
          </Field>
          <Field label="Assigned To">
            <SearchableSelect value={form.assignedTo} onChange={onAssignedToChange} options={assignedToOptions} placeholder="Search or type a name" />
          </Field>
          <div className="form-full">
            <Field label="Location">
              <SearchableSelect
                value={form.locationId}
                onChange={(v) => set("locationId", v)}
                options={locationSelectOptions}
                placeholder="Search locations"
                allowCustom={false}
                disabled={!!asset._lockLocationField}
              />
              {!!asset._lockLocationField && (
                <span className="field-hint">Fixed to your assigned location.</span>
              )}
            </Field>
          </div>
        </FormSection>

        <div className="form-section optional-section">
          <button type="button" className="optional-section-toggle" onClick={() => setOptionalOpen((o) => !o)}>
            <span className="form-section-head" style={{ marginBottom: 0 }}>
              <span className="form-section-head-icon"><Info size={14} /></span>
              <h4>Optional Information</h4>
            </span>
            {optionalOpen ? <ChevronUp size={16} /> : <ChevronDown size={16} />}
          </button>
          {optionalOpen && (
            <div className="section-grid optional-section-body">
              <div className="form-full">
                <ToggleSwitch checked={hasPurchaseInfo} onChange={(yes) => {
                  setHasPurchaseInfo(yes);
                  if (yes && !form.purchaseDate) set("purchaseDate", todayISO());
                }} label="Add purchase info" />
              </div>
              {hasPurchaseInfo && (
                <>
                  <Field label="Purchase Date"><input type="date" value={form.purchaseDate} onChange={(e) => set("purchaseDate", e.target.value)} /></Field>
                  <Field label="Purchase Cost"><input type="number" value={form.purchaseCost} onChange={(e) => set("purchaseCost", e.target.value)} /></Field>
                </>
              )}
              {form.assetType === "IT" && (
                <div className="form-full">
                  <ToggleSwitch checked={hasWarrantyExpiry} onChange={(yes) => {
                    setHasWarrantyExpiry(yes);
                    if (yes && (!form.warrantyExpiry || form.warrantyExpiry === "N/A")) set("warrantyExpiry", todayISO());
                    if (!yes) set("warrantyExpiry", "");
                  }} label="Add warranty expiry" />
                </div>
              )}
              {form.assetType === "IT" && hasWarrantyExpiry && (
                <Field label="Warranty Expiry"><input type="date" value={form.warrantyExpiry} onChange={(e) => set("warrantyExpiry", e.target.value)} /></Field>
              )}
              <div className="form-full">
                <Field label="Notes">
                  <textarea value={form.notes} onChange={(e) => set("notes", e.target.value)} rows={2} />
                </Field>
              </div>
            </div>
          )}
        </div>

        <div className="modal-footer-sticky">
          <button type="button" className="btn ghost" onClick={onClose}>Cancel</button>
          <button type="submit" className="btn primary">{asset.id ? "Save Changes" : "Create Asset"}</button>
        </div>
      </form>
    </Modal>
  );
}

/* ---------------------------------------------------------
   Asset Detail Modal (read-only view, opened by clicking the tag)
--------------------------------------------------------- */
function AssetDetailModal({ asset, categories, locations, isAdmin, comments, maintenance, currentUser, onAddComment, onClose, onDelete }) {
  const cat = categories.find((c) => c.id === asset.categoryId);
  const loc = locations.find((l) => l.id === asset.locationId);
  const [commentText, setCommentText] = useState("");
  // Most recent maintenance first, so the newest work on this device is
  // what you see right away.
  const maintHistory = [...(maintenance || [])].sort((a, b) => new Date(b.date) - new Date(a.date));
  const maintStatusColor = { "Not Started": "#9CA3AF", "In Progress": "#F59E0B", "Done": "#10B981" };
  // Chronological order, oldest first, like a chat thread. Whoever started
  // the conversation on this asset anchors to the left; anyone else who
  // replies shows up on the right — a simple two-side chat layout.
  const chronComments = [...(comments || [])].sort((a, b) => new Date(a.at) - new Date(b.at));
  const firstAuthorId = chronComments[0]?.authorId;
  const submitComment = () => {
    if (!commentText.trim()) return;
    onAddComment(commentText);
    setCommentText("");
  };
  const row = (label, value, full) => (
    <div className={`detail-row${full ? " detail-full" : ""}`}>
      <span className="detail-label">{label}</span>
      <span className="detail-value">{value || "—"}</span>
    </div>
  );
  // Lightweight section wrapper — reuses the same header style as the New
  // Asset window's FormSection, but keeps the two-column detail-grid body
  // (rather than FormSection's input-oriented section-grid) since this is
  // a read-only label/value layout, not a form.
  const detailSection = (Icon, title, content) => (
    <div className="form-section detail-section">
      <div className="form-section-head">
        <span className="form-section-head-icon"><Icon size={14} /></span>
        <h4>{title}</h4>
      </div>
      <div className="detail-grid">{content}</div>
    </div>
  );
  return (
    <Modal title={`Asset Details — ${asset.tag}`} onClose={onClose} width={760}>
      {detailSection(Package, "Asset Information", <>
        {row("Name", asset.name)}
        {row("Department", asset.department)}
        {row("Category", cat && (
          <span style={{ display: "inline-flex", alignItems: "center" }}>
            <span className="cat-dot" style={{ background: categoryColor(categories, asset.categoryId) }} />
            {cat.name}
          </span>
        ))}
        {row("Asset Type", asset.assetType)}
      </>)}

      {detailSection(Tags, "Device Details", <>
        {row("Brand / Model", [asset.brand, asset.model].filter(Boolean).join(" / "))}
        {row("Manufactured Year / Year Model", asset.yearModel)}
        {row("Serial Number", asset.serial)}
        {row("Status", <Badge color={STATUS_COLORS[asset.status] || "#6B7280"}>{asset.status}</Badge>)}
        {row("Condition", asset.condition)}
        {asset.assetType === "Non-IT" && row("Requires Calibration?", asset.requiresCalibration ? "Yes" : "No")}
        {asset.assetType === "Non-IT" && asset.requiresCalibration && row("Calibration Date", asset.calibrationDate)}
        {asset.assetType === "Non-IT" && asset.requiresCalibration && row("Next Recalibration Date", asset.nextCalibrationDate)}
      </>)}

      {detailSection(User, "Assignment", <>
        {row("Location", loc?.name)}
        {row("Assigned To", asset.assignedTo)}
        {row("Added By", asset.createdByName)}
      </>)}

      {detailSection(Info, "Purchase & Warranty", <>
        {row("Purchase Date", asset.purchaseDate)}
        {row("Purchase Cost", asset.purchaseCost ? `$${asset.purchaseCost}` : "")}
        {asset.assetType === "IT" && row("Warranty Expiry", asset.warrantyExpiry)}
      </>)}

      <div className="detail-full">
        <div className="notif-section-title" style={{ marginTop: 14 }}>Notes</div>
        <div className="notes-highlight">
          {asset.notes || "No notes yet."}
        </div>
      </div>

      <div className="detail-full">
        <div className="notif-section-title" style={{ marginTop: 14 }}>Service &amp; Maintenance History</div>
        {maintHistory.length === 0 ? (
          <div className="notif-empty">No maintenance history yet for this asset.</div>
        ) : (
          <ul className="maint-history-list">
            {maintHistory.map((m) => (
              <li key={m.id} className="maint-history-item">
                <span className="maint-history-date">{m.date}</span>
                <span className="maint-history-desc">{m.description}</span>
                {m.cost ? <span className="maint-history-cost">${m.cost}</span> : null}
                <Badge color={maintStatusColor[m.status] || "#6B7280"}>{m.status}</Badge>
              </li>
            ))}
          </ul>
        )}
      </div>

      <div className="detail-full comments-section">
        <div className="notif-section-title" style={{ marginTop: 14 }}>Comments</div>
        <div className="comment-list chat-list">
          {chronComments.length === 0 && (
            <div className="notif-empty">No comments yet — use this to clarify something with the assigned user.</div>
          )}
          {chronComments.map((c) => {
            const side = c.authorId === firstAuthorId ? "left" : "right";
            return (
              <div key={c.id} className={`chat-row chat-row-${side}`}>
                <div className={`chat-bubble chat-bubble-${side}`}>
                  <div className="comment-meta">
                    <span className="comment-author">{c.authorName}</span>
                    <span className="comment-time">{new Date(c.at).toLocaleString([], { dateStyle: "medium", timeStyle: "short" })}</span>
                  </div>
                  <div className="comment-text">{c.message}</div>
                </div>
              </div>
            );
          })}
        </div>
        <div className="comment-composer">
          <textarea
            value={commentText}
            onChange={(e) => setCommentText(e.target.value)}
            rows={2}
            placeholder="Add a comment…"
          />
          <button type="button" className="btn primary" onClick={submitComment} disabled={!commentText.trim()}>
            Send
          </button>
        </div>
      </div>

      {asset.transferHistory && asset.transferHistory.length > 0 && (
        <div className="detail-transfer-history">
          <div className="notif-section-title">Transfer History</div>
          {asset.transferHistory.map((t) => (
            <div key={t.id} className="notif-item">
              {t.fromLocationName} → {t.toLocationName} — {t.reason} ({t.by}, {new Date(t.at).toLocaleString([], { dateStyle: "medium", timeStyle: "short" })})
            </div>
          ))}
        </div>
      )}

      <div className="modal-actions" style={{ marginTop: 18, justifyContent: "space-between" }}>
        <button
          type="button"
          className="btn danger-outline"
          onClick={onDelete}
          title={!isAdmin && asset.pendingDeletion ? "Awaiting Admin approval" : "Delete"}
          disabled={!isAdmin && !!asset.pendingDeletion}
        >
          <Trash2 size={14} /> Delete
        </button>
        <button type="button" className="btn ghost" onClick={onClose}>Close</button>
      </div>
    </Modal>
  );
}

/* ---------------------------------------------------------
   Maintenance View
--------------------------------------------------------- */
function emptyMaint(assetId) {
  return { id: null, assetId: assetId || "", description: "", cost: "", date: todayISO(), status: "Not Started" };
}

function MaintenanceView({ data, persist, showToast, scopedLocationId, currentUser }) {
  const [editing, setEditing] = useState(null);
  const [confirmDelete, setConfirmDelete] = useState(null);

  const assetsInScope = scopedLocationId ? data.assets.filter((a) => a.locationId === scopedLocationId) : data.assets;
  const assetIds = new Set(assetsInScope.map((a) => a.id));
  const logs = data.maintenance.filter((m) => assetIds.has(m.assetId));

  // Once an entry is marked Done it moves out of the working list and into
  // a read-only history section below — like an activity-log record, not
  // something you keep clicking into.
  const activeLogs = logs.filter((m) => m.status !== "Done");
  const doneLogs = logs.filter((m) => m.status === "Done").sort((a, b) => new Date(b.date) - new Date(a.date));

  const [selected, setSelected] = useState([]);

  const assetLabel = (assetId) => data.assets.find((a) => a.id === assetId)?.name || data.assets.find((a) => a.id === assetId)?.tag || "an asset";

  const save = async (log) => {
    let next;
    if (log.id) {
      const maintenance = data.maintenance.map((m) => (m.id === log.id ? log : m));
      const assets = log.status === "Done"
        ? maybeRestoreStatus(data.assets, maintenance, log.assetId)
        : markUnderRepair(data.assets, log.assetId);
      next = withLog({ ...data, assets, maintenance }, currentUser, `Edited maintenance entry for ${assetLabel(log.assetId)}`);
    } else {
      const newEntry = { ...log, id: uid("maint") };
      const maintenance = [newEntry, ...data.maintenance];
      const assets = newEntry.status === "Done" ? data.assets : markUnderRepair(data.assets, newEntry.assetId);
      const becameUnderRepair = newEntry.status !== "Done" && data.assets.find((a) => a.id === newEntry.assetId)?.status !== "Under Repair";
      next = withLog({ ...data, assets, maintenance }, currentUser, `Added maintenance entry for ${assetLabel(log.assetId)}${becameUnderRepair ? " — asset set to Under Repair" : ""}`);
    }
    persist(next);
    setEditing(null);
    showToast("Maintenance entry saved.");
  };

  const remove = async (id) => {
    const log = data.maintenance.find((m) => m.id === id);
    const maintenance = data.maintenance.filter((m) => m.id !== id);
    const assets = maybeRestoreStatus(data.assets, maintenance, log?.assetId);
    const next = withLog({ ...data, assets, maintenance }, currentUser, `Deleted maintenance entry for ${assetLabel(log?.assetId)}`);
    persist(next);
    setConfirmDelete(null);
    showToast("Maintenance entry deleted.");
  };

  const bulkDelete = async () => {
    const affectedAssetIds = [...new Set(data.maintenance.filter((m) => selected.includes(m.id)).map((m) => m.assetId))];
    const maintenance = data.maintenance.filter((m) => !selected.includes(m.id));
    let assets = data.assets;
    affectedAssetIds.forEach((aid) => { assets = maybeRestoreStatus(assets, maintenance, aid); });
    const next = withLog({ ...data, assets, maintenance }, currentUser, `Deleted ${selected.length} maintenance entry(ies) in bulk`);
    persist(next);
    showToast(`${selected.length} entry(ies) deleted.`);
    setSelected([]);
  };

  const quickStatus = async (id, status) => {
    const log = data.maintenance.find((m) => m.id === id);
    const maintenance = data.maintenance.map((m) => (m.id === id ? { ...m, status } : m));
    const assets = status === "Done"
      ? maybeRestoreStatus(data.assets, maintenance, log?.assetId)
      : markUnderRepair(data.assets, log?.assetId);
    const next = withLog({ ...data, assets, maintenance }, currentUser, `Changed maintenance status for ${assetLabel(log?.assetId)} to "${status}"`);
    persist(next);
  };

  const statusColor = { "Not Started": "#9CA3AF", "In Progress": "#F59E0B", "Done": "#10B981" };

  return (
    <div>
      <div className="view-head">
        <h2 className="view-title">Service &amp; Maintenance</h2>
        <div className="view-actions">
          {selected.length > 0 && (
            <button className="btn danger" onClick={bulkDelete}><Trash2 size={14} /> Delete ({selected.length})</button>
          )}
          <button className="btn primary" onClick={() => setEditing(emptyMaint())}><Plus size={14} /> New Entry</button>
        </div>
      </div>

      <div className="panel">
        <div className="panel-head"><h3>In Progress</h3></div>
        <div className="table-wrap">
          <table>
            <thead>
              <tr>
                <th style={{ width: 32 }}>
                  <input
                    type="checkbox"
                    checked={activeLogs.length > 0 && selected.length === activeLogs.length}
                    onChange={(e) => setSelected(e.target.checked ? activeLogs.map((m) => m.id) : [])}
                  />
                </th>
                <th>Asset</th>
                <th>Description</th>
                <th>Date</th>
                <th>Cost</th>
                <th>Status</th>
                <th style={{ width: 90 }}></th>
              </tr>
            </thead>
            <tbody>
              {activeLogs.length === 0 && (
                <tr><td colSpan={7} className="empty-cell">Nothing in progress — click "New Entry" to log one.</td></tr>
              )}
              {activeLogs.map((m) => {
                const asset = data.assets.find((a) => a.id === m.assetId);
                return (
                  <tr key={m.id}>
                    <td className="checkbox-cell">
                      <input
                        type="checkbox"
                        checked={selected.includes(m.id)}
                        onChange={(e) => setSelected((s) => (e.target.checked ? [...s, m.id] : s.filter((x) => x !== m.id)))}
                      />
                    </td>
                    <td className="mono" data-label="Asset">{asset?.tag || "—"}</td>
                    <td data-label="Description">{m.description}</td>
                    <td data-label="Date">{m.date}</td>
                    <td data-label="Cost">{m.cost ? `$${m.cost}` : "—"}</td>
                    <td data-label="Status">
                      <select
                        className="status-select"
                        value={m.status}
                        onChange={(e) => quickStatus(m.id, e.target.value)}
                        style={{ color: statusColor[m.status], borderColor: `${statusColor[m.status]}55` }}
                      >
                        {MAINT_STATUS.map((s) => <option key={s} value={s}>{s}</option>)}
                      </select>
                    </td>
                    <td className="actions-cell">
                      <div className="row-actions">
                        <IconBtn icon={Pencil} title="Edit" onClick={() => setEditing(m)} />
                        <IconBtn icon={Trash2} title="Delete" danger onClick={() => setConfirmDelete(m.id)} />
                      </div>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </div>

      <div className="panel">
        <div className="panel-head"><h3>History (Done)</h3></div>
        <p className="export-hint" style={{ padding: "0 18px" }}>
          Completed entries are read-only, like an activity log — mark something
          "Done" above and it moves here automatically.
        </p>
        <div className="table-wrap">
          <table>
            <thead>
              <tr>
                <th>Asset</th>
                <th>Description</th>
                <th>Date</th>
                <th>Cost</th>
                <th>Status</th>
              </tr>
            </thead>
            <tbody>
              {doneLogs.length === 0 && (
                <tr><td colSpan={5} className="empty-cell">No completed entries yet.</td></tr>
              )}
              {doneLogs.map((m) => {
                const asset = data.assets.find((a) => a.id === m.assetId);
                return (
                  <tr key={m.id} className="history-row">
                    <td className="mono" data-label="Asset">{asset?.tag || "—"}</td>
                    <td data-label="Description">{m.description}</td>
                    <td data-label="Date">{m.date}</td>
                    <td data-label="Cost">{m.cost ? `$${m.cost}` : "—"}</td>
                    <td data-label="Status">
                      <Badge color={statusColor.Done}>Done</Badge>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </div>

      {editing && (
        <Modal title={editing.id ? "Edit Maintenance Entry" : "New Maintenance Entry"} onClose={() => setEditing(null)} width={760}>
          <MaintForm entry={editing} assets={assetsInScope} onSave={save} onClose={() => setEditing(null)} />
        </Modal>
      )}
      {confirmDelete && (
        <ConfirmDialog message="Delete this maintenance entry?" onCancel={() => setConfirmDelete(null)} onConfirm={() => remove(confirmDelete)} />
      )}
    </div>
  );
}

function MaintForm({ entry, assets, onSave, onClose }) {
  const [form, setForm] = useState(entry);
  const set = (k, v) => setForm((f) => ({ ...f, [k]: v }));
  const submit = () => {
    if (!form.assetId) { alert("Please select an asset."); return; }
    if (!form.description.trim()) { alert("Please enter a description."); return; }
    onSave(form);
  };
  return (
    <div className="form-grid">
      <Field label="Asset">
        <select value={form.assetId} onChange={(e) => set("assetId", e.target.value)}>
          <option value="">Select asset</option>
          {assets.map((a) => <option key={a.id} value={a.id}>{a.tag} — {a.name}</option>)}
        </select>
      </Field>
      <Field label="Status">
        <select value={form.status} onChange={(e) => set("status", e.target.value)}>
          {MAINT_STATUS.map((s) => <option key={s} value={s}>{s}</option>)}
        </select>
      </Field>
      <div className="form-full">
        <Field label="Description">
          <textarea rows={2} value={form.description} onChange={(e) => set("description", e.target.value)} />
        </Field>
      </div>
      <Field label="Date"><input type="date" value={form.date} onChange={(e) => set("date", e.target.value)} /></Field>
      <Field label="Cost"><input type="number" value={form.cost} onChange={(e) => set("cost", e.target.value)} /></Field>
      <div className="form-full modal-actions">
        <button type="button" className="btn ghost" onClick={onClose}>Cancel</button>
        <button type="button" className="btn primary" onClick={submit}>Save Entry</button>
      </div>
    </div>
  );
}

/* ---------------------------------------------------------
   Categories View
--------------------------------------------------------- */
function CategoriesView({ data, persist, showToast, currentUser }) {
  const [editing, setEditing] = useState(null);
  const [confirmDelete, setConfirmDelete] = useState(null);

  const save = async (cat) => {
    let next;
    if (cat.id) next = withLog({ ...data, categories: data.categories.map((c) => (c.id === cat.id ? cat : c)) }, currentUser, `Edited category "${cat.name}"`);
    else next = withLog({ ...data, categories: [...data.categories, { ...cat, id: uid("cat") }] }, currentUser, `Added category "${cat.name}"`);
    persist(next);
    setEditing(null);
    showToast("Category saved.");
  };

  const remove = async (id) => {
    const inUse = data.assets.some((a) => a.categoryId === id);
    if (inUse) {
      showToast("Can't delete — one or more assets still use this category.");
      setConfirmDelete(null);
      return;
    }
    const cat = data.categories.find((c) => c.id === id);
    persist(withLog({ ...data, categories: data.categories.filter((c) => c.id !== id) }, currentUser, `Deleted category "${cat?.name}"`));
    setConfirmDelete(null);
    showToast("Category deleted.");
  };

  return (
    <div>
      <div className="view-head">
        <h2 className="view-title">Categories</h2>
        <button className="btn primary" onClick={() => setEditing({ id: null, name: "", type: "IT", usefulLife: 3 })}>
          <Plus size={14} /> New Category
        </button>
      </div>
      <div className="panel">
        <div className="table-wrap">
          <table>
            <thead><tr><th>Name</th><th>Type</th><th>Useful Life (yrs)</th><th style={{ width: 90 }}></th></tr></thead>
            <tbody>
              {data.categories.map((c) => (
                <tr key={c.id}>
                  <td data-label="Name">{c.name}</td>
                  <td data-label="Type"><Badge color={c.type === "IT" ? "#6366F1" : "#F59E0B"}>{c.type}</Badge></td>
                  <td data-label="Useful Life (yrs)">{c.usefulLife}</td>
                  <td className="actions-cell">
                    <div className="row-actions">
                      <IconBtn icon={Pencil} title="Edit" onClick={() => setEditing(c)} />
                      <IconBtn icon={Trash2} title="Delete" danger onClick={() => setConfirmDelete(c.id)} />
                    </div>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
      {editing && (
        <Modal title={editing.id ? "Edit Category" : "New Category"} onClose={() => setEditing(null)}>
          <div className="form-grid">
            <Field label="Name"><input value={editing.name} onChange={(e) => setEditing({ ...editing, name: e.target.value })} /></Field>
            <Field label="Type">
              <select value={editing.type} onChange={(e) => setEditing({ ...editing, type: e.target.value })}>
                <option value="IT">IT</option>
                <option value="Non-IT">Non-IT</option>
              </select>
            </Field>
            <Field label="Useful Life (years)">
              <input type="number" value={editing.usefulLife} onChange={(e) => setEditing({ ...editing, usefulLife: Number(e.target.value) })} />
            </Field>
            <div className="form-full modal-actions">
              <button type="button" className="btn ghost" onClick={() => setEditing(null)}>Cancel</button>
              <button type="button" className="btn primary" onClick={() => { if (!editing.name.trim()) { alert("Please enter a category name."); return; } save(editing); }}>Save Category</button>
            </div>
          </div>
        </Modal>
      )}
      {confirmDelete && (
        <ConfirmDialog message="Delete this category?" onCancel={() => setConfirmDelete(null)} onConfirm={() => remove(confirmDelete)} />
      )}
    </div>
  );
}

/* ---------------------------------------------------------
   Locations View
--------------------------------------------------------- */
function LocationsView({ data, persist, showToast, currentUser }) {
  const [editing, setEditing] = useState(null);
  const [confirmDelete, setConfirmDelete] = useState(null);

  const save = async (loc) => {
    let next;
    if (loc.id) next = withLog({ ...data, locations: data.locations.map((l) => (l.id === loc.id ? loc : l)) }, currentUser, `Edited location "${loc.name}"`);
    else next = withLog({ ...data, locations: [...data.locations, { ...loc, id: uid("loc") }] }, currentUser, `Added location "${loc.name}"`);
    persist(next);
    setEditing(null);
    showToast("Location saved.");
  };

  const remove = async (id) => {
    const inUse = data.assets.some((a) => a.locationId === id) || data.users.some((u) => u.locationId === id);
    if (inUse) {
      showToast("Can't delete — assets or users are assigned to this location.");
      setConfirmDelete(null);
      return;
    }
    const loc = data.locations.find((l) => l.id === id);
    persist(withLog({ ...data, locations: data.locations.filter((l) => l.id !== id) }, currentUser, `Deleted location "${loc?.name}"`));
    setConfirmDelete(null);
    showToast("Location deleted.");
  };

  return (
    <div>
      <div className="view-head">
        <h2 className="view-title">Locations</h2>
        <button className="btn primary" onClick={() => setEditing({ id: null, name: "" })}><Plus size={14} /> New Location</button>
      </div>
      <div className="panel">
        <div className="table-wrap">
          <table>
            <thead><tr><th>Name</th><th style={{ width: 90 }}></th></tr></thead>
            <tbody>
              {data.locations.map((l) => (
                <tr key={l.id}>
                  <td data-label="Name">{l.name}</td>
                  <td className="actions-cell">
                    <div className="row-actions">
                      <IconBtn icon={Pencil} title="Edit" onClick={() => setEditing(l)} />
                      <IconBtn icon={Trash2} title="Delete" danger onClick={() => setConfirmDelete(l.id)} />
                    </div>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
      {editing && (
        <Modal title={editing.id ? "Edit Location" : "New Location"} onClose={() => setEditing(null)} width={380}>
          <div className="form-grid">
            <Field label="Name"><input value={editing.name} onChange={(e) => setEditing({ ...editing, name: e.target.value })} /></Field>
            <div className="form-full modal-actions">
              <button type="button" className="btn ghost" onClick={() => setEditing(null)}>Cancel</button>
              <button type="button" className="btn primary" onClick={() => { if (!editing.name.trim()) { alert("Please enter a location name."); return; } save(editing); }}>Save Location</button>
            </div>
          </div>
        </Modal>
      )}
      {confirmDelete && (
        <ConfirmDialog message="Delete this location?" onCancel={() => setConfirmDelete(null)} onConfirm={() => remove(confirmDelete)} />
      )}
    </div>
  );
}

/* ---------------------------------------------------------
   Users View
--------------------------------------------------------- */
function emptyUser(locations) {
  return { id: null, name: "", username: "", email: "", position: "", role: "Regional Staff", locationId: locations[0]?.id || "", password: "" };
}

function UsersView({ data, persist, showToast, currentUser }) {
  const [editing, setEditing] = useState(null);
  const [resetTarget, setResetTarget] = useState(null);
  const [confirmDelete, setConfirmDelete] = useState(null);

  const save = async (form) => {
    if (form.id) {
      const next = withLog({ ...data, users: data.users.map((u) => (u.id === form.id ? { ...u, name: form.name, username: form.username, email: form.email, position: form.position, role: form.role, locationId: form.role === "Admin" ? null : form.locationId } : u)) }, currentUser, `Edited user "${form.name}"`);
      persist(next);
    } else {
      const hash = await sha256(form.password || "changeme123");
      const newUser = { id: uid("usr"), name: form.name, username: form.username, email: form.email, position: form.position, role: form.role, locationId: form.role === "Admin" ? null : form.locationId, passwordHash: hash };
      persist(withLog({ ...data, users: [...data.users, newUser] }, currentUser, `Added user "${newUser.name}" (${newUser.role})`));
    }
    setEditing(null);
    showToast("User saved.");
  };

  const remove = async (id) => {
    const u = data.users.find((x) => x.id === id);
    persist(withLog({ ...data, users: data.users.filter((u) => u.id !== id) }, currentUser, `Removed user "${u?.name}"`));
    setConfirmDelete(null);
    showToast("User removed.");
  };

  const resetPassword = async (id, newPass) => {
    const hash = await sha256(newPass);
    const u = data.users.find((x) => x.id === id);
    persist(withLog({ ...data, users: data.users.map((u) => (u.id === id ? { ...u, passwordHash: hash } : u)) }, currentUser, `Reset password for "${u?.name}"`));
    setResetTarget(null);
    showToast("Password reset.");
  };

  return (
    <div>
      <div className="view-head">
        <h2 className="view-title">User Accounts</h2>
        <button className="btn primary" onClick={() => setEditing(emptyUser(data.locations))}><Plus size={14} /> New User</button>
      </div>
      <div className="panel">
        <div className="table-wrap">
          <table>
            <thead><tr><th>Name</th><th>Position</th><th>Username</th><th>Role</th><th>Location</th><th style={{ width: 130 }}></th></tr></thead>
            <tbody>
              {data.users.map((u) => {
                const loc = data.locations.find((l) => l.id === u.locationId);
                return (
                  <tr key={u.id}>
                    <td data-label="Name">{u.name}</td>
                    <td data-label="Position">{u.position || "—"}</td>
                    <td className="mono" data-label="Username">{u.username}</td>
                    <td data-label="Role"><Badge color={u.role === "Admin" ? "#6366F1" : "#10B981"}>{u.role}</Badge></td>
                    <td data-label="Location">{loc?.name || "— (Global)"}</td>
                    <td className="actions-cell">
                      <div className="row-actions">
                        <IconBtn icon={Pencil} title="Edit" onClick={() => setEditing({ ...u })} />
                        <IconBtn icon={KeyRound} title="Reset Password" onClick={() => setResetTarget(u.id)} />
                        <IconBtn icon={Trash2} title="Delete" danger onClick={() => setConfirmDelete(u.id)} />
                      </div>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </div>

      {editing && (
        <Modal title={editing.id ? "Edit User" : "New User"} onClose={() => setEditing(null)}>
          <div className="form-grid">
            <Field label="Full Name"><input value={editing.name} onChange={(e) => setEditing({ ...editing, name: e.target.value })} /></Field>
            <Field label="Position / Title"><input value={editing.position || ""} onChange={(e) => setEditing({ ...editing, position: e.target.value })} placeholder="e.g. HR Manager" /></Field>
            <Field label="Username"><input value={editing.username} onChange={(e) => setEditing({ ...editing, username: e.target.value })} /></Field>
            <Field label="Email"><input type="email" value={editing.email || ""} onChange={(e) => setEditing({ ...editing, email: e.target.value })} placeholder="name@company.com" /></Field>
            <Field label="Role">
              <select value={editing.role} onChange={(e) => setEditing({ ...editing, role: e.target.value })}>
                <option value="Admin">Admin</option>
                <option value="Regional Staff">Regional Staff</option>
              </select>
            </Field>
            {editing.role !== "Admin" && (
              <Field label="Location">
                <select value={editing.locationId} onChange={(e) => setEditing({ ...editing, locationId: e.target.value })}>
                  {data.locations.map((l) => <option key={l.id} value={l.id}>{l.name}</option>)}
                </select>
              </Field>
            )}
            {!editing.id && (
              <Field label="Temporary Password">
                <input type="text" value={editing.password} onChange={(e) => setEditing({ ...editing, password: e.target.value })} placeholder="e.g. changeme123" />
              </Field>
            )}
            <div className="form-full modal-actions">
              <button type="button" className="btn ghost" onClick={() => setEditing(null)}>Cancel</button>
              <button
                type="button"
                className="btn primary"
                onClick={() => {
                  if (!editing.name.trim() || !editing.username.trim()) {
                    alert("Please enter both a full name and username.");
                    return;
                  }
                  save(editing);
                }}
              >
                Save User
              </button>
            </div>
          </div>
        </Modal>
      )}

      {resetTarget && (
        <ResetPasswordModal onClose={() => setResetTarget(null)} onConfirm={(pass) => resetPassword(resetTarget, pass)} />
      )}
      {confirmDelete && (
        <ConfirmDialog message="Remove this user's access?" onCancel={() => setConfirmDelete(null)} onConfirm={() => remove(confirmDelete)} />
      )}
    </div>
  );
}

/* ---------------------------------------------------------
   Backup & Restore — full-data Excel workbook (one sheet per table)
--------------------------------------------------------- */
function buildBackupWorkbook(data) {
  const wb = XLSX.utils.book_new();

  const addSheet = (name, rows) => {
    const sheet = rows.length ? XLSX.utils.json_to_sheet(rows) : XLSX.utils.aoa_to_sheet([[]]);
    XLSX.utils.book_append_sheet(wb, sheet, name);
  };

  addSheet("Locations", (data.locations || []).map((l) => ({ id: l.id, name: l.name })));

  addSheet("Categories", (data.categories || []).map((c) => ({
    id: c.id, name: c.name, type: c.type, usefulLife: c.usefulLife,
  })));

  addSheet("Assets", (data.assets || []).map((a) => ({
    id: a.id, tag: a.tag, name: a.name, department: a.department, categoryId: a.categoryId,
    assetType: a.assetType, brand: a.brand, model: a.model, yearModel: a.yearModel, serial: a.serial,
    status: a.status, condition: a.condition, locationId: a.locationId, assignedTo: a.assignedTo,
    createdById: a.createdById, createdByName: a.createdByName,
    purchaseDate: a.purchaseDate, purchaseCost: a.purchaseCost, warrantyExpiry: a.warrantyExpiry,
    requiresCalibration: a.requiresCalibration ? "Yes" : "No",
    calibrationDate: a.calibrationDate, nextCalibrationDate: a.nextCalibrationDate,
    preRepairStatus: a.preRepairStatus || "", notes: a.notes,
    transferHistory: cellify(a.transferHistory || []),
    pendingDeletion: cellify(a.pendingDeletion || null),
  })));

  addSheet("Maintenance", (data.maintenance || []).map((m) => ({
    id: m.id, assetId: m.assetId, description: m.description, cost: m.cost, date: m.date, status: m.status,
  })));

  addSheet("Users", (data.users || []).map((u) => ({
    id: u.id, name: u.name, username: u.username, email: u.email, position: u.position,
    passwordHash: u.passwordHash, role: u.role, locationId: u.locationId || "",
  })));

  addSheet("Comments", (data.comments || []).map((c) => ({
    id: c.id, assetId: c.assetId, authorId: c.authorId, authorName: c.authorName,
    message: c.message, at: c.at,
    targetUserIds: cellify(c.targetUserIds || []), readBy: cellify(c.readBy || []),
  })));

  addSheet("AuditLog", (data.auditLog || []).map((l) => ({
    id: l.id, at: l.at, userId: l.userId, userName: l.userName, message: l.message,
  })));

  return wb;
}

function parseBackupWorkbook(wb) {
  const locations = sheetRows(wb, "Locations").map((r) => ({ id: String(r.id), name: r.name }));

  const categories = sheetRows(wb, "Categories").map((r) => ({
    id: String(r.id), name: r.name, type: r.type, usefulLife: Number(r.usefulLife) || 0,
  }));

  const assets = sheetRows(wb, "Assets").map((r) => ({
    id: String(r.id), tag: r.tag, name: r.name, department: r.department || "",
    categoryId: String(r.categoryId || ""), assetType: r.assetType, brand: r.brand, model: r.model,
    yearModel: r.yearModel, serial: r.serial, status: r.status, condition: r.condition,
    locationId: String(r.locationId || ""), assignedTo: r.assignedTo,
    createdById: r.createdById ? String(r.createdById) : null, createdByName: r.createdByName || "",
    purchaseDate: r.purchaseDate, purchaseCost: Number(r.purchaseCost) || 0, warrantyExpiry: r.warrantyExpiry,
    requiresCalibration: String(r.requiresCalibration).toLowerCase() === "yes",
    calibrationDate: r.calibrationDate, nextCalibrationDate: r.nextCalibrationDate,
    preRepairStatus: r.preRepairStatus || null, notes: r.notes,
    transferHistory: parseCell(r.transferHistory, []),
    pendingDeletion: parseCell(r.pendingDeletion, null),
  }));

  const maintenance = sheetRows(wb, "Maintenance").map((r) => ({
    id: String(r.id), assetId: String(r.assetId), description: r.description,
    cost: Number(r.cost) || 0, date: r.date, status: r.status,
  }));

  const users = sheetRows(wb, "Users").map((r) => ({
    id: String(r.id), name: r.name, username: r.username, email: r.email, position: r.position,
    passwordHash: r.passwordHash, role: r.role, locationId: r.locationId ? String(r.locationId) : null,
  }));

  const comments = sheetRows(wb, "Comments").map((r) => ({
    id: String(r.id), assetId: String(r.assetId), authorId: String(r.authorId || ""),
    authorName: r.authorName, message: r.message, at: r.at,
    targetUserIds: parseCell(r.targetUserIds, []), readBy: parseCell(r.readBy, []),
  }));

  const auditLog = sheetRows(wb, "AuditLog").map((r) => ({
    id: String(r.id), at: r.at, userId: r.userId ? String(r.userId) : null,
    userName: r.userName, message: r.message,
  }));

  return { locations, categories, assets, maintenance, users, comments, auditLog };
}

/* ---------------------------------------------------------
   Backup & Restore
--------------------------------------------------------- */
function BackupView({ data, persist, showToast, currentUser }) {
  const inputRef = React.useRef(null);
  const [pendingRestore, setPendingRestore] = useState(null); // parsed backup awaiting confirm
  const [restoreError, setRestoreError] = useState("");

  const orphanCount = useMemo(() => {
    const assetIds = new Set(data.assets.map((a) => a.id));
    return data.maintenance.filter((m) => !assetIds.has(m.assetId)).length;
  }, [data.assets, data.maintenance]);

  const cleanupOrphans = async () => {
    const assetIds = new Set(data.assets.map((a) => a.id));
    const removed = data.maintenance.filter((m) => !assetIds.has(m.assetId)).length;
    const next = withLog(
      { ...data, maintenance: data.maintenance.filter((m) => assetIds.has(m.assetId)) },
      currentUser,
      `Cleaned up ${removed} leftover maintenance record(s) not linked to any asset`
    );
    await persist(next);
    showToast(`Removed ${removed} leftover record(s).`);
  };

  const downloadBackup = () => {
    const wb = buildBackupWorkbook(data);
    const stamp = new Date().toISOString().slice(0, 19).replace(/[:T]/g, "-");
    downloadWorkbook(wb, `asset-manager-backup-${stamp}.xlsx`);
    showToast("Backup downloaded.");
  };

  const triggerRestore = () => inputRef.current?.click();

  const handleFile = async (e) => {
    const file = e.target.files?.[0];
    e.target.value = "";
    if (!file) return;
    setRestoreError("");
    try {
      const wb = await readWorkbookFile(file);
      const parsed = parseBackupWorkbook(wb);
      if (!Array.isArray(parsed.assets) || !Array.isArray(parsed.locations) || parsed.locations.length === 0) {
        setRestoreError("This doesn't look like a valid backup file.");
        return;
      }
      setPendingRestore(parsed);
    } catch {
      setRestoreError("Couldn't read that file — make sure it's a backup .xlsx file.");
    }
  };

  const confirmRestore = async () => {
    await persist(pendingRestore);
    setPendingRestore(null);
    showToast("Data restored from backup.");
  };

  return (
    <div>
      <div className="view-head">
        <h2 className="view-title">Backup &amp; Restore</h2>
      </div>

      {/* TEMPORARY: one-time cleanup for leftover demo maintenance records.
          Remove this panel once orphanCount reaches 0 for you. */}
      {orphanCount > 0 && (
        <div className="panel">
          <div className="panel-head"><h3>One-Time Cleanup</h3></div>
          <div style={{ padding: "18px" }}>
            <p className="login-sub" style={{ margin: "0 0 14px" }}>
              Found <strong>{orphanCount}</strong> maintenance record(s) left over from
              deleted assets (from before deletions started cleaning these up
              automatically). This removes just those — nothing else is touched.
            </p>
            <button className="btn danger" onClick={cleanupOrphans}>
              <Trash2 size={14} /> Remove {orphanCount} Leftover Record(s)
            </button>
          </div>
        </div>
      )}

      <div className="panel">
        <div className="panel-head"><h3>Download a Backup</h3></div>
        <div style={{ padding: "18px" }}>
          <p className="login-sub" style={{ margin: "0 0 14px" }}>
            Saves everything — assets, categories, locations, users, and maintenance
            records — as one file on your device. The shared database has no automatic
            backups, so it's worth doing this occasionally.
          </p>
          <button className="btn primary" onClick={downloadBackup}>
            <Download size={14} /> Download Backup (.xlsx)
          </button>
        </div>
      </div>

      <div className="panel">
        <div className="panel-head"><h3>Restore from a Backup</h3></div>
        <div style={{ padding: "18px" }}>
          <p className="login-sub" style={{ margin: "0 0 14px" }}>
            This replaces <strong>all</strong> current data — everyone's data — with
            what's in the file you choose. Use this to undo a mistake or recover from
            a lost project.
          </p>
          <input ref={inputRef} type="file" accept=".xlsx" style={{ display: "none" }} onChange={handleFile} />
          <button className="btn ghost" onClick={triggerRestore}>
            <Upload size={14} /> Choose Backup File
          </button>
          {restoreError && <div className="form-error" style={{ marginTop: 12 }}>{restoreError}</div>}
        </div>
      </div>

      {pendingRestore && (
        <div className="modal-overlay">
          <div className="modal confirm">
            <div className="confirm-icon"><AlertTriangle size={20} /></div>
            <h3>Replace all current data?</h3>
            <p className="login-sub">
              This will overwrite the live data for every user with the contents of
              this backup file. This can't be undone unless you have another backup.
            </p>
            <div className="confirm-actions">
              <button className="btn ghost" onClick={() => setPendingRestore(null)}>Cancel</button>
              <button className="btn danger" onClick={confirmRestore}>Yes, Restore</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

/* ---------------------------------------------------------
   Activity Log
--------------------------------------------------------- */
function ActivityLogView({ data, isAdmin, scopedLocationId }) {
  const [search, setSearch] = useState("");
  const [userFilter, setUserFilter] = useState("all");

  // Regional Staff only ever see activity tagged with their own location
  // (see withLog's locationId param) — Admins see the full, unfiltered log.
  const scopedLog = useMemo(
    () => (scopedLocationId ? (data.auditLog || []).filter((l) => l.locationId === scopedLocationId) : (data.auditLog || [])),
    [data.auditLog, scopedLocationId]
  );

  const userOptions = useMemo(() => {
    const names = new Set(scopedLog.map((l) => l.userName));
    return Array.from(names).sort();
  }, [scopedLog]);

  const entries = useMemo(() => {
    let list = scopedLog;
    if (userFilter !== "all") list = list.filter((l) => l.userName === userFilter);
    if (search.trim()) {
      const q = search.toLowerCase();
      list = list.filter((l) => l.message.toLowerCase().includes(q) || l.userName.toLowerCase().includes(q));
    }
    return list;
  }, [scopedLog, search, userFilter]);

  const formatWhen = (iso) => {
    const d = new Date(iso);
    return d.toLocaleString([], { month: "short", day: "numeric", hour: "2-digit", minute: "2-digit" });
  };

  return (
    <div>
      <div className="view-head">
        <h2 className="view-title">Activity Log</h2>
        <div className="view-actions">
          <div className="search-box">
            <Search size={14} />
            <input placeholder="Search activity..." value={search} onChange={(e) => setSearch(e.target.value)} />
          </div>
          <select className="sort-select" value={userFilter} onChange={(e) => setUserFilter(e.target.value)}>
            <option value="all">All users</option>
            {userOptions.map((n) => <option key={n} value={n}>{n}</option>)}
          </select>
        </div>
      </div>

      <p className="export-hint">
        {isAdmin
          ? "Every add, edit, and delete made by any user, most recent first."
          : "Activity for assets in your location, most recent first."}
      </p>

      <div className="panel">
        <div className="table-wrap">
          <table>
            <thead>
              <tr>
                <th style={{ width: 150 }}>When</th>
                <th style={{ width: 160 }}>User</th>
                <th>Action</th>
              </tr>
            </thead>
            <tbody>
              {entries.length === 0 ? (
                <tr><td colSpan={3} className="empty-cell">No activity recorded yet.</td></tr>
              ) : entries.map((l) => (
                <tr key={l.id}>
                  <td className="mono">{formatWhen(l.at)}</td>
                  <td>{l.userName}</td>
                  <td>{l.message}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

/* ---------------------------------------------------------
   Approvals (pending asset-deletion requests)
--------------------------------------------------------- */
function ApprovalsView({ data, persist, showToast, currentUser }) {
  const pending = data.assets.filter((a) => a.pendingDeletion);

  const approve = async (asset) => {
    const removedLogs = data.maintenance.filter((m) => m.assetId === asset.id).length;
    const suffix = removedLogs > 0 ? ` (and ${removedLogs} maintenance record${removedLogs > 1 ? "s" : ""})` : "";
    const next = withLog({
      ...data,
      assets: data.assets.filter((a) => a.id !== asset.id),
      maintenance: data.maintenance.filter((m) => m.assetId !== asset.id),
    }, currentUser, `Approved deletion of asset "${asset.name || asset.tag}" — requested by ${asset.pendingDeletion.requestedByName}${suffix}`);
    persist(next);
    showToast("Deletion approved.");
  };

  const reject = async (asset) => {
    const next = withLog({
      ...data,
      assets: data.assets.map((a) => (a.id === asset.id ? { ...a, pendingDeletion: null } : a)),
    }, currentUser, `Rejected deletion request for asset "${asset.name || asset.tag}" — requested by ${asset.pendingDeletion.requestedByName}`);
    persist(next);
    showToast("Deletion request rejected — asset kept.");
  };

  const formatWhen = (iso) => {
    if (!iso) return "—";
    return new Date(iso).toLocaleString([], { month: "short", day: "numeric", hour: "2-digit", minute: "2-digit" });
  };

  return (
    <div>
      <div className="view-head">
        <h2 className="view-title">Approvals</h2>
      </div>
      <p className="export-hint">
        Deletion requests from Regional Staff wait here until you approve or reject them.
        Nothing is removed until you say so.
      </p>

      {pending.length === 0 ? (
        <div className="panel">
          <div className="empty-cell">No pending requests right now.</div>
        </div>
      ) : (
        pending.map((a) => (
          <div className="panel" key={a.id}>
            <div className="panel-head" style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
              <h3>{a.name || a.tag}</h3>
              <Badge color="#EF4444">Pending Deletion</Badge>
            </div>
            <div style={{ padding: "18px" }}>
              <p className="login-sub" style={{ margin: "0 0 6px" }}>
                <strong>Requested by:</strong> {a.pendingDeletion.requestedByName} on {formatWhen(a.pendingDeletion.requestedAt)}
              </p>
              <p className="login-sub" style={{ margin: "0 0 16px" }}>
                <strong>Reason:</strong> {a.pendingDeletion.reason}
              </p>
              <div className="row-actions">
                <button className="btn ghost" onClick={() => reject(a)}>Reject — Keep Asset</button>
                <button className="btn danger" onClick={() => approve(a)}>Approve — Delete Asset</button>
              </div>
            </div>
          </div>
        ))
      )}
    </div>
  );
}

function ResetPasswordModal({ onClose, onConfirm }) {
  const [pass, setPass] = useState("");
  const submit = () => {
    if (!pass) { alert("Please enter a new password."); return; }
    onConfirm(pass);
  };
  return (
    <Modal title="Reset Password" onClose={onClose} width={380}>
      <div className="form-grid" onKeyDown={(e) => { if (e.key === "Enter") submit(); }}>
        <Field label="New Password"><input type="text" value={pass} onChange={(e) => setPass(e.target.value)} autoFocus /></Field>
        <div className="form-full modal-actions">
          <button type="button" className="btn ghost" onClick={onClose}>Cancel</button>
          <button type="button" className="btn primary" onClick={submit}>Set Password</button>
        </div>
      </div>
    </Modal>
  );
}

/* ---------------------------------------------------------
   Global styles
--------------------------------------------------------- */
function GlobalStyles() {
  return (
    <style>{`
      @import url('https://fonts.googleapis.com/css2?family=Poppins:wght@400;500;600;700&display=swap');

      .theme-light {
        --bg: #F7F8FA; --surface: #FFFFFF; --border: #E5E7EB; --text: #111827;
        --text-soft: #6B7280; --accent: #4F46E5; --accent-soft: #EEF2FF; --danger: #EF4444;
      }
      .theme-dark {
        --bg: #12141A; --surface: #1A1D24; --border: #2A2E38; --text: #F3F4F6;
        --text-soft: #9CA3AF; --accent: #818CF8; --accent-soft: #262B45; --danger: #F87171;
      }
      .theme-light, .theme-dark { min-height: 100vh; background: var(--bg); color: var(--text); font-family: 'Poppins', system-ui, sans-serif; }
      * { box-sizing: border-box; }
      button, input, select, textarea { font-family: inherit; }
      h1,h2,h3 { font-family: 'Poppins', sans-serif; margin: 0; }

      .boot { min-height: 100vh; display: flex; align-items: center; justify-content: center; }
      .spinner { width: 28px; height: 28px; border-radius: 999px; border: 3px solid #ddd; border-top-color: #4F46E5; animation: spin 0.8s linear infinite; }
      @keyframes spin { to { transform: rotate(360deg); } }

      .login-wrap { min-height: 100vh; display: flex; align-items: center; justify-content: center; padding: 20px; }
      .login-card { width: 360px; background: var(--surface); border: 1px solid var(--border); border-radius: 16px; padding: 32px; }
      .login-logo { display: flex; align-items: center; gap: 8px; font-family: 'Poppins', sans-serif; font-weight: 700; font-size: 18px; color: var(--accent); }
      .login-sub { color: var(--text-soft); font-size: 13px; margin: 6px 0 22px; }
      .login-hint { margin-top: 16px; font-size: 12px; color: var(--text-soft); text-align: center; }
      .form-error { background: #FEE2E2; color: #B91C1C; padding: 8px 10px; border-radius: 8px; font-size: 12.5px; margin-bottom: 12px; }

      .shell { display: flex; min-height: 100vh; }
      .sidebar { width: 220px; background: var(--surface); border-right: 1px solid var(--border); display: flex; flex-direction: column; transition: width 0.18s ease; flex-shrink: 0; overflow: hidden; }
      .sidebar.collapsed { width: 64px; }
      .sidebar-top { display: flex; align-items: center; padding: 16px; border-bottom: 1px solid var(--border); flex-shrink: 0; }
      .brand { display: flex; align-items: center; gap: 8px; font-family: 'Poppins', sans-serif; font-weight: 700; color: var(--accent); white-space: nowrap; }
      .brand-badge { width: 28px; height: 28px; border-radius: 9px; background: linear-gradient(135deg, var(--accent), color-mix(in srgb, var(--accent) 55%, #7C3AED)); color: white; display: flex; align-items: center; justify-content: center; flex-shrink: 0; }
      .brand-mini { display: flex; align-items: center; justify-content: center; color: var(--accent); width: 100%; }
      nav { padding: 10px; display: flex; flex-direction: column; gap: 2px; flex: 1; overflow-y: auto; -webkit-overflow-scrolling: touch; overscroll-behavior: contain; }
      .nav-item { display: flex; align-items: center; gap: 10px; padding: 9px 12px; border-radius: 8px; border: none; background: none; color: var(--text-soft); cursor: pointer; font-size: 13.5px; font-weight: 500; white-space: nowrap; overflow: hidden; }
      .nav-item:hover { background: var(--accent-soft); color: var(--accent); }
      .nav-item.active { background: var(--accent-soft); color: var(--accent); }

      .main { flex: 1; display: flex; flex-direction: column; min-width: 0; }
      .topbar { display: flex; align-items: center; justify-content: space-between; padding: 14px 24px; border-bottom: 1px solid var(--border); background: var(--surface); gap: 10px; }
      .topbar-left { display: flex; align-items: center; gap: 10px; min-width: 0; }
      .topbar-region { display: flex; align-items: center; gap: 6px; font-size: 13px; color: var(--text-soft); background: var(--bg); padding: 6px 12px; border-radius: 999px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
      .topbar-right { display: flex; align-items: center; gap: 10px; }
      .mobile-menu-btn { display: flex; flex-shrink: 0; }
      .user-chip { display: flex; align-items: center; gap: 9px; padding: 4px 14px 4px 4px; border-radius: 999px; background: var(--surface); border: 1px solid var(--border); box-shadow: 0 1px 2px rgba(0,0,0,0.04); transition: box-shadow 0.15s ease, border-color 0.15s ease; }
      .user-chip:hover { box-shadow: 0 2px 8px rgba(0,0,0,0.08); border-color: color-mix(in srgb, var(--accent) 35%, var(--border)); }
      .avatar { width: 32px; height: 32px; border-radius: 999px; background: linear-gradient(135deg, var(--accent), color-mix(in srgb, var(--accent) 55%, #7C3AED)); color: white; display: flex; align-items: center; justify-content: center; font-size: 12px; font-weight: 700; letter-spacing: 0.3px; flex-shrink: 0; box-shadow: 0 0 0 2px var(--surface), 0 0 0 3.5px color-mix(in srgb, var(--accent) 45%, transparent); }
      .user-name { font-size: 12.5px; font-weight: 600; line-height: 1.2; }
      .user-role { font-size: 11px; color: var(--text-soft); }
      .sidebar-backdrop { display: none; }

      .link-tag { background: none; border: none; padding: 0; font-family: ui-monospace, monospace; font-size: 12.5px; color: var(--accent); font-weight: 600; cursor: pointer; text-decoration: underline; text-underline-offset: 2px; }
      .link-tag:hover { opacity: 0.8; }
      .comment-flag { color: var(--danger); margin-left: 6px; vertical-align: middle; }
      .alert-flag { color: #F59E0B; margin-left: 6px; vertical-align: middle; }
      .alert-flag-urgent { color: var(--danger); }
      .table-legend { display: flex; flex-wrap: wrap; gap: 16px; align-items: center; font-size: 12px; color: var(--text-soft); margin: 0 0 10px 2px; }
      .table-legend-item { display: inline-flex; align-items: center; gap: 5px; }
      .table-legend-item svg { margin-left: 0 !important; }

      .notif-wrap { position: relative; }
      .notif-dot { position: absolute; top: -3px; right: -3px; background: var(--danger); color: white; font-size: 9.5px; font-weight: 700; border-radius: 999px; min-width: 15px; height: 15px; display: flex; align-items: center; justify-content: center; padding: 0 3px; }
      .notif-backdrop { position: fixed; inset: 0; z-index: 60; }
      .notif-panel { position: absolute; top: calc(100% + 8px); right: 0; width: 300px; max-height: 360px; overflow-y: auto; background: var(--surface); border: 1px solid var(--border); border-radius: 12px; box-shadow: 0 12px 28px rgba(0,0,0,0.18); padding: 12px; z-index: 61; scrollbar-width: none; }
      .notif-panel::-webkit-scrollbar { display: none; }
      .notif-section-title { font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.03em; color: var(--text-soft); margin-bottom: 6px; }
      .notif-item { font-size: 12.5px; padding: 6px 0; border-bottom: 1px solid var(--border); line-height: 1.4; }
      .notif-item:last-child { border-bottom: none; }
      .notif-item.urgent { color: var(--danger); font-weight: 600; }
      .notif-empty { font-size: 12.5px; color: var(--text-soft); padding: 4px 0 8px; }
      .notif-item-btn { display: block; width: 100%; text-align: left; background: none; border: none; font-family: inherit; font-size: 12.5px; line-height: 1.4; color: inherit; cursor: pointer; padding: 6px 0; }
      .notif-item-btn:hover { color: var(--accent); }

      .detail-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 0 32px; }
      .detail-row { display: flex; justify-content: space-between; gap: 12px; padding: 9px 0; border-bottom: 1px solid var(--border); font-size: 13px; }
      .detail-row:last-child { border-bottom: none; }
      .detail-row.detail-full { grid-column: 1 / -1; }
      .detail-label { color: var(--text-soft); font-weight: 500; }
      .detail-value { font-weight: 600; text-align: right; }
      .detail-transfer-history { margin-top: 14px; padding-top: 12px; border-top: 1px solid var(--border); }

      .notes-highlight { background: var(--accent-soft); border-left: 3px solid var(--accent); border-radius: 8px; padding: 10px 12px; font-size: 13px; line-height: 1.5; white-space: pre-wrap; }

      .maint-history-list { list-style: none; margin: 0; padding: 0; display: flex; flex-direction: column; gap: 6px; max-height: 220px; overflow-y: auto; }
      .maint-history-item { display: flex; align-items: center; gap: 10px; background: var(--bg); border: 1px solid var(--border); border-radius: 8px; padding: 8px 10px; font-size: 12.5px; }
      .maint-history-date { color: var(--text-soft); font-family: ui-monospace, monospace; font-size: 12px; white-space: nowrap; }
      .maint-history-desc { flex: 1; min-width: 0; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
      .maint-history-cost { color: var(--text-soft); white-space: nowrap; }

      .comments-section { margin-top: 14px; padding-top: 12px; border-top: 1px solid var(--border); }
      .comment-list { display: flex; flex-direction: column; gap: 8px; max-height: 260px; overflow-y: auto; margin-bottom: 10px; padding: 4px 2px; }
      .chat-row { display: flex; }
      .chat-row-left { justify-content: flex-start; }
      .chat-row-right { justify-content: flex-end; }
      .chat-bubble { max-width: 78%; border-radius: 14px; padding: 8px 12px; }
      .chat-bubble-left { background: var(--bg); border: 1px solid var(--border); border-bottom-left-radius: 4px; }
      .chat-bubble-right { background: var(--accent); color: white; border-bottom-right-radius: 4px; }
      .comment-meta { display: flex; justify-content: space-between; gap: 10px; font-size: 11px; margin-bottom: 4px; color: var(--text-soft); }
      .chat-bubble-right .comment-meta { color: rgba(255,255,255,0.8); }
      .comment-author { font-weight: 700; }
      .comment-text { font-size: 13px; line-height: 1.4; white-space: pre-wrap; }
      .comment-composer { display: flex; gap: 8px; align-items: flex-end; }
      .comment-composer textarea { flex: 1; border: 1px solid var(--border); border-radius: 8px; padding: 8px 10px; font-size: 13px; background: var(--bg); color: var(--text); font-family: inherit; resize: vertical; }

      .content { padding: 24px; flex: 1; overflow-y: auto; scrollbar-width: none; -ms-overflow-style: none; }
      .content::-webkit-scrollbar { display: none; }

      .welcome-banner { display: flex; align-items: center; justify-content: space-between; background: var(--accent-soft); border-radius: 16px; padding: 20px 24px; margin-bottom: 18px; gap: 16px; }
      .welcome-banner h2 { font-size: 20px; color: var(--text); }
      .welcome-sub { font-size: 13px; color: var(--text-soft); margin: 4px 0 0; }
      .welcome-icon { width: 52px; height: 52px; border-radius: 14px; background: var(--accent); color: white; display: flex; align-items: center; justify-content: center; flex-shrink: 0; }

      .metrics-row { display: grid; grid-template-columns: repeat(4, 1fr); gap: 14px; margin-bottom: 18px; }
      .metric { background: var(--surface); border-radius: 16px; padding: 16px 18px; box-shadow: 0 2px 8px rgba(0,0,0,0.06); }
      .metric-top { display: flex; align-items: center; gap: 8px; margin-bottom: 10px; }
      .metric-icon { width: 30px; height: 30px; border-radius: 9px; display: flex; align-items: center; justify-content: center; flex-shrink: 0; }
      .metric-value { font-size: 24px; font-weight: 700; font-family: 'Poppins', sans-serif; }
      .metric-label { font-size: 12px; color: var(--text-soft); font-weight: 600; }
      .metric-sub { font-size: 11px; color: var(--text-soft); margin-top: 3px; }

      .charts-row { display: grid; grid-template-columns: 1fr 1fr; gap: 14px; margin-bottom: 18px; }
      .charts-row-3 { grid-template-columns: repeat(3, 1fr); }
      .charts-row-4 { grid-template-columns: repeat(4, 1fr); }
      .chart-card { min-height: 230px; }
      .charts-row-3 .chart-card, .charts-row-4 .chart-card { min-height: 210px; padding-bottom: 4px; }
      .empty-chart { display: flex; align-items: center; justify-content: center; height: 220px; color: var(--text-soft); font-size: 13px; }
      .donut-body { display: flex; align-items: center; justify-content: center; gap: 14px; padding: 6px 18px 16px; }

      .panel { background: var(--surface); border: 1px solid var(--border); border-radius: 12px; overflow: hidden; margin-bottom: 14px; }
      .panel-head { padding: 14px 18px; border-bottom: 1px solid var(--border); display: flex; align-items: center; justify-content: space-between; }
      .panel-head h3 { font-size: 14.5px; }
      .panel-head-sub { padding: 10px 18px 8px; border-top: 1px solid var(--border); border-bottom: none; margin-top: 4px; }
      .panel-head-sub h3 { font-size: 12.5px; color: var(--text-soft); text-transform: uppercase; letter-spacing: 0.03em; }
      .panel-link { background: none; border: none; color: var(--accent); font-size: 12px; font-weight: 600; cursor: pointer; font-family: inherit; padding: 0; }
      .panel-link:hover { text-decoration: underline; }

      .legend-scroll-wrap { position: relative; flex: 0 1 190px; min-width: 0; max-width: 190px; }
      .legend-list {
        list-style: none; margin: 0; padding: 0; display: flex; flex-direction: column; gap: 8px; min-width: 0;
        max-height: 132px; overflow-y: auto; scrollbar-width: none; -ms-overflow-style: none;
      }
      .legend-list::-webkit-scrollbar { display: none; }
      .legend-list li { display: flex; align-items: center; gap: 7px; font-size: 12.5px; flex-shrink: 0; }
      .legend-dot { width: 8px; height: 8px; border-radius: 999px; flex-shrink: 0; }
      .legend-name { flex: 1; color: var(--text); white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
      .legend-count { color: var(--text-soft); font-weight: 600; white-space: nowrap; }
      .legend-pct { font-weight: 400; }
      .legend-scroll-hint {
        position: absolute; left: 50%; transform: translateX(-50%); color: var(--accent);
        background: var(--surface); border: 1px solid var(--border); border-radius: 999px;
        padding: 1px; pointer-events: none; animation: legendBob 1.3s ease-in-out infinite;
      }
      .legend-scroll-hint-down { bottom: -3px; }
      .legend-scroll-hint-up { top: -3px; }
      @keyframes legendBob { 0%, 100% { opacity: 0.6; transform: translate(-50%, 0); } 50% { opacity: 1; transform: translate(-50%, 2px); } }

      .bottom-row { display: grid; grid-template-columns: 1fr 1fr 1.3fr; gap: 14px; align-items: start; }
      .bottom-row-2 { grid-template-columns: 1.1fr 1fr; }
      .health-body { display: flex; align-items: center; gap: 12px; padding: 8px 14px 16px; }
      .health-ring-wrap { position: relative; width: 110px; height: 110px; flex-shrink: 0; }
      .health-ring-center { position: absolute; inset: 0; display: flex; flex-direction: column; align-items: center; justify-content: center; }
      .health-ring-pct { font-size: 19px; font-weight: 700; font-family: 'Poppins', sans-serif; color: #10B981; }
      .health-ring-label { font-size: 10px; color: var(--text-soft); }

      .warranty-stats { display: flex; padding: 8px 12px 18px; }
      .warranty-stat { flex: 1; text-align: center; padding: 8px; }
      .warranty-stat-icon { width: 30px; height: 30px; border-radius: 9px; display: flex; align-items: center; justify-content: center; margin: 0 auto 8px; }
      .warranty-stat-value { font-size: 19px; font-weight: 700; font-family: 'Poppins', sans-serif; }
      .warranty-stat-label { font-size: 10.5px; color: var(--text-soft); line-height: 1.3; margin-top: 3px; }

      .activity-list { list-style: none; margin: 0; padding: 6px 0; }
      .activity-item { display: flex; align-items: center; gap: 10px; padding: 9px 18px; }
      .activity-icon { width: 26px; height: 26px; border-radius: 999px; display: flex; align-items: center; justify-content: center; flex-shrink: 0; }
      .activity-text { flex: 1; font-size: 12.5px; color: var(--text); min-width: 0; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
      .activity-time { font-size: 10.5px; color: var(--text-soft); white-space: nowrap; flex-shrink: 0; }
      .activity-empty { padding: 14px 18px; font-size: 12.5px; color: var(--text-soft); }

      .table-wrap { overflow-x: auto; scrollbar-width: none; -ms-overflow-style: none; }
      .table-wrap::-webkit-scrollbar { display: none; }
      table { width: 100%; border-collapse: collapse; font-size: 13px; }
      thead th { text-align: left; padding: 11px 16px; color: var(--text-soft); font-weight: 700; font-size: 11px; text-transform: uppercase; letter-spacing: 0.04em; border-bottom: 1px solid var(--border); white-space: nowrap; background: color-mix(in srgb, var(--accent) 6%, var(--surface)); }
      tbody td { padding: 10px 16px; border-bottom: 1px solid var(--border); white-space: nowrap; }
      tbody tr:last-child td { border-bottom: none; }
      .mono { font-family: ui-monospace, monospace; font-size: 12.5px; }
      .cat-dot { display: inline-block; width: 8px; height: 8px; border-radius: 999px; margin-right: 7px; vertical-align: middle; }
      tbody tr:hover { background: color-mix(in srgb, var(--accent) 4%, transparent); }
      .asset-row { cursor: pointer; }
      .asset-row .checkbox-cell { cursor: default; }
      .history-row { cursor: default; opacity: 0.75; }
      .history-row:hover { background: transparent !important; }
      .empty-cell { text-align: center; color: var(--text-soft); padding: 28px !important; white-space: normal; }

      /* Keep row quick-actions reachable without horizontal scrolling on
         wide tables — pinned to the right edge of the table, above the
         other cells as the table scrolls under it. */
      .actions-th, td.actions-cell { position: sticky; right: 0; z-index: 3; }
      .actions-th { background: color-mix(in srgb, var(--accent) 6%, var(--surface)); }
      td.actions-cell { background: var(--surface); }
      tbody tr:hover td.actions-cell { background: color-mix(in srgb, var(--accent) 4%, var(--surface)); }

      .badge { padding: 3px 10px; border-radius: 999px; font-size: 11.5px; font-weight: 600; white-space: nowrap; display: inline-block; }
      .status-select { font-size: 12px; font-weight: 600; border-radius: 999px; padding: 4px 10px; background: var(--surface); border: 1px solid; }

      .view-head { display: flex; align-items: center; justify-content: space-between; gap: 12px; margin-bottom: 16px; flex-wrap: wrap; }
      .view-title { font-size: 18px; }
      .view-actions { display: flex; gap: 8px; }
      .search-box { display: flex; align-items: center; gap: 8px; background: var(--surface); border: 1px solid var(--border); border-radius: 9px; padding: 8px 12px; width: 280px; color: var(--text-soft); }
      .search-box input { border: none; outline: none; background: none; color: var(--text); font-size: 13px; width: 100%; }
      .sort-select { border: 1px solid var(--border); border-radius: 9px; padding: 8px 12px; font-size: 13px; background: var(--surface); color: var(--text); }
      .filter-group { display: flex; gap: 8px; flex-wrap: wrap; }
      .export-hint { font-size: 12.5px; color: var(--text-soft); margin-bottom: 10px; }
      .export-textarea { width: 100%; font-family: ui-monospace, monospace; font-size: 12px; border: 1px solid var(--border); border-radius: 8px; padding: 10px; background: var(--bg); color: var(--text); resize: vertical; }

      .btn { display: inline-flex; align-items: center; gap: 6px; padding: 8px 14px; border-radius: 9px; font-size: 13px; font-weight: 600; cursor: pointer; border: 1px solid transparent; appearance: none; -webkit-appearance: none; font-family: inherit; }
      .btn.primary { background: var(--accent); color: white; }
      .btn.ghost { background: var(--surface); border-color: var(--border); color: var(--text); }
      .btn.danger-outline { background: var(--surface); border-color: var(--danger); color: var(--danger); }
      .btn.danger { background: var(--danger); color: white; }
      .btn.full { width: 100%; justify-content: center; margin-top: 6px; }
      .btn:disabled { opacity: 0.5; cursor: not-allowed; }

      .icon-btn { display: flex; align-items: center; justify-content: center; width: 30px; height: 30px; border-radius: 8px; border: 1px solid var(--border); background: var(--surface); color: var(--text-soft); cursor: pointer; transition: background 0.15s ease, border-color 0.15s ease, color 0.15s ease, transform 0.1s ease; }
      .icon-btn:hover { color: var(--accent); border-color: var(--accent); }
      .icon-btn.danger { color: var(--danger); border-color: color-mix(in srgb, var(--danger) 45%, transparent); background: color-mix(in srgb, var(--danger) 8%, var(--surface)); }
      .icon-btn.danger:hover { color: white; background: var(--danger); border-color: var(--danger); }
      .icon-btn:active { transform: scale(0.94); }
      .row-actions { display: flex; gap: 6px; }

      .modal-head { display: flex; align-items: center; justify-content: space-between; padding: 16px 20px; border-bottom: 1px solid var(--border); }
      .modal-head h3 { font-size: 15px; }
      .modal-body { padding: 18px; }
      .modal.confirm { max-width: 340px; padding: 22px; text-align: center; }
      .confirm-icon { width: 40px; height: 40px; border-radius: 999px; background: #FEE2E2; color: #DC2626; display: flex; align-items: center; justify-content: center; margin: 0 auto 12px; }
      .confirm-actions { display: flex; gap: 8px; justify-content: center; margin-top: 16px; }

      .form-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 14px; }
      .form-full { grid-column: 1 / -1; }
      .field { display: flex; flex-direction: column; gap: 4px; font-size: 12.5px; color: var(--text-soft); font-weight: 500; }
      .field input, .field select, .field textarea {
        border: 1px solid var(--border); border-radius: 7px; padding: 7px 10px; font-size: 13.5px;
        background: var(--bg); color: var(--text); font-family: inherit;
      }
      .field input:disabled, .field select:disabled, .field textarea:disabled { opacity: 0.55; cursor: not-allowed; }
      .field-hint { display: block; margin-top: 5px; font-size: 11.5px; color: var(--text-soft); }
      .field-inline { display: flex; gap: 6px; flex-wrap: wrap; }
      .field-inline input { flex: 1; min-width: 120px; }
      .sn-check-btn { white-space: nowrap; padding: 0 10px; font-size: 12px; }
      .modal-actions { display: flex; justify-content: flex-end; gap: 8px; margin-top: 6px; }
      .field-label-row { display: flex; align-items: center; justify-content: space-between; gap: 8px; }
      .required-mark { color: var(--danger); margin-left: 2px; }

      /* New Asset form — compact sections separated by hairline dividers
         (not nested cards), Linear/Notion-style density. Divider is drawn
         between consecutive sections via the adjacent-sibling combinator,
         so it never mistakenly lands above the sticky footer. */
      .form-section { padding: 14px 0; }
      .form-section + .form-section { border-top: 1px solid var(--border); }
      .form-section-head { display: flex; align-items: center; gap: 7px; margin-bottom: 10px; }
      .form-section-head-icon { width: 15px; height: 15px; color: var(--accent); display: flex; align-items: center; justify-content: center; flex-shrink: 0; }
      .form-section-head h4 { font-size: 11px; font-weight: 700; letter-spacing: 0.05em; text-transform: uppercase; color: var(--text-soft); }
      .form-section-head-action { margin-left: auto; }
      .section-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 12px 14px; }
      .section-grid .form-full { grid-column: 1 / -1; }

      /* Toggle switch */
      .toggle-switch-wrap { display: inline-flex; align-items: center; gap: 6px; cursor: pointer; font-size: 11px; color: var(--text-soft); font-weight: 600; user-select: none; }
      .toggle-switch { position: relative; display: inline-block; width: 28px; height: 16px; flex-shrink: 0; }
      .toggle-switch input { opacity: 0; width: 0; height: 0; position: absolute; }
      .toggle-slider { position: absolute; inset: 0; background: var(--border); border-radius: 999px; transition: background 0.15s ease; }
      .toggle-slider::before { content: ""; position: absolute; width: 12px; height: 12px; left: 2px; top: 2px; background: #fff; border-radius: 999px; transition: transform 0.15s ease; box-shadow: 0 1px 2px rgba(0,0,0,0.25); }
      .toggle-switch input:checked + .toggle-slider { background: var(--accent); }
      .toggle-switch input:checked + .toggle-slider::before { transform: translateX(12px); }

      /* Status picker (colored badge dropdown) */
      .status-picker { position: relative; }
      .status-picker-trigger { display: flex; align-items: center; justify-content: space-between; width: 100%; border: 1px solid var(--border); border-radius: 7px; padding: 5px 9px; background: var(--bg); cursor: pointer; color: var(--text-soft); font-family: inherit; font-size: 13px; }
      .status-picker-placeholder { color: var(--text-soft); }
      .status-dot-badge { display: inline-flex; align-items: center; gap: 5px; padding: 2px 9px; border-radius: 999px; font-size: 11.5px; font-weight: 600; }
      .status-dot { width: 6px; height: 6px; border-radius: 999px; flex-shrink: 0; }
      .status-picker-menu { position: absolute; top: calc(100% + 4px); left: 0; right: 0; background: var(--surface); border: 1px solid var(--border); border-radius: 10px; box-shadow: 0 10px 26px rgba(0,0,0,0.16); padding: 4px; z-index: 20; }
      .status-picker-option { display: flex; align-items: center; gap: 7px; width: 100%; padding: 6px 7px; border-radius: 6px; border: none; background: none; cursor: pointer; font-family: inherit; font-size: 12.5px; color: var(--text); text-align: left; }
      .status-picker-option:hover { background: var(--accent-soft); color: var(--accent); }

      /* Searchable select */
      .searchable-select { position: relative; }
      .searchable-select input { width: 100%; }
      .searchable-select-menu { position: absolute; top: calc(100% + 4px); left: 0; right: 0; max-height: 190px; overflow-y: auto; background: var(--surface); border: 1px solid var(--border); border-radius: 10px; box-shadow: 0 10px 26px rgba(0,0,0,0.16); padding: 4px; z-index: 20; scrollbar-width: none; }
      .searchable-select-menu::-webkit-scrollbar { display: none; }
      .searchable-select-option { padding: 6px 8px; border-radius: 6px; cursor: pointer; font-size: 12.5px; }
      .searchable-select-option:hover { background: var(--accent-soft); color: var(--accent); }
      .searchable-select-clear { position: absolute; right: 6px; top: 50%; transform: translateY(-50%); background: none; border: none; color: var(--text-soft); cursor: pointer; padding: 2px; display: flex; }

      /* Repair-reason callout + optional collapsible section */
      .repair-reason-box { background: color-mix(in srgb, var(--danger) 6%, var(--bg)); border: 1px solid color-mix(in srgb, var(--danger) 30%, var(--border)); border-radius: 9px; padding: 10px 12px; }
      .optional-section-toggle { display: flex; align-items: center; justify-content: space-between; width: 100%; background: none; border: none; cursor: pointer; padding: 0; color: var(--text-soft); font-family: inherit; }
      .optional-section-toggle:hover .form-section-head h4 { color: var(--accent); }
      .optional-section-body { margin-top: 10px; }

      /* Sticky footer for the New/Edit Asset form */
      .modal-footer-sticky { position: sticky; bottom: 0; margin: 14px -18px -18px; padding: 12px 18px; background: var(--surface); border-top: 1px solid var(--border); display: flex; justify-content: flex-end; gap: 8px; border-radius: 0 0 14px 14px; }
      .sort-th-btn { display: inline-flex; align-items: center; gap: 4px; background: none; border: none; padding: 0; font: inherit; font-weight: inherit; color: inherit; cursor: pointer; }
      .sort-th-idle { opacity: 0.35; }
      .name-subtext { font-size: 11.5px; color: var(--text-soft); font-weight: 400; margin-top: 2px; }
      .hint-box { background: var(--accent-soft); color: var(--accent); font-size: 12px; padding: 10px 12px; border-radius: 8px; }

      .spin { animation: spin 0.8s linear infinite; }
      @keyframes spin { from { transform: rotate(0deg); } to { transform: rotate(360deg); } }

      .toast { position: fixed; bottom: 20px; left: 50%; background: var(--text); color: var(--bg); padding: 10px 18px; border-radius: 999px; font-size: 13px; z-index: 100; box-shadow: 0 8px 24px rgba(0,0,0,0.18); }
      .toast-in { animation: toastIn 0.2s ease-out forwards; }
      .toast-out { animation: toastOut 0.2s ease-in forwards; }
      @keyframes toastIn { from { opacity: 0; transform: translate(-50%, 8px); } to { opacity: 1; transform: translate(-50%, 0); } }
      @keyframes toastOut { from { opacity: 1; transform: translate(-50%, 0); } to { opacity: 0; transform: translate(-50%, 8px); } }

      .modal-overlay { position: fixed; inset: 0; background: rgba(0,0,0,0.4); display: flex; align-items: center; justify-content: center; z-index: 50; padding: 16px; animation: overlayIn 0.15s ease-out; }
      .modal { background: var(--surface); border-radius: 14px; width: 100%; max-height: 88vh; overflow-y: auto; animation: modalIn 0.16s ease-out; scrollbar-width: none; -ms-overflow-style: none; box-shadow: 0 16px 40px rgba(0,0,0,0.2); }
      .fab-add { display: none; }
      .modal::-webkit-scrollbar { display: none; }
      @keyframes overlayIn { from { opacity: 0; } to { opacity: 1; } }
      @keyframes modalIn { from { opacity: 0; transform: scale(0.97) translateY(4px); } to { opacity: 1; transform: scale(1) translateY(0); } }

      @media (max-width: 1100px) {
        .charts-row-3 { grid-template-columns: 1fr 1fr; }
        .charts-row-4 { grid-template-columns: 1fr 1fr; }
        .bottom-row { grid-template-columns: 1fr 1fr; }
      }
      @media (max-width: 860px) {
        .sidebar { position: fixed; z-index: 70; height: 100vh; height: 100dvh; top: 0; left: 0; width: 240px; transform: translateX(-100%); box-shadow: 0 0 0 rgba(0,0,0,0); transition: transform 0.2s ease; }
        .sidebar:not(.collapsed) { transform: translateX(0); box-shadow: 12px 0 32px rgba(0,0,0,0.18); }
        .sidebar.collapsed { width: 240px; }
        .sidebar-backdrop { display: block; position: fixed; inset: 0; background: rgba(0,0,0,0.4); z-index: 65; }
        .main { width: 100%; }
        .content { padding: 14px; }
        .topbar { padding: 10px 12px; }
        .topbar-region span { display: none; }
        .synced-label { display: none; }
        .user-meta { display: none; }
        .user-chip { padding: 4px; }
        .metrics-row { grid-template-columns: repeat(2, 1fr); gap: 10px; }
        .bottom-row { grid-template-columns: 1fr; }
        .welcome-banner { padding: 16px 18px; border-radius: 14px; }
        .welcome-banner h2 { font-size: 17px; }
        .welcome-icon { width: 42px; height: 42px; border-radius: 12px; }
        .metric { padding: 12px 14px; }
        .metric-value { font-size: 21px; }
        .charts-row { grid-template-columns: 1fr; }
        .form-grid { grid-template-columns: 1fr; }
        .section-grid { grid-template-columns: 1fr; }
        .form-section { padding: 14px; }
        .modal-footer-sticky { border-radius: 0; }
        .detail-grid { grid-template-columns: 1fr; }
        .search-box { width: 100%; }
        .view-head { flex-direction: column; align-items: stretch; }
        .view-actions { flex-wrap: wrap; }
        .view-actions .btn { flex: 1; justify-content: center; }
        .new-asset-btn { display: none; }
        .fab-add {
          display: flex; align-items: center; justify-content: center;
          position: fixed; bottom: 20px; right: 20px; width: 52px; height: 52px;
          border-radius: 999px; background: var(--accent); color: white; border: none;
          box-shadow: 0 8px 24px rgba(0,0,0,0.28); z-index: 40; cursor: pointer;
        }
        .fab-add:active { transform: scale(0.94); }
        .modal-overlay { padding: 0; align-items: flex-end; }
        .modal { max-width: 100% !important; width: 100%; max-height: 92vh; border-radius: 16px 16px 0 0; }
        .notif-panel { width: calc(100vw - 24px); right: -12px; }

        /* Card-style tables: each row becomes a stacked card */
        .table-wrap table, .table-wrap thead, .table-wrap tbody, .table-wrap th, .table-wrap td, .table-wrap tr {
          display: block;
        }
        .table-wrap thead { position: absolute; top: -9999px; left: -9999px; }
        .table-wrap tbody tr {
          background: var(--surface); border: 1px solid var(--border); border-radius: 10px;
          margin-bottom: 10px; padding: 4px 12px;
        }
        .table-wrap tbody tr:last-child { margin-bottom: 0; }
        .table-wrap td {
          display: flex; align-items: center; justify-content: space-between; gap: 10px;
          padding: 8px 0; border-bottom: 1px solid var(--border); white-space: normal; text-align: right;
        }
        .table-wrap td:last-child { border-bottom: none; }
        .table-wrap td[data-label]::before {
          content: attr(data-label); font-size: 11px; font-weight: 700; text-transform: uppercase;
          letter-spacing: 0.03em; color: var(--text-soft); text-align: left; margin-right: 10px;
        }
        .table-wrap td:not([data-label]) { justify-content: flex-end; }
        .table-wrap td.checkbox-cell { justify-content: flex-start; }
        .table-wrap td.actions-cell { justify-content: flex-end; position: static; }
        .row-actions { flex-wrap: wrap; justify-content: flex-end; }
      }

      @media (max-width: 480px) {
        .metrics-row { grid-template-columns: 1fr 1fr; }
        .login-card { width: 100%; padding: 24px; }
      }
    `}</style>
  );
}
