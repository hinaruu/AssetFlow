import React, { useState, useEffect, useCallback, useMemo, useRef } from "react";
import {
  LayoutDashboard, Package, Wrench, MapPin, Tags, Users, User, LogOut,
  Menu, Sun, Moon, Plus, Pencil, Trash2, Download, Upload, X, Search,
  KeyRound, ShieldCheck, AlertTriangle, Info, Clock,
  Bell, Copy, Truck, CheckSquare, Archive, ExternalLink,
  ChevronUp, ChevronDown, ChevronsUpDown, MessageCircle, Check,
} from "lucide-react";
import { PieChart, Pie, Cell, ResponsiveContainer, Legend, Tooltip } from "recharts";
import * as XLSX from "xlsx";
import { fetchOrgData, saveOrgData, subscribeToOrgData, supabase } from "./lib/supabase.js";

/* ---------------------------------------------------------
   Helpers
--------------------------------------------------------- */
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
// assetId (when the entry is about a single asset) is only ever used
// internally to power the asset-activity notification feed below — it's
// never shown in the Activity Log itself, so that view is unchanged.
function withLog(data, currentUser, message, locationId = null, assetId = null) {
  const entry = {
    id: uid("log"),
    at: new Date().toISOString(),
    userId: currentUser?.id || null,
    userName: currentUser?.name || "Unknown",
    message,
    locationId: locationId || null,
    assetId: assetId || null,
  };
  const auditLog = [entry, ...(data.auditLog || [])].slice(0, 300);
  return { ...data, auditLog };
}

/* ---------------------------------------------------------
   Asset activity notifications (Overall Admin / Regional Admin)

   Complements the Activity Log without changing it. Every asset-tagged
   audit log entry (status/assignment/disposal/edit/transfer/maintenance,
   etc — see the assetId argument on withLog calls below) plus every
   comment is treated as one "activity event" for its asset. An asset
   counts as unread for a user when it has an event newer than that
   user's notification_reads row for it (see withAssetRead) — so multiple
   updates naturally accumulate as unread until the asset is opened again,
   with no per-event bookkeeping needed.

   Comment events deliberately come only from the comments table (the
   "Commented on asset" audit log entry for the same action is left
   untagged — see addComment) so a single comment isn't counted twice.
--------------------------------------------------------- */
function computeAssetActivityFeed(auditLog, comments) {
  const events = [];
  (auditLog || []).forEach((e) => {
    if (!e.assetId) return;
    events.push({ assetId: e.assetId, at: e.at, message: e.message });
  });
  (comments || []).forEach((c) => {
    if (!c.assetId) return;
    events.push({ assetId: c.assetId, at: c.at, message: `${c.authorName || "Someone"} commented: "${c.message}"` });
  });
  events.sort((a, b) => new Date(b.at) - new Date(a.at));
  return events;
}

// Groups activity events by asset, keeping only ones newer than the
// user's last-read timestamp for that asset (unread) and within scope
// (inScope receives an assetId and returns whether the current user
// should be notified about that asset at all — see NOTIFICATION SCOPE
// in TopBar/AssetsView). Returns a Map<assetId, { latest, count }>.
function computeUnreadAssetActivity(events, notificationReads, userId, inScope) {
  const readMap = new Map(
    (notificationReads || [])
      .filter((r) => r.userId === userId)
      .map((r) => [r.assetId, r.lastReadAt])
  );
  const byAsset = new Map();
  events.forEach((e) => {
    if (!inScope(e.assetId)) return;
    const lastRead = readMap.get(e.assetId);
    if (lastRead && new Date(e.at) <= new Date(lastRead)) return;
    const entry = byAsset.get(e.assetId);
    if (!entry) byAsset.set(e.assetId, { latest: e, count: 1 });
    else entry.count += 1;
  });
  return byAsset;
}

// Marks an asset as "seen" by a user right now — used both when a
// notification is clicked and when the asset's detail view is opened.
function withAssetRead(data, userId, assetId) {
  if (!userId || !assetId) return data;
  const id = `ntr-${userId}-${assetId}`;
  const now = new Date().toISOString();
  const exists = (data.notificationReads || []).some((r) => r.id === id);
  const notificationReads = exists
    ? (data.notificationReads || []).map((r) => (r.id === id ? { ...r, lastReadAt: now } : r))
    : [...(data.notificationReads || []), { id, userId, assetId, lastReadAt: now }];
  return { ...data, notificationReads };
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
// Disposed also drops condition to Poor since it's no longer in service at
// all, while In Stock just means it's back on the shelf.
function applyStatusSideEffects(asset) {
  if (asset.status === "Disposed") {
    return { ...asset, assignedTo: "", condition: "Poor", department: "" };
  }
  // Any other status means the asset is back in active service — clear
  // disposal details left over from a previous Disposed period.
  const cleared = asset.disposalInfo ? { ...asset, disposalInfo: null } : asset;
  if (asset.status === "In Stock") {
    return { ...cleared, assignedTo: "", department: "" };
  }
  return cleared;
}

// One-time migration: Retired and Disposed used to be separate statuses;
// they've since been merged into a single "Disposed" status. Any asset
// still carrying the old "Retired" value (from before this change, or
// pushed by a stale client) gets converted the moment it's loaded. Returns
// the same object unchanged (changed: false) when there's nothing to do,
// so callers can skip re-saving.
function migrateRetiredToDisposed(orgData) {
  if (!orgData?.assets?.some((a) => a.status === "Retired")) return { data: orgData, changed: false };
  const assets = orgData.assets.map((a) => (a.status === "Retired" ? { ...a, status: "Disposed" } : a));
  return { data: { ...orgData, assets }, changed: true };
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
  let list = scopedLocationId ? assets.filter((a) => a.locationId === scopedLocationId) : assets;
  list = list.filter((a) => a.status !== "Disposed");
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

// Per-asset version of the same warranty/calibration windows used by
// computeAlerts(), plus condition — powers the "Needs Attention" banner
// in the Asset Details modal. Data-driven on purpose: nobody sets this by
// hand, it's derived fresh from the asset's own fields every time.
function getAssetIssues(asset) {
  const issues = [];
  if (!asset || asset.status === "Disposed") return issues;
  const now = new Date();
  const in30 = new Date(now.getTime() + 30 * 24 * 60 * 60 * 1000);
  const plainDate = (v) => {
    const d = new Date(v);
    return isNaN(d) ? v : d.toLocaleDateString([], { month: "short", day: "numeric", year: "numeric" });
  };

  if (asset.assetType === "IT" && asset.warrantyExpiry) {
    const d = new Date(asset.warrantyExpiry);
    if (!isNaN(d) && d <= in30) {
      const expired = d < now;
      issues.push({
        severity: expired ? "critical" : "warning",
        label: expired ? "Warranty Expired" : "Warranty Expiring Soon",
        detail: `Warranty ${expired ? "expired" : "expires"} on ${plainDate(asset.warrantyExpiry)}.`,
      });
    }
  }

  if (asset.assetType === "Non-IT" && asset.requiresCalibration) {
    const checkDate = asset.nextCalibrationDate || asset.calibrationDate;
    if (checkDate) {
      const d = new Date(checkDate);
      if (!isNaN(d) && d <= in30) {
        const overdue = d < now;
        issues.push({
          severity: overdue ? "critical" : "warning",
          label: overdue ? "Calibration Overdue" : "Calibration Due Soon",
          detail: `Calibration ${overdue ? "was due" : "is due"} on ${plainDate(checkDate)}.`,
        });
      }
    }
  }

  if (asset.condition === "Poor") {
    issues.push({ severity: "critical", label: "Poor Condition", detail: "This asset is recorded in poor condition and may need service or replacement." });
  } else if (asset.condition === "Fair") {
    issues.push({ severity: "warning", label: "Fair Condition", detail: "This asset is recorded in fair condition — worth keeping an eye on." });
  }

  return issues;
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
const CONDITION_COLORS = { "New": "#3B82F6", "Good": "#10B981", "Fair": "#F59E0B", "Poor": "#EF4444" };
const DISPOSAL_REASON_OPTIONS = [
  "End of Life (EOL)", "Hardware failure", "Motherboard failure", "Damaged beyond repair",
  "Lost / Stolen", "Obsolete / Replaced", "Other",
];
const MAINT_STATUS = ["Not Started", "In Progress", "Done"];

const STATUS_COLORS = {
  "In Stock": "#3B82F6",
  "In Use": "#10B981",
  "Under Repair": "#F59E0B",
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
function LoginScreen() {
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [error, setError] = useState("");
  const [busy, setBusy] = useState(false);

  const submit = async (e) => {
    e.preventDefault();
    setError("");
    setBusy(true);
    try {
      if (!email.trim() || !password) {
        setError("Enter both email and password.");
        return;
      }
      // Real Supabase Auth sign-in — no passwords ever compared in the
      // browser. A successful call updates the session, which the app
      // picks up automatically via onAuthStateChange (see App()).
      const { error: authError } = await supabase.auth.signInWithPassword({
        email: email.trim(),
        password,
      });
      if (authError) {
        setError("Incorrect email or password.");
      }
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
          <Field label="Email">
            <input type="email" value={email} onChange={(e) => setEmail(e.target.value)} autoFocus placeholder="name@company.com" />
          </Field>
          <Field label="Password">
            <input type="password" value={password} onChange={(e) => setPassword(e.target.value)} placeholder="••••••••" />
          </Field>
          {error && <div className="form-error">{error}</div>}
          <button className="btn primary full" disabled={busy} type="button" onClick={submit}>
            {busy ? "Signing in…" : "Sign In"}
          </button>
        </div>
      </div>
    </div>
  );
}

// Shown when someone has a valid Supabase Auth session but no matching row
// in the `users` profile table yet (e.g. their Auth account was just
// created and hasn't been linked). See CRITICAL-SECURITY-STEPS.md, Phase 3.
function NoProfileScreen({ email, onSignOut }) {
  return (
    <div className="login-wrap">
      <div className="login-card" style={{ textAlign: "center" }}>
        <div className="confirm-icon"><AlertTriangle size={20} /></div>
        <h3 style={{ marginBottom: 8 }}>Account not set up yet</h3>
        <p className="login-sub">
          You're signed in as <strong>{email}</strong>, but there's no matching profile for
          you in this app yet. Ask an Admin to add your name/role/location under
          <strong> User Accounts</strong> and link it to this login.
        </p>
        <button className="btn primary full" onClick={onSignOut}>Sign Out</button>
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
  // `session` is Supabase Auth's session object — undefined until the first
  // check completes, null when signed out, an object when signed in.
  // `currentUser` is derived from it below (the matching row in
  // data.users), not stored as its own state — there is exactly one
  // source of truth for "who's logged in" now: the Auth session.
  const [session, setSession] = useState(undefined);
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
      // No demo user accounts are seeded here anymore — logins are real
      // Supabase Auth accounts now, which can't be created from the
      // browser with just the anon key. Create your first Admin account
      // per CRITICAL-SECURITY-STEPS.md (Phase 3) instead.
      orgData.users = [];
      await saveOrgData(orgData);
    }
    return orgData;
  }, []);

  // Tracks the Supabase Auth session — this is the actual security
  // boundary now (paired with RLS policies keyed off auth.uid()), not
  // just app state. onAuthStateChange fires on sign-in, sign-out, and
  // token refresh, so this always reflects the real session.
  useEffect(() => {
    supabase.auth.getSession().then(({ data: { session } }) => setSession(session));
    const { data: sub } = supabase.auth.onAuthStateChange((_event, session) => setSession(session));
    return () => sub.subscription.unsubscribe();
  }, []);

  // The signed-in user's profile — found by matching the Auth session's
  // user id against users.authUserId (see UsersView for how that link is
  // made). null while there's no session, or if the session has no
  // matching profile row yet.
  const currentUser = useMemo(() => {
    if (!session || !data) return null;
    return data.users.find((u) => u.authUserId === session.user.id) || null;
  }, [session, data]);

  // Load persisted UI prefs on mount — independent of auth, always safe.
  useEffect(() => {
    try {
      const t = localStorage.getItem("theme-pref");
      if (t) setTheme(t);
      const s = localStorage.getItem("sidebar-pref");
      if (s) setSidebarOpen(s === "open");
    } catch {}
  }, []);

  // Load the org's data — only once there's a real signed-in session.
  // RLS now requires auth.uid() to be set for reads on every table, so
  // fetching before sign-in would just get blocked; there's also nothing
  // useful to show pre-login anymore since the login screen no longer
  // needs `data.users` (it calls Supabase Auth directly). Resets back to
  // "not loaded" on sign-out so a different account's session doesn't
  // reuse stale data.
  useEffect(() => {
    if (session === undefined) return; // still checking for a session
    if (!session) {
      setData(null);
      setConnectionError(null);
      setLoaded(true);
      return;
    }
    let cancelled = false;
    setLoaded(false);
    (async () => {
      try {
        const orgData = await loadFromCloud({ seedIfEmpty: true });
        if (cancelled) return;
        const { data: migrated, changed } = migrateRetiredToDisposed(orgData);
        setData(migrated);
        setConnectionError(null);
        setLastSynced(new Date());
        if (changed) saveOrgData(migrated, orgData).catch(() => {});
      } catch (err) {
        if (!cancelled) setConnectionError(err?.message || "Could not connect to the database.");
      } finally {
        if (!cancelled) setLoaded(true);
      }
    })();
    return () => { cancelled = true; };
  }, [session, loadFromCloud]);

  // Live feed: any change anyone saves is pushed here automatically —
  // no manual sync needed. The Sync button still works as a manual
  // fallback (e.g. right after reconnecting). Only subscribes once
  // signed in — Realtime enforces the same RLS policies as regular
  // queries, so an anonymous subscription wouldn't receive anything
  // anyway, and would just churn reconnect attempts.
  useEffect(() => {
    if (!session) return;
    const unsubscribe = subscribeToOrgData((next) => {
      const { data: migrated, changed } = migrateRetiredToDisposed(next);
      setData(migrated);
      setLastSynced(new Date());
      if (changed) saveOrgData(migrated, next).catch(() => {});
    });
    return unsubscribe;
  }, [session]);

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

  if (!loaded || session === undefined) {
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
              onClick={() => { setLoaded(false); setConnectionError(null); loadFromCloud({ seedIfEmpty: true }).then((d) => { const { data: migrated, changed } = migrateRetiredToDisposed(d); setData(migrated); setLastSynced(new Date()); if (changed) saveOrgData(migrated, d).catch(() => {}); }).catch((e) => setConnectionError(e?.message || "Could not connect.")).finally(() => setLoaded(true)); }}
            >
              Retry
            </button>
          </div>
        </div>
      </div>
    );
  }

  if (session && !currentUser) {
    return (
      <div className={theme === "dark" ? "theme-dark" : "theme-light"}>
        <GlobalStyles />
        <NoProfileScreen email={session.user.email} onSignOut={() => supabase.auth.signOut()} />
      </div>
    );
  }

  if (!currentUser) {
    return (
      <div className={theme === "dark" ? "theme-dark" : "theme-light"}>
        <GlobalStyles />
        <LoginScreen />
      </div>
    );
  }

  const isAdmin = currentUser.role === "Admin";
  // A Regional Admin is scoped to a single location just like Regional Staff,
  // but is trusted to delete assets in that location without anyone's
  // approval, and is the one who approves/rejects Regional Staff deletion
  // requests for that same location (see locationApprover / ApprovalsView).
  const isRegionalAdmin = currentUser.role === "Regional Admin";
  const canDeleteDirectly = isAdmin || isRegionalAdmin;
  const scopedLocationId = isAdmin ? null : currentUser.locationId;

  // Who a pending deletion request for a given location should be routed
  // to: the Regional Admin assigned to that location if one exists,
  // otherwise the Overall Admin as a fallback. Recomputed live off the
  // current user list, so adding/removing a Regional Admin immediately
  // changes where new AND already-pending requests are routed.
  const locationApprover = (locationId) => data.users.find((u) => u.role === "Regional Admin" && u.locationId === locationId) || null;

  const pendingCount = isAdmin
    ? data.assets.filter((a) => a.pendingDeletion && !locationApprover(a.locationId)).length
    : isRegionalAdmin
    ? data.assets.filter((a) => a.pendingDeletion && a.locationId === currentUser.locationId).length
    : 0;

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
          isRegionalAdmin={isRegionalAdmin}
          pendingCount={pendingCount}
        />
        <div className="main">
          <TopBar
            theme={theme}
            toggleTheme={toggleTheme}
            currentUser={currentUser}
            onLogout={() => supabase.auth.signOut()}
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
                isRegionalAdmin={isRegionalAdmin}
                canDeleteDirectly={canDeleteDirectly}
                scopedLocationId={scopedLocationId}
                showToast={showToast}
                currentUser={currentUser}
                focusAssetId={focusAssetId}
                onFocusHandled={() => setFocusAssetId(null)}
                applyLocalOnly={setData}
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
              <ActivityLogView data={data} isAdmin={isAdmin} scopedLocationId={scopedLocationId} persist={persist} showToast={showToast} currentUser={currentUser} />
            )}
            {view === "approvals" && (isAdmin || isRegionalAdmin) && (
              <ApprovalsView data={data} persist={persist} showToast={showToast} currentUser={currentUser} isAdmin={isAdmin} isRegionalAdmin={isRegionalAdmin} />
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
function Sidebar({ open, onToggle, view, setView, isAdmin, isRegionalAdmin, pendingCount }) {
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
    ...(isAdmin || isRegionalAdmin ? [
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
  // NOTIFICATION SCOPE — Overall Admin gets asset activity across every
  // location; Regional Admin gets it for their own location (their own
  // actions plus Regional Staff's within it). Regional Staff keeps the
  // original targeted-comment notifications below, unchanged.
  const isRegionalAdmin = currentUser.role === "Regional Admin";
  const broadNotifScope = isAdmin || isRegionalAdmin;
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
    () => (broadNotifScope ? [] : (data.comments || []).filter((c) => {
      if (!(c.targetUserIds || []).includes(currentUser.id) || (c.readBy || []).includes(currentUser.id)) return false;
      // Live access check: even if targetUserIds was set at comment time,
      // an asset that has since been transferred away from this user's
      // location should no longer surface a notification for it.
      const asset = data.assets.find((a) => a.id === c.assetId);
      return !!asset && asset.locationId === scopedLocationId;
    }).sort((a, b) => new Date(b.at) - new Date(a.at))),
    [data.comments, data.assets, currentUser.id, scopedLocationId, broadNotifScope]
  );

  // Overall Admin / Regional Admin: any significant update on any asset in
  // scope (comments, notes, chat, status/assignment changes, disposal —
  // see computeAssetActivityFeed), grouped per asset so it reads as "this
  // asset has N new updates" rather than one row per event.
  const assetActivityEvents = useMemo(
    () => (broadNotifScope ? computeAssetActivityFeed(data.auditLog, data.comments) : []),
    [data.auditLog, data.comments, broadNotifScope]
  );
  const activityScopeAssetIds = useMemo(() => {
    if (!broadNotifScope || isAdmin) return null; // null = every asset (Overall Admin)
    return new Set(data.assets.filter((a) => a.locationId === scopedLocationId).map((a) => a.id));
  }, [data.assets, broadNotifScope, isAdmin, scopedLocationId]);
  const unreadActivityMap = useMemo(() => {
    if (!broadNotifScope) return new Map();
    return computeUnreadAssetActivity(
      assetActivityEvents,
      data.notificationReads,
      currentUser.id,
      (assetId) => activityScopeAssetIds === null || activityScopeAssetIds.has(assetId)
    );
  }, [broadNotifScope, assetActivityEvents, data.notificationReads, currentUser.id, activityScopeAssetIds]);
  const activityNotifs = useMemo(
    () => Array.from(unreadActivityMap.entries())
      .map(([assetId, info]) => ({ assetId, ...info }))
      .sort((a, b) => new Date(b.latest.at) - new Date(a.latest.at)),
    [unreadActivityMap]
  );

  const notifCount = alerts.length + myComments.length + activityNotifs.length;

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

  const openAssetActivity = (assetId) => {
    if (persist) persist(withAssetRead(data, currentUser.id, assetId));
    setNotifOpen(false);
    if (onOpenAsset) onOpenAsset(assetId);
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
                {alerts.length === 0 && myComments.length === 0 && activityNotifs.length === 0 && (
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
                {activityNotifs.slice(0, 10).map((n) => {
                  const asset = data.assets.find((a) => a.id === n.assetId);
                  return (
                    <button
                      key={n.assetId}
                      type="button"
                      className="notif-item notif-item-btn"
                      onClick={() => openAssetActivity(n.assetId)}
                      title="View asset"
                    >
                      {n.count > 1
                        ? <>Asset #{asset?.tag || "—"} — {n.count} new updates, latest: {n.latest.message}</>
                        : <>Asset #{asset?.tag || "—"} — {n.latest.message}</>}
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
  // Disposed assets are excluded from every dashboard figure below —
  // they're historical records, not part of the active fleet being measured.
  const assets = (scopedLocationId
    ? data.assets.filter((a) => a.locationId === scopedLocationId)
    : data.assets
  ).filter((a) => a.status !== "Disposed");

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
    data.assets.filter((a) => a.status !== "Disposed").forEach((a) => {
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

  // Asset Condition: a direct read of the Condition field, nothing else —
  // Status already has its own "Assets by Status" chart and warranty/
  // calibration already has its own overview panel, so mixing those into
  // this too just made "Critical" mean two different things at once.
  const conditionData = useMemo(() => {
    const counts = {};
    assets.forEach((a) => { counts[a.condition] = (counts[a.condition] || 0) + 1; });
    return CONDITION_OPTIONS.map((name) => ({ name, value: counts[name] || 0 })).filter((d) => d.value > 0);
  }, [assets]);

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

  // Purchase / Asset Value: total recorded purchase cost, broken down by
  // category. Assets with no purchase cost on file simply don't contribute
  // — this is a "what we've recorded" figure, not an estimate.
  const valueByCategory = useMemo(() => {
    const totals = {};
    assets.forEach((a) => {
      const cost = Number(a.purchaseCost) || 0;
      if (!cost) return;
      const cat = data.categories.find((c) => c.id === a.categoryId);
      const name = cat ? cat.name : "Uncategorized";
      totals[name] = (totals[name] || 0) + cost;
    });
    return Object.entries(totals).map(([name, value]) => ({ name, value })).sort((a, b) => b.value - a.value);
  }, [assets, data.categories]);
  const totalValue = valueByCategory.reduce((sum, c) => sum + c.value, 0);
  const maxCatValue = valueByCategory.length ? valueByCategory[0].value : 0;

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
        <DonutCard title="Asset Condition" data={conditionData} palette={CONDITION_COLORS} total={assets.length} />
        {isAdmin && (
          <DonutCard title="Assets by Location" data={locationData} palette={locationPalette} total={data.assets.length} />
        )}
      </div>

      <div className="bottom-row">
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
          <div className="panel-head"><h3>Asset Value</h3></div>
          <div className="value-total">
            <span className="value-total-amount">${totalValue.toLocaleString()}</span>
            <span className="value-total-label">Total recorded value{isAdmin ? ", all locations" : ""}</span>
          </div>
          {valueByCategory.length === 0 ? (
            <div className="empty-chart" style={{ padding: "0 18px 18px" }}>No purchase cost recorded yet.</div>
          ) : (
            <ul className="value-bar-list">
              {valueByCategory.slice(0, 5).map((c, i) => (
                <li key={c.name} className="value-bar-row">
                  <span className="value-bar-label" title={c.name}>{c.name}</span>
                  <div className="value-bar-track">
                    <div className="value-bar-fill" style={{ width: `${maxCatValue ? (c.value / maxCatValue) * 100 : 0}%`, background: categoryPalette[c.name] || CAT_PALETTE[i % CAT_PALETTE.length] }} />
                  </div>
                  <span className="value-bar-amount">${c.value.toLocaleString()}</span>
                </li>
              ))}
            </ul>
          )}
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
    notes: "", notesLog: [], transferHistory: [], department: "", disposalInfo: null,
  };
}

function AssetsView({ data, persist, isAdmin, isRegionalAdmin, canDeleteDirectly, scopedLocationId, showToast, currentUser, focusAssetId, onFocusHandled, applyLocalOnly }) {
  // Regional Staff can export their location's assets but never import —
  // only Admin and Regional Admin are trusted to write bulk data in.
  const canImport = isAdmin || isRegionalAdmin;
  const [search, setSearch] = useState("");
  const [locationFilter, setLocationFilter] = useState("all");
  const [categoryFilter, setCategoryFilter] = useState("all");
  const [userFilter, setUserFilter] = useState("all");
  const [presetFilter, setPresetFilter] = useState("all"); // one-tap shortcuts: under repair / needs attention / pending approval
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

  // NOTIFICATION SCOPE — Overall Admin and Regional Admin see a badge for
  // ANY significant update on an asset in their scope (comments, notes,
  // chat, status/assignment changes, disposal — see
  // computeAssetActivityFeed), not just comments addressed to them.
  // Regional Staff's notifications are unchanged: still just comments
  // that targeted them specifically.
  const broadActivityScope = isAdmin || isRegionalAdmin;

  const unreadCommentAssetIds = useMemo(() => {
    if (broadActivityScope) return new Set();
    const set = new Set();
    (data.comments || []).forEach((c) => {
      if ((c.targetUserIds || []).includes(currentUser.id) && !(c.readBy || []).includes(currentUser.id)) {
        set.add(c.assetId);
      }
    });
    return set;
  }, [data.comments, currentUser.id, broadActivityScope]);

  const assetActivityEvents = useMemo(
    () => (broadActivityScope ? computeAssetActivityFeed(data.auditLog, data.comments) : []),
    [data.auditLog, data.comments, broadActivityScope]
  );
  // Regional Admin's activity feed is limited to assets currently in their
  // own location — recomputed live off data.assets, same as everywhere
  // else location scoping happens in this app.
  const regionalActivityAssetIds = useMemo(() => {
    if (isAdmin || !broadActivityScope) return null;
    return new Set(data.assets.filter((a) => a.locationId === scopedLocationId).map((a) => a.id));
  }, [data.assets, isAdmin, broadActivityScope, scopedLocationId]);
  const unreadActivityAssetIds = useMemo(() => {
    if (!broadActivityScope) return unreadCommentAssetIds;
    const inScope = (assetId) => isAdmin || (regionalActivityAssetIds && regionalActivityAssetIds.has(assetId));
    const map = computeUnreadAssetActivity(assetActivityEvents, data.notificationReads, currentUser.id, inScope);
    return new Set(map.keys());
  }, [broadActivityScope, assetActivityEvents, data.notificationReads, currentUser.id, isAdmin, regionalActivityAssetIds, unreadCommentAssetIds]);

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

  // Same as above, but for the broader Overall Admin / Regional Admin
  // asset-activity notifications (see broadActivityScope) — opening the
  // asset counts as having seen its latest activity.
  useEffect(() => {
    if (!viewing || !broadActivityScope) return;
    if (!unreadActivityAssetIds.has(viewing.id)) return;
    persist(withAssetRead(data, currentUser.id, viewing.id));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [viewing?.id, broadActivityScope]);
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
  const [disposePromptAsset, setDisposePromptAsset] = useState(null); // asset awaiting disposal details before going Disposed
  const [disposedBy, setDisposedBy] = useState("");
  const [disposeReason, setDisposeReason] = useState("");
  const [disposeDate, setDisposeDate] = useState("");
  const [bulkDeleteText, setBulkDeleteText] = useState("");
  const [transferTarget, setTransferTarget] = useState(null); // asset id
  const [transferLocationId, setTransferLocationId] = useState("");
  const [transferNewLocationName, setTransferNewLocationName] = useState("");
  const [transferAssignedTo, setTransferAssignedTo] = useState("");
  const [transferReason, setTransferReason] = useState("");
  const [sort, setSort] = useState({ key: null, dir: "asc" });
  const fileInputRef = React.useRef(null);
  const [pendingImport, setPendingImport] = useState(null); // staged parsed rows awaiting the replace-confirmation below

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

  // Disposed assets are kept for records but shouldn't clutter the everyday
  // working view — "All assets" (and every other preset except the
  // dedicated one below) only ever looks at assets still in active use.
  const activeAssetsBase = useMemo(
    () => scopedAssetsBase.filter((a) => a.status !== "Disposed"),
    [scopedAssetsBase]
  );
  const disposedAssetsBase = useMemo(
    () => scopedAssetsBase.filter((a) => a.status === "Disposed"),
    [scopedAssetsBase]
  );

  // Quick-filter presets — one-tap shortcuts for the views managers check
  // most often, built from data already computed above (status + the same
  // alert map the table's warning icons use) rather than new tracking.
  const presetOptions = useMemo(() => {
    const inStockCount = activeAssetsBase.filter((a) => a.status === "In Stock").length;
    const inUseCount = activeAssetsBase.filter((a) => a.status === "In Use").length;
    const underRepairCount = activeAssetsBase.filter((a) => a.status === "Under Repair").length;
    const needsAttentionCount = activeAssetsBase.filter((a) => assetAlertMap.has(a.id)).length;
    const pendingApprovalCount = activeAssetsBase.filter((a) => !!a.pendingDeletion).length;
    const disposedCount = disposedAssetsBase.length;
    return [
      { id: "all", label: "All assets", color: null },
      { id: "inStock", label: "In Stock", count: inStockCount, color: STATUS_COLORS["In Stock"] },
      { id: "inUse", label: "In Use", count: inUseCount, color: STATUS_COLORS["In Use"] },
      { id: "underRepair", label: "Under repair", count: underRepairCount, color: STATUS_COLORS["Under Repair"] },
      { id: "needsAttention", label: "Needs attention", count: needsAttentionCount, color: "#EF4444" },
      ...(isAdmin || isRegionalAdmin ? [{ id: "pendingApproval", label: "Pending approval", count: pendingApprovalCount, color: "#8B5CF6" }] : []),
      { id: "disposed", label: "Disposed", count: disposedCount, color: STATUS_COLORS["Disposed"] },
    ].filter((p) => p.id === "all" || p.count > 0 || presetFilter === p.id);
  }, [activeAssetsBase, disposedAssetsBase, assetAlertMap, isAdmin, isRegionalAdmin, presetFilter]);

  // If the active preset's items all get resolved (e.g. every repair is
  // closed out) and the chip disappears, fall back to "all" instead of
  // silently showing an empty table.
  useEffect(() => {
    if (presetFilter !== "all" && !presetOptions.some((p) => p.id === presetFilter)) {
      setPresetFilter("all");
    }
  }, [presetOptions, presetFilter]);

  const visibleAssets = useMemo(() => {
    let list = presetFilter === "disposed" ? disposedAssetsBase : activeAssetsBase;
    if (presetFilter === "inStock") list = list.filter((a) => a.status === "In Stock");
    else if (presetFilter === "inUse") list = list.filter((a) => a.status === "In Use");
    else if (presetFilter === "underRepair") list = list.filter((a) => a.status === "Under Repair");
    else if (presetFilter === "needsAttention") list = list.filter((a) => assetAlertMap.has(a.id));
    else if (presetFilter === "pendingApproval") list = list.filter((a) => !!a.pendingDeletion);
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
  }, [activeAssetsBase, disposedAssetsBase, data.locations, data.categories, search, locationFilter, categoryFilter, userFilter, presetFilter, assetAlertMap]);

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
    const { repairReason, disposalBy, disposalReason, disposalDate, _lockLocationField, ...assetFields } = asset;
    const buildDisposalInfo = () => ({
      by: (disposalBy || currentUser.name).trim() || currentUser.name,
      reason: (disposalReason || "").trim(),
      date: disposalDate || todayISO(),
      at: new Date().toISOString(),
    });
    let next;
    let autoMaint = null;
    if (assetFields.id) {
      const prev = data.assets.find((a) => a.id === assetFields.id);
      const prevWasDisposed = prev?.status === "Disposed";
      let finalAsset = assetFields;
      if (assetFields.status === "Under Repair" && prev?.status !== "Under Repair") {
        finalAsset = { ...assetFields, preRepairStatus: prev?.status || "In Use" };
        autoMaint = { id: uid("maint"), assetId: assetFields.id, description: (repairReason || "").trim() || "Marked Under Repair from Assets", status: "Not Started", date: todayISO(), cost: "" };
      } else if (assetFields.status !== "Under Repair" && prev?.status === "Under Repair") {
        finalAsset = { ...assetFields, preRepairStatus: null };
      }
      if ((assetFields.status === "Disposed") && !prevWasDisposed) {
        finalAsset = { ...finalAsset, disposalInfo: buildDisposalInfo() };
      }
      finalAsset = applyStatusSideEffects(finalAsset);
      next = withLog({
        ...data,
        assets: data.assets.map((a) => (a.id === assetFields.id ? finalAsset : a)),
        maintenance: autoMaint ? [autoMaint, ...data.maintenance] : data.maintenance,
      }, currentUser, `Edited asset "${assetFields.name || assetFields.tag}"${autoMaint ? " — added maintenance entry (status: Under Repair)" : ""}`, finalAsset.locationId, finalAsset.id);
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
      if (newAsset.status === "Disposed") {
        newAsset = { ...newAsset, disposalInfo: buildDisposalInfo() };
      }
      newAsset = applyStatusSideEffects(newAsset);
      next = withLog({
        ...data,
        assets: [newAsset, ...data.assets],
        maintenance: autoMaint ? [autoMaint, ...data.maintenance] : data.maintenance,
      }, currentUser, `Added asset "${newAsset.name || newAsset.tag}"`, newAsset.locationId, newAsset.id);
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
  // (see submitInUseAssignment below). Disposed is the same idea —
  // it pops up a form for the disposal details; the status only actually
  // changes once that form is saved (see submitDisposal below).
  const quickStatusChange = async (asset, newStatus) => {
    if (newStatus === asset.status) return;
    const wasDisposedAsset = asset.status === "Disposed";
    if (wasDisposedAsset) {
      const ok = window.confirm(
        `This asset is currently marked "${asset.status}". Changing its status to "${newStatus}" will restore it to active use and clear its disposal record. Continue?`
      );
      if (!ok) return;
    }
    if (newStatus === "Under Repair") {
      setMaintPromptAsset(asset);
      return;
    }
    if (newStatus === "In Use") {
      setInUseDepartment(asset.department || "");
      setInUseAssignedTo(asset.assignedTo || "");
      // Regional Admin/Staff can only ever act on their own location, so
      // default (and later lock) it to that — only the Overall Admin picks
      // freely, defaulting to wherever the asset currently sits.
      setInUseLocationId(scopedLocationId || asset.locationId || "");
      setInUsePromptAsset(asset);
      return;
    }
    if (newStatus === "Disposed") {
      setDisposedBy(currentUser.name || "");
      setDisposeReason("");
      setDisposeDate(todayISO());
      setDisposePromptAsset(asset);
      return;
    }
    let finalAsset = { ...asset, status: newStatus };
    if (asset.status === "Under Repair") finalAsset.preRepairStatus = null;
    finalAsset = applyStatusSideEffects(finalAsset);
    const next = withLog({
      ...data,
      assets: data.assets.map((a) => (a.id === asset.id ? finalAsset : a)),
    }, currentUser, `Changed status of asset "${asset.name || asset.tag}" to "${newStatus}"`, finalAsset.locationId, finalAsset.id);
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
    }, currentUser, `Changed status of asset "${asset.name || asset.tag}" to "In Use"`, finalAsset.locationId, finalAsset.id);
    persist(next);
    setInUsePromptAsset(null);
    showToast("Status updated.");
  };

  // Saves the disposal details collected from the Disposed popup, and only
  // now actually flips the asset's status — cancelling that popup leaves
  // the asset untouched.
  const submitDisposal = () => {
    const asset = disposePromptAsset;
    if (!asset) return;
    if (!disposeReason.trim()) { alert("Please select or enter a reason."); return; }
    let finalAsset = {
      ...asset,
      status: "Disposed",
      disposalInfo: {
        by: disposedBy.trim() || currentUser.name,
        reason: disposeReason.trim(),
        date: disposeDate || todayISO(),
        at: new Date().toISOString(),
      },
    };
    if (asset.status === "Under Repair") finalAsset.preRepairStatus = null;
    finalAsset = applyStatusSideEffects(finalAsset);
    const next = withLog({
      ...data,
      assets: data.assets.map((a) => (a.id === asset.id ? finalAsset : a)),
    }, currentUser, `Changed status of asset "${asset.name || asset.tag}" to "Disposed" — reason: ${disposeReason.trim()}`, finalAsset.locationId, finalAsset.id);
    persist(next);
    setDisposePromptAsset(null);
    showToast("Status updated.");
  };

  // Saves the maintenance entry collected from the Under Repair popup, and
  // only now actually flips the asset's status — cancelling that popup
  // leaves the asset untouched.
  const submitRepairMaintenance = (entry) => {
    const asset = maintPromptAsset;
    if (!asset) return;
    const newEntry = { ...entry, id: uid("maint"), assetId: asset.id };
    const finalAsset = { ...asset, status: "Under Repair", preRepairStatus: asset.status, disposalInfo: null };
    const next = withLog({
      ...data,
      assets: data.assets.map((a) => (a.id === asset.id ? finalAsset : a)),
      maintenance: [newEntry, ...data.maintenance],
    }, currentUser, `Changed status of asset "${asset.name || asset.tag}" to "Under Repair" — added maintenance entry`, finalAsset.locationId, finalAsset.id);
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
  // Singapore) plus every Overall Admin, and anyone who has already
  // commented in this asset's thread AND still has access to the asset's
  // current location (so a reply reaches whoever it's replying to) —
  // since those are the only people with app access who should hear
  // about it. The "Assigned To" field is just free text for an employee
  // with no app account, so it's never notified.
  //
  // Overall Admins are always included, regardless of location, per the
  // asset-notification scope: Admin sees activity everywhere. Deliberately
  // does NOT unconditionally include the asset's original creator or
  // every past commenter: once an asset is transferred to a different
  // location, whoever isn't an admin or currently assigned to the new
  // location should stop hearing about it.
  const addComment = (assetId, message) => {
    const text = (message || "").trim();
    if (!text) return;
    const asset = data.assets.find((a) => a.id === assetId);
    const hasCurrentAccess = (userId) => {
      const u = data.users.find((x) => x.id === userId);
      return !!u && (u.role === "Admin" || (u.locationId && u.locationId === asset?.locationId));
    };
    const adminUserIds = data.users.filter((u) => u.role === "Admin").map((u) => u.id);
    const locationUserIds = data.users
      .filter((u) => u.locationId && u.locationId === asset?.locationId)
      .map((u) => u.id);
    const priorParticipantIds = (data.comments || [])
      .filter((c) => c.assetId === assetId)
      .map((c) => c.authorId)
      .filter(hasCurrentAccess);
    const createdById = hasCurrentAccess(asset?.createdById) ? asset?.createdById : null;
    const targetUserIds = Array.from(new Set(
      [createdById, ...adminUserIds, ...locationUserIds, ...priorParticipantIds].filter((id) => id && id !== currentUser.id)
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
    // Deliberately not tagged with assetId here — the comment row above
    // already feeds the asset-activity notification system (see
    // computeAssetActivityFeed), so tagging this entry too would double
    // up the same comment as two separate activity events.
    const next = withLog({
      ...data,
      comments: [comment, ...(data.comments || [])],
    }, currentUser, `Commented on asset "${asset?.name || asset?.tag}"`, asset?.locationId);
    persist(next);
    showToast("Comment sent.");
  };

  // Adds or edits an entry in an asset's Notes — a lightweight,
  // timestamped operational log distinct from the Comments thread. This
  // is a same-location field edit like any other asset update, so it
  // goes through the normal save path rather than a special RPC.
  const saveAssetNote = (assetId, noteId, text) => {
    const trimmed = (text || "").trim();
    if (!trimmed) return;
    const asset = data.assets.find((a) => a.id === assetId);
    if (!asset) return;
    const now = new Date().toISOString();
    const existing = asset.notesLog || [];
    const notesLog = noteId
      ? existing.map((n) => (n.id === noteId ? { ...n, text: trimmed, editedAt: now } : n))
      : [{ id: uid("note"), text: trimmed, authorId: currentUser.id, authorName: currentUser.name, at: now, editedAt: null }, ...existing];
    const next = withLog({
      ...data,
      assets: data.assets.map((a) => (a.id === assetId ? { ...a, notesLog } : a)),
    }, currentUser, `${noteId ? "Edited a note" : "Added a note"} on asset "${asset.name || asset.tag}"`, asset.locationId, asset.id);
    persist(next);
    showToast(noteId ? "Note updated." : "Note added.");
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
      notesLog: [],
      disposalInfo: null,
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

  const submitTransfer = async () => {
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

    // A transfer moves the asset OUT of the mover's own location — for a
    // non-admin, that means the row won't satisfy the "must stay in your
    // own location" rule the moment it's saved, even though the move
    // itself is exactly what's supposed to be allowed. Rather than fight
    // that at the RLS-policy level, non-admin transfers go through a
    // dedicated database function (transfer_asset — see
    // critical-security-migration.sql) that does its own permission
    // check in plain code and applies the change directly, sidestepping
    // the conflict entirely. Admins don't hit this at all — an Admin's
    // regular save already works, so their path is unchanged below.
    if (!isAdmin) {
      const { error } = await supabase.rpc("transfer_asset", {
        p_asset_id: asset.id,
        p_new_location_id: destLocationId,
        p_new_assigned_to: newAssignedTo,
        p_reason: transferReason.trim(),
        p_by_name: currentUser.name,
      });
      if (error) {
        showToast("Could not transfer this asset — check your connection and try again.");
        return;
      }
      // The asset just moved outside this user's own location, so their
      // own view can no longer include it — reflect that immediately
      // rather than waiting on the next realtime refresh.
      applyLocalOnly({ ...data, assets: data.assets.filter((a) => a.id !== asset.id) });
      setTransferTarget(null);
      setTransferReason("");
      setTransferNewLocationName("");
      showToast("Asset transferred.");
      return;
    }

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
    const assignedChanged = newAssignedTo !== (asset.assignedTo || "");
    const logBits = [`Transferred asset "${asset.name || asset.tag}" from ${fromLoc} to ${toLoc}`];
    if (assignedChanged) logBits.push(`reassigned to ${newAssignedTo || "Unassigned"}`);
    logBits.push(`reason: ${transferReason.trim()}`);

    const next = withLog({
      ...data,
      locations,
      assets: data.assets.map((a) => (a.id === transferTarget
        ? { ...a, locationId: destLocationId, assignedTo: newAssignedTo, transferHistory: [transferEntry, ...(a.transferHistory || [])] }
        : a)),
    }, currentUser, logBits.join(" — "), destLocationId, asset.id);
    persist(next);
    setTransferTarget(null);
    setTransferReason("");
    setTransferNewLocationName("");
    showToast("Asset transferred.");
  };

  // Regional Staff can't delete outright — they submit a reason, and the
  // asset is flagged pending. Who that flag is actually waiting on is
  // resolved dynamically (see locationApprover in the parent component):
  // the Regional Admin assigned to the asset's location if one exists,
  // otherwise the Overall Admin as a fallback. We don't bake a fixed
  // approver into the request itself, so if a Regional Admin is added or
  // removed later, an already-pending request re-routes automatically.
  const submitDeleteRequest = async () => {
    if (!deleteReason.trim()) { alert("Please enter a reason for this deletion request."); return; }
    const asset = data.assets.find((a) => a.id === requestDeleteTarget);
    const approverName = data.users.find((u) => u.role === "Regional Admin" && u.locationId === asset?.locationId)?.name;
    const next = withLog({
      ...data,
      assets: data.assets.map((a) => (a.id === requestDeleteTarget
        ? { ...a, pendingDeletion: { requestedBy: currentUser.id, requestedByName: currentUser.name, reason: deleteReason.trim(), requestedAt: new Date().toISOString() } }
        : a)),
    }, currentUser, `Requested deletion of asset "${asset?.name || asset?.tag}" — reason: ${deleteReason.trim()}`, asset?.locationId, asset?.id);
    persist(next);
    setRequestDeleteTarget(null);
    setDeleteReason("");
    showToast(approverName ? `Deletion request sent to ${approverName} for approval.` : "Deletion request sent for Admin approval.");
  };

  // Overall Admin and Regional Admin can delete an asset in their own
  // location immediately, at their own discretion — no approval needed.
  // Regional Staff always has to submit a request instead (see above).
  // This intentionally only touches the single-asset delete path — Bulk
  // Delete (further below) stays Admin-only and unchanged.
  const startDelete = (asset) => {
    if (canDeleteDirectly) {
      setConfirmDelete(asset.id);
    } else if (!asset.pendingDeletion) {
      setRequestDeleteTarget(asset.id);
    }
  };

  const EXPORT_COLS = ["tag", "name", "department", "assetType", "brand", "model", "yearModel", "serial", "status", "condition", "location", "assignedTo", "purchaseDate", "purchaseCost", "warrantyExpiry", "requiresCalibration", "calibrationDate", "nextCalibrationDate", "notes"];

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
        notes: a.notes || "",
      };
    });
    const wb = XLSX.utils.book_new();
    const sheet = XLSX.utils.json_to_sheet(rows, { header: EXPORT_COLS });
    XLSX.utils.book_append_sheet(wb, sheet, "Assets");

    // A second sheet with every comment on the exported assets, keyed by
    // the human-readable Asset Tag rather than an internal id — this
    // export is meant to be read directly, not restored from.
    const exportedAssetIds = new Set(visibleAssets.map((a) => a.id));
    const commentRows = (data.comments || [])
      .filter((c) => exportedAssetIds.has(c.assetId))
      .map((c) => {
        const asset = visibleAssets.find((a) => a.id === c.assetId);
        return {
          assetTag: asset?.tag || "", assetName: asset?.name || "",
          author: c.authorName || "", message: c.message || "",
          at: c.at ? new Date(c.at).toLocaleString() : "",
        };
      })
      .sort((a, b) => a.assetTag.localeCompare(b.assetTag) || new Date(a.at) - new Date(b.at));
    const commentSheet = XLSX.utils.json_to_sheet(commentRows, { header: ["assetTag", "assetName", "author", "message", "at"] });
    XLSX.utils.book_append_sheet(wb, commentSheet, "Comments");

    downloadWorkbook(wb, "assets-export.xlsx");
    showToast("Assets exported.");
  };

  const triggerImport = () => fileInputRef.current?.click();

  // Regional Staff never reach this — the Import button and file input are
  // only rendered for canImport (Admin or Regional Admin) below. Regional
  // Admin's imported rows are always forced into their own assigned
  // location (see locationId below), regardless of what the sheet's
  // "location" column says, so they can never write into another country's
  // data. Parsing only stages the result — nothing is saved until the
  // confirmation dialog (which spells out exactly what will be replaced)
  // is accepted.
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
          // Regional Admin (scoped) can only ever import into their own
          // location — the sheet's location column is ignored for them.
          // Admin (unscoped) keeps whatever location the sheet specifies,
          // falling back to the first location if it doesn't match one.
          locationId: !isAdmin && scopedLocationId ? scopedLocationId : (loc?.id || data.locations[0]?.id || ""),
          assignedTo: row.assignedTo || "",
          purchaseDate: row.purchaseDate || todayISO(),
          purchaseCost: Number(row.purchaseCost) || 0,
          warrantyExpiry: (row.warrantyExpiry && row.warrantyExpiry !== "N/A") ? row.warrantyExpiry : "",
          requiresCalibration: String(row.requiresCalibration || "").toLowerCase() === "yes",
          calibrationDate: row.calibrationDate || "",
          nextCalibrationDate: row.nextCalibrationDate || "",
          notes: row.notes || "",
          transferHistory: [],
          disposalInfo: null,
        };
      });
      // Every location touched by this import will have its existing
      // assets replaced by the new rows — locations the file doesn't
      // mention are left completely untouched.
      const affectedLocationIds = new Set(newAssets.map((a) => a.locationId).filter(Boolean));
      const existingInScope = data.assets.filter((a) => affectedLocationIds.has(a.locationId));
      const affectedLocationNames = data.locations
        .filter((l) => affectedLocationIds.has(l.id))
        .map((l) => l.name);
      setPendingImport({
        newAssets,
        affectedLocationIds,
        affectedLocationNames,
        existingCount: existingInScope.length,
      });
    } catch {
      showToast("Could not parse this Excel file.");
    }
  };

  const confirmImport = () => {
    if (!pendingImport) return;
    const { newAssets, affectedLocationIds, affectedLocationNames } = pendingImport;
    const removedAssetIds = new Set(data.assets.filter((a) => affectedLocationIds.has(a.locationId)).map((a) => a.id));
    const keptAssets = data.assets.filter((a) => !affectedLocationIds.has(a.locationId));
    const keptMaintenance = data.maintenance.filter((m) => !removedAssetIds.has(m.assetId));
    const scopeLabel = affectedLocationNames.length ? affectedLocationNames.join(", ") : "the selected location(s)";
    persist(withLog(
      { ...data, assets: [...newAssets, ...keptAssets], maintenance: keptMaintenance },
      currentUser,
      `Imported ${newAssets.length} asset(s) via Excel, replacing existing asset data for ${scopeLabel}`,
      !isAdmin ? scopedLocationId : null
    ));
    showToast(`Imported ${newAssets.length} asset(s).`);
    setPendingImport(null);
  };

  return (
    <div>
      <div className="preset-filters">
        {presetOptions.map((p) => (
          <button
            key={p.id}
            type="button"
            className={`preset-chip${!p.color ? " preset-chip-all" : ""}${presetFilter === p.id ? " active" : ""}`}
            style={p.color ? {
              background: presetFilter === p.id ? p.color : `${p.color}1a`,
              color: presetFilter === p.id ? "#fff" : p.color,
              borderColor: presetFilter === p.id ? p.color : `${p.color}40`,
            } : undefined}
            onClick={() => setPresetFilter(p.id)}
          >
            {p.label}{typeof p.count === "number" && ` · ${p.count}`}
          </button>
        ))}
      </div>
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
          {canImport && (
            <>
              <input ref={fileInputRef} type="file" accept=".xlsx" style={{ display: "none" }} onChange={handleImportFile} />
              <button className="btn ghost" onClick={triggerImport}><Upload size={14} /> Import Excel</button>
            </>
          )}
          <button className="btn ghost" onClick={exportExcel}><Download size={14} /> Export Excel</button>
          <button className="btn primary new-asset-btn" onClick={() => setEditing(newAssetDraft())}>
            <Plus size={14} /> New Asset
          </button>
        </div>
      </div>

      <button className="fab-add" onClick={() => setEditing(newAssetDraft())} title="New Asset">
        <Plus size={22} />
      </button>

      {(assetAlertMap.size > 0 || unreadActivityAssetIds.size > 0) && (
        <div className="table-legend">
          {unreadActivityAssetIds.size > 0 && (
            <span className="table-legend-item">
              <MessageCircle size={13} className="comment-flag" /> {broadActivityScope ? "New activity" : "New comment"}
            </span>
          )}
          {assetAlertMap.size > 0 && (
            <>
              <span className="table-legend-item"><Clock size={13} className="alert-flag" /> Warranty/calibration due soon</span>
              <span className="table-legend-item"><AlertTriangle size={13} className="alert-flag alert-flag-urgent" /> Expired/overdue</span>
            </>
          )}
        </div>
      )}

      {/* Built straight off STATUS_COLORS/CONDITION_COLORS, so adding a new
          status or condition anywhere in the app automatically shows up
          here too — nothing to remember to update by hand. */}
      <div className="table-legend status-legend">
        <span className="table-legend-label">Status:</span>
        {STATUS_OPTIONS.map((s) => (
          <span key={s} className="table-legend-item">
            <span className="legend-dot" style={{ background: STATUS_COLORS[s] }} /> {s}
          </span>
        ))}
        <span className="table-legend-label">Condition:</span>
        {CONDITION_OPTIONS.map((c) => (
          <span key={c} className="table-legend-item">
            <span className="legend-dot" style={{ background: CONDITION_COLORS[c] }} /> {c}
          </span>
        ))}
      </div>

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
                      {unreadActivityAssetIds.has(a.id) && (
                        <MessageCircle size={13} className="comment-flag" title={broadActivityScope ? "New activity on this asset" : "New comment on this asset"} />
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
                    <td data-label="Condition"><Badge color={CONDITION_COLORS[a.condition] || "#6B7280"}>{a.condition}</Badge></td>
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
          currentUser={currentUser}
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
                disabled={!isAdmin}
              />
              {!isAdmin && (
                <span className="field-hint">Locked to your assigned location — only the Overall Admin can change this.</span>
              )}
            </Field>
          </div>
          <div className="form-full modal-actions">
            <button type="button" className="btn ghost" onClick={() => setInUsePromptAsset(null)}>Cancel</button>
            <button type="button" className="btn primary" onClick={submitInUseAssignment}>Save</button>
          </div>
        </Modal>
      )}
      {disposePromptAsset && (
        <Modal title={`Mark Disposed — ${disposePromptAsset.tag}`} onClose={() => setDisposePromptAsset(null)} width={480}>
          <div className="form-full hint-box" style={{ marginBottom: 14 }}>
            This asset will be marked Disposed once you save these details.
          </div>
          <div className="form-grid">
            <Field label="Disposed By" required>
              <input value={disposedBy} onChange={(e) => setDisposedBy(e.target.value)} placeholder="Name of person handling disposal" />
            </Field>
            <Field label="Reason" required>
              <SearchableSelect
                value={disposeReason}
                onChange={setDisposeReason}
                options={DISPOSAL_REASON_OPTIONS.map((v) => ({ value: v, label: v }))}
                placeholder="Search or type a reason"
              />
            </Field>
            <Field label="Date of Disposal" required>
              <input type="date" value={disposeDate} onChange={(e) => setDisposeDate(e.target.value)} />
            </Field>
          </div>
          <div className="form-full modal-actions">
            <button type="button" className="btn ghost" onClick={() => setDisposePromptAsset(null)}>Cancel</button>
            <button type="button" className="btn primary" onClick={submitDisposal}>Save</button>
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
      {pendingImport && (
        <ConfirmDialog
          message={`This will replace the existing asset data for ${
            pendingImport.affectedLocationNames.length ? pendingImport.affectedLocationNames.join(", ") : "the imported location(s)"
          }: ${pendingImport.existingCount} existing asset(s) there will be removed and replaced with the ${pendingImport.newAssets.length} row(s) from this file. Assets in other locations are not affected. This cannot be undone. Continue?`}
          confirmLabel="Replace & Import"
          onCancel={() => setPendingImport(null)}
          onConfirm={confirmImport}
          maxWidth={480}
        />
      )}
      {requestDeleteTarget && (
        <Modal title="Request Deletion" onClose={() => { setRequestDeleteTarget(null); setDeleteReason(""); }} width={760}>
          <div className="form-grid">
            <div className="form-full hint-box">
              This asset won't be deleted right away — your request and reason go to{" "}
              {(() => {
                const asset = data.assets.find((a) => a.id === requestDeleteTarget);
                const approver = data.users.find((u) => u.role === "Regional Admin" && u.locationId === asset?.locationId);
                return approver ? `${approver.name} (Regional Admin)` : "the Overall Admin";
              })()} for approval first.
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
          asset={data.assets.find((a) => a.id === viewing.id) || viewing}
          categories={data.categories}
          locations={data.locations}
          isAdmin={isAdmin}
          canDeleteDirectly={canDeleteDirectly}
          comments={(data.comments || []).filter((c) => c.assetId === viewing.id)}
          maintenance={data.maintenance.filter((m) => m.assetId === viewing.id)}
          currentUser={currentUser}
          onAddComment={(message) => addComment(viewing.id, message)}
          onSaveNote={(noteId, text) => saveAssetNote(viewing.id, noteId, text)}
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

function AssetModal({ asset, categories, locations, isAdmin, scopedLocationId, existingAssets, departmentOptions, currentUser, onClose, onSave }) {
  const [form, setForm] = useState(asset);
  const [hasPurchaseInfo, setHasPurchaseInfo] = useState(!!(asset.purchaseDate || asset.purchaseCost));
  const [hasWarrantyExpiry, setHasWarrantyExpiry] = useState(!!(asset.warrantyExpiry && asset.warrantyExpiry !== "N/A"));
  // ^ still checks for legacy "N/A" values from before this fix, so old assets open with the toggle correctly off
  // "New Asset" starts with the tag auto-generated (blank tag → server
  // assigns the next ASTUTE### on save). Editing an asset always shows its
  // real tag — auto-generate only applies at creation time.
  const [autoTag, setAutoTag] = useState(!asset.id && !asset.tag);
  const set = (k, v) => setForm((f) => ({ ...f, [k]: v }));

  // A brand-new asset walks through the steps one at a time — that's the
  // whole point of the wizard, so someone isn't handed all four sections'
  // worth of fields at once. Editing an existing asset already has every
  // field filled in, so there's nothing to "walk through" — every step tab
  // is unlocked immediately for quick jumping around.
  const isNew = !asset.id;
  const WIZARD_STEPS = [
    { key: "info", label: "Asset Information", icon: Package },
    { key: "device", label: "Device Details", icon: Tags },
    { key: "assignment", label: "Assignment", icon: User },
    { key: "optional", label: "Optional Information", icon: Info },
  ];
  const [step, setStep] = useState(0);
  const [maxStepReached, setMaxStepReached] = useState(isNew ? 0 : WIZARD_STEPS.length - 1);

  const wasUnderRepair = asset.status === "Under Repair";
  const needsReason = form.status === "Under Repair" && !wasUnderRepair;

  const wasDisposed = asset.status === "Disposed";
  const needsDisposalReason = (form.status === "Disposed") && !wasDisposed;

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

  // Checks only the fields that belong to a given step — this is what lets
  // "Next" catch a missing asset name before letting someone wander off to
  // Device Details, without also demanding they've filled in Assignment.
  const stepError = (i) => {
    if (i === 0) {
      if (!form.name.trim()) return "Please enter an asset name.";
      if (!form.categoryId) return "Please select a category.";
    }
    if (i === 1) {
      if (!form.status) return "Please select a status.";
      if (needsReason && !(form.repairReason || "").trim()) return "Please enter a reason — this creates the matching Maintenance entry.";
      if (needsDisposalReason && !(form.disposalReason || "").trim()) return `Please select or enter a reason for marking this asset "${form.status}".`;
    }
    return null;
  };

  const goToStep = (i) => {
    // New assets can only jump to a step they've already reached via Next —
    // no skipping ahead to Assignment with a blank name. Editing an
    // existing asset has no such restriction (maxStepReached starts maxed).
    if (i > maxStepReached) return;
    setStep(i);
  };

  const goNext = () => {
    const err = stepError(step);
    if (err) { alert(err); return; }
    const next = Math.min(step + 1, WIZARD_STEPS.length - 1);
    setStep(next);
    setMaxStepReached((m) => Math.max(m, next));
  };

  const goBack = () => setStep((s) => Math.max(s - 1, 0));

  const submit = (e) => {
    if (e && e.preventDefault) e.preventDefault();
    // Enter-key inside any step's inputs fires the form's onSubmit — on any
    // step but the last, that should just behave like clicking "Next"
    // rather than trying to save early.
    if (step !== WIZARD_STEPS.length - 1) {
      goNext();
      return;
    }
    for (let i = 0; i < WIZARD_STEPS.length - 1; i++) {
      const err = stepError(i);
      if (err) { alert(err); setStep(i); setMaxStepReached((m) => Math.max(m, i)); return; }
    }
    const tag = String(form.tag || "").trim();
    if (tag) {
      const dup = (existingAssets || []).find(
        (a) => a.id !== form.id && String(a.tag || "").trim().toLowerCase() === tag.toLowerCase()
      );
      if (dup) {
        alert(`Asset Tag "${tag}" is already used by "${dup.name || "another asset"}". Please use a unique tag.`);
        setStep(0);
        return;
      }
    }
    if (wasDisposed && form.status !== "Disposed") {
      const ok = window.confirm(
        `This asset is currently marked "${asset.status}". Changing its status to "${form.status}" will restore it to active use and clear its disposal record. Continue?`
      );
      if (!ok) return;
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
        <div className="wizard-steps">
          {WIZARD_STEPS.map((s, i) => {
            const Icon = s.icon;
            const locked = i > maxStepReached;
            return (
              <React.Fragment key={s.key}>
                {i > 0 && <span className={`wizard-step-connector ${i <= maxStepReached ? "done" : ""}`} />}
                <button
                  type="button"
                  className={`wizard-step ${step === i ? "active" : ""} ${i < maxStepReached ? "done" : ""} ${locked ? "locked" : ""}`}
                  onClick={() => goToStep(i)}
                  disabled={locked}
                  title={locked ? "Complete the earlier steps first" : s.label}
                >
                  <span className="wizard-step-icon">{i < maxStepReached ? <Check size={13} /> : <Icon size={13} />}</span>
                  <span className="wizard-step-label">{s.label}</span>
                </button>
              </React.Fragment>
            );
          })}
        </div>

        {step === 0 && (
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
        )}

        {step === 1 && (
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

          {needsDisposalReason && (
            <div className="form-full repair-reason-box">
              <div className="form-grid">
                <Field label="Disposed By">
                  <input
                    value={form.disposalBy ?? (currentUser?.name || "")}
                    onChange={(e) => set("disposalBy", e.target.value)}
                    placeholder="Name of person handling disposal"
                  />
                </Field>
                <Field label="Reason">
                  <SearchableSelect
                    value={form.disposalReason || ""}
                    onChange={(v) => set("disposalReason", v)}
                    options={DISPOSAL_REASON_OPTIONS.map((v) => ({ value: v, label: v }))}
                    placeholder="Search or type a reason"
                  />
                </Field>
                <Field label="Date of Disposal">
                  <input
                    type="date"
                    value={form.disposalDate ?? todayISO()}
                    onChange={(e) => set("disposalDate", e.target.value)}
                  />
                </Field>
              </div>
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
        )}

        {step === 2 && (
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
        )}

        {step === 3 && (
        <FormSection icon={Info} title="Optional Information">
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
        </FormSection>
        )}

        <div className="modal-footer-sticky">
          <span className="wizard-step-count">Step {step + 1} of {WIZARD_STEPS.length}</span>
          <div style={{ flex: 1 }} />
          <button type="button" className="btn ghost" onClick={onClose}>Cancel</button>
          {step > 0 && <button type="button" className="btn ghost" onClick={goBack}>Back</button>}
          {step < WIZARD_STEPS.length - 1
            ? <button key="next-btn" type="button" className="btn primary" onClick={goNext}>Next</button>
            : <button key="submit-btn" type="submit" className="btn primary">{asset.id ? "Save Changes" : "Create Asset"}</button>}
        </div>
      </form>
    </Modal>
  );
}

/* ---------------------------------------------------------
   Asset Detail Modal (read-only view, opened by clicking the tag)
--------------------------------------------------------- */
function AssetDetailModal({ asset, categories, locations, isAdmin, canDeleteDirectly, comments, maintenance, currentUser, onAddComment, onSaveNote, onClose, onDelete }) {
  const cat = categories.find((c) => c.id === asset.categoryId);
  const loc = locations.find((l) => l.id === asset.locationId);
  const [commentText, setCommentText] = useState("");
  const [noteText, setNoteText] = useState("");
  const [editingNoteId, setEditingNoteId] = useState(null);
  const [editingNoteText, setEditingNoteText] = useState("");
  const issues = useMemo(() => getAssetIssues(asset), [asset]);
  const notesLog = [...(asset.notesLog || [])].sort((a, b) => new Date(b.at) - new Date(a.at));
  const canEditNote = (note) => isAdmin || note.authorId === currentUser.id;
  const submitNote = () => {
    if (!noteText.trim()) return;
    onSaveNote(null, noteText);
    setNoteText("");
  };
  const startEditNote = (note) => { setEditingNoteId(note.id); setEditingNoteText(note.text); };
  const saveEditNote = () => {
    if (!editingNoteText.trim()) return;
    onSaveNote(editingNoteId, editingNoteText);
    setEditingNoteId(null);
    setEditingNoteText("");
  };
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
  // Compact group — a slim uppercase caption plus a tight 2-col grid of
  // label/value rows. Several of these stack inside one continuous card
  // (separated by a hairline, not a full boxed section each) so the
  // overview reads as one dense spec sheet instead of four separate cards.
  const detailGroup = (Icon, title, content) => (
    <div className="detail-group">
      <div className="detail-group-label"><Icon size={12} /> {title}</div>
      <div className="detail-grid">{content}</div>
    </div>
  );
  // Right-rail sections (Notes / History / Comments) share the same slim
  // caption style, kept visually distinct from the overview via the pane
  // divider rather than their own card chrome.
  const sideSection = (title, content) => (
    <div className="detail-side-section">
      <div className="notif-section-title">{title}</div>
      {content}
    </div>
  );

  return (
    <Modal title={`Asset Details — ${asset.tag}`} onClose={onClose} width={1040}>
      <div className="detail-hero">
        <Badge color={STATUS_COLORS[asset.status] || "#6B7280"}>{asset.status}</Badge>
        {cat && (
          <span className="detail-hero-chip">
            <span className="cat-dot" style={{ background: categoryColor(categories, asset.categoryId) }} />
            {cat.name}
          </span>
        )}
        <span className="detail-hero-chip"><MapPin size={12} /> {loc?.name || "No location"}</span>
        {asset.assignedTo && <span className="detail-hero-chip"><User size={12} /> {asset.assignedTo}</span>}
      </div>

      {issues.length > 0 && (
        <div className="attention-banner">
          <div className="attention-banner-title">
            <AlertTriangle size={15} />
            {issues.length === 1 ? issues[0].label : `${issues.length} Issues Require Attention`}
          </div>
          <div className="attention-banner-list">
            {issues.map((issue, i) => (
              <div key={i} className={`attention-item attention-${issue.severity}`}>
                <span className="attention-dot" />
                <div>
                  {issues.length > 1 && <div className="attention-item-label">{issue.label}</div>}
                  <div className="attention-item-detail">{issue.detail}</div>
                </div>
              </div>
            ))}
          </div>
        </div>
      )}

      <div className="detail-layout">
        <div className="detail-main">
          {detailGroup(Package, "Asset Information", <>
            {row("Name", asset.name)}
            {row("Department", asset.department)}
            {row("Category", cat?.name)}
            {row("Asset Type", asset.assetType)}
          </>)}

          {detailGroup(Tags, "Device Details", <>
            {row("Brand / Model", [asset.brand, asset.model].filter(Boolean).join(" / "))}
            {row("Year Model", asset.yearModel)}
            {row("Serial Number", asset.serial)}
            {row("Condition", asset.condition)}
            {asset.assetType === "Non-IT" && row("Requires Calibration?", asset.requiresCalibration ? "Yes" : "No")}
            {asset.assetType === "Non-IT" && asset.requiresCalibration && row("Calibration Date", asset.calibrationDate)}
            {asset.assetType === "Non-IT" && asset.requiresCalibration && row("Next Recalibration", asset.nextCalibrationDate)}
          </>)}

          {detailGroup(User, "Assignment", <>
            {row("Location", loc?.name)}
            {row("Assigned To", asset.assignedTo)}
            {row("Added By", asset.createdByName)}
          </>)}

          {detailGroup(Info, "Purchase & Warranty", <>
            {row("Purchase Date", asset.purchaseDate)}
            {row("Purchase Cost", asset.purchaseCost ? `$${asset.purchaseCost}` : "")}
            {asset.assetType === "IT" && row("Warranty Expiry", asset.warrantyExpiry)}
          </>)}

          {asset.transferHistory && asset.transferHistory.length > 0 && (
            <div className="detail-group">
              <div className="detail-group-label"><Truck size={12} /> Transfer History</div>
              {asset.transferHistory.map((t) => (
                <div key={t.id} className="notif-item">
                  {t.fromLocationName} → {t.toLocationName} — {t.reason} ({t.by}, {new Date(t.at).toLocaleString([], { dateStyle: "medium", timeStyle: "short" })})
                </div>
              ))}
            </div>
          )}

          {asset.disposalInfo && (
            <div className="detail-group">
              <div className="detail-group-label"><Archive size={12} /> Disposal Information</div>
              <div className="detail-grid">
                {row("Disposed By", asset.disposalInfo.by)}
                {row("Reason", asset.disposalInfo.reason)}
                {row("Date of Disposal", asset.disposalInfo.date)}
              </div>
            </div>
          )}
        </div>

        <div className="detail-side">
          {sideSection("Notes", (
            <div className="comments-section">
              <div className="notes-log-list">
                {notesLog.length === 0 && (
                  <div className="notif-empty">No notes yet — use this to log anything worth remembering about this asset.</div>
                )}
                {notesLog.map((n) => (
                  <div key={n.id} className="note-item">
                    {editingNoteId === n.id ? (
                      <>
                        <textarea
                          value={editingNoteText}
                          onChange={(e) => setEditingNoteText(e.target.value)}
                          rows={2}
                          autoFocus
                        />
                        <div className="note-edit-actions">
                          <button type="button" className="btn ghost" onClick={() => setEditingNoteId(null)}>Cancel</button>
                          <button type="button" className="btn primary" onClick={saveEditNote} disabled={!editingNoteText.trim()}>Save</button>
                        </div>
                      </>
                    ) : (
                      <>
                        <div className="note-text">{n.text}</div>
                        <div className="note-meta">
                          <span>{n.authorName}</span>
                          <span>·</span>
                          <span>{formatLogTime(n.at)}{n.editedAt ? " (edited)" : ""}</span>
                          {canEditNote(n) && (
                            <button type="button" className="note-edit-btn" onClick={() => startEditNote(n)} title="Edit note">
                              <Pencil size={11} />
                            </button>
                          )}
                        </div>
                      </>
                    )}
                  </div>
                ))}
              </div>
              <div className="comment-composer">
                <textarea
                  value={noteText}
                  onChange={(e) => setNoteText(e.target.value)}
                  rows={2}
                  placeholder="Add a note…"
                />
                <button type="button" className="btn primary" onClick={submitNote} disabled={!noteText.trim()}>
                  Add Note
                </button>
              </div>
            </div>
          ))}

          {sideSection("Service & Maintenance History", (
            maintHistory.length === 0 ? (
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
            )
          ))}

          {sideSection("Comments", (
            <div className="comments-section">
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
          ))}
        </div>
      </div>

      <div className="modal-footer-sticky" style={{ justifyContent: "space-between" }}>
        <button
          type="button"
          className="btn danger-outline"
          onClick={onDelete}
          title={!canDeleteDirectly && asset.pendingDeletion ? "Awaiting approval" : "Delete"}
          disabled={!canDeleteDirectly && !!asset.pendingDeletion}
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
    const logAsset = data.assets.find((a) => a.id === log.assetId);
    if (log.id) {
      const maintenance = data.maintenance.map((m) => (m.id === log.id ? log : m));
      const assets = log.status === "Done"
        ? maybeRestoreStatus(data.assets, maintenance, log.assetId)
        : markUnderRepair(data.assets, log.assetId);
      next = withLog({ ...data, assets, maintenance }, currentUser, `Edited maintenance entry for ${assetLabel(log.assetId)}`, logAsset?.locationId, log.assetId);
    } else {
      const newEntry = { ...log, id: uid("maint") };
      const maintenance = [newEntry, ...data.maintenance];
      const assets = newEntry.status === "Done" ? data.assets : markUnderRepair(data.assets, newEntry.assetId);
      const becameUnderRepair = newEntry.status !== "Done" && data.assets.find((a) => a.id === newEntry.assetId)?.status !== "Under Repair";
      next = withLog({ ...data, assets, maintenance }, currentUser, `Added maintenance entry for ${assetLabel(log.assetId)}${becameUnderRepair ? " — asset set to Under Repair" : ""}`, logAsset?.locationId, log.assetId);
    }
    persist(next);
    setEditing(null);
    showToast("Maintenance entry saved.");
  };

  const remove = async (id) => {
    const log = data.maintenance.find((m) => m.id === id);
    const logAsset = data.assets.find((a) => a.id === log?.assetId);
    const maintenance = data.maintenance.filter((m) => m.id !== id);
    const assets = maybeRestoreStatus(data.assets, maintenance, log?.assetId);
    const next = withLog({ ...data, assets, maintenance }, currentUser, `Deleted maintenance entry for ${assetLabel(log?.assetId)}`, logAsset?.locationId, log?.assetId);
    persist(next);
    setConfirmDelete(null);
    showToast("Maintenance entry deleted.");
  };

  const bulkDelete = async () => {
    const affectedAssetIds = [...new Set(data.maintenance.filter((m) => selected.includes(m.id)).map((m) => m.assetId))];
    const maintenance = data.maintenance.filter((m) => !selected.includes(m.id));
    let assets = data.assets;
    affectedAssetIds.forEach((aid) => { assets = maybeRestoreStatus(assets, maintenance, aid); });
    const next = withLog({ ...data, assets, maintenance }, currentUser, `Deleted ${selected.length} maintenance entry(ies) in bulk`, scopedLocationId);
    persist(next);
    showToast(`${selected.length} entry(ies) deleted.`);
    setSelected([]);
  };

  const quickStatus = async (id, status) => {
    const log = data.maintenance.find((m) => m.id === id);
    const logAsset = data.assets.find((a) => a.id === log?.assetId);
    const maintenance = data.maintenance.map((m) => (m.id === id ? { ...m, status } : m));
    const assets = status === "Done"
      ? maybeRestoreStatus(data.assets, maintenance, log?.assetId)
      : markUnderRepair(data.assets, log?.assetId);
    const next = withLog({ ...data, assets, maintenance }, currentUser, `Changed maintenance status for ${assetLabel(log?.assetId)} to "${status}"`, logAsset?.locationId, log?.assetId);
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
  return { id: null, name: "", username: "", email: "", position: "", role: "Regional Staff", locationId: locations[0]?.id || "" };
}

function UsersView({ data, persist, showToast, currentUser }) {
  const [editing, setEditing] = useState(null);
  const [confirmDelete, setConfirmDelete] = useState(null);

  const save = async (form) => {
    if (form.id) {
      const next = withLog({ ...data, users: data.users.map((u) => (u.id === form.id ? { ...u, name: form.name, username: form.username, email: form.email, position: form.position, role: form.role, locationId: form.role === "Admin" ? null : form.locationId } : u)) }, currentUser, `Edited user "${form.name}"`);
      persist(next);
    } else {
      // This only creates the profile row (name/role/location) — it does
      // NOT create a login. Creating real Auth accounts requires the
      // service_role key, which must never reach the browser, so that
      // step happens in the Supabase dashboard (Authentication → Users →
      // Invite) using the same email, then the resulting account is
      // linked to this profile — see the hint shown below the form and
      // CRITICAL-SECURITY-STEPS.md (Phase 3/6).
      const newUser = { id: uid("usr"), name: form.name, username: form.username, email: form.email, position: form.position, role: form.role, locationId: form.role === "Admin" ? null : form.locationId, authUserId: null };
      persist(withLog({ ...data, users: [...data.users, newUser] }, currentUser, `Added user "${newUser.name}" (${newUser.role})`));
    }
    setEditing(null);
    showToast(form.id ? "User saved." : "Profile created — invite them in Supabase to finish setup.");
  };

  const remove = async (id) => {
    const u = data.users.find((x) => x.id === id);
    persist(withLog({ ...data, users: data.users.filter((u) => u.id !== id) }, currentUser, `Removed user "${u?.name}"`));
    setConfirmDelete(null);
    showToast("User removed.");
  };

  // Sends a real Supabase Auth password-reset email — this is a safe,
  // anon-key-compatible call (no service_role needed), unlike creating an
  // account outright. Requires the user to already have a linked Auth
  // account (authUserId set) with this email.
  const sendPasswordReset = async (u) => {
    if (!u.email) { alert("This user has no email on file — add one first."); return; }
    if (!u.authUserId) { alert("This account isn't linked to a login yet — see the setup hint when adding a user."); return; }
    const { error } = await supabase.auth.resetPasswordForEmail(u.email);
    if (error) { showToast("Couldn't send reset email — check your connection."); return; }
    showToast(`Password reset email sent to ${u.email}.`);
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
                    <td data-label="Role"><Badge color={u.role === "Admin" ? "#6366F1" : u.role === "Regional Admin" ? "#F59E0B" : "#10B981"}>{u.role}</Badge></td>
                    <td data-label="Location">{loc?.name || "— (Global)"}</td>
                    <td className="actions-cell">
                      <div className="row-actions">
                        <IconBtn icon={Pencil} title="Edit" onClick={() => setEditing({ ...u })} />
                        <IconBtn icon={KeyRound} title={u.authUserId ? "Send Password Reset Email" : "Not linked to a login yet"} onClick={() => sendPasswordReset(u)} />
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
                <option value="Regional Admin">Regional Admin</option>
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
            {editing.role === "Regional Admin" && (
              <div className="form-full hint-box">
                Deletion requests from Regional Staff assigned to this same location will be sent to this person for approval, instead of the Overall Admin.
              </div>
            )}
            {!editing.id && (
              <div className="form-full hint-box">
                This creates their profile only. To give them a real login: after saving,
                invite <strong>{editing.email || "their email"}</strong> in Supabase →
                Authentication → Users → Invite, then run the linking step from
                CRITICAL-SECURITY-STEPS.md (Phase 6) so this profile connects to that login.
              </div>
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
                  if (!editing.id && !editing.email.trim()) {
                    alert("Please enter an email — it's needed to link their login.");
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
    notesLog: cellify(a.notesLog || []),
    transferHistory: cellify(a.transferHistory || []),
    pendingDeletion: cellify(a.pendingDeletion || null),
    disposalInfo: cellify(a.disposalInfo || null),
  })));

  addSheet("Maintenance", (data.maintenance || []).map((m) => ({
    id: m.id, assetId: m.assetId, description: m.description, cost: m.cost, date: m.date, status: m.status,
  })));

  addSheet("Users", (data.users || []).map((u) => ({
    id: u.id, name: u.name, username: u.username, email: u.email, position: u.position,
    authUserId: u.authUserId || "", role: u.role, locationId: u.locationId || "",
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
    notesLog: parseCell(r.notesLog, []),
    transferHistory: parseCell(r.transferHistory, []),
    pendingDeletion: parseCell(r.pendingDeletion, null),
    disposalInfo: parseCell(r.disposalInfo, null),
  }));

  const maintenance = sheetRows(wb, "Maintenance").map((r) => ({
    id: String(r.id), assetId: String(r.assetId), description: r.description,
    cost: Number(r.cost) || 0, date: r.date, status: r.status,
  }));

  const users = sheetRows(wb, "Users").map((r) => ({
    id: String(r.id), name: r.name, username: r.username, email: r.email, position: r.position,
    authUserId: r.authUserId || null, role: r.role, locationId: r.locationId ? String(r.locationId) : null,
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
function ActivityLogView({ data, isAdmin, scopedLocationId, persist, showToast, currentUser }) {
  const [search, setSearch] = useState("");
  const [userFilter, setUserFilter] = useState("all");
  const [clearConfirmOpen, setClearConfirmOpen] = useState(false);
  const [clearConfirmText, setClearConfirmText] = useState("");

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

  const clearLog = () => {
    const count = (data.auditLog || []).length;
    persist(withLog({ ...data, auditLog: [] }, currentUser, `Cleared the activity log (${count} entries removed)`));
    setClearConfirmOpen(false);
    setClearConfirmText("");
    showToast?.("Activity log cleared.");
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
          {isAdmin && (data.auditLog || []).length > 0 && (
            <button type="button" className="btn ghost" onClick={() => setClearConfirmOpen(true)}>
              <Trash2 size={14} /> Clear Log
            </button>
          )}
        </div>
      </div>

      <p className="export-hint">
        {isAdmin
          ? "Every add, edit, and delete made by any user, most recent first."
          : "Activity for assets in your location, most recent first."}
      </p>

      {clearConfirmOpen && (
        <TypeToConfirmDialog
          title="Clear Activity Log"
          message="This permanently clears the entire activity log for every location. If you need to keep a copy, back it up first from Backup & Restore. This cannot be undone."
          confirmWord="CLEAR"
          value={clearConfirmText}
          onChange={setClearConfirmText}
          onCancel={() => { setClearConfirmOpen(false); setClearConfirmText(""); }}
          onConfirm={clearLog}
        />
      )}

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
function ApprovalsView({ data, persist, showToast, currentUser, isAdmin, isRegionalAdmin }) {
  // Each pending request is routed dynamically: a Regional Admin only ever
  // sees (and can act on) requests for their own location, while the
  // Overall Admin only sees requests for locations that currently have no
  // Regional Admin assigned — i.e. the ones nobody else can approve. This
  // is recomputed from the live user list every render, so adding or
  // removing a Regional Admin immediately changes who a request appears
  // in front of, even for requests that were already pending.
  const pending = isRegionalAdmin
    ? data.assets.filter((a) => a.pendingDeletion && a.locationId === currentUser.locationId)
    : data.assets.filter((a) => a.pendingDeletion && !data.users.some((u) => u.role === "Regional Admin" && u.locationId === a.locationId));

  const approve = async (asset) => {
    const removedLogs = data.maintenance.filter((m) => m.assetId === asset.id).length;
    const suffix = removedLogs > 0 ? ` (and ${removedLogs} maintenance record${removedLogs > 1 ? "s" : ""})` : "";
    const next = withLog({
      ...data,
      assets: data.assets.filter((a) => a.id !== asset.id),
      maintenance: data.maintenance.filter((m) => m.assetId !== asset.id),
    }, currentUser, `Approved deletion of asset "${asset.name || asset.tag}" — requested by ${asset.pendingDeletion.requestedByName}${suffix}`, asset.locationId);
    persist(next);
    showToast("Deletion approved.");
  };

  const reject = async (asset) => {
    const next = withLog({
      ...data,
      assets: data.assets.map((a) => (a.id === asset.id ? { ...a, pendingDeletion: null } : a)),
    }, currentUser, `Rejected deletion request for asset "${asset.name || asset.tag}" — requested by ${asset.pendingDeletion.requestedByName}`, asset.locationId, asset.id);
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
        {isRegionalAdmin
          ? "Deletion requests from Regional Staff in your location wait here until you approve or reject them. Nothing is removed until you say so."
          : "Deletion requests from Regional Staff in locations without a Regional Admin wait here until you approve or reject them. Locations with a Regional Admin are handled by them directly. Nothing is removed until you say so."}
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
              {!isRegionalAdmin && (
                <p className="login-sub" style={{ margin: "0 0 6px" }}>
                  <strong>Location:</strong> {data.locations.find((l) => l.id === a.locationId)?.name || "Unknown"} (no Regional Admin assigned)
                </p>
              )}
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
      .table-legend-label { font-weight: 700; color: var(--text); margin-right: -8px; }
      .status-legend { row-gap: 8px; margin-bottom: 12px; }

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

      /* Asset Details modal — wide two-pane layout instead of one long
         stacked column. A hero strip surfaces status/category/location at
         a glance, then the overview (left) and notes/history/comments
         (right) run side by side so the modal reads across, not down. */
      .detail-hero { display: flex; flex-wrap: wrap; align-items: center; gap: 8px; margin-bottom: 14px; }
      .detail-hero-chip { display: inline-flex; align-items: center; gap: 5px; font-size: 12px; font-weight: 600; color: var(--text-soft); background: var(--bg); border: 1px solid var(--border); border-radius: 999px; padding: 4px 10px; }

      .attention-banner { background: var(--bg); border: 1px solid var(--border); border-left: 3px solid #EF4444; border-radius: 10px; padding: 12px 14px; margin-bottom: 16px; }
      .attention-banner-title { display: flex; align-items: center; gap: 7px; font-size: 12.5px; font-weight: 700; color: #B91C1C; text-transform: uppercase; letter-spacing: 0.02em; }
      .theme-dark .attention-banner-title { color: #FCA5A5; }
      .attention-banner-list { display: flex; flex-direction: column; gap: 8px; margin-top: 10px; }
      .attention-item { display: flex; align-items: flex-start; gap: 8px; }
      .attention-dot { width: 7px; height: 7px; border-radius: 999px; margin-top: 5px; flex-shrink: 0; }
      .attention-critical .attention-dot { background: #EF4444; }
      .attention-warning .attention-dot { background: #F59E0B; }
      .attention-item-label { font-size: 12.5px; font-weight: 700; }
      .attention-critical .attention-item-label { color: #B91C1C; }
      .attention-warning .attention-item-label { color: #B45309; }
      .theme-dark .attention-critical .attention-item-label { color: #FCA5A5; }
      .theme-dark .attention-warning .attention-item-label { color: #FCD34D; }
      .attention-item-detail { font-size: 12.5px; color: var(--text-soft); line-height: 1.4; }
      .detail-layout { display: grid; grid-template-columns: 1.25fr 1fr; gap: 0 28px; align-items: start; }
      .detail-main, .detail-side { min-width: 0; }
      .detail-side { border-left: 1px solid var(--border); padding-left: 24px; }

      .detail-group { padding: 10px 0; }
      .detail-group + .detail-group { border-top: 1px solid var(--border); }
      .detail-group-label { display: flex; align-items: center; gap: 6px; font-size: 10.5px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.04em; color: var(--text-soft); margin-bottom: 6px; }
      .detail-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 0 20px; }
      .detail-row { display: flex; justify-content: space-between; gap: 10px; padding: 5px 0; font-size: 12.5px; }
      .detail-row.detail-full { grid-column: 1 / -1; }
      .detail-label { color: var(--text-soft); font-weight: 500; }
      .detail-value { font-weight: 600; text-align: right; }

      .detail-side-section { padding: 10px 0; }
      .detail-side-section + .detail-side-section { border-top: 1px solid var(--border); }
      .detail-side-section:first-child { padding-top: 0; }

      .notes-highlight {
        position: relative; background: #FEF9C3; color: #713F12; border: 1px solid #FDE68A;
        border-radius: 3px 10px 10px 10px; padding: 12px 14px; font-size: 12.5px; line-height: 1.55;
        white-space: pre-wrap; box-shadow: 0 3px 8px rgba(0,0,0,0.10), 0 1px 2px rgba(0,0,0,0.06);
        transform: rotate(-0.6deg);
      }
      .notes-highlight::before {
        content: ""; position: absolute; top: 0; right: 0; width: 0; height: 0;
        border-style: solid; border-width: 0 14px 14px 0; border-color: transparent rgba(113,63,18,0.16) transparent transparent;
        border-radius: 0 3px 0 10px;
      }

      .notes-log-list { display: flex; flex-direction: column; gap: 8px; max-height: 220px; overflow-y: auto; margin-bottom: 10px; padding: 4px 2px; }
      .note-item { background: var(--bg); border: 1px solid var(--border); border-radius: 8px; padding: 8px 10px; }
      .note-text { font-size: 13px; line-height: 1.4; white-space: pre-wrap; margin-bottom: 5px; }
      .note-meta { display: flex; align-items: center; gap: 6px; font-size: 11px; color: var(--text-soft); }
      .note-edit-btn { border: none; background: none; color: var(--text-soft); cursor: pointer; padding: 2px 4px; margin-left: auto; border-radius: 4px; display: inline-flex; }
      .note-edit-btn:hover { background: var(--surface); color: var(--text); }
      .note-item textarea { width: 100%; border: 1px solid var(--border); border-radius: 8px; padding: 8px 10px; font-size: 13px; background: var(--bg); color: var(--text); font-family: inherit; resize: vertical; margin-bottom: 6px; }
      .note-edit-actions { display: flex; justify-content: flex-end; gap: 8px; }

      .maint-history-list { list-style: none; margin: 0; padding: 0; display: flex; flex-direction: column; gap: 6px; max-height: 180px; overflow-y: auto; }
      .maint-history-item { display: flex; align-items: center; gap: 10px; background: var(--bg); border: 1px solid var(--border); border-radius: 8px; padding: 8px 10px; font-size: 12.5px; }
      .maint-history-date { color: var(--text-soft); font-family: ui-monospace, monospace; font-size: 12px; white-space: nowrap; }
      .maint-history-desc { flex: 1; min-width: 0; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
      .maint-history-cost { color: var(--text-soft); white-space: nowrap; }

      .comments-section { display: flex; flex-direction: column; }
      .comment-list { display: flex; flex-direction: column; gap: 8px; max-height: 220px; overflow-y: auto; margin-bottom: 10px; padding: 4px 2px; }
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

      .warranty-stats { display: flex; padding: 6px 8px 14px; }
      .warranty-stat { flex: 1; text-align: center; padding: 4px; min-width: 0; }
      .warranty-stat-icon { width: 24px; height: 24px; border-radius: 7px; display: flex; align-items: center; justify-content: center; margin: 0 auto 6px; }
      .warranty-stat-value { font-size: 16px; font-weight: 700; font-family: 'Poppins', sans-serif; }
      .warranty-stat-label { font-size: 9.5px; color: var(--text-soft); line-height: 1.25; margin-top: 2px; }

      /* Purchase / Asset Value panel */
      .value-total { padding: 14px 18px 2px; }
      .value-total-amount { display: block; font-size: 21px; font-weight: 700; font-family: 'Poppins', sans-serif; }
      .value-total-label { font-size: 10.5px; color: var(--text-soft); }
      .value-bar-list { list-style: none; margin: 0; padding: 12px 18px 16px; display: flex; flex-direction: column; gap: 9px; }
      .value-bar-row { display: grid; grid-template-columns: 78px 1fr 56px; align-items: center; gap: 8px; font-size: 11px; }
      .value-bar-label { color: var(--text-soft); overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
      .value-bar-track { height: 6px; border-radius: 999px; background: var(--bg); overflow: hidden; }
      .value-bar-fill { height: 100%; border-radius: 999px; }
      .value-bar-amount { text-align: right; font-weight: 600; }

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
      .preset-filters { display: flex; gap: 8px; flex-wrap: wrap; margin-bottom: 12px; }
      .preset-chip { display: inline-flex; align-items: center; font-size: 12px; font-weight: 600; padding: 5px 12px; border-radius: 999px; border: 1px solid var(--border); background: var(--surface); color: var(--text-soft); cursor: pointer; font-family: inherit; transition: background 0.15s ease, border-color 0.15s ease, color 0.15s ease; }
      .preset-chip-all.active { background: var(--accent); border-color: var(--accent); color: #fff; }
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
        background: var(--bg); color: var(--text); font-family: inherit; transition: border-color 0.15s ease, box-shadow 0.15s ease;
      }
      .field input:hover, .field select:hover, .field textarea:hover, .sort-select:hover { border-color: color-mix(in srgb, var(--accent) 40%, var(--border)); }
      .field input:focus, .field select:focus, .field textarea:focus, .sort-select:focus {
        outline: none; border-color: var(--accent); box-shadow: 0 0 0 3px color-mix(in srgb, var(--accent) 18%, transparent);
      }
      .field input:disabled, .field select:disabled, .field textarea:disabled { opacity: 0.55; cursor: not-allowed; }
      /* Every native <select> in the app — Category/Status/Type fields,
         table filters, etc. — gets the same custom chevron and hover/focus
         treatment as the rest of the UI instead of the browser default. */
      .field select, .sort-select {
        appearance: none; -webkit-appearance: none; -moz-appearance: none;
        cursor: pointer; padding-right: 30px;
        background-repeat: no-repeat; background-position: right 10px center; background-size: 15px;
      }
      .theme-light .field select, .theme-light .sort-select {
        background-image: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='24' height='24' viewBox='0 0 24 24' fill='none' stroke='%236B7280' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'%3E%3Cpolyline points='6 9 12 15 18 9'%3E%3C/polyline%3E%3C/svg%3E");
      }
      .theme-dark .field select, .theme-dark .sort-select {
        background-image: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='24' height='24' viewBox='0 0 24 24' fill='none' stroke='%239CA3AF' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'%3E%3Cpolyline points='6 9 12 15 18 9'%3E%3C/polyline%3E%3C/svg%3E");
      }
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

      /* Repair-reason callout */
      .repair-reason-box { background: color-mix(in srgb, var(--danger) 6%, var(--bg)); border: 1px solid color-mix(in srgb, var(--danger) 30%, var(--border)); border-radius: 9px; padding: 10px 12px; }

      /* New/Edit Asset wizard stepper */
      .wizard-steps { display: flex; align-items: center; margin-bottom: 16px; }
      .wizard-step { display: flex; align-items: center; gap: 6px; background: none; border: none; padding: 6px 4px; cursor: pointer; color: var(--text-soft); font-family: inherit; font-size: 12px; font-weight: 600; white-space: nowrap; }
      .wizard-step:disabled { cursor: not-allowed; opacity: 0.45; }
      .wizard-step-icon { width: 20px; height: 20px; border-radius: 999px; display: flex; align-items: center; justify-content: center; flex-shrink: 0; background: var(--bg); border: 1px solid var(--border); color: var(--text-soft); }
      .wizard-step.active { color: var(--text); }
      .wizard-step.active .wizard-step-icon { background: var(--accent); border-color: var(--accent); color: white; }
      .wizard-step.done .wizard-step-icon { background: var(--accent-soft); border-color: var(--accent); color: var(--accent); }
      .wizard-step-connector { flex: 1 1 16px; height: 1px; background: var(--border); min-width: 10px; }
      .wizard-step-connector.done { background: var(--accent); }
      .wizard-step-count { font-size: 11.5px; color: var(--text-soft); }
      @media (max-width: 640px) { .wizard-step-label { display: none; } }

      /* Sticky footer for the New/Edit Asset form */
      .modal-footer-sticky { position: sticky; bottom: 0; margin: 14px -18px -18px; padding: 12px 18px; background: var(--surface); border-top: 1px solid var(--border); display: flex; align-items: center; justify-content: flex-end; gap: 8px; border-radius: 0 0 14px 14px; }
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
        .detail-layout { grid-template-columns: 1fr; gap: 0; }
        .detail-side { border-left: none; padding-left: 0; border-top: 1px solid var(--border); margin-top: 4px; padding-top: 4px; }
        .icon-btn { width: 38px; height: 38px; }
        .notif-dot { min-width: 17px; height: 17px; font-size: 10px; }
        .notif-panel { width: calc(100vw - 24px); right: -12px; max-height: min(360px, 70vh); }
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
