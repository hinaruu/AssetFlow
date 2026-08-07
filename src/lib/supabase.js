import { createClient } from "@supabase/supabase-js";

const supabaseUrl = import.meta.env.VITE_SUPABASE_URL;
const supabaseAnonKey = import.meta.env.VITE_SUPABASE_ANON_KEY;

if (!supabaseUrl || !supabaseAnonKey) {
  // eslint-disable-next-line no-console
  console.warn(
    "Missing VITE_SUPABASE_URL / VITE_SUPABASE_ANON_KEY. " +
    "Create a .env file (see .env.example) with your Supabase project credentials."
  );
}

export const supabase = createClient(supabaseUrl || "", supabaseAnonKey || "");

// A blank form field comes through as "" (empty string), not null/undefined.
// Postgres numeric columns reject "" outright, so any optional number field
// must be funneled through this before hitting the database.
function num(v) {
  if (v === "" || v === null || v === undefined) return null;
  const n = Number(v);
  return Number.isNaN(n) ? null : n;
}

/* ---------------------------------------------------------
   Row <-> app-object mapping.
   The app still works with the exact same camelCase shape it
   always has ({ locations, categories, users, assets, maintenance,
   auditLog }) — only how it gets in/out of the database changed.
--------------------------------------------------------- */

const TABLES = {
  locations: {
    table: "locations",
    toRow: (l) => ({ id: l.id, name: l.name }),
    fromRow: (r) => ({ id: r.id, name: r.name }),
  },
  categories: {
    table: "categories",
    toRow: (c) => ({ id: c.id, name: c.name, type: c.type, useful_life: c.usefulLife ?? null }),
    fromRow: (r) => ({ id: r.id, name: r.name, type: r.type, usefulLife: r.useful_life }),
  },
  users: {
    table: "users",
    // Credentials now live in Supabase Auth (auth.users), not here — this
    // table only stores the profile (role/location) and a link to the
    // matching Auth account via auth_user_id. See CRITICAL-SECURITY-STEPS.md.
    toRow: (u) => ({
      id: u.id, name: u.name, username: u.username, email: u.email || null,
      position: u.position || null, role: u.role,
      location_id: u.locationId || null, auth_user_id: u.authUserId || null,
    }),
    fromRow: (r) => ({
      id: r.id, name: r.name, username: r.username, email: r.email,
      position: r.position, role: r.role, locationId: r.location_id, authUserId: r.auth_user_id,
    }),
  },
  assets: {
    table: "assets",
    toRow: (a) => ({
      id: a.id, tag: a.tag || null, name: a.name || null, department: a.department || null,
      category_id: a.categoryId || null, asset_type: a.assetType || null,
      brand: a.brand || null, model: a.model || null, year_model: a.yearModel || null, serial: a.serial || null,
      status: a.status || null, condition: a.condition || null,
      location_id: a.locationId || null, assigned_to: a.assignedTo || null,
      purchase_date: a.purchaseDate || null, purchase_cost: num(a.purchaseCost),
      warranty_expiry: (a.warrantyExpiry && a.warrantyExpiry !== "N/A") ? a.warrantyExpiry : null,
      requires_calibration: !!a.requiresCalibration,
      calibration_date: a.calibrationDate || null,
      next_calibration_date: a.nextCalibrationDate || null,
      notes: a.notes || null, pre_repair_status: a.preRepairStatus || null,
      transfer_history: a.transferHistory || [],
      disposed_by: a.disposalInfo?.by || null,
      disposal_reason: a.disposalInfo?.reason || null,
      disposed_at: a.disposalInfo?.date || null,
      disposal_logged_at: a.disposalInfo?.at || null,
      created_by_id: a.createdById || null, created_by_name: a.createdByName || null,
      pending_deletion: a.pendingDeletion || null,
      updated_at: new Date().toISOString(),
    }),
    fromRow: (r) => ({
      id: r.id, tag: r.tag, name: r.name, department: r.department, categoryId: r.category_id, assetType: r.asset_type,
      brand: r.brand, model: r.model, yearModel: r.year_model, serial: r.serial, status: r.status, condition: r.condition,
      locationId: r.location_id, assignedTo: r.assigned_to,
      purchaseDate: r.purchase_date, purchaseCost: r.purchase_cost, warrantyExpiry: r.warranty_expiry,
      requiresCalibration: r.requires_calibration, calibrationDate: r.calibration_date,
      nextCalibrationDate: r.next_calibration_date, notes: r.notes,
      preRepairStatus: r.pre_repair_status, transferHistory: r.transfer_history || [],
      disposalInfo: r.disposed_by || r.disposal_reason || r.disposed_at
        ? { by: r.disposed_by, reason: r.disposal_reason, date: r.disposed_at, at: r.disposal_logged_at }
        : null,
      createdById: r.created_by_id, createdByName: r.created_by_name,
      pendingDeletion: r.pending_deletion || null,
    }),
  },
  maintenance: {
    table: "maintenance",
    toRow: (m) => ({
      id: m.id, asset_id: m.assetId || null, description: m.description || null,
      cost: num(m.cost), date: m.date || null, status: m.status || null,
    }),
    fromRow: (r) => ({
      id: r.id, assetId: r.asset_id, description: r.description, cost: r.cost, date: r.date, status: r.status,
    }),
  },
  auditLog: {
    table: "audit_log",
    // asset_id is only ever set for entries that relate to a single asset
    // (see withLog in App.jsx) — it powers the broader Overall Admin /
    // Regional Admin asset-activity notification feed and is never shown
    // in the Activity Log itself, so that view's behavior is unchanged.
    toRow: (e) => ({ id: e.id, at: e.at, user_id: e.userId || null, user_name: e.userName || null, message: e.message || null, location_id: e.locationId || null, asset_id: e.assetId || null }),
    fromRow: (r) => ({ id: r.id, at: r.at, userId: r.user_id, userName: r.user_name, message: r.message, locationId: r.location_id, assetId: r.asset_id }),
  },
  comments: {
    table: "comments",
    toRow: (c) => ({
      id: c.id, asset_id: c.assetId || null, at: c.at,
      author_id: c.authorId || null, author_name: c.authorName || null,
      message: c.message || null,
      target_user_ids: c.targetUserIds || [],
      read_by: c.readBy || [],
    }),
    fromRow: (r) => ({
      id: r.id, assetId: r.asset_id, at: r.at,
      authorId: r.author_id, authorName: r.author_name,
      message: r.message,
      targetUserIds: r.target_user_ids || [],
      readBy: r.read_by || [],
    }),
  },
  // Per-user "last viewed" timestamp for a given asset. This is what
  // powers the asset-activity notification badges/bell for Overall Admin
  // and Regional Admin (see computeAssetActivityFeed / withAssetRead in
  // App.jsx) — comparing an asset's latest activity timestamp against the
  // current user's row here is how "unread" is determined, and it's what
  // gets updated the moment they open that asset.
  notificationReads: {
    table: "notification_reads",
    toRow: (r) => ({ id: r.id, user_id: r.userId || null, asset_id: r.assetId || null, last_read_at: r.lastReadAt || null }),
    fromRow: (r) => ({ id: r.id, userId: r.user_id, assetId: r.asset_id, lastReadAt: r.last_read_at }),
  },
};

/**
 * Loads all six tables and assembles them into the same shape the app
 * has always used. Returns null only if every table is empty (first run).
 */
export async function fetchOrgData() {
  const entries = Object.entries(TABLES);
  const results = await Promise.all(
    entries.map(([, def]) =>
      def.table === "audit_log"
        ? supabase.from("audit_log").select("*").order("at", { ascending: false }).limit(300)
        : supabase.from(def.table).select("*")
    )
  );

  results.forEach((r, i) => {
    if (r.error) throw r.error;
  });

  const data = {};
  entries.forEach(([key, def], i) => {
    data[key] = (results[i].data || []).map(def.fromRow);
  });

  const isEmpty = Object.values(data).every((arr) => arr.length === 0);
  return isEmpty ? null : data;
}

/**
 * Saves org data by diffing `next` against `prev` (the last known state)
 * per entity, per row — so a single edit only writes the row(s) that
 * actually changed instead of rewriting everyone's data.
 *
 * `prev` should be the data object you last loaded/saved. Pass null/undefined
 * on first save (e.g. seeding) — everything will be treated as new.
 */
export async function saveOrgData(next, prev) {
  const before = prev || {};

  for (const [key, def] of Object.entries(TABLES)) {
    const nextList = next[key] || [];
    const prevList = before[key] || [];
    const prevById = new Map(prevList.map((r) => [r.id, r]));
    const nextIds = new Set(nextList.map((r) => r.id));

    const toUpsert = nextList.filter((r) => {
      const p = prevById.get(r.id);
      return !p || JSON.stringify(p) !== JSON.stringify(r);
    });
    const toDeleteIds = prevList
      .filter((r) => !nextIds.has(r.id))
      .map((r) => r.id);

    if (toUpsert.length) {
      const { error } = await supabase.from(def.table).upsert(toUpsert.map(def.toRow));
      if (error) throw error;
    }
    if (toDeleteIds.length) {
      const { error } = await supabase.from(def.table).delete().in("id", toDeleteIds);
      if (error) throw error;
    }
  }
}

/**
 * Subscribes to live changes across all six tables via Supabase Realtime.
 * On any change from any user, refetches the full data set and calls
 * onChange with it — no polling, no manual sync button needed.
 *
 * Returns an unsubscribe function — call it on unmount to clean up.
 */
export function subscribeToOrgData(onChange) {
  let cancelled = false;
  const refetch = async () => {
    try {
      const data = await fetchOrgData();
      if (!cancelled && data) onChange(data);
    } catch {
      // Transient errors on a live-update refetch aren't worth surfacing —
      // the next successful change event (or manual Sync) will catch up.
    }
  };

  const channel = supabase.channel("org_data_live");
  Object.values(TABLES).forEach((def) => {
    channel.on("postgres_changes", { event: "*", schema: "public", table: def.table }, refetch);
  });
  channel.subscribe();

  return () => {
    cancelled = true;
    supabase.removeChannel(channel);
  };
}
