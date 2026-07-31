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

const ROW_ID = 1;

/**
 * Fetch the single shared org-data record.
 * Returns null if the table is empty (first run) so the caller can seed it.
 * Throws if the request fails (offline, misconfigured, etc.) — callers
 * should catch this and show a connection-error state.
 */
export async function fetchOrgData() {
  const { data, error } = await supabase
    .from("app_data")
    .select("payload")
    .eq("id", ROW_ID)
    .maybeSingle();

  if (error) throw error;
  if (!data?.payload) return null;
  // Backward compatibility: older saved data won't have an auditLog yet.
  return { auditLog: [], ...data.payload };
}

/**
 * Write the full org-data object back to the shared database row.
 * Uses upsert so the very first save (seeding) works even if the row
 * doesn't exist yet.
 */
export async function saveOrgData(payload) {
  const { error } = await supabase
    .from("app_data")
    .upsert({ id: ROW_ID, payload, updated_at: new Date().toISOString() });

  if (error) throw error;
}
