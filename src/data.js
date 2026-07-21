import { sb } from "./supabase.js";

export const getConfig = () =>
  sb.from("app_config").select("*").eq("id", 1).single();

export const saveConfig = (patch) =>
  sb
    .from("app_config")
    .update({ ...patch, updated_at: new Date() })
    .eq("id", 1);
// Treasurer yang mencoba ini akan ditolak database, bukan sekadar disembunyikan tombolnya.

export const getTx = () =>
  sb.from("transactions").select("*").order("tx_date", { ascending: false });

export const addTx = (t, me) =>
  sb
    .from("transactions")
    .insert({ ...t, created_by: me.id, created_by_name: me.name });

export const updateTx = (id, patch) =>
  sb
    .from("transactions")
    .update({ ...patch, updated_at: new Date() })
    .eq("id", id);

export const verifyTx = (id, me) =>
  sb
    .from("transactions")
    .update({ status: "verified", verified_by_name: me.name })
    .eq("id", id);

export const deleteTx = (id) => sb.from("transactions").delete().eq("id", id);

export const importTx = (rows, me) =>
  sb
    .from("transactions")
    .insert(
      rows.map((r) => ({ ...r, created_by: me.id, created_by_name: me.name })),
    );

export const log = (me, action, detail) =>
  sb.from("activity_logs").insert({
    user_id: me.id,
    user_name: me.name,
    user_email: me.email,
    role: me.role,
    action,
    detail,
  });

export function subscribeAll(onChange) {
  return sb
    .channel("kas")
    .on(
      "postgres_changes",
      { event: "*", schema: "public", table: "transactions" },
      onChange,
    )
    .on(
      "postgres_changes",
      { event: "*", schema: "public", table: "app_config" },
      onChange,
    )
    .on(
      "postgres_changes",
      { event: "*", schema: "public", table: "profiles" },
      onChange,
    )
    .subscribe();
}

export const listUsers = () =>
  sb.from("profiles").select("*").order("created_at");

export const setRole = (id, role) =>
  sb.from("profiles").update({ role }).eq("id", id);
export const setActive = (id, active) =>
  sb.from("profiles").update({ active }).eq("id", id);
