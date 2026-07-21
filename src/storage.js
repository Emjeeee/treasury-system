import { supabase } from "./supabaseClient";

const TABLE = "kv_store";

async function sbGet(key) {
  const { data, error } = await supabase
    .from(TABLE)
    .select("value")
    .eq("key", key)
    .maybeSingle();
  if (error) throw error;
  return data ? { value: data.value } : null;
}
async function sbSet(key, value) {
  const { error } = await supabase
    .from(TABLE)
    .upsert({ key, value, updated_at: new Date().toISOString() });
  if (error) throw error;
}
function lsGet(key) {
  const v = localStorage.getItem(key);
  return v == null ? null : { value: v };
}
function lsSet(key, value) {
  localStorage.setItem(key, value);
}

// shared=true -> persisted in Supabase (synced across every user/device)
// shared=false -> kept in this browser only (e.g. which account is signed in here)
window.storage = {
  get: (key, shared) => (shared ? sbGet(key) : Promise.resolve(lsGet(key))),
  set: (key, value, shared) =>
    shared ? sbSet(key, value) : Promise.resolve(lsSet(key, value)),
};
