import { sb } from "./supabase.js";

export const signUp = (email, password, name) =>
  sb.auth.signUp({ email, password, options: { data: { full_name: name } } });

export const signIn = (email, password) =>
  sb.auth.signInWithPassword({ email, password });

export const signInGoogle = () =>
  sb.auth.signInWithOAuth({
    provider: "google",
    options: { redirectTo: window.location.origin },
  });

export const resetPassword = (email) =>
  sb.auth.resetPasswordForEmail(email, {
    redirectTo: window.location.origin + "/#reset",
  });

export const updatePassword = (password) => sb.auth.updateUser({ password });

export const signOut = () => sb.auth.signOut();

// Profil = peran + status aktif
export async function loadProfile() {
  const {
    data: { user },
  } = await sb.auth.getUser();
  if (!user) return null;
  const { data } = await sb
    .from("profiles")
    .select("*")
    .eq("id", user.id)
    .single();
  if (!data || !data.active) {
    await signOut();
    return null;
  }
  return data;
}

// Menjaga sesi tetap sinkron di semua tab dan perangkat
sb.auth.onAuthStateChange((event) => {
  if (event === "SIGNED_OUT") location.reload();
});
