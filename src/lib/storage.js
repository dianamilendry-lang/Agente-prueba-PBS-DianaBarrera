/**
 * Storage adapter — usa Supabase si las variables de entorno están configuradas,
 * localStorage como fallback para desarrollo local sin credenciales.
 */
import { supabase } from "./supabase";

// ─── Supabase adapter ────────────────────────────────────────────────────────

async function sbGetIndex() {
  const { data, error } = await supabase.from("projects").select("id, name").order("created_at");
  if (error) throw error;
  return (data || []).map((r) => ({ id: r.id, name: r.name }));
}

async function sbGetProject(id) {
  const { data, error } = await supabase.from("projects").select("*").eq("id", id).single();
  if (error) throw error;
  return data;
}

async function sbSaveProject(id, payload) {
  const { error } = await supabase.from("projects").upsert({
    id,
    name:      payload.name,
    messages:  payload.messages  ?? [],
    sales:     payload.sales     ?? null,
    budget:    payload.budget    ?? null,
    inventory: payload.inventory ?? null,
  });
  if (error) throw error;
}

async function sbDeleteProject(id) {
  const { error } = await supabase.from("projects").delete().eq("id", id);
  if (error) throw error;
}

// ─── localStorage adapter ────────────────────────────────────────────────────

function lsGetIndex() {
  try { return JSON.parse(localStorage.getItem("proj_index") || "[]"); } catch { return []; }
}

function lsSetIndex(list) {
  localStorage.setItem("proj_index", JSON.stringify(list));
}

function lsGetProject(id) {
  try { return JSON.parse(localStorage.getItem("proj_" + id) || "null"); } catch { return null; }
}

function lsSaveProject(id, payload) {
  localStorage.setItem("proj_" + id, JSON.stringify(payload));
  // keep index in sync
  const idx = lsGetIndex();
  if (!idx.find((p) => p.id === id)) {
    idx.push({ id, name: payload.name });
    lsSetIndex(idx);
  }
}

function lsDeleteProject(id) {
  localStorage.removeItem("proj_" + id);
  lsSetIndex(lsGetIndex().filter((p) => p.id !== id));
}

// ─── Public API ──────────────────────────────────────────────────────────────

export async function loadProjectIndex() {
  if (supabase) return sbGetIndex();
  return lsGetIndex();
}

export async function loadProject(id) {
  if (supabase) return sbGetProject(id);
  return lsGetProject(id);
}

export async function saveProject(id, payload) {
  if (supabase) return sbSaveProject(id, payload);
  lsSaveProject(id, payload);
}

export async function deleteProject(id) {
  if (supabase) return sbDeleteProject(id);
  lsDeleteProject(id);
}

export async function saveProjectIndex(list) {
  // Solo necesario para localStorage; Supabase mantiene su propio índice
  if (!supabase) lsSetIndex(list);
}

export const storageMode = supabase ? "supabase" : "localStorage";
