import type { ColorTheme } from "../model/types";

/**
 * User-defined color palettes and font pairs (Design tab "Customize…").
 * Persisted in localStorage so they survive reloads; Google-font pairs
 * re-inject their stylesheet on boot so the editor can render them.
 */

const THEMES_KEY = "pe.customThemes";
const FONTS_KEY = "pe.customFontPairs";

export interface CustomFontPair {
  name: string;   // display label
  major: string;  // headings font family
  minor: string;  // body font family
  google?: boolean; // load both families from Google Fonts
}

function read<T>(key: string): T[] {
  if (typeof localStorage === "undefined") return [];
  try {
    const v = JSON.parse(localStorage.getItem(key) ?? "[]");
    return Array.isArray(v) ? v : [];
  } catch {
    return [];
  }
}

function write(key: string, value: unknown) {
  if (typeof localStorage === "undefined") return;
  try { localStorage.setItem(key, JSON.stringify(value)); } catch { /* quota */ }
}

// ---------- color palettes ----------
export function loadCustomThemes(): ColorTheme[] {
  return read<ColorTheme>(THEMES_KEY).filter(t => t && typeof t.name === "string" && typeof t.accent1 === "string");
}

export function saveCustomTheme(t: ColorTheme) {
  const list = loadCustomThemes().filter(x => x.name !== t.name);
  list.push(t);
  write(THEMES_KEY, list);
}

export function deleteCustomTheme(name: string) {
  write(THEMES_KEY, loadCustomThemes().filter(x => x.name !== name));
}

// ---------- font pairs ----------
export function loadCustomFontPairs(): CustomFontPair[] {
  return read<CustomFontPair>(FONTS_KEY).filter(p => p && p.major && p.minor);
}

export function saveCustomFontPair(p: CustomFontPair) {
  const list = loadCustomFontPairs().filter(x => x.name !== p.name);
  list.push(p);
  write(FONTS_KEY, list);
  if (p.google) { injectGoogleFont(p.major); injectGoogleFont(p.minor); }
}

export function deleteCustomFontPair(name: string) {
  write(FONTS_KEY, loadCustomFontPairs().filter(x => x.name !== name));
}

/** Distinct custom family names — merged into the Home-tab font combo. */
export function customFontNames(): string[] {
  const out = new Set<string>();
  for (const p of loadCustomFontPairs()) { out.add(p.major); out.add(p.minor); }
  return [...out].sort();
}

/** Load a family from Google Fonts (regular + bold); editor-side rendering only. */
export function injectGoogleFont(family: string) {
  if (typeof document === "undefined" || !family.trim()) return;
  const id = "gf-" + family.trim().replace(/\s+/g, "-").toLowerCase();
  if (document.getElementById(id)) return;
  const link = document.createElement("link");
  link.id = id;
  link.rel = "stylesheet";
  link.href = `https://fonts.googleapis.com/css2?family=${family.trim().replace(/\s+/g, "+")}:ital,wght@0,400;0,700;1,400&display=swap`;
  document.head.appendChild(link);
}

/** Re-inject Google-font stylesheets for saved pairs (call once on boot). */
export function initCustomFonts() {
  for (const p of loadCustomFontPairs()) {
    if (p.google) { injectGoogleFont(p.major); injectGoogleFont(p.minor); }
  }
}
