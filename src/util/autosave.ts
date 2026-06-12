import type { MediaItem, Presentation } from "../model/types";

/**
 * Crash/refresh recovery: the whole document (model + media) is written to
 * IndexedDB shortly after every change and offered for restore on next launch.
 */

const DB_NAME = "presentation-editor";
const STORE = "autosave";
const KEY = "latest";
const MAX_AGE_MS = 7 * 24 * 3600 * 1000;

interface SavedMedia {
  id: string;
  mime: string;
  dataUrl: string;
  pngFallbackB64?: string;
}

export interface Snapshot {
  time: number;
  title: string;
  slideCount: number;
  pres: Presentation;
  media: SavedMedia[];
}

function openDb(): Promise<IDBDatabase> {
  return new Promise((res, rej) => {
    const req = indexedDB.open(DB_NAME, 1);
    req.onupgradeneeded = () => {
      if (!req.result.objectStoreNames.contains(STORE)) req.result.createObjectStore(STORE);
    };
    req.onsuccess = () => res(req.result);
    req.onerror = () => rej(req.error);
  });
}

function b64encode(bytes: Uint8Array): string {
  let bin = "";
  const chunk = 0x8000;
  for (let i = 0; i < bytes.length; i += chunk) {
    bin += String.fromCharCode.apply(null, Array.from(bytes.subarray(i, i + chunk)));
  }
  return btoa(bin);
}

function b64decode(b64: string): Uint8Array {
  const bin = atob(b64);
  const out = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
  return out;
}

export async function saveSnapshot(pres: Presentation, media: Map<string, MediaItem>): Promise<void> {
  try {
    const db = await openDb();
    const snap: Snapshot = {
      time: Date.now(),
      title: pres.title,
      slideCount: pres.slides.length,
      pres,
      media: [...media.values()].map(m => ({
        id: m.id, mime: m.mime, dataUrl: m.dataUrl,
        pngFallbackB64: m.pngFallback ? b64encode(m.pngFallback) : undefined,
      })),
    };
    await new Promise<void>((res, rej) => {
      const tx = db.transaction(STORE, "readwrite");
      tx.objectStore(STORE).put(snap, KEY);
      tx.oncomplete = () => res();
      tx.onerror = () => rej(tx.error);
    });
    db.close();
  } catch { /* quota/permission errors are non-fatal */ }
}

export async function loadSnapshot(): Promise<Snapshot | null> {
  try {
    const db = await openDb();
    const snap = await new Promise<Snapshot | undefined>((res, rej) => {
      const tx = db.transaction(STORE, "readonly");
      const req = tx.objectStore(STORE).get(KEY);
      req.onsuccess = () => res(req.result as Snapshot | undefined);
      req.onerror = () => rej(req.error);
    });
    db.close();
    if (!snap || Date.now() - snap.time > MAX_AGE_MS) return null;
    return snap;
  } catch {
    return null;
  }
}

export async function clearSnapshot(): Promise<void> {
  try {
    const db = await openDb();
    await new Promise<void>((res, rej) => {
      const tx = db.transaction(STORE, "readwrite");
      tx.objectStore(STORE).delete(KEY);
      tx.oncomplete = () => res();
      tx.onerror = () => rej(tx.error);
    });
    db.close();
  } catch { /* ignore */ }
}

/** Rebuild the live media map from a snapshot. */
export function mediaFromSnapshot(snap: Snapshot): Map<string, MediaItem> {
  const map = new Map<string, MediaItem>();
  for (const m of snap.media) {
    const b64 = m.dataUrl.split(",")[1] ?? "";
    map.set(m.id, {
      id: m.id,
      mime: m.mime,
      dataUrl: m.dataUrl,
      bytes: b64 ? b64decode(b64) : new Uint8Array(),
      pngFallback: m.pngFallbackB64 ? b64decode(m.pngFallbackB64) : undefined,
    });
  }
  return map;
}
