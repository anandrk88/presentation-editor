import { nextId } from "../model/defaults";
import { store } from "../state/store";

/**
 * Reads an image file into the media registry (browser only).
 * SVG: root width/height normalized from viewBox + PNG fallback rasterized
 * (PowerPoint needs the bitmap blip next to the asvg extension).
 */
export async function loadImageFile(file: File): Promise<{ mediaId: string; natW: number; natH: number }> {
  let bytes = new Uint8Array(await file.arrayBuffer());
  let mime = file.type || "image/png";
  let dataUrl: string;
  let pngFallback: Uint8Array | undefined;

  if (mime === "image/svg+xml" || file.name.toLowerCase().endsWith(".svg")) {
    mime = "image/svg+xml";
    const fixed = normalizeSvg(new TextDecoder().decode(bytes));
    bytes = new TextEncoder().encode(fixed);
    dataUrl = "data:image/svg+xml;base64," + btoa(unescape(encodeURIComponent(fixed)));
    pngFallback = await rasterizeSvg(dataUrl);
  } else {
    dataUrl = await new Promise<string>(res => {
      const r = new FileReader();
      r.onload = () => res(r.result as string);
      r.readAsDataURL(file);
    });
  }

  const img = new Image();
  await new Promise<void>(res => { img.onload = () => res(); img.onerror = () => res(); img.src = dataUrl; });
  const natW = img.naturalWidth || 400, natH = img.naturalHeight || 300;
  const mediaId = nextId("media");
  store.media.set(mediaId, { id: mediaId, mime, bytes, dataUrl, pngFallback });
  return { mediaId, natW, natH };
}

/** Fetch an image from a URL into the media registry (CORS must permit it). */
export async function loadImageFromUrl(url: string): Promise<{ mediaId: string; natW: number; natH: number }> {
  const res = await fetch(url, { mode: "cors" });
  if (!res.ok) throw new Error(`Could not fetch image (HTTP ${res.status})`);
  const blob = await res.blob();
  if (!/^image\//.test(blob.type) && !/\.(png|jpe?g|gif|svg|bmp|webp)(\?|$)/i.test(url)) {
    throw new Error("That URL does not point to an image");
  }
  const name = (url.split("/").pop() || "image").split("?")[0] || "image";
  const file = new File([blob], name, { type: blob.type || "image/png" });
  return loadImageFile(file);
}

/** Ensure the svg root has explicit width/height (derived from viewBox when absent). */
export function normalizeSvg(text: string): string {
  try {
    const doc = new DOMParser().parseFromString(text, "image/svg+xml");
    const root = doc.documentElement;
    if (root.localName !== "svg") return text;
    if (!root.getAttribute("width") || !root.getAttribute("height")) {
      const vb = (root.getAttribute("viewBox") ?? "").trim().split(/[\s,]+/).map(Number);
      const w = vb.length === 4 && vb[2] > 0 ? vb[2] : 300;
      const h = vb.length === 4 && vb[3] > 0 ? vb[3] : 300;
      root.setAttribute("width", String(w));
      root.setAttribute("height", String(h));
      if (vb.length !== 4) root.setAttribute("viewBox", `0 0 ${w} ${h}`);
    }
    return new XMLSerializer().serializeToString(root);
  } catch {
    return text;
  }
}

/** Rasterize an svg data url to PNG bytes (PowerPoint's fallback blip). */
export async function rasterizeSvg(svgDataUrl: string): Promise<Uint8Array | undefined> {
  try {
    const img = new Image();
    await new Promise<void>((res, rej) => {
      img.onload = () => res();
      img.onerror = () => rej(new Error("svg load failed"));
      img.src = svgDataUrl;
    });
    const natW = img.naturalWidth || 300, natH = img.naturalHeight || 300;
    const scale = Math.min(2, 2048 / Math.max(natW, natH));
    const canvas = document.createElement("canvas");
    canvas.width = Math.max(1, Math.round(natW * scale));
    canvas.height = Math.max(1, Math.round(natH * scale));
    const ctx = canvas.getContext("2d");
    if (!ctx) return undefined;
    ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
    const blob = await new Promise<Blob | null>(res => canvas.toBlob(res, "image/png"));
    return blob ? new Uint8Array(await blob.arrayBuffer()) : undefined;
  } catch {
    return undefined; // export will fall back to embedding the bare svg
  }
}
