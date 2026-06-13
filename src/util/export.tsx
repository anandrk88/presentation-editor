/**
 * Slide → PNG / PDF export (100% client-side).
 *
 * Each slide is rendered offscreen with the real SVG renderer (SlideSVG),
 * serialized to a standalone SVG string, rasterized through an <img> onto a
 * canvas, and emitted as a PNG. The PDF path lays one rasterized page per slide
 * via jsPDF (lazy-loaded, so it never weighs down the main bundle).
 *
 * Font note: text is rasterized with whatever font the browser has for each
 * family. System/common families render exactly; an exotic bundled family that
 * the export <img> can't see falls back — the same substitution PowerPoint does
 * for a missing font. Geometry, fills, images, charts and tables are pixel-exact.
 */
import React from "react";
import { createRoot } from "react-dom/client";
import { flushSync } from "react-dom";
import { SlideSVG, px } from "../render/SlideView";
import type { MediaItem, Presentation, SlideModel } from "../model/types";

export interface ExportOpts {
  /** Raster scale (device px per slide px). Default 2 — crisp at typical sizes. */
  scale?: number;
}

/** Render a slide to a standalone, self-contained SVG string (text laid out). */
function slideToSvgMarkup(pres: Presentation, slide: SlideModel, media: Map<string, MediaItem>): string {
  const W = px(pres.slideWidth), H = px(pres.slideHeight);
  const host = document.createElement("div");
  host.setAttribute("aria-hidden", "true");
  host.style.cssText = `position:fixed; left:-100000px; top:0; width:${W}px; height:${H}px; pointer-events:none;`;
  document.body.appendChild(host);
  const root = createRoot(host);
  try {
    // flushSync commits synchronously so the laid-out <svg> is in the DOM right away
    flushSync(() => {
      root.render(<SlideSVG pres={pres} slide={slide} media={media} width={W} height={H} />);
    });
    const svg = host.querySelector("svg");
    if (!svg) throw new Error("slide did not render");
    svg.setAttribute("xmlns", "http://www.w3.org/2000/svg");
    return new XMLSerializer().serializeToString(svg);
  } finally {
    root.unmount();
    host.remove();
  }
}

/** Rasterize a standalone SVG string onto a white canvas at the given scale. */
async function rasterize(svgMarkup: string, w: number, h: number, scale: number): Promise<HTMLCanvasElement> {
  const url = "data:image/svg+xml;charset=utf-8," + encodeURIComponent(svgMarkup);
  const img = new Image();
  await new Promise<void>((resolve, reject) => {
    img.onload = () => resolve();
    img.onerror = () => reject(new Error("slide rasterization failed"));
    img.src = url;
  });
  const canvas = document.createElement("canvas");
  canvas.width = Math.max(1, Math.round(w * scale));
  canvas.height = Math.max(1, Math.round(h * scale));
  const ctx = canvas.getContext("2d");
  if (!ctx) throw new Error("canvas 2d context unavailable");
  ctx.fillStyle = "#FFFFFF"; // PNGs are opaque; slide background paints over this
  ctx.fillRect(0, 0, canvas.width, canvas.height);
  ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
  return canvas;
}

/** Render one slide to a PNG canvas. */
export async function slideToCanvas(pres: Presentation, slide: SlideModel, media: Map<string, MediaItem>, opts: ExportOpts = {}): Promise<HTMLCanvasElement> {
  const W = px(pres.slideWidth), H = px(pres.slideHeight);
  // give any already-loaded webfonts a chance to be ready before we snapshot
  if (document.fonts?.ready) { try { await document.fonts.ready; } catch { /* ignore */ } }
  return rasterize(slideToSvgMarkup(pres, slide, media), W, H, opts.scale ?? 2);
}

/** Render one slide to a PNG Blob. */
export async function slideToPngBlob(pres: Presentation, slide: SlideModel, media: Map<string, MediaItem>, opts: ExportOpts = {}): Promise<Blob> {
  const canvas = await slideToCanvas(pres, slide, media, opts);
  return new Promise<Blob>((resolve, reject) =>
    canvas.toBlob(b => (b ? resolve(b) : reject(new Error("PNG encode failed"))), "image/png"));
}

/** Render the whole deck to a multi-page PDF Blob (one slide per page). */
export async function exportPdfBlob(pres: Presentation, media: Map<string, MediaItem>, opts: ExportOpts = {}): Promise<Blob> {
  const { jsPDF } = await import("jspdf");
  const W = px(pres.slideWidth), H = px(pres.slideHeight);
  const orientation = W >= H ? "landscape" : "portrait";
  const pdf = new jsPDF({ orientation, unit: "px", format: [W, H], compress: true });
  for (let i = 0; i < pres.slides.length; i++) {
    const canvas = await slideToCanvas(pres, pres.slides[i], media, opts);
    if (i > 0) pdf.addPage([W, H], orientation);
    pdf.addImage(canvas.toDataURL("image/png"), "PNG", 0, 0, W, H);
  }
  return pdf.output("blob");
}

/** Zip every slide's PNG into one archive Blob. */
export async function exportPngZipBlob(pres: Presentation, media: Map<string, MediaItem>, opts: ExportOpts = {}): Promise<Blob> {
  const JSZip = (await import("jszip")).default;
  const zip = new JSZip();
  const pad = String(pres.slides.length).length;
  for (let i = 0; i < pres.slides.length; i++) {
    const blob = await slideToPngBlob(pres, pres.slides[i], media, opts);
    zip.file(`slide-${String(i + 1).padStart(pad, "0")}.png`, blob);
  }
  return zip.generateAsync({ type: "blob" });
}

/** Trigger a browser download of a Blob. */
export function downloadBlob(blob: Blob, filename: string): void {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  a.click();
  setTimeout(() => URL.revokeObjectURL(url), 5000);
}
