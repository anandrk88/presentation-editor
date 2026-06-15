import React, { useEffect, useLayoutEffect, useRef, useState } from "react";
import type { ChartShape, PicShape, Shape, SpShape, TableShape } from "../model/types";
import { EMU_PER_PX, makeShape, makeTextBox } from "../model/defaults";
import { isLineGeom, presetPath } from "../render/geometry";
import { SlideSVG, px, shapeTransform } from "../render/SlideView";
import { chartPartRegions, tableGrid } from "../render/GraphicViews";
import { store } from "../state/store";
import { useEditorState } from "../state/useStore";
import { loadImageFile, loadImageFromUrl } from "../util/loadImage";
import { TextEditOverlay } from "./TextEditOverlay";

type HandleId = "n" | "s" | "e" | "w" | "ne" | "nw" | "se" | "sw";
const HANDLE_VEC: Record<HandleId, [number, number]> = {
  nw: [-1, -1], n: [0, -1], ne: [1, -1], e: [1, 0], se: [1, 1], s: [0, 1], sw: [-1, 1], w: [-1, 0],
};

interface DragMove { kind: "move"; start: Pt; originals: Map<string, Shape>; moved: boolean; downHit?: string; wasFullGroup?: boolean }
/** Resize a multi-shape selection (group): scale every member about the union-bbox anchor. */
interface DragGroupResize {
  kind: "gresize"; handle: HandleId; hx: number; hy: number;
  o: Pt; b0: { w: number; h: number };
  originals: Map<string, Shape>; moved: boolean;
}
/** Rotate a multi-shape selection: every member orbits the union-bbox center and spins itself. */
interface DragGroupRotate {
  kind: "grotate"; center: Pt; startAng: number;
  originals: Map<string, Shape>; moved: boolean;
}
interface DragResize {
  kind: "resize"; handle: HandleId; o: Pt; rot: number; hx: number; hy: number;
  w0: number; h0: number; id: string; moved: boolean;
}
interface DragRotate { kind: "rotate"; center: Pt; startAng: number; rot0: number; id: string; moved: boolean }
interface DragMarquee { kind: "marquee"; start: Pt; additive: boolean }
interface DragDraw { kind: "draw"; start: Pt; geom: string }
/** Crop-mode: drag a crop handle — the image extent stays fixed, the frame edge follows. */
interface DragCrop {
  kind: "crop"; handle: HandleId; id: string; rot: number;
  c0: Pt; fw: number; fh: number;
  ext: { x: number; y: number; w: number; h: number }; // extent in start-frame-local px
  moved: boolean;
}
/** Crop-mode: drag the picture — the frame stays fixed, the image slides underneath. */
interface DragCropPan {
  kind: "croppan"; id: string; rot: number; start: Pt;
  fw: number; fh: number; ext0: { x: number; y: number; w: number; h: number };
  moved: boolean;
}
/** Drag a rectangular cell-range selection inside an already-selected table. */
interface DragCellSel { kind: "cellsel"; id: string; anchorR: number; anchorC: number }
type Drag = DragMove | DragResize | DragRotate | DragMarquee | DragDraw | DragCrop | DragCropPan | DragCellSel | DragGroupResize | DragGroupRotate | null;

interface Pt { x: number; y: number }

const rotate = (x: number, y: number, deg: number): Pt => {
  const r = (deg * Math.PI) / 180, c = Math.cos(r), s = Math.sin(r);
  return { x: x * c - y * s, y: x * s + y * c };
};

function shapeAabb(s: Shape): { x1: number; y1: number; x2: number; y2: number } {
  const w = px(s.w), h = px(s.h), cx = px(s.x) + w / 2, cy = px(s.y) + h / 2;
  const corners = [[-w / 2, -h / 2], [w / 2, -h / 2], [w / 2, h / 2], [-w / 2, h / 2]]
    .map(([dx, dy]) => rotate(dx, dy, s.rot));
  const xs = corners.map(c => cx + c.x), ys = corners.map(c => cy + c.y);
  return { x1: Math.min(...xs), y1: Math.min(...ys), x2: Math.max(...xs), y2: Math.max(...ys) };
}

function cursorFor(handle: HandleId, rot: number): string {
  const base: Record<HandleId, number> = { e: 0, ne: 45, n: 90, nw: 135, w: 180, sw: 225, s: 270, se: 315 };
  const ang = ((base[handle] - rot) % 360 + 360) % 360;
  const idx = Math.round(ang / 45) % 8;
  return ["ew-resize", "nesw-resize", "ns-resize", "nwse-resize", "ew-resize", "nesw-resize", "ns-resize", "nwse-resize"][idx];
}

/** Frame size + full-image extent (in frame-local px) for a picture's current crop. */
function cropGeom(s: PicShape) {
  const fw = px(s.w), fh = px(s.h);
  const c = s.srcRect ?? { l: 0, t: 0, r: 0, b: 0 };
  const ew = fw / Math.max(0.01, 1 - c.l - c.r);
  const eh = fh / Math.max(0.01, 1 - c.t - c.b);
  return { fw, fh, ext: { x: -c.l * ew, y: -c.t * eh, w: ew, h: eh } };
}

/** Clamp tiny negatives, round, and drop the srcRect entirely when it is a no-op. */
function normSrcRect(c: { l: number; t: number; r: number; b: number }) {
  const f = (v: number) => Math.max(0, Math.round(v * 10000) / 10000);
  const o = { l: f(c.l), t: f(c.t), r: f(c.r), b: f(c.b) };
  return o.l || o.t || o.r || o.b ? o : undefined;
}

export function EditorCanvas() {
  const state = useEditorState();
  const containerRef = useRef<HTMLDivElement>(null);
  const svgRef = useRef<SVGSVGElement>(null);
  const dragRef = useRef<Drag>(null);
  // pointer capture retargets the derived dblclick to the svg, losing the shape
  // target — remember what the last pointerdown actually hit
  const lastDownHit = useRef<string | null>(null);
  const [containerSize, setContainerSize] = useState({ w: 800, h: 500 });
  const [ctxMenu, setCtxMenu] = useState<{ x: number; y: number; shapeId: string | null; cell?: { row: number; col: number } } | null>(null);
  const [marquee, setMarquee] = useState<{ x: number; y: number; w: number; h: number } | null>(null);
  const [guides, setGuides] = useState<{ v?: boolean; h?: boolean }>({});
  const [ghost, setGhost] = useState<{ x: number; y: number; w: number; h: number } | null>(null);

  const { pres, selection, editingShapeId, pendingShape } = state;
  const slide = pres.slides[selection.slideIndex];
  const W = px(pres.slideWidth), H = px(pres.slideHeight);

  // Re-bind on showRuler: toggling the ruler re-parents .workspace into/out of
  // the grid wrapper, so React remounts it — the observer must follow the live
  // node, and we ignore 0×0 (the detached old node reports zero and would clamp
  // fit-zoom to its 5% floor).
  useLayoutEffect(() => {
    const el = containerRef.current;
    if (!el) return;
    const measure = () => {
      const w = el.clientWidth, h = el.clientHeight;
      if (w > 0 && h > 0) setContainerSize({ w, h });
    };
    const ro = new ResizeObserver(measure);
    ro.observe(el);
    measure();
    return () => ro.disconnect();
  }, [state.showRuler]);

  const PAD = 28;
  const zoom = state.zoom === "fit"
    ? Math.max(0.05, Math.min((containerSize.w - PAD * 2) / W, (containerSize.h - PAD * 2) / H))
    : state.zoom === "fitw"
      ? Math.max(0.05, (containerSize.w - PAD * 2) / W)
      : state.zoom;

  useEffect(() => {
    if (Math.abs(state.effectiveZoom - zoom) > 0.0001) store.setState({ effectiveZoom: zoom });
  }, [zoom, state.effectiveZoom]);

  const toSlide = (e: { clientX: number; clientY: number }): Pt => {
    const rect = svgRef.current!.getBoundingClientRect();
    return { x: ((e.clientX - rect.left) / rect.width) * W, y: ((e.clientY - rect.top) / rect.height) * H };
  };

  const editingShape = editingShapeId
    ? (slide.shapes.find(s => s.id === editingShapeId && (s.kind === "sp" || s.kind === "table")) as SpShape | TableShape | undefined)
    : undefined;

  // ---------- pointer handlers ----------
  const onPointerDown = (e: React.PointerEvent<SVGSVGElement>) => {
    if (e.button !== 0) return;
    const p = toSlide(e);
    const target = e.target as Element;
    lastDownHit.current = target.closest("[data-hit-id]")?.getAttribute("data-hit-id") ?? null;
    (e.currentTarget as SVGSVGElement).setPointerCapture(e.pointerId);

    if (pendingShape) {
      dragRef.current = { kind: "draw", start: p, geom: pendingShape };
      setGhost({ x: p.x, y: p.y, w: 0, h: 0 });
      return;
    }

    // chart text-element selection: clicking a part region (over a selected chart)
    // selects it for per-element formatting (Home tab + chart pane)
    const partEl = target.closest("[data-chart-part]");
    if (partEl) {
      const cid = partEl.getAttribute("data-chart-id")!;
      const part = partEl.getAttribute("data-chart-part") as import("../model/types").ChartPart;
      store.setChartPart(cid, part);
      return;
    }

    // crop mode: handles + pan surface take priority; any other press exits crop
    if (state.croppingId) {
      const cropPic = slide.shapes.find(sh => sh.id === state.croppingId && sh.kind === "pic") as PicShape | undefined;
      const cropHit = target.closest("[data-crop]")?.getAttribute("data-crop");
      if (cropPic && cropHit) {
        const g = cropGeom(cropPic);
        if (cropHit === "pan") {
          dragRef.current = {
            kind: "croppan", id: cropPic.id, rot: cropPic.rot, start: p,
            fw: g.fw, fh: g.fh, ext0: g.ext, moved: false,
          };
        } else {
          dragRef.current = {
            kind: "crop", handle: cropHit as HandleId, id: cropPic.id, rot: cropPic.rot,
            c0: { x: px(cropPic.x) + g.fw / 2, y: px(cropPic.y) + g.fh / 2 },
            fw: g.fw, fh: g.fh, ext: g.ext, moved: false,
          };
        }
        return;
      }
      store.setState({ croppingId: null }); // click elsewhere commits + leaves crop mode
    }

    const handle = target.closest("[data-handle]")?.getAttribute("data-handle");
    if (handle === "rot" && selection.shapeIds.length > 1) {
      // group rotation: spin every member around the union-bbox center
      const members = slide.shapes.filter(s => selection.shapeIds.includes(s.id));
      if (!members.length) return;
      let x1 = Infinity, y1 = Infinity, x2 = -Infinity, y2 = -Infinity;
      members.forEach(s => {
        const b = shapeAabb(s);
        x1 = Math.min(x1, b.x1); y1 = Math.min(y1, b.y1);
        x2 = Math.max(x2, b.x2); y2 = Math.max(y2, b.y2);
      });
      const center = { x: (x1 + x2) / 2, y: (y1 + y2) / 2 };
      const originals = new Map<string, Shape>();
      members.forEach(s => originals.set(s.id, s));
      dragRef.current = {
        kind: "grotate", center,
        startAng: (Math.atan2(p.y - center.y, p.x - center.x) * 180) / Math.PI,
        originals, moved: false,
      };
      return;
    }
    if (handle && handle !== "rot" && selection.shapeIds.length > 1) {
      // group/multi-selection resize: scale every member about the union-bbox anchor
      const members = slide.shapes.filter(s => selection.shapeIds.includes(s.id));
      if (!members.length) return;
      let x1 = Infinity, y1 = Infinity, x2 = -Infinity, y2 = -Infinity;
      members.forEach(s => {
        const b = shapeAabb(s);
        x1 = Math.min(x1, b.x1); y1 = Math.min(y1, b.y1);
        x2 = Math.max(x2, b.x2); y2 = Math.max(y2, b.y2);
      });
      const [hx, hy] = HANDLE_VEC[handle as HandleId];
      const originals = new Map<string, Shape>();
      members.forEach(s => originals.set(s.id, s));
      dragRef.current = {
        kind: "gresize", handle: handle as HandleId, hx, hy,
        o: { x: hx > 0 ? x1 : hx < 0 ? x2 : (x1 + x2) / 2, y: hy > 0 ? y1 : hy < 0 ? y2 : (y1 + y2) / 2 },
        b0: { w: Math.max(1, x2 - x1), h: Math.max(1, y2 - y1) },
        originals, moved: false,
      };
      return;
    }
    if (handle && selection.shapeIds.length === 1) {
      const s = slide.shapes.find(sh => sh.id === selection.shapeIds[0]);
      if (!s) return;
      const w0 = px(s.w), h0 = px(s.h);
      const cx = px(s.x) + w0 / 2, cy = px(s.y) + h0 / 2;
      if (handle === "rot") {
        dragRef.current = {
          kind: "rotate", center: { x: cx, y: cy },
          startAng: (Math.atan2(p.y - cy, p.x - cx) * 180) / Math.PI,
          rot0: s.rot, id: s.id, moved: false,
        };
      } else {
        const [hx, hy] = HANDLE_VEC[handle as HandleId];
        const off = rotate((-hx * w0) / 2, (-hy * h0) / 2, s.rot);
        dragRef.current = {
          kind: "resize", handle: handle as HandleId,
          o: { x: cx + off.x, y: cy + off.y }, rot: s.rot, hx, hy, w0, h0, id: s.id, moved: false,
        };
      }
      return;
    }

    const hit = target.closest("[data-hit-id]")?.getAttribute("data-hit-id");

    // already-selected table: an interior press selects cells (drag = range, shift = extend);
    // a press near the outer border falls through to a normal move, like PowerPoint's frame.
    if (hit && !editingShapeId) {
      const hitShape = slide.shapes.find(s => s.id === hit);
      if (hitShape?.kind === "table" && selection.shapeIds.length === 1 && selection.shapeIds[0] === hit) {
        const { row, col, nearEdge } = cellAt(hitShape, p);
        if (!nearEdge) {
          const prev = state.tableSel?.id === hit ? state.tableSel : null;
          const aR = e.shiftKey && prev ? prev.r0 : row;
          const aC = e.shiftKey && prev ? prev.c0 : col;
          store.setState({
            tableSel: { id: hit, r0: Math.min(aR, row), c0: Math.min(aC, col), r1: Math.max(aR, row), c1: Math.max(aC, col) },
            editingShapeId: null, editingCell: null,
          });
          dragRef.current = { kind: "cellsel", id: hit, anchorR: aR, anchorC: aC };
          return;
        }
      }
    }

    if (hit) {
      let ids = selection.shapeIds;
      const groupIds = store.expandToGroups([hit]); // whole group when the shape is grouped
      const fullGroupSelected = groupIds.length > 1 && ids.length === groupIds.length && groupIds.every(g => ids.includes(g));
      if (e.shiftKey || e.ctrlKey) {
        ids = ids.includes(hit)
          ? ids.filter(i => !groupIds.includes(i))
          : [...new Set([...ids, ...groupIds])];
        store.selectShapes(ids);
      } else if (!ids.includes(hit)) {
        ids = groupIds;
        store.selectShapes(ids);
      }
      if (state.tableSel && state.tableSel.id !== hit) store.setState({ tableSel: null });
      const originals = new Map<string, Shape>();
      slide.shapes.forEach(s => { if (ids.includes(s.id)) originals.set(s.id, s); });
      dragRef.current = { kind: "move", start: p, originals, moved: false, downHit: hit, wasFullGroup: fullGroupSelected };
      return;
    }

    // empty canvas: marquee select
    dragRef.current = { kind: "marquee", start: p, additive: e.shiftKey };
    if (!e.shiftKey) { store.selectShapes([]); if (state.tableSel) store.setState({ tableSel: null }); }
  };

  const onPointerMove = (e: React.PointerEvent<SVGSVGElement>) => {
    const drag = dragRef.current;
    if (!drag) return;
    const p = toSlide(e);

    if (drag.kind === "draw") {
      setGhost({
        x: Math.min(drag.start.x, p.x), y: Math.min(drag.start.y, p.y),
        w: Math.abs(p.x - drag.start.x), h: Math.abs(p.y - drag.start.y),
      });
      return;
    }

    if (drag.kind === "marquee") {
      setMarquee({
        x: Math.min(drag.start.x, p.x), y: Math.min(drag.start.y, p.y),
        w: Math.abs(p.x - drag.start.x), h: Math.abs(p.y - drag.start.y),
      });
      return;
    }

    if (drag.kind === "cellsel") {
      const tbl = slide.shapes.find(s => s.id === drag.id);
      if (tbl?.kind !== "table") return;
      const { row, col } = cellAt(tbl, p);
      store.setState({
        tableSel: {
          id: drag.id,
          r0: Math.min(drag.anchorR, row), c0: Math.min(drag.anchorC, col),
          r1: Math.max(drag.anchorR, row), c1: Math.max(drag.anchorC, col),
        },
      });
      return;
    }

    if (drag.kind === "move") {
      let dx = p.x - drag.start.x, dy = p.y - drag.start.y;
      if (Math.abs(dx) + Math.abs(dy) > 1.5 / zoom) drag.moved = true;
      if (!drag.moved) return;
      // smart guides: snap selection center to slide center
      const ids = Array.from(drag.originals.keys());
      let bx1 = Infinity, by1 = Infinity, bx2 = -Infinity, by2 = -Infinity;
      drag.originals.forEach(s => {
        const b = shapeAabb(s);
        bx1 = Math.min(bx1, b.x1); by1 = Math.min(by1, b.y1);
        bx2 = Math.max(bx2, b.x2); by2 = Math.max(by2, b.y2);
      });
      const thr = 4 / zoom;
      const cX = (bx1 + bx2) / 2 + dx, cY = (by1 + by2) / 2 + dy;
      const g: { v?: boolean; h?: boolean } = {};
      if (Math.abs(cX - W / 2) < thr) { dx += W / 2 - cX; g.v = true; }
      if (Math.abs(cY - H / 2) < thr) { dy += H / 2 - cY; g.h = true; }
      setGuides(g);
      const dEmuX = Math.round(dx * EMU_PER_PX), dEmuY = Math.round(dy * EMU_PER_PX);
      store.updateShapes(ids, s => {
        const o = drag.originals.get(s.id)!;
        return { ...s, x: o.x + dEmuX, y: o.y + dEmuY };
      }, { historic: false });
      return;
    }

    if (drag.kind === "grotate") {
      drag.moved = true;
      const ang = (Math.atan2(p.y - drag.center.y, p.x - drag.center.x) * 180) / Math.PI;
      let delta = ang - drag.startAng;
      if (e.shiftKey) delta = Math.round(delta / 15) * 15;
      else {
        // soft-snap the group to upright multiples
        const firstRot = drag.originals.values().next().value?.rot ?? 0;
        for (const snap of [0, 90, 180, 270, 360]) {
          const total = ((firstRot + delta) % 360 + 360) % 360;
          if (Math.abs(total - snap) < 3) { delta += snap % 360 - total; break; }
        }
      }
      store.updateShapes([...drag.originals.keys()], s => {
        const o = drag.originals.get(s.id)!;
        const ocx = px(o.x) + px(o.w) / 2, ocy = px(o.y) + px(o.h) / 2;
        const nc = rotate(ocx - drag.center.x, ocy - drag.center.y, delta);
        const ncx = drag.center.x + nc.x, ncy = drag.center.y + nc.y;
        return {
          ...s,
          x: Math.round((ncx - px(o.w) / 2) * EMU_PER_PX),
          y: Math.round((ncy - px(o.h) / 2) * EMU_PER_PX),
          rot: ((o.rot + delta) % 360 + 360) % 360,
        };
      }, { historic: false });
      return;
    }

    if (drag.kind === "gresize") {
      drag.moved = true;
      const newW = drag.hx === 1 ? p.x - drag.o.x : drag.hx === -1 ? drag.o.x - p.x : drag.b0.w;
      const newH = drag.hy === 1 ? p.y - drag.o.y : drag.hy === -1 ? drag.o.y - p.y : drag.b0.h;
      let kx = drag.hx !== 0 ? Math.max(0.02, newW / drag.b0.w) : 1;
      let ky = drag.hy !== 0 ? Math.max(0.02, newH / drag.b0.h) : 1;
      if (e.shiftKey && drag.hx !== 0 && drag.hy !== 0) { kx = ky = Math.max(kx, ky); }
      store.updateShapes([...drag.originals.keys()], s => {
        const o = drag.originals.get(s.id)!;
        const nx = drag.o.x + (px(o.x) - drag.o.x) * kx;
        const ny = drag.o.y + (px(o.y) - drag.o.y) * ky;
        return {
          ...s,
          x: Math.round(nx * EMU_PER_PX), y: Math.round(ny * EMU_PER_PX),
          w: Math.max(1, Math.round(px(o.w) * kx * EMU_PER_PX)),
          h: Math.max(1, Math.round(px(o.h) * ky * EMU_PER_PX)),
        };
      }, { historic: false });
      return;
    }

    if (drag.kind === "resize") {
      drag.moved = true;
      const v = rotate(p.x - drag.o.x, p.y - drag.o.y, -drag.rot);
      const MIN = 4;
      let w = drag.hx !== 0 ? Math.max(MIN, v.x / drag.hx) : drag.w0;
      let h = drag.hy !== 0 ? Math.max(MIN, v.y / drag.hy) : drag.h0;
      if (e.shiftKey && drag.hx !== 0 && drag.hy !== 0) {
        const k = Math.max(w / drag.w0, h / drag.h0);
        w = drag.w0 * k; h = drag.h0 * k;
      }
      const centerLocal = { x: drag.hx !== 0 ? (drag.hx * w) / 2 : 0, y: drag.hy !== 0 ? (drag.hy * h) / 2 : 0 };
      const cw = rotate(centerLocal.x, centerLocal.y, drag.rot);
      const cx = drag.o.x + cw.x, cy = drag.o.y + cw.y;
      store.updateShapes([drag.id], s => ({
        ...s,
        x: Math.round((cx - w / 2) * EMU_PER_PX),
        y: Math.round((cy - h / 2) * EMU_PER_PX),
        w: Math.round(w * EMU_PER_PX),
        h: Math.round(h * EMU_PER_PX),
      }), { historic: false });
      return;
    }

    if (drag.kind === "rotate") {
      drag.moved = true;
      const ang = (Math.atan2(p.y - drag.center.y, p.x - drag.center.x) * 180) / Math.PI;
      let rot = (drag.rot0 + ang - drag.startAng) % 360;
      if (rot < 0) rot += 360;
      if (e.shiftKey) rot = Math.round(rot / 15) * 15 % 360;
      else {
        for (const snap of [0, 90, 180, 270, 360]) {
          if (Math.abs(rot - snap) < 3) { rot = snap % 360; break; }
        }
      }
      store.updateShapes([drag.id], s => ({ ...s, rot }), { historic: false });
      return;
    }

    if (drag.kind === "croppan") {
      drag.moved = true;
      const d = rotate(p.x - drag.start.x, p.y - drag.start.y, -drag.rot);
      // extent must always cover the frame
      const nx = Math.min(0, Math.max(drag.fw - drag.ext0.w, drag.ext0.x + d.x));
      const ny = Math.min(0, Math.max(drag.fh - drag.ext0.h, drag.ext0.y + d.y));
      const srcRect = normSrcRect({
        l: -nx / drag.ext0.w, r: (nx + drag.ext0.w - drag.fw) / drag.ext0.w,
        t: -ny / drag.ext0.h, b: (ny + drag.ext0.h - drag.fh) / drag.ext0.h,
      });
      store.updateShapes([drag.id], s => (s.kind === "pic" ? { ...s, srcRect } : s), { historic: false });
      return;
    }

    if (drag.kind === "crop") {
      drag.moved = true;
      // pointer in start-frame-local coords
      const lp = rotate(p.x - drag.c0.x, p.y - drag.c0.y, -drag.rot);
      const lx = lp.x + drag.fw / 2, ly = lp.y + drag.fh / 2;
      const [hx, hy] = HANDLE_VEC[drag.handle];
      const MIN = 8;
      let x1 = 0, y1 = 0, x2 = drag.fw, y2 = drag.fh;
      if (hx < 0) x1 = Math.min(Math.max(drag.ext.x, lx), x2 - MIN);
      if (hx > 0) x2 = Math.max(Math.min(drag.ext.x + drag.ext.w, lx), x1 + MIN);
      if (hy < 0) y1 = Math.min(Math.max(drag.ext.y, ly), y2 - MIN);
      if (hy > 0) y2 = Math.max(Math.min(drag.ext.y + drag.ext.h, ly), y1 + MIN);
      const w = x2 - x1, h = y2 - y1;
      // the image extent stays fixed on screen; the frame (and its center) moves
      const nc = rotate(x1 + w / 2 - drag.fw / 2, y1 + h / 2 - drag.fh / 2, drag.rot);
      const cx = drag.c0.x + nc.x, cy = drag.c0.y + nc.y;
      const srcRect = normSrcRect({
        l: (x1 - drag.ext.x) / drag.ext.w, r: (drag.ext.x + drag.ext.w - x2) / drag.ext.w,
        t: (y1 - drag.ext.y) / drag.ext.h, b: (drag.ext.y + drag.ext.h - y2) / drag.ext.h,
      });
      store.updateShapes([drag.id], s => (s.kind === "pic" ? {
        ...s,
        x: Math.round((cx - w / 2) * EMU_PER_PX), y: Math.round((cy - h / 2) * EMU_PER_PX),
        w: Math.round(w * EMU_PER_PX), h: Math.round(h * EMU_PER_PX),
        srcRect,
      } : s), { historic: false });
    }
  };

  const onPointerUp = (e: React.PointerEvent<SVGSVGElement>) => {
    const drag = dragRef.current;
    dragRef.current = null;
    setGuides({});
    if (!drag) return;

    if (drag.kind === "draw") {
      const g = ghost;
      setGhost(null);
      store.setState({ pendingShape: null });
      if (!g) return;
      let { x, y, w, h } = g;
      const tooSmall = w < 8 / zoom && h < 8 / zoom;
      if (drag.geom === "textbox") {
        if (tooSmall) { w = 240; h = 45; }
        const sh = makeTextBox(Math.round(x * EMU_PER_PX), Math.round(y * EMU_PER_PX), Math.round(w * EMU_PER_PX), Math.round(h * EMU_PER_PX));
        store.addShape(sh);
        store.setState({ editingShapeId: sh.id });
      } else {
        if (tooSmall) { w = 152; h = isLineGeom(drag.geom as any) ? 0 : 114; }
        if (e.shiftKey && w && h) { w = h = Math.max(w, h); }
        const sh = makeShape(drag.geom as SpShape["geom"], Math.round(x * EMU_PER_PX), Math.round(y * EMU_PER_PX), Math.round(Math.max(w, 1) * EMU_PER_PX), Math.round(Math.max(h, isLineGeom(drag.geom as any) ? 0 : 1) * EMU_PER_PX));
        store.addShape(sh);
      }
      return;
    }

    if (drag.kind === "marquee") {
      const m = marquee;
      setMarquee(null);
      if (m && (m.w > 2 || m.h > 2)) {
        const hitIds = slide.shapes.filter(s => {
          const b = shapeAabb(s);
          return b.x1 < m.x + m.w && b.x2 > m.x && b.y1 < m.y + m.h && b.y2 > m.y;
        }).map(s => s.id);
        store.selectShapes(store.expandToGroups(hitIds), drag.additive);
      } else if (!drag.additive) {
        // a plain click on empty canvas (no drag, no shift) advances to the next
        // slide — deselect already happened on pointer-down; Esc clears selection
        const idx = store.getState().selection.slideIndex;
        if (idx < pres.slides.length - 1) store.selectSlide(idx + 1);
      }
      return;
    }

    if (drag.kind === "move" || drag.kind === "resize" || drag.kind === "rotate" || drag.kind === "crop" || drag.kind === "croppan" || drag.kind === "gresize" || drag.kind === "grotate") {
      store.endPreview(drag.moved);
      // PowerPoint drill-in: a stationary click on a fully-selected group selects just that member
      if (drag.kind === "move" && !drag.moved && drag.wasFullGroup && drag.downHit) {
        store.selectShapes([drag.downHit]);
      }
    }
  };

  /** Cell under a slide-space point (same resolution as dblclick, incl. merge anchors). */
  const cellAt = (s: TableShape, p: Pt): { row: number; col: number; nearEdge: boolean } => {
    const tw = px(s.w), th = px(s.h);
    const cx = px(s.x) + tw / 2, cy = px(s.y) + th / 2;
    const lp = rotate(p.x - cx, p.y - cy, -s.rot);
    const lx = lp.x + tw / 2, ly = lp.y + th / 2;
    const grid = tableGrid(s);
    let col = grid.colW.length - 1, row = grid.rowH.length - 1;
    for (let c = 0; c < grid.colW.length; c++) if (lx < grid.colX[c + 1]) { col = c; break; }
    for (let r = 0; r < grid.rowH.length; r++) if (ly < grid.rowY[r + 1]) { row = r; break; }
    while (col > 0 && s.cells[row]?.[col]?.merged === "h") col--;
    while (row > 0 && s.cells[row]?.[col]?.merged === "v") row--;
    // within this margin of the outer border → grab-to-move (like PowerPoint's frame)
    const edge = 6 / zoom;
    const nearEdge = lx < edge || ly < edge || lx > tw - edge || ly > th - edge;
    return { row, col, nearEdge };
  };

  // ---------- replace image (file / url) ----------
  const replaceInputRef = useRef<HTMLInputElement>(null);
  const replaceTarget = useRef<string | null>(null);
  const replaceFromFile = (picId: string) => {
    replaceTarget.current = picId;
    replaceInputRef.current?.click();
  };
  const replaceFromUrl = async (picId: string) => {
    const url = window.prompt("Image URL (must allow cross-origin access):", "https://");
    if (!url || url.trim() === "https://") return;
    try {
      store.setStatus("Fetching image…");
      const { mediaId } = await loadImageFromUrl(url.trim());
      store.replacePicMedia(picId, mediaId);
    } catch (err) {
      store.setStatus(null);
      alert("Could not load image: " + (err as Error).message);
    }
  };

  // right-click: select what's under the cursor and show object/canvas actions
  const onContextMenu = (e: React.MouseEvent) => {
    e.preventDefault();
    if (state.editingShapeId) return; // native menu is useful while editing text
    const hit = (e.target as Element).closest?.("[data-hit-id]")?.getAttribute("data-hit-id") ?? null;
    if (hit && !selection.shapeIds.includes(hit)) store.selectShapes([hit]);
    if (!hit) store.selectShapes([]);
    const hitShape = hit ? slide.shapes.find(s => s.id === hit) : undefined;
    const cell = hitShape?.kind === "table" ? cellAt(hitShape, toSlide(e)) : undefined;
    setCtxMenu({ x: e.clientX, y: e.clientY, shapeId: hit, cell });
  };

  useEffect(() => {
    if (!ctxMenu) return;
    const close = () => setCtxMenu(null);
    document.addEventListener("pointerdown", close);
    window.addEventListener("blur", close);
    return () => {
      document.removeEventListener("pointerdown", close);
      window.removeEventListener("blur", close);
    };
  }, [ctxMenu]);

  const onDoubleClick = (e: React.MouseEvent) => {
    const hit = (e.target as Element).closest?.("[data-hit-id]")?.getAttribute("data-hit-id")
      ?? lastDownHit.current;
    if (!hit) return;
    const s = slide.shapes.find(sh => sh.id === hit);
    if (!s) return;
    if (s.kind === "sp" && !isLineGeom(s.geom)) {
      store.selectShapes([s.id]);
      store.setState({ editingShapeId: s.id, editingCell: null });
    } else if (s.kind === "table") {
      const p = toSlide(e);
      // into shape-local coords (undo rotation around center)
      const cx = px(s.x) + px(s.w) / 2, cy = px(s.y) + px(s.h) / 2;
      const lp = rotate(p.x - cx, p.y - cy, -s.rot);
      const lx = lp.x + px(s.w) / 2, ly = lp.y + px(s.h) / 2;
      const grid = tableGrid(s);
      let col = grid.colW.length - 1, row = grid.rowH.length - 1;
      for (let c = 0; c < grid.colW.length; c++) if (lx < grid.colX[c + 1]) { col = c; break; }
      for (let r = 0; r < grid.rowH.length; r++) if (ly < grid.rowY[r + 1]) { row = r; break; }
      // spanned-away cell -> jump to its anchor
      while (col > 0 && s.cells[row]?.[col]?.merged === "h") col--;
      while (row > 0 && s.cells[row]?.[col]?.merged === "v") row--;
      store.selectShapes([s.id]);
      store.setState({ editingShapeId: s.id, editingCell: { row, col }, tableSel: null });
    } else if (s.kind === "chart") {
      store.selectShapes([s.id]);
      store.setState({ chartEditId: s.id });
    } else if (s.kind === "pic") {
      store.selectShapes([s.id]);
      store.setState({ croppingId: s.id });
    }
  };

  // keep zoom anchored when using ctrl+wheel
  const onWheel = (e: React.WheelEvent) => {
    if (!e.ctrlKey) return;
    e.preventDefault();
    const cur = zoom;
    const next = Math.min(4, Math.max(0.1, cur * (e.deltaY > 0 ? 0.9 : 1.1)));
    store.setState({ zoom: Math.round(next * 100) / 100 });
  };

  useEffect(() => {
    const el = containerRef.current;
    if (!el) return;
    const handler = (e: WheelEvent) => { if (e.ctrlKey) e.preventDefault(); };
    el.addEventListener("wheel", handler, { passive: false });
    return () => el.removeEventListener("wheel", handler);
  }, []);

  const selected = slide.shapes.filter(s => selection.shapeIds.includes(s.id));
  const single = selected.length === 1 ? selected[0] : null;
  const hs = 4 / zoom; // half handle size (slide px)

  // crop mode follows the selection: deselecting (or losing) the picture exits
  const cropShape = !editingShapeId && state.croppingId
    ? (slide.shapes.find(s => s.id === state.croppingId && s.kind === "pic") as PicShape | undefined)
    : undefined;
  useEffect(() => {
    if (state.croppingId && (!selection.shapeIds.includes(state.croppingId)
      || !slide.shapes.some(s => s.id === state.croppingId && s.kind === "pic"))) {
      store.setState({ croppingId: null });
    }
  }, [state.croppingId, selection.shapeIds, slide]);

  // cell-range selection follows the table: clear it once the table stops being the sole selection
  const cellSelTable = !editingShapeId && state.tableSel
    ? (slide.shapes.find(s => s.id === state.tableSel!.id && s.kind === "table") as TableShape | undefined)
    : undefined;
  useEffect(() => {
    if (state.tableSel && !(selection.shapeIds.length === 1 && selection.shapeIds[0] === state.tableSel.id
      && slide.shapes.some(s => s.id === state.tableSel!.id && s.kind === "table"))) {
      store.setState({ tableSel: null });
    }
  }, [state.tableSel, selection.shapeIds, slide]);

  // ---------- rulers ----------
  const [slideOrigin, setSlideOrigin] = useState({ x: 0, y: 0 });
  const measureOrigin = () => {
    const ws = containerRef.current;
    const wrap = ws?.querySelector(".slide-wrap");
    if (!ws || !wrap) return;
    const a = ws.getBoundingClientRect();
    const b = wrap.getBoundingClientRect();
    setSlideOrigin(prev => {
      const next = { x: b.left - a.left, y: b.top - a.top };
      return Math.abs(prev.x - next.x) > 0.5 || Math.abs(prev.y - next.y) > 0.5 ? next : prev;
    });
  };
  useLayoutEffect(measureOrigin, [zoom, containerSize, state.showRuler, selection.slideIndex]);

  const canvasBody = (
    <div className="workspace" ref={containerRef} onWheel={onWheel} onScroll={measureOrigin}>
      <div
        className="slide-wrap"
        style={{ width: W * zoom, height: H * zoom, cursor: pendingShape ? "crosshair" : undefined }}
        onContextMenu={onContextMenu}
      >
        <SlideSVG
          pres={pres}
          slide={slide}
          media={store.media}
          width={W * zoom}
          height={H * zoom}
          hiddenShapeId={editingShapeId}
          hiddenCell={state.editingCell}
          showPrompts
          svgRef={svgRef}
          onPointerDown={onPointerDown}
        >
          {/* hit layer */}
          {!editingShapeId && slide.shapes.map(s => (
            <g key={`hit-${s.id}`} transform={shapeTransform(s)} data-hit-id={s.id} style={{ cursor: pendingShape ? "crosshair" : "move" }}>
              <rect x={0} y={0} width={Math.max(px(s.w), 8)} height={Math.max(px(s.h), 8)} fill="transparent" stroke="none" />
            </g>
          ))}
          {/* smart guides */}
          {guides.v && <line x1={W / 2} y1={0} x2={W / 2} y2={H} stroke="#BF4A2B" strokeWidth={1 / zoom} strokeDasharray={`${5 / zoom} ${3 / zoom}`} />}
          {guides.h && <line x1={0} y1={H / 2} x2={W} y2={H / 2} stroke="#BF4A2B" strokeWidth={1 / zoom} strokeDasharray={`${5 / zoom} ${3 / zoom}`} />}
          {/* selection chrome */}
          {!editingShapeId && selected.filter(s => s.id !== cropShape?.id).map(s => (
            <g key={`sel-${s.id}`} transform={shapeTransform(s)} pointerEvents="none">
              <rect
                x={0} y={0} width={px(s.w)} height={px(s.h)}
                fill="none" stroke="#888888" strokeWidth={1 / zoom} strokeDasharray={`${4 / zoom} ${3 / zoom}`}
              />
            </g>
          ))}
          {/* multi-selection / group: resize handles around the union bbox */}
          {selected.length > 1 && !editingShapeId && (() => {
            let x1 = Infinity, y1 = Infinity, x2 = -Infinity, y2 = -Infinity;
            selected.forEach(s => {
              const b = shapeAabb(s);
              x1 = Math.min(x1, b.x1); y1 = Math.min(y1, b.y1);
              x2 = Math.max(x2, b.x2); y2 = Math.max(y2, b.y2);
            });
            const isGroup = selected.every(s => s.groupId && s.groupId === selected[0].groupId);
            return (
              <g>
                <rect
                  x={x1} y={y1} width={x2 - x1} height={y2 - y1}
                  fill="none" stroke={isGroup ? "#2B5797" : "#888888"} strokeWidth={1 / zoom}
                  strokeDasharray={`${5 / zoom} ${3 / zoom}`} pointerEvents="none"
                />
                {/* rotation handle above the union frame */}
                <line x1={(x1 + x2) / 2} y1={y1} x2={(x1 + x2) / 2} y2={y1 - 18 / zoom} stroke="#777" strokeWidth={1 / zoom} />
                <circle
                  data-handle="rot"
                  cx={(x1 + x2) / 2} cy={y1 - 18 / zoom} r={4.5 / zoom}
                  fill="#fff" stroke="#777" strokeWidth={1 / zoom}
                  style={{ cursor: "grab" }}
                />
                {(Object.keys(HANDLE_VEC) as HandleId[]).map(hid => {
                  const [hx, hy] = HANDLE_VEC[hid];
                  const cx = x1 + ((hx + 1) / 2) * (x2 - x1);
                  const cy = y1 + ((hy + 1) / 2) * (y2 - y1);
                  return (
                    <rect
                      key={hid}
                      data-handle={hid}
                      x={cx - hs} y={cy - hs} width={hs * 2} height={hs * 2}
                      fill="#FFFFFF" stroke="#777777" strokeWidth={1 / zoom}
                      style={{ cursor: cursorFor(hid, 0) }}
                    />
                  );
                })}
              </g>
            );
          })()}
          {single && !editingShapeId && single.id !== cropShape?.id && (
            <g transform={shapeTransform(single)}>
              {/* rotation handle */}
              <line x1={px(single.w) / 2} y1={0} x2={px(single.w) / 2} y2={-18 / zoom} stroke="#777" strokeWidth={1 / zoom} />
              <circle
                data-handle="rot"
                cx={px(single.w) / 2} cy={-18 / zoom} r={4.5 / zoom}
                fill="#fff" stroke="#777" strokeWidth={1 / zoom}
                style={{ cursor: "grab" }}
              />
              {(Object.keys(HANDLE_VEC) as HandleId[]).map(hid => {
                const [hx, hy] = HANDLE_VEC[hid];
                const cx = ((hx + 1) / 2) * px(single.w);
                const cy = ((hy + 1) / 2) * px(single.h);
                return (
                  <rect
                    key={hid}
                    data-handle={hid}
                    x={cx - hs} y={cy - hs} width={hs * 2} height={hs * 2}
                    fill="#FFFFFF" stroke="#777777" strokeWidth={1 / zoom}
                    style={{ cursor: cursorFor(hid, single.rot) }}
                  />
                );
              })}
            </g>
          )}
          {/* chart text-element overlay: clickable part regions + highlight for the selected part */}
          {single && single.kind === "chart" && !editingShapeId && (
            <g transform={shapeTransform(single)}>
              {chartPartRegions(single as ChartShape).map(r => {
                const sel = state.chartPartSel?.id === single.id && state.chartPartSel.part === r.part;
                return (
                  <rect
                    key={r.part}
                    data-chart-part={r.part}
                    data-chart-id={single.id}
                    x={r.x} y={r.y} width={r.w} height={r.h}
                    fill={sel ? "rgba(43,87,151,0.12)" : "transparent"}
                    stroke={sel ? "#2B5797" : "none"}
                    strokeWidth={sel ? 1 / zoom : 0}
                    strokeDasharray={sel ? `${3 / zoom} ${2 / zoom}` : undefined}
                    style={{ cursor: "text" }}
                  />
                );
              })}
            </g>
          )}
          {/* crop mode chrome: ghosted full image, pan surface, frame + black crop handles */}
          {cropShape && (() => {
            const g = cropGeom(cropShape);
            const url = store.media.get(cropShape.mediaId)?.dataUrl;
            const leg = 14 / zoom, thick = 4 / zoom, pad = 9 / zoom;
            const maskId = `crop-mask-${cropShape.id}`;
            // corner bracket: two legs meeting at (cx,cy), opening toward (sx,sy)
            const corner = (hid: HandleId, cx: number, cy: number, sx: number, sy: number) => (
              <g key={hid}>
                <path
                  d={`M${cx + sx * leg} ${cy + sy * (thick / 2)} L${cx + sx * (thick / 2)} ${cy + sy * (thick / 2)} L${cx + sx * (thick / 2)} ${cy + sy * leg}`}
                  fill="none" stroke="#1F1F1F" strokeWidth={thick} pointerEvents="none"
                />
                <rect
                  data-crop={hid}
                  x={Math.min(cx, cx + sx * leg) - pad / 2} y={Math.min(cy, cy + sy * leg) - pad / 2}
                  width={leg + pad} height={leg + pad}
                  fill="transparent" style={{ cursor: cursorFor(hid, cropShape.rot) }}
                />
              </g>
            );
            // center a bar of `thick` on the frame edge coordinate
            const onEdge = (v: number) => v - thick / 2;
            const edge = (hid: HandleId, cx: number, cy: number, horiz: boolean) => (
              <g key={hid}>
                <rect
                  x={horiz ? cx - leg / 2 : onEdge(cx)} y={horiz ? onEdge(cy) : cy - leg / 2}
                  width={horiz ? leg : thick} height={horiz ? thick : leg}
                  fill="#1F1F1F" pointerEvents="none"
                />
                <rect
                  data-crop={hid}
                  x={(horiz ? cx - leg / 2 : onEdge(cx)) - pad / 2} y={(horiz ? onEdge(cy) : cy - leg / 2) - pad / 2}
                  width={(horiz ? leg : thick) + pad} height={(horiz ? thick : leg) + pad}
                  fill="transparent" style={{ cursor: cursorFor(hid, cropShape.rot) }}
                />
              </g>
            );
            return (
              <g transform={shapeTransform(cropShape)}>
                <defs>
                  <mask id={maskId}>
                    <rect x={g.ext.x} y={g.ext.y} width={g.ext.w} height={g.ext.h} fill="#fff" />
                    <rect x={0} y={0} width={g.fw} height={g.fh} fill="#000" />
                  </mask>
                </defs>
                {url && (
                  <image
                    href={url} x={g.ext.x} y={g.ext.y} width={g.ext.w} height={g.ext.h}
                    preserveAspectRatio="none" opacity={0.4} mask={`url(#${maskId})`}
                    pointerEvents="none"
                  />
                )}
                <rect
                  x={g.ext.x} y={g.ext.y} width={g.ext.w} height={g.ext.h}
                  fill="none" stroke="#9A9A9A" strokeWidth={1 / zoom} strokeDasharray={`${3 / zoom} ${2 / zoom}`}
                  pointerEvents="none"
                />
                {/* pan surface: drag anywhere on the picture to slide it under the frame */}
                <rect data-crop="pan" x={g.ext.x} y={g.ext.y} width={g.ext.w} height={g.ext.h} fill="transparent" style={{ cursor: "move" }} />
                <rect x={0} y={0} width={g.fw} height={g.fh} fill="none" stroke="#1F1F1F" strokeWidth={1.2 / zoom} pointerEvents="none" />
                {corner("nw", 0, 0, 1, 1)}
                {corner("ne", g.fw, 0, -1, 1)}
                {corner("se", g.fw, g.fh, -1, -1)}
                {corner("sw", 0, g.fh, 1, -1)}
                {edge("n", g.fw / 2, 0, true)}
                {edge("s", g.fw / 2, g.fh, true)}
                {edge("w", 0, g.fh / 2, false)}
                {edge("e", g.fw, g.fh / 2, false)}
              </g>
            );
          })()}
          {/* table cell-range selection highlight (union of the selected rectangle) */}
          {cellSelTable && state.tableSel && (() => {
            const t = cellSelTable;
            const { r0, c0, r1, c1 } = state.tableSel;
            const grid = tableGrid(t);
            const x0 = grid.colX[Math.min(c0, grid.colX.length - 1)] ?? 0;
            const x1 = grid.colX[Math.min(c1 + 1, grid.colX.length - 1)] ?? grid.w;
            const y0 = grid.rowY[Math.min(r0, grid.rowY.length - 1)] ?? 0;
            const y1 = grid.rowY[Math.min(r1 + 1, grid.rowY.length - 1)] ?? grid.h;
            const multi = r0 !== r1 || c0 !== c1;
            return (
              <g transform={shapeTransform(t)} pointerEvents="none">
                <rect
                  x={x0} y={y0} width={Math.max(0, x1 - x0)} height={Math.max(0, y1 - y0)}
                  fill="rgba(43,87,151,0.22)" stroke="#2B5797" strokeWidth={(multi ? 1.6 : 1) / zoom}
                />
              </g>
            );
          })()}
          {/* marquee */}
          {marquee && (
            <rect
              x={marquee.x} y={marquee.y} width={marquee.w} height={marquee.h}
              fill="rgba(183,91,68,0.08)" stroke="#B75B44" strokeWidth={1 / zoom} strokeDasharray={`${4 / zoom} ${2 / zoom}`}
            />
          )}
          {/* draw ghost */}
          {ghost && pendingShape && (
            pendingShape === "textbox" || pendingShape === "rect"
              ? <rect x={ghost.x} y={ghost.y} width={ghost.w} height={ghost.h} fill="none" stroke="#B75B44" strokeWidth={1 / zoom} strokeDasharray={`${4 / zoom} ${2 / zoom}`} />
              : <g transform={`translate(${ghost.x} ${ghost.y})`}>
                  <path d={presetPath(pendingShape as any, Math.max(ghost.w, 1), Math.max(ghost.h, 1))} fill="none" stroke="#B75B44" strokeWidth={1 / zoom} strokeDasharray={`${4 / zoom} ${2 / zoom}`} />
                </g>
          )}
        </SlideSVG>
        {/* React needs the move/up handlers on the svg element — attach via wrapper */}
        <PointerProxy svgRef={svgRef} onMove={onPointerMove} onUp={onPointerUp} onDblClick={onDoubleClick} />
        {editingShape && (
          <TextEditOverlay
            key={`${editingShape.id}:${state.editingCell?.row ?? -1}:${state.editingCell?.col ?? -1}`}
            shape={editingShape}
            cell={state.editingCell}
            zoom={zoom}
            theme={pres.theme}
          />
        )}
      </div>
      {ctxMenu && (
        <div className="ctx-menu" style={{ left: ctxMenu.x, top: ctxMenu.y, position: "fixed" }} onPointerDown={e => e.stopPropagation()}>
          {ctxMenu.shapeId ? (
            <>
              {ctxMenu.cell && (() => {
                const tid = ctxMenu.shapeId!;
                const { row, col } = ctxMenu.cell;
                const tbl = slide.shapes.find(s => s.id === tid);
                const anchor = tbl?.kind === "table" ? tbl.cells[row]?.[col] : undefined;
                const canSplit = !!anchor && ((anchor.gridSpan ?? 1) > 1 || (anchor.rowSpan ?? 1) > 1);
                const sel = state.tableSel?.id === tid ? state.tableSel : null;
                const isRange = !!sel && (sel.r0 !== sel.r1 || sel.c0 !== sel.c1);
                return (
                  <>
                    <button className="menu-item" onClick={() => { store.insertTableRow(tid, row, true); setCtxMenu(null); }}>Insert Row Above</button>
                    <button className="menu-item" onClick={() => { store.insertTableRow(tid, row, false); setCtxMenu(null); }}>Insert Row Below</button>
                    <button className="menu-item" onClick={() => { store.insertTableCol(tid, col, true); setCtxMenu(null); }}>Insert Column Left</button>
                    <button className="menu-item" onClick={() => { store.insertTableCol(tid, col, false); setCtxMenu(null); }}>Insert Column Right</button>
                    <button className="menu-item" onClick={() => { store.deleteTableRow(tid, row); setCtxMenu(null); }}>Delete Row</button>
                    <button className="menu-item" onClick={() => { store.deleteTableCol(tid, col); setCtxMenu(null); }}>Delete Column</button>
                    <div className="menu-div" />
                    {isRange ? (
                      <button className="menu-item" onClick={() => { store.mergeTableCells(tid, sel!.r0, sel!.c0, sel!.r1, sel!.c1); setCtxMenu(null); }}>Merge Cells</button>
                    ) : (
                      <>
                        <button className="menu-item" onClick={() => { store.mergeTableRight(tid, row, col); setCtxMenu(null); }}>Merge with Right</button>
                        <button className="menu-item" onClick={() => { store.mergeTableDown(tid, row, col); setCtxMenu(null); }}>Merge with Below</button>
                      </>
                    )}
                    <button className="menu-item" disabled={!canSplit} onClick={() => { store.splitTableCell(tid, row, col); setCtxMenu(null); }}>Split Cells</button>
                    <div className="menu-div" />
                  </>
                );
              })()}
              {selection.shapeIds.length > 1 && (
                <button className="menu-item" onClick={() => { store.groupSelection(); setCtxMenu(null); }}>Group</button>
              )}
              {store.selectionHasGroup && (
                <button className="menu-item" onClick={() => { store.ungroupSelection(); setCtxMenu(null); }}>Ungroup</button>
              )}
              {slide.shapes.find(s => s.id === ctxMenu.shapeId)?.kind === "pic" && (
                <>
                  <button className="menu-item" onClick={() => { replaceFromFile(ctxMenu.shapeId!); setCtxMenu(null); }}>Replace Image (from file)…</button>
                  <button className="menu-item" onClick={() => { replaceFromUrl(ctxMenu.shapeId!); setCtxMenu(null); }}>Replace Image (from URL)…</button>
                  <div className="menu-div" />
                </>
              )}
              {(selection.shapeIds.length > 1 || store.selectionHasGroup) && <div className="menu-div" />}
              <button className="menu-item" onClick={() => { store.copySelection(); store.deleteSelectedShapes(); setCtxMenu(null); }}>Cut</button>
              <button className="menu-item" onClick={() => { store.copySelection(); setCtxMenu(null); }}>Copy</button>
              <button className="menu-item" disabled={!store.clipboard.length} onClick={() => { store.pasteClipboard(); setCtxMenu(null); }}>Paste</button>
              <button className="menu-item" onClick={() => { store.copySelection(); store.pasteClipboard(); setCtxMenu(null); }}>Duplicate</button>
              <div className="menu-div" />
              <button className="menu-item" onClick={() => { store.reorderShape(ctxMenu.shapeId!, "front"); setCtxMenu(null); }}>Bring to Front</button>
              <button className="menu-item" onClick={() => { store.reorderShape(ctxMenu.shapeId!, "forward"); setCtxMenu(null); }}>Bring Forward</button>
              <button className="menu-item" onClick={() => { store.reorderShape(ctxMenu.shapeId!, "backward"); setCtxMenu(null); }}>Send Backward</button>
              <button className="menu-item" onClick={() => { store.reorderShape(ctxMenu.shapeId!, "back"); setCtxMenu(null); }}>Send to Back</button>
              <div className="menu-div" />
              <button className="menu-item" onClick={() => { store.deleteSelectedShapes(); setCtxMenu(null); }}>Delete</button>
            </>
          ) : (
            <>
              <button className="menu-item" disabled={!store.clipboard.length} onClick={() => { store.pasteClipboard(); setCtxMenu(null); }}>Paste</button>
              <button className="menu-item" onClick={() => { store.selectShapes(slide.shapes.map(s => s.id)); setCtxMenu(null); }}>Select All</button>
              <div className="menu-div" />
              <button className="menu-item" onClick={() => { store.addSlide("titleContent"); setCtxMenu(null); }}>New Slide</button>
            </>
          )}
        </div>
      )}
      <input
        ref={replaceInputRef}
        type="file"
        accept="image/png,image/jpeg,image/gif,image/svg+xml,image/webp,image/bmp"
        style={{ display: "none" }}
        onChange={async e => {
          const f = e.target.files?.[0];
          e.target.value = "";
          const id = replaceTarget.current;
          replaceTarget.current = null;
          if (!f || !id) return;
          try {
            const { mediaId } = await loadImageFile(f);
            store.replacePicMedia(id, mediaId);
          } catch (err) {
            alert("Could not load image: " + (err as Error).message);
          }
        }}
      />
    </div>
  );

  if (!state.showRuler) return canvasBody;
  return (
    <div className="canvas-area">
      <div className="ruler-corner" />
      <Ruler dir="h" lengthPx={containerSize.w} origin={slideOrigin.x} slideLenPx={W * zoom} zoom={zoom} />
      <Ruler dir="v" lengthPx={containerSize.h} origin={slideOrigin.y} slideLenPx={H * zoom} zoom={zoom} />
      {canvasBody}
    </div>
  );
}

/** PowerPoint-style center-zero ruler in centimeters. */
function Ruler({ dir, lengthPx, origin, slideLenPx, zoom }: {
  dir: "h" | "v";
  lengthPx: number;
  origin: number;     // slide edge offset within the viewport, px
  slideLenPx: number;
  zoom: number;
}) {
  const SIZE = 18;
  const pxPerCm = (360000 / 9525) * zoom; // 1cm in EMU -> px at zoom
  const center = origin + slideLenPx / 2;
  const ticks: React.ReactNode[] = [];
  const kMax = Math.ceil(Math.max(center, lengthPx - center) / pxPerCm) + 1;
  for (let k = -kMax; k <= kMax; k++) {
    const pos = center + k * pxPerCm;
    if (pos < -20 || pos > lengthPx + 20) continue;
    const label = Math.abs(k);
    ticks.push(dir === "h" ? (
      <g key={k}>
        <line x1={pos} y1={SIZE - 5} x2={pos} y2={SIZE} stroke="#9A9A9A" strokeWidth={1} />
        <text x={pos + 2.5} y={SIZE - 6.5} fontSize={8} fill="#8A8A8A" fontFamily="Segoe UI, Arial">{label}</text>
      </g>
    ) : (
      <g key={k}>
        <line x1={SIZE - 5} y1={pos} x2={SIZE} y2={pos} stroke="#9A9A9A" strokeWidth={1} />
        <text x={SIZE - 6.5} y={pos + 2.5} fontSize={8} fill="#8A8A8A" fontFamily="Segoe UI, Arial" transform={`rotate(-90 ${SIZE - 6.5} ${pos + 2.5})`} textAnchor="start">{label}</text>
      </g>
    ));
    // half-cm minor tick
    const mid = pos + pxPerCm / 2;
    if (mid > -20 && mid < lengthPx + 20) {
      ticks.push(dir === "h"
        ? <line key={`m${k}`} x1={mid} y1={SIZE - 3} x2={mid} y2={SIZE} stroke="#BDBDBD" strokeWidth={1} />
        : <line key={`m${k}`} x1={SIZE - 3} y1={mid} x2={SIZE} y2={mid} stroke="#BDBDBD" strokeWidth={1} />);
    }
  }
  return (
    <svg
      className={`ruler ruler-${dir}`}
      width={dir === "h" ? lengthPx : SIZE}
      height={dir === "h" ? SIZE : lengthPx}
      shapeRendering="crispEdges"
    >
      {/* slide extent shading */}
      {dir === "h"
        ? <rect x={origin} y={0} width={slideLenPx} height={SIZE} fill="#FFFFFF" />
        : <rect x={0} y={origin} width={SIZE} height={slideLenPx} fill="#FFFFFF" />}
      {ticks}
      {dir === "h"
        ? <line x1={0} y1={SIZE - 0.5} x2={lengthPx} y2={SIZE - 0.5} stroke="#D8D8D8" />
        : <line x1={SIZE - 0.5} y1={0} x2={SIZE - 0.5} y2={lengthPx} stroke="#D8D8D8" />}
    </svg>
  );
}

/** Attaches move/up/dblclick listeners directly to the already-rendered svg element. */
function PointerProxy({ svgRef, onMove, onUp, onDblClick }: {
  svgRef: React.RefObject<SVGSVGElement>;
  onMove: (e: any) => void;
  onUp: (e: any) => void;
  onDblClick: (e: any) => void;
}) {
  useEffect(() => {
    const el = svgRef.current;
    if (!el) return;
    const move = (e: PointerEvent) => onMove(e as any);
    const up = (e: PointerEvent) => onUp(e as any);
    const dbl = (e: MouseEvent) => onDblClick(e as any);
    el.addEventListener("pointermove", move);
    el.addEventListener("pointerup", up);
    el.addEventListener("dblclick", dbl);
    return () => {
      el.removeEventListener("pointermove", move);
      el.removeEventListener("pointerup", up);
      el.removeEventListener("dblclick", dbl);
    };
  });
  return null;
}
