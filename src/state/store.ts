import type { Fill, MediaItem, Paragraph, Presentation, Run, Shape, SlideModel, SlideTransition, TableCell, TableShape, TextAnchor } from "../model/types";
import { makeSlide, newPresentation, nextId, tableCell, tableCellStyle, type LayoutKind } from "../model/defaults";

/** Clamp spans to the grid and recompute which cells are covered by a span. */
function sanitizeSpans(cells: TableCell[][]): TableCell[][] {
  const out = cells.map(row => row.map(cell => ({ ...cell })));
  const R = out.length, C = out[0]?.length ?? 0;
  for (let r = 0; r < R; r++) {
    for (let c = 0; c < C; c++) {
      const cell = out[r][c];
      if (cell.gridSpan) cell.gridSpan = Math.min(cell.gridSpan, C - c);
      if (cell.rowSpan) cell.rowSpan = Math.min(cell.rowSpan, R - r);
      if ((cell.gridSpan ?? 1) <= 1) delete cell.gridSpan;
      if ((cell.rowSpan ?? 1) <= 1) delete cell.rowSpan;
    }
  }
  const covered: ("h" | "v" | null)[][] = Array.from({ length: R }, () => Array(C).fill(null));
  for (let r = 0; r < R; r++) {
    for (let c = 0; c < C; c++) {
      const cell = out[r][c];
      if (cell.merged) continue;
      const gs = cell.gridSpan ?? 1, rs = cell.rowSpan ?? 1;
      for (let rr = r; rr < Math.min(r + rs, R); rr++) {
        for (let cc = c; cc < Math.min(c + gs, C); cc++) {
          if (rr === r && cc === c) continue;
          covered[rr][cc] = rr === r ? "h" : "v";
        }
      }
    }
  }
  for (let r = 0; r < R; r++) {
    for (let c = 0; c < C; c++) {
      const cell = out[r][c];
      if (covered[r][c]) cell.merged = covered[r][c]!;
      else if (cell.merged) delete cell.merged;
    }
  }
  return out;
}

export interface Selection {
  slideIndex: number;
  shapeIds: string[];
}

export type ZoomMode = number | "fit" | "fitw";

export interface EditorState {
  pres: Presentation;
  selection: Selection;
  editingShapeId: string | null;
  /** When editingShapeId is a table: the cell being edited. */
  editingCell: { row: number; col: number } | null;
  /** Rectangular cell-range selection within a table (PowerPoint-style), normalized r0<=r1, c0<=c1. */
  tableSel: { id: string; r0: number; c0: number; r1: number; c1: number } | null;
  /** Chart whose data is being edited in the data dialog. */
  chartEditId: string | null;
  zoom: ZoomMode;
  presenting: boolean;
  presentIndex: number;
  dirty: boolean;
  canUndo: boolean;
  canRedo: boolean;
  statusMessage: string | null;
  /** Armed by the shape gallery / Text Box button: next canvas drag draws this shape. */
  pendingShape: import("../model/types").PresetGeom | "textbox" | null;
  /** Picture in on-canvas crop mode (PowerPoint-style handles). */
  croppingId: string | null;
  /** Resolved zoom factor (canvas publishes it so the status bar can show fit-mode %). */
  effectiveZoom: number;
  /** View-tab toggle: PowerPoint-style center-zero rulers around the canvas. */
  showRuler: boolean;
}

type Listener = () => void;

const MAX_HISTORY = 200;

export class EditorStore {
  private listeners = new Set<Listener>();
  private past: Presentation[] = [];
  private future: Presentation[] = [];
  private previewBase: Presentation | null = null;
  readonly media = new Map<string, MediaItem>();
  clipboard: Shape[] = [];

  state: EditorState = {
    pres: newPresentation(),
    selection: { slideIndex: 0, shapeIds: [] },
    editingShapeId: null,
    editingCell: null,
    tableSel: null,
    chartEditId: null,
    zoom: "fit",
    presenting: false,
    presentIndex: 0,
    dirty: false,
    canUndo: false,
    canRedo: false,
    statusMessage: null,
    pendingShape: null,
    croppingId: null,
    effectiveZoom: 1,
    showRuler: true,
  };

  subscribe = (fn: Listener) => {
    this.listeners.add(fn);
    return () => this.listeners.delete(fn);
  };

  getState = () => this.state;

  private emit(partial: Partial<EditorState>) {
    this.state = {
      ...this.state,
      ...partial,
      canUndo: this.past.length > 0,
      canRedo: this.future.length > 0,
    };
    this.listeners.forEach(l => l());
  }

  setState(partial: Partial<EditorState>) {
    this.emit(partial);
  }

  // ---------- history ----------
  /** Commit a new presentation state (undoable). */
  commit(next: Presentation, extra: Partial<EditorState> = {}) {
    this.past.push(this.state.pres);
    if (this.past.length > MAX_HISTORY) this.past.shift();
    this.future = [];
    this.emit({ pres: next, dirty: true, ...extra });
  }

  /** Live update during drag/resize — no history entry until endPreview. */
  preview(next: Presentation) {
    if (!this.previewBase) this.previewBase = this.state.pres;
    this.emit({ pres: next });
  }

  endPreview(keep = true) {
    const base = this.previewBase;
    this.previewBase = null;
    if (!base) return;
    if (keep && base !== this.state.pres) {
      this.past.push(base);
      if (this.past.length > MAX_HISTORY) this.past.shift();
      this.future = [];
      this.emit({ dirty: true });
    } else if (!keep) {
      this.emit({ pres: base });
    }
  }

  undo = () => {
    const prev = this.past.pop();
    if (!prev) return;
    this.future.push(this.state.pres);
    const sel = this.clampSelection(prev, this.state.selection);
    this.emit({ pres: prev, selection: sel, editingShapeId: null, editingCell: null, dirty: true });
  };

  redo = () => {
    const next = this.future.pop();
    if (!next) return;
    this.past.push(this.state.pres);
    const sel = this.clampSelection(next, this.state.selection);
    this.emit({ pres: next, selection: sel, editingShapeId: null, editingCell: null, dirty: true });
  };

  private clampSelection(pres: Presentation, sel: Selection): Selection {
    const slideIndex = Math.min(sel.slideIndex, pres.slides.length - 1);
    const slide = pres.slides[slideIndex];
    const ids = new Set(slide.shapes.map(s => s.id));
    return { slideIndex, shapeIds: sel.shapeIds.filter(id => ids.has(id)) };
  }

  // ---------- convenience accessors ----------
  get pres() { return this.state.pres; }
  get currentSlide(): SlideModel { return this.state.pres.slides[this.state.selection.slideIndex]; }
  get selectedShapes(): Shape[] {
    const slide = this.currentSlide;
    return slide ? slide.shapes.filter(s => this.state.selection.shapeIds.includes(s.id)) : [];
  }

  // ---------- slide ops ----------
  selectSlide(i: number) {
    const idx = Math.max(0, Math.min(i, this.pres.slides.length - 1));
    this.emit({ selection: { slideIndex: idx, shapeIds: [] }, editingShapeId: null });
  }

  addSlide(layout: LayoutKind = "titleContent", after?: number) {
    const at = (after ?? this.state.selection.slideIndex) + 1;
    const slides = [...this.pres.slides];
    slides.splice(at, 0, makeSlide(layout));
    this.commit({ ...this.pres, slides }, { selection: { slideIndex: at, shapeIds: [] }, editingShapeId: null });
  }

  duplicateSlide(i?: number) {
    const idx = i ?? this.state.selection.slideIndex;
    const src = this.pres.slides[idx];
    const copy: SlideModel = {
      id: nextId("slide"),
      background: src.background,
      shapes: src.shapes.map(s => ({ ...structuredClone(s), id: nextId(s.kind) })),
    };
    const slides = [...this.pres.slides];
    slides.splice(idx + 1, 0, copy);
    this.commit({ ...this.pres, slides }, { selection: { slideIndex: idx + 1, shapeIds: [] } });
  }

  deleteSlide(i?: number) {
    const idx = i ?? this.state.selection.slideIndex;
    if (this.pres.slides.length <= 1) {
      // never leave zero slides — replace with a blank one (PowerPoint behavior)
      this.commit({ ...this.pres, slides: [makeSlide("blank")] }, { selection: { slideIndex: 0, shapeIds: [] } });
      return;
    }
    const slides = this.pres.slides.filter((_, k) => k !== idx);
    this.commit({ ...this.pres, slides }, {
      selection: { slideIndex: Math.min(idx, slides.length - 1), shapeIds: [] },
      editingShapeId: null,
    });
  }

  moveSlide(from: number, to: number) {
    if (from === to || to < 0 || to >= this.pres.slides.length) return;
    const slides = [...this.pres.slides];
    const [s] = slides.splice(from, 1);
    slides.splice(to, 0, s);
    this.commit({ ...this.pres, slides }, { selection: { slideIndex: to, shapeIds: [] } });
  }

  // ---------- shape ops ----------
  private withSlide(idx: number, fn: (slide: SlideModel) => SlideModel): Presentation {
    const slides = this.pres.slides.map((s, k) => (k === idx ? fn(s) : s));
    return { ...this.pres, slides };
  }

  selectShapes(ids: string[], additive = false) {
    const cur = this.state.selection;
    const shapeIds = additive ? Array.from(new Set([...cur.shapeIds, ...ids])) : ids;
    this.emit({ selection: { ...cur, shapeIds } });
  }

  addShape(shape: Shape, select = true) {
    const idx = this.state.selection.slideIndex;
    const next = this.withSlide(idx, sl => ({ ...sl, shapes: [...sl.shapes, shape] }));
    this.commit(next, select ? { selection: { slideIndex: idx, shapeIds: [shape.id] } } : {});
  }

  updateShapes(ids: string[], fn: (s: Shape) => Shape, opts: { historic?: boolean } = { historic: true }) {
    const idx = this.state.selection.slideIndex;
    const next = this.withSlide(idx, sl => ({
      ...sl,
      shapes: sl.shapes.map(s => (ids.includes(s.id) ? fn(s) : s)),
    }));
    if (opts.historic === false) this.preview(next);
    else this.commit(next);
  }

  deleteSelectedShapes() {
    const { slideIndex, shapeIds } = this.state.selection;
    if (!shapeIds.length) return;
    const next = this.withSlide(slideIndex, sl => ({
      ...sl,
      shapes: sl.shapes.filter(s => !shapeIds.includes(s.id)),
    }));
    this.commit(next, { selection: { slideIndex, shapeIds: [] }, editingShapeId: null });
  }

  reorderShape(id: string, dir: "front" | "back" | "forward" | "backward") {
    const idx = this.state.selection.slideIndex;
    const next = this.withSlide(idx, sl => {
      const shapes = [...sl.shapes];
      const i = shapes.findIndex(s => s.id === id);
      if (i < 0) return sl;
      const [sh] = shapes.splice(i, 1);
      let at = dir === "front" ? shapes.length : dir === "back" ? 0 : dir === "forward" ? Math.min(i + 1, shapes.length) : Math.max(i - 1, 0);
      shapes.splice(at, 0, sh);
      return { ...sl, shapes };
    });
    this.commit(next);
  }

  /** Align selection: single shape -> against the slide; multiple -> against their bounding box. */
  alignSelected(mode: "l" | "c" | "r" | "t" | "m" | "b") {
    const sel = this.selectedShapes;
    if (!sel.length) return;
    const bounds = sel.length === 1
      ? { x1: 0, y1: 0, x2: this.pres.slideWidth, y2: this.pres.slideHeight }
      : {
          x1: Math.min(...sel.map(s => s.x)),
          y1: Math.min(...sel.map(s => s.y)),
          x2: Math.max(...sel.map(s => s.x + s.w)),
          y2: Math.max(...sel.map(s => s.y + s.h)),
        };
    this.updateShapes(sel.map(s => s.id), s => {
      switch (mode) {
        case "l": return { ...s, x: bounds.x1 };
        case "c": return { ...s, x: Math.round((bounds.x1 + bounds.x2) / 2 - s.w / 2) };
        case "r": return { ...s, x: bounds.x2 - s.w };
        case "t": return { ...s, y: bounds.y1 };
        case "m": return { ...s, y: Math.round((bounds.y1 + bounds.y2) / 2 - s.h / 2) };
        case "b": return { ...s, y: bounds.y2 - s.h };
      }
    });
  }

  /** Even gaps between 3+ selected shapes along an axis. */
  distributeSelected(axis: "h" | "v") {
    const sel = [...this.selectedShapes];
    if (sel.length < 3) return;
    sel.sort((a, b) => (axis === "h" ? a.x - b.x : a.y - b.y));
    const first = sel[0], last = sel[sel.length - 1];
    const targets = new Map<string, number>();
    if (axis === "h") {
      const span = last.x + last.w - first.x;
      const totalW = sel.reduce((a, s) => a + s.w, 0);
      const gap = (span - totalW) / (sel.length - 1);
      let at = first.x;
      for (const s of sel) { targets.set(s.id, Math.round(at)); at += s.w + gap; }
      this.updateShapes(sel.map(s => s.id), s => ({ ...s, x: targets.get(s.id) ?? s.x }));
    } else {
      const span = last.y + last.h - first.y;
      const totalH = sel.reduce((a, s) => a + s.h, 0);
      const gap = (span - totalH) / (sel.length - 1);
      let at = first.y;
      for (const s of sel) { targets.set(s.id, Math.round(at)); at += s.h + gap; }
      this.updateShapes(sel.map(s => s.id), s => ({ ...s, y: targets.get(s.id) ?? s.y }));
    }
  }

  copySelection() {
    this.clipboard = this.selectedShapes.map(s => structuredClone(s));
  }

  pasteClipboard() {
    if (!this.clipboard.length) return;
    const idx = this.state.selection.slideIndex;
    const off = 457200; // 0.5cm-ish offset
    // pasted groups stay grouped, but as a NEW group independent of the source
    const gidMap = new Map<string, string>();
    const copies = this.clipboard.map(s => {
      const copy = { ...structuredClone(s), id: nextId(s.kind), x: s.x + off, y: s.y + off };
      if (copy.groupId) {
        if (!gidMap.has(copy.groupId)) gidMap.set(copy.groupId, nextId("grp"));
        copy.groupId = gidMap.get(copy.groupId)!;
      }
      return copy;
    });
    this.clipboard = copies.map(s => structuredClone(s)); // next paste offsets further
    const next = this.withSlide(idx, sl => ({ ...sl, shapes: [...sl.shapes, ...copies] }));
    this.commit(next, { selection: { slideIndex: idx, shapeIds: copies.map(c => c.id) } });
  }

  // ---------- grouping ----------
  /** All shape ids on the current slide expanded to whole groups. */
  expandToGroups(ids: string[]): string[] {
    const slide = this.currentSlide;
    if (!slide) return ids;
    const gids = new Set(slide.shapes.filter(s => ids.includes(s.id) && s.groupId).map(s => s.groupId!));
    if (!gids.size) return ids;
    const out = new Set(ids);
    for (const s of slide.shapes) if (s.groupId && gids.has(s.groupId)) out.add(s.id);
    return [...out];
  }

  /** Group the selected shapes (≥2): shared groupId + members pulled adjacent in z-order. */
  groupSelection() {
    const { slideIndex, shapeIds } = this.state.selection;
    if (shapeIds.length < 2) return;
    const gid = nextId("grp");
    const next = this.withSlide(slideIndex, sl => {
      const members = sl.shapes.filter(s => shapeIds.includes(s.id)).map(s => ({ ...s, groupId: gid }));
      // members become one contiguous z-block at the topmost member's position (like PowerPoint)
      const topIdx = Math.max(...shapeIds.map(id => sl.shapes.findIndex(s => s.id === id)));
      const rest = sl.shapes.map((s, i) => ({ s, i })).filter(r => !shapeIds.includes(r.s.id));
      const lower = rest.filter(r => r.i < topIdx).map(r => r.s);
      const upper = rest.filter(r => r.i >= topIdx).map(r => r.s);
      return { ...sl, shapes: [...lower, ...members, ...upper] };
    });
    this.commit(next, { selection: { slideIndex, shapeIds } });
    this.setStatus("Grouped");
    setTimeout(() => this.setStatus(null), 1500);
  }

  /** Remove group membership from every selected shape. */
  ungroupSelection() {
    const { slideIndex, shapeIds } = this.state.selection;
    if (!shapeIds.length) return;
    const next = this.withSlide(slideIndex, sl => ({
      ...sl,
      shapes: sl.shapes.map(s => (shapeIds.includes(s.id) && s.groupId ? { ...s, groupId: undefined } : s)),
    }));
    this.commit(next, { selection: { slideIndex, shapeIds } });
    this.setStatus("Ungrouped");
    setTimeout(() => this.setStatus(null), 1500);
  }

  /** True when the current selection contains at least one grouped shape. */
  get selectionHasGroup(): boolean {
    return this.selectedShapes.some(s => s.groupId);
  }

  /** Rotation-aware bounding box of a shape in EMU (mirrors the canvas AABB). */
  private static rotAabb(s: Shape): { x1: number; y1: number; x2: number; y2: number } {
    const cx = s.x + s.w / 2, cy = s.y + s.h / 2;
    const rad = (s.rot * Math.PI) / 180, c = Math.cos(rad), si = Math.sin(rad);
    const xs: number[] = [], ys: number[] = [];
    for (const [dx, dy] of [[-s.w / 2, -s.h / 2], [s.w / 2, -s.h / 2], [s.w / 2, s.h / 2], [-s.w / 2, s.h / 2]]) {
      xs.push(cx + dx * c - dy * si);
      ys.push(cy + dx * si + dy * c);
    }
    return { x1: Math.min(...xs), y1: Math.min(...ys), x2: Math.max(...xs), y2: Math.max(...ys) };
  }

  private selectionBounds(): { cx: number; cy: number } | null {
    const shapes = this.selectedShapes;
    if (!shapes.length) return null;
    let x1 = Infinity, y1 = Infinity, x2 = -Infinity, y2 = -Infinity;
    for (const s of shapes) {
      const b = EditorStore.rotAabb(s);
      x1 = Math.min(x1, b.x1); y1 = Math.min(y1, b.y1);
      x2 = Math.max(x2, b.x2); y2 = Math.max(y2, b.y2);
    }
    return { cx: (x1 + x2) / 2, cy: (y1 + y2) / 2 };
  }

  /** Rotate the whole selection by delta degrees around its combined center (group rotate). */
  rotateSelected(delta: number) {
    const c = this.selectionBounds();
    if (!c) return;
    const single = this.state.selection.shapeIds.length === 1;
    const rad = (delta * Math.PI) / 180, cos = Math.cos(rad), sin = Math.sin(rad);
    this.updateShapes(this.state.selection.shapeIds, s => {
      const ocx = s.x + s.w / 2, ocy = s.y + s.h / 2;
      const dx = ocx - c.cx, dy = ocy - c.cy;
      const ncx = single ? ocx : c.cx + dx * cos - dy * sin;
      const ncy = single ? ocy : c.cy + dx * sin + dy * cos;
      return {
        ...s,
        x: Math.round(ncx - s.w / 2), y: Math.round(ncy - s.h / 2),
        rot: ((s.rot + delta) % 360 + 360) % 360,
      };
    });
  }

  /** Mirror the whole selection about its combined center axis (group flip). */
  flipSelected(axis: "h" | "v") {
    const c = this.selectionBounds();
    if (!c) return;
    this.updateShapes(this.state.selection.shapeIds, s => ({
      ...s,
      x: axis === "h" ? Math.round(2 * c.cx - (s.x + s.w)) : s.x,
      y: axis === "v" ? Math.round(2 * c.cy - (s.y + s.h)) : s.y,
      flipH: axis === "h" ? (!s.flipH || undefined) : s.flipH,
      flipV: axis === "v" ? (!s.flipV || undefined) : s.flipV,
      rot: ((-s.rot) % 360 + 360) % 360,
    }));
  }

  /** Apply run-level formatting to every run in the selected shapes (or a specific shape id list). */
  formatRuns(fn: (r: Run) => Run, ids?: string[]) {
    const targets = ids ?? this.state.selection.shapeIds;
    if (!targets.length) return;
    this.updateShapes(targets, s => {
      if (s.kind === "table") {
        return {
          ...s,
          cells: s.cells.map(row => row.map(c => ({
            ...c,
            text: { ...c.text, paragraphs: c.text.paragraphs.map(p => ({ ...p, runs: p.runs.map(fn) })) },
          }))),
        };
      }
      if (s.kind !== "sp" || !s.text) return s;
      return {
        ...s,
        text: {
          ...s.text,
          paragraphs: s.text.paragraphs.map(p => ({ ...p, runs: p.runs.map(fn) })),
        },
      };
    });
  }

  formatParagraphs(fn: (p: Paragraph) => Paragraph, ids?: string[]) {
    const targets = ids ?? this.state.selection.shapeIds;
    if (!targets.length) return;
    this.updateShapes(targets, s => {
      if (s.kind === "table") {
        return {
          ...s,
          cells: s.cells.map(row => row.map(c => ({
            ...c,
            text: { ...c.text, paragraphs: c.text.paragraphs.map(fn) },
          }))),
        };
      }
      if (s.kind !== "sp" || !s.text) return s;
      return { ...s, text: { ...s.text, paragraphs: s.text.paragraphs.map(fn) } };
    });
  }

  /** Apply a length-preserving text transform (case changes) per paragraph, keeping run boundaries. */
  transformText(fn: (s: string) => string, ids?: string[]) {
    const mapPara = (p: Paragraph): Paragraph => {
      const full = p.runs.map(r => r.text).join("");
      if (!full) return p;
      const out = fn(full);
      if (out.length !== full.length) {
        return { ...p, runs: [{ ...p.runs[0], text: out }, ...p.runs.slice(1).map(r => ({ ...r, text: "" }))] };
      }
      let at = 0;
      return {
        ...p,
        runs: p.runs.map(r => {
          const t = out.slice(at, at + r.text.length);
          at += r.text.length;
          return { ...r, text: t };
        }),
      };
    };
    this.formatParagraphs(mapPara, ids);
  }

  // ---------- table ops ----------
  modifyTable(id: string, fn: (t: import("../model/types").TableShape) => import("../model/types").TableShape) {
    this.updateShapes([id], s => (s.kind === "table" ? fn(s) : s));
  }

  /** Styled empty cell matching the table style at (r,c). */
  private styledCell(t: TableShape, r: number, c: number): TableCell {
    const st = tableCellStyle(t, r, c);
    return tableCell("", { bold: st.bold || undefined, color: st.textColor });
  }

  insertTableRow(id: string, at: number, before: boolean) {
    this.modifyTable(id, t => {
      const idx = Math.max(0, Math.min(before ? at : at + 1, t.cells.length));
      const nCols = t.cells[0]?.length ?? 1;
      const cells = t.cells.map(row => row.slice());
      cells.splice(idx, 0, Array.from({ length: nCols }, (_, c) => this.styledCell(t, idx, c)));
      const rowH = t.rowH.slice();
      rowH.splice(idx, 0, t.rowH[Math.min(at, t.rowH.length - 1)] ?? 370840);
      return { ...t, cells: sanitizeSpans(cells), rowH, h: Math.round(t.h * (cells.length) / Math.max(1, t.cells.length)) };
    });
  }

  insertTableCol(id: string, at: number, before: boolean) {
    this.modifyTable(id, t => {
      const idx = Math.max(0, Math.min(before ? at : at + 1, t.cells[0]?.length ?? 0));
      const cells = t.cells.map((row, r) => {
        const next = row.slice();
        next.splice(idx, 0, this.styledCell(t, r, idx));
        return next;
      });
      const colW = t.colW.slice();
      colW.splice(idx, 0, t.colW[Math.min(at, t.colW.length - 1)] ?? 914400);
      return { ...t, cells: sanitizeSpans(cells), colW };
    });
  }

  deleteTableRow(id: string, at: number) {
    this.modifyTable(id, t => {
      if (t.cells.length <= 1) return t;
      const cells = t.cells.filter((_, r) => r !== at);
      const rowH = t.rowH.filter((_, r) => r !== at);
      return { ...t, cells: sanitizeSpans(cells), rowH, h: Math.round(t.h * cells.length / t.cells.length) };
    });
  }

  deleteTableCol(id: string, at: number) {
    this.modifyTable(id, t => {
      if ((t.cells[0]?.length ?? 0) <= 1) return t;
      const cells = t.cells.map(row => row.filter((_, c) => c !== at));
      const colW = t.colW.filter((_, c) => c !== at);
      return { ...t, cells: sanitizeSpans(cells), colW };
    });
  }

  distributeTableRows(id: string) {
    this.modifyTable(id, t => {
      const avg = Math.max(1, Math.round(t.rowH.reduce((a, b) => a + b, 0) / Math.max(1, t.rowH.length)));
      return { ...t, rowH: t.rowH.map(() => avg) };
    });
  }

  distributeTableCols(id: string) {
    this.modifyTable(id, t => {
      const avg = Math.max(1, Math.round(t.colW.reduce((a, b) => a + b, 0) / Math.max(1, t.colW.length)));
      return { ...t, colW: t.colW.map(() => avg) };
    });
  }

  /** Merge the cell at (r,c) with its right neighbor (single-row-height cells only). */
  mergeTableRight(id: string, r: number, c: number) {
    this.modifyTable(id, t => {
      const cell = t.cells[r]?.[c];
      const span = Math.max(1, cell?.gridSpan ?? 1);
      const next = t.cells[r]?.[c + span];
      if (!cell || cell.merged || !next || next.merged) return t;
      if ((cell.rowSpan ?? 1) > 1 || (next.rowSpan ?? 1) > 1) return t;
      const cells = t.cells.map(row => row.map(x => ({ ...x })));
      const extraParas = next.text.paragraphs.filter(p => p.runs.some(run => run.text));
      cells[r][c] = {
        ...cell,
        gridSpan: span + Math.max(1, next.gridSpan ?? 1),
        text: extraParas.length ? { ...cell.text, paragraphs: [...cell.text.paragraphs, ...extraParas] } : cell.text,
      };
      cells[r][c + span] = { ...tableCell(""), merged: "h" };
      return { ...t, cells: sanitizeSpans(cells) };
    });
  }

  /** Merge the cell at (r,c) with the cell below (single-column-width cells only). */
  mergeTableDown(id: string, r: number, c: number) {
    this.modifyTable(id, t => {
      const cell = t.cells[r]?.[c];
      const span = Math.max(1, cell?.rowSpan ?? 1);
      const below = t.cells[r + span]?.[c];
      if (!cell || cell.merged || !below || below.merged) return t;
      if ((cell.gridSpan ?? 1) > 1 || (below.gridSpan ?? 1) > 1) return t;
      const cells = t.cells.map(row => row.map(x => ({ ...x })));
      const extraParas = below.text.paragraphs.filter(p => p.runs.some(run => run.text));
      cells[r][c] = {
        ...cell,
        rowSpan: span + Math.max(1, below.rowSpan ?? 1),
        text: extraParas.length ? { ...cell.text, paragraphs: [...cell.text.paragraphs, ...extraParas] } : cell.text,
      };
      cells[r + span][c] = { ...tableCell(""), merged: "v" };
      return { ...t, cells: sanitizeSpans(cells) };
    });
  }

  /** Merge a rectangular range of cells into the top-left anchor (PowerPoint "Merge Cells"). */
  mergeTableCells(id: string, r0: number, c0: number, r1: number, c1: number) {
    this.modifyTable(id, t => {
      const R = t.cells.length, C = t.cells[0]?.length ?? 0;
      r0 = Math.max(0, Math.min(r0, R - 1)); r1 = Math.max(0, Math.min(r1, R - 1));
      c0 = Math.max(0, Math.min(c0, C - 1)); c1 = Math.max(0, Math.min(c1, C - 1));
      if (r0 > r1) [r0, r1] = [r1, r0];
      if (c0 > c1) [c0, c1] = [c1, c0];
      // grow the rectangle to fully cover any span that pokes out of it
      let grew = true;
      while (grew) {
        grew = false;
        for (let r = r0; r <= r1; r++) for (let c = c0; c <= c1; c++) {
          const cell = t.cells[r]?.[c];
          if (!cell || cell.merged) continue;
          const er = r + (cell.rowSpan ?? 1) - 1, ec = c + (cell.gridSpan ?? 1) - 1;
          if (er > r1) { r1 = er; grew = true; }
          if (ec > c1) { c1 = ec; grew = true; }
        }
      }
      if (r0 === r1 && c0 === c1) return t;
      const cells = t.cells.map(row => row.map(x => ({ ...x })));
      const extraParas = [];
      for (let r = r0; r <= r1; r++) for (let c = c0; c <= c1; c++) {
        if (r === r0 && c === c0) continue;
        const cell = cells[r][c];
        if (!cell.merged) extraParas.push(...cell.text.paragraphs.filter(p => p.runs.some(run => run.text)));
        cells[r][c] = { ...tableCell("") };
      }
      const anchor = cells[r0][c0];
      cells[r0][c0] = {
        ...anchor,
        gridSpan: c1 - c0 + 1,
        rowSpan: r1 - r0 + 1,
        text: extraParas.length ? { ...anchor.text, paragraphs: [...anchor.text.paragraphs, ...extraParas] } : anchor.text,
      };
      return { ...t, cells: sanitizeSpans(cells) };
    });
    this.setState({ tableSel: { id, r0, c0, r1: r0, c1: c0 } });
  }

  /** Undo a merge: clear the spans on the anchor at (r,c) and free the covered cells. */
  splitTableCell(id: string, r: number, c: number) {
    this.modifyTable(id, t => {
      const cell = t.cells[r]?.[c];
      if (!cell || ((cell.gridSpan ?? 1) <= 1 && (cell.rowSpan ?? 1) <= 1)) return t;
      const cells = t.cells.map(row => row.map(x => ({ ...x })));
      const gs = cell.gridSpan ?? 1, rs = cell.rowSpan ?? 1;
      for (let rr = r; rr < Math.min(r + rs, cells.length); rr++) {
        for (let cc = c; cc < Math.min(c + gs, cells[rr].length); cc++) {
          if (rr === r && cc === c) continue;
          cells[rr][cc] = this.styledCell(t, rr, cc);
        }
      }
      cells[r][c] = { ...cell, gridSpan: undefined, rowSpan: undefined };
      return { ...t, cells: sanitizeSpans(cells) };
    });
  }

  /**
   * Apply a built-in style (family/accent/flags): like PowerPoint, this resets
   * cell shading and recolors text to the style's colors.
   */
  applyTableStyle(id: string, patch: Partial<TableShape>) {
    this.modifyTable(id, t => {
      const next = { ...t, ...patch };
      const cells = next.cells.map((row, r) => row.map((cell, c) => {
        if (cell.merged) return cell;
        const st = tableCellStyle(next, r, c);
        return {
          ...cell,
          fill: undefined,
          text: {
            ...cell.text,
            paragraphs: cell.text.paragraphs.map(p => ({
              ...p,
              runs: p.runs.map(run => ({ ...run, color: st.textColor, bold: st.bold || undefined })),
            })),
          },
        };
      }));
      return { ...next, cells };
    });
  }

  setTableCellFill(id: string, r: number, c: number, fill: Fill | undefined) {
    this.modifyTable(id, t => ({
      ...t,
      cells: t.cells.map((row, rr) => row.map((cell, cc) => (rr === r && cc === c ? { ...cell, fill } : cell))),
    }));
  }

  /** Apply a fill to every cell in a rectangular range (undefined = back to style). */
  setCellRangeFill(id: string, r0: number, c0: number, r1: number, c1: number, fill: Fill | undefined) {
    this.modifyTable(id, t => ({
      ...t,
      cells: t.cells.map((row, rr) => row.map((cell, cc) =>
        rr >= r0 && rr <= r1 && cc >= c0 && cc <= c1 ? { ...cell, fill } : cell)),
    }));
  }

  /** Vertical alignment for a rectangular range of cells. */
  setCellRangeAnchor(id: string, r0: number, c0: number, r1: number, c1: number, anchor: TextAnchor) {
    this.modifyTable(id, t => ({
      ...t,
      cells: t.cells.map((row, rr) => row.map((cell, cc) =>
        rr >= r0 && rr <= r1 && cc >= c0 && cc <= c1 ? { ...cell, text: { ...cell.text, anchor } } : cell)),
    }));
  }

  /** Vertical alignment for every cell (table-wide, like the Layout tab buttons). */
  setTableCellsAnchor(id: string, anchor: TextAnchor) {
    this.modifyTable(id, t => ({
      ...t,
      cells: t.cells.map(row => row.map(cell => ({ ...cell, text: { ...cell.text, anchor } }))),
    }));
  }

  /** Cell margins for every cell ([l,t,r,b] EMU). */
  setTableCellsInsets(id: string, insets: [number, number, number, number]) {
    this.modifyTable(id, t => ({
      ...t,
      cells: t.cells.map(row => row.map(cell => ({ ...cell, text: { ...cell.text, insets } }))),
    }));
  }

  // ---------- find & replace ----------
  /** All shapes (across slides) whose text contains the query; tables search cell text. */
  findMatches(query: string, matchCase: boolean): { slideIndex: number; shapeId: string; hits: number }[] {
    if (!query) return [];
    const q = matchCase ? query : query.toLowerCase();
    const out: { slideIndex: number; shapeId: string; hits: number }[] = [];
    const countIn = (texts: string[]): number => {
      let n = 0;
      for (const t of texts) {
        const s = matchCase ? t : t.toLowerCase();
        let at = s.indexOf(q);
        while (at >= 0) { n++; at = s.indexOf(q, at + q.length); }
      }
      return n;
    };
    this.pres.slides.forEach((slide, si) => {
      for (const sh of slide.shapes) {
        let hits = 0;
        if (sh.kind === "sp" && sh.text) {
          hits = countIn(sh.text.paragraphs.map(p => p.runs.map(r => r.text).join("")));
        } else if (sh.kind === "table") {
          hits = countIn(sh.cells.flat().flatMap(c => c.text.paragraphs.map(p => p.runs.map(r => r.text).join(""))));
        }
        if (hits > 0) out.push({ slideIndex: si, shapeId: sh.id, hits });
      }
    });
    return out;
  }

  /** Replace within runs across ALL slides (matches spanning run boundaries are left alone). */
  replaceAll(query: string, replacement: string, matchCase: boolean): number {
    if (!query) return 0;
    let count = 0;
    const rx = new RegExp(query.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"), matchCase ? "g" : "gi");
    const mapRun = (r: Run): Run => {
      const next = r.text.replace(rx, () => { count++; return replacement; });
      return next === r.text ? r : { ...r, text: next };
    };
    const slides = this.pres.slides.map(slide => ({
      ...slide,
      shapes: slide.shapes.map(sh => {
        if (sh.kind === "sp" && sh.text) {
          return { ...sh, text: { ...sh.text, paragraphs: sh.text.paragraphs.map(p => ({ ...p, runs: p.runs.map(mapRun) })) } };
        }
        if (sh.kind === "table") {
          return {
            ...sh,
            cells: sh.cells.map(row => row.map(c => ({
              ...c,
              text: { ...c.text, paragraphs: c.text.paragraphs.map(p => ({ ...p, runs: p.runs.map(mapRun) })) },
            }))),
          };
        }
        return sh;
      }),
    }));
    if (count > 0) this.commit({ ...this.pres, slides });
    return count;
  }

  setStatus(msg: string | null) {
    this.emit({ statusMessage: msg });
  }

  // ---------- slide-level properties ----------
  updateSlideBg(bg: Fill | undefined, opts: { historic?: boolean } = {}) {
    const idx = this.state.selection.slideIndex;
    const next = this.withSlide(idx, sl => ({ ...sl, background: bg }));
    if (opts.historic === false) this.preview(next);
    else this.commit(next);
  }

  applyBgToAll() {
    const bg = this.currentSlide?.background;
    this.commit({ ...this.pres, slides: this.pres.slides.map(sl => ({ ...sl, background: bg })) });
  }

  setTransition(t: SlideTransition | undefined) {
    const idx = this.state.selection.slideIndex;
    this.commit(this.withSlide(idx, sl => ({ ...sl, transition: t })));
  }

  applyTransitionToAll() {
    const t = this.currentSlide?.transition;
    this.commit({ ...this.pres, slides: this.pres.slides.map(sl => ({ ...sl, transition: t })) });
  }

  setSlideSize(w: number, h: number) {
    this.commit({ ...this.pres, slideWidth: w, slideHeight: h });
  }

  /** Insert a parsed slide (e.g. pattern import) after the current one, merging its media. */
  addImportedSlide(slide: SlideModel, media: Map<string, MediaItem>) {
    media.forEach((v, k) => this.media.set(k, v));
    const at = this.state.selection.slideIndex + 1;
    const slides = [...this.pres.slides];
    slides.splice(at, 0, slide);
    this.commit({ ...this.pres, slides }, { selection: { slideIndex: at, shapeIds: [] }, editingShapeId: null });
  }

  newPresentation() {
    this.past = [];
    this.future = [];
    this.emit({
      pres: newPresentation(),
      selection: { slideIndex: 0, shapeIds: [] },
      editingShapeId: null,
      dirty: false,
      zoom: "fit",
      pendingShape: null,
    });
  }

  loadPresentation(pres: Presentation, media: Map<string, MediaItem>) {
    this.past = [];
    this.future = [];
    this.media.clear();
    media.forEach((v, k) => this.media.set(k, v));
    this.emit({
      pres,
      selection: { slideIndex: 0, shapeIds: [] },
      editingShapeId: null,
      dirty: false,
      zoom: "fit",
    });
  }
}

export const store = new EditorStore();

// dev console access
if (typeof window !== "undefined") (window as unknown as { __store: EditorStore }).__store = store;
