import React, { useEffect, useRef, useState } from "react";
import { ChartDataDialog } from "./components/ChartDataDialog";
import { EditorCanvas } from "./components/EditorCanvas";
import { FindDialog } from "./components/FindDialog";
import { NotesBar } from "./components/NotesBar";
import { PatternDialog } from "./components/PatternDialog";
import { PresentMode } from "./components/PresentMode";
import { Ribbon } from "./components/Ribbon";
import { RightPanel } from "./components/RightPanel";
import { SlidePanel } from "./components/SlidePanel";
import { StatusBar } from "./components/StatusBar";
import { Icon } from "./components/icons";
import { EMU_PER_PX, nextId } from "./model/defaults";
import { exportPptxBlob } from "./ooxml/write";
import { parsePptx } from "./ooxml/parse";
import { importPatternSlide } from "./ooxml/pattern";
import { store } from "./state/store";
import { useEditorState } from "./state/useStore";
import { saveSnapshot } from "./util/autosave";
import { initCustomFonts } from "./util/custom";
import { embedConfig, initEmbedBridge, notifyHost, uploadTo } from "./util/embed";
import { loadImageFile } from "./util/loadImage";

// host-integration settings, fixed for the lifetime of the page
const EMBED = embedConfig();

export default function App() {
  const state = useEditorState();
  const openInputRef = useRef<HTMLInputElement>(null);
  const [patternOpen, setPatternOpen] = useState(false);
  const [showThumbs, setShowThumbs] = useState(true);
  const [findOpen, setFindOpen] = useState(false);

  const chartShape = state.chartEditId
    ? store.currentSlide?.shapes.find(s => s.id === state.chartEditId && s.kind === "chart")
    : undefined;

  // ---------- autosave (debounced snapshot after every committed change) ----------
  useEffect(() => {
    let timer: number | undefined;
    const unsub = store.subscribe(() => {
      if (store.state.presenting) return;
      window.clearTimeout(timer);
      timer = window.setTimeout(() => saveSnapshot(store.pres, store.media), 1500);
    });
    return () => { unsub(); window.clearTimeout(timer); };
  }, []);

  // re-load Google fonts for saved custom font pairs
  useEffect(() => { initCustomFonts(); }, []);

  // dev/demo hook: window.__importPattern(jsonText, values?)
  useEffect(() => {
    (window as unknown as { __importPattern: unknown }).__importPattern =
      async (json: string, values?: Record<string, string>) => {
        const r = await importPatternSlide(json, values);
        store.addImportedSlide(r.slide, r.media);
        return { warnings: r.warnings, values: r.values, shapes: r.slide.shapes.length };
      };
  }, []);

  const onSave = async () => {
    try {
      store.setStatus("Saving…");
      const blob = await exportPptxBlob(store.pres, store.media);
      if (EMBED.saveUrl) {
        // host integration: upload back to the presigned URL (R2/S3)
        await uploadTo(EMBED.saveUrl, blob);
        store.setState({ dirty: false });
        store.setStatus("Saved to storage");
        notifyHost({ type: "pe:saved", via: "upload" });
      } else if (EMBED.embed && EMBED.parentOrigin) {
        // host integration: hand the document to the embedding page
        const data = await blob.arrayBuffer();
        notifyHost({ type: "pe:document", data, fileName: `${store.pres.title || "Presentation"}.pptx` }, [data]);
        notifyHost({ type: "pe:saved", via: "message" });
        store.setState({ dirty: false });
        store.setStatus("Saved");
      } else {
        const a = document.createElement("a");
        a.href = URL.createObjectURL(blob);
        a.download = `${store.pres.title || "Presentation"}.pptx`;
        a.click();
        setTimeout(() => URL.revokeObjectURL(a.href), 5000);
        store.setStatus("Saved as .pptx");
      }
      setTimeout(() => store.setStatus(null), 3000);
    } catch (err) {
      console.error(err);
      store.setStatus("Save failed: " + (err as Error).message);
      notifyHost({ type: "pe:error", message: (err as Error).message });
    }
  };

  const onOpenFile = () => openInputRef.current?.click();

  const openBuffer = async (buf: ArrayBuffer, name: string) => {
    store.setStatus(`Opening ${name}…`);
    const { pres, media, warnings } = await parsePptx(buf, name);
    store.loadPresentation(pres, media);
    store.setStatus(warnings.length
      ? `Opened ${name} — ${warnings.length} item(s) imported approximately`
      : `Opened ${name}`);
    setTimeout(() => store.setStatus(null), 5000);
    if (warnings.length) console.warn("Import warnings:", warnings);
    notifyHost({ type: "pe:loaded", title: pres.title, slideCount: pres.slides.length });
  };

  const handleOpen = async (file: File) => {
    try {
      await openBuffer(await file.arrayBuffer(), file.name);
    } catch (err) {
      console.error(err);
      alert("Could not open file: " + (err as Error).message);
      store.setStatus(null);
    }
  };

  // ---------- host integration: ?file= boot load + postMessage bridge ----------
  useEffect(() => {
    initEmbedBridge(EMBED, {
      load: async (buf, title) => { await openBuffer(buf, title); },
      exportPptx: () => exportPptxBlob(store.pres, store.media),
      title: () => store.pres.title,
    });
    if (EMBED.fileUrl) {
      (async () => {
        try {
          store.setStatus("Fetching document…");
          const res = await fetch(EMBED.fileUrl!);
          if (!res.ok) throw new Error(`HTTP ${res.status} fetching the document`);
          const buf = await res.arrayBuffer();
          await openBuffer(buf, EMBED.title ?? decodeURIComponent(EMBED.fileUrl!.split("/").pop() ?? "Presentation"));
          if (EMBED.title) store.setState({ pres: { ...store.pres, title: EMBED.title } });
        } catch (err) {
          console.error(err);
          store.setStatus("Could not open document: " + (err as Error).message);
          notifyHost({ type: "pe:error", message: (err as Error).message });
        }
      })();
    }
    // surface unsaved-changes state to the host
    let lastDirty = store.state.dirty;
    const unsub = store.subscribe(() => {
      if (store.state.dirty !== lastDirty) {
        lastDirty = store.state.dirty;
        notifyHost({ type: "pe:dirty", dirty: lastDirty });
      }
    });
    return () => { unsub(); };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const handleImage = async (file: File) => {
    const { mediaId, natW, natH } = await loadImageFile(file);
    const pres = store.pres;
    const maxW = pres.slideWidth * 0.5, maxH = pres.slideHeight * 0.5;
    let w = natW * EMU_PER_PX, h = natH * EMU_PER_PX;
    const k = Math.min(1, maxW / w, maxH / h);
    w *= k; h *= k;
    store.addShape({
      kind: "pic", id: nextId("pic"), name: file.name, mediaId,
      x: Math.round((pres.slideWidth - w) / 2), y: Math.round((pres.slideHeight - h) / 2),
      w: Math.round(w), h: Math.round(h), rot: 0,
    });
  };

  const present = (fromStart = false) => {
    store.setState({
      presenting: true,
      presentIndex: fromStart ? 0 : state.selection.slideIndex,
      editingShapeId: null,
      pendingShape: null,
    });
  };

  // ---------- global keyboard ----------
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      const s = store.getState();
      if (s.presenting) return; // PresentMode handles its own keys
      const tgt = e.target;
      const inField = tgt instanceof Element && tgt.closest("input, select, textarea, [contenteditable=true]");
      const ctrl = e.ctrlKey || e.metaKey;

      if (e.key === "F5") { e.preventDefault(); present(!e.shiftKey); return; }
      if (ctrl && e.key.toLowerCase() === "s") { e.preventDefault(); onSave(); return; }
      if (ctrl && e.key.toLowerCase() === "f") { e.preventDefault(); setFindOpen(true); return; }
      if (inField) return;

      if (e.key === "Escape") {
        if (s.croppingId) store.setState({ croppingId: null });
        else if (s.tableSel) store.setState({ tableSel: null });
        else if (s.pendingShape) store.setState({ pendingShape: null });
        else store.selectShapes([]);
        return;
      }
      if (e.key === "Enter" && s.croppingId) {
        store.setState({ croppingId: null });
        return;
      }
      if (ctrl && e.key.toLowerCase() === "z") { e.preventDefault(); store.undo(); return; }
      if (ctrl && (e.key.toLowerCase() === "y" || (e.shiftKey && e.key.toLowerCase() === "z"))) { e.preventDefault(); store.redo(); return; }
      if (ctrl && e.key.toLowerCase() === "c") { store.copySelection(); return; }
      if (ctrl && e.key.toLowerCase() === "x") { store.copySelection(); store.deleteSelectedShapes(); return; }
      if (ctrl && e.key.toLowerCase() === "v") { store.pasteClipboard(); return; }
      if (ctrl && e.key.toLowerCase() === "d") { e.preventDefault(); store.copySelection(); store.pasteClipboard(); return; }
      if (ctrl && e.key.toLowerCase() === "g") {
        e.preventDefault();
        if (e.shiftKey) store.ungroupSelection();
        else store.groupSelection();
        return;
      }
      if (ctrl && e.key.toLowerCase() === "a") {
        e.preventDefault();
        store.selectShapes(store.currentSlide.shapes.map(sh => sh.id));
        return;
      }
      if (e.key === "Delete" || e.key === "Backspace") {
        if (s.selection.shapeIds.length) { e.preventDefault(); store.deleteSelectedShapes(); }
        return;
      }
      if (e.key === "F2" && s.selection.shapeIds.length === 1) {
        const sh = store.selectedShapes[0];
        if (sh?.kind === "sp") store.setState({ editingShapeId: sh.id });
        return;
      }
      if (e.key.startsWith("Arrow") && s.selection.shapeIds.length) {
        e.preventDefault();
        const step = (e.shiftKey ? 10 : 1) * EMU_PER_PX;
        const dx = e.key === "ArrowLeft" ? -step : e.key === "ArrowRight" ? step : 0;
        const dy = e.key === "ArrowUp" ? -step : e.key === "ArrowDown" ? step : 0;
        store.updateShapes(s.selection.shapeIds, sh => ({ ...sh, x: sh.x + dx, y: sh.y + dy }));
        return;
      }
      if (e.key === "PageDown") { e.preventDefault(); store.selectSlide(s.selection.slideIndex + 1); return; }
      if (e.key === "PageUp") { e.preventDefault(); store.selectSlide(s.selection.slideIndex - 1); return; }
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // warn before leaving with unsaved changes
  useEffect(() => {
    const onBeforeUnload = (e: BeforeUnloadEvent) => {
      if (store.getState().dirty) { e.preventDefault(); e.returnValue = ""; }
    };
    window.addEventListener("beforeunload", onBeforeUnload);
    return () => window.removeEventListener("beforeunload", onBeforeUnload);
  }, []);

  return (
    <div className="app">
      <Ribbon onOpenFile={onOpenFile} onSave={onSave} onPresent={present} onImportPattern={() => setPatternOpen(true)} />
      <div className="main-row">
        <div className="left-rail">
          <button className={`rail-btn ${findOpen ? "active" : ""}`} title="Find and replace (Ctrl+F)" onClick={() => setFindOpen(v => !v)}><SearchIcon /></button>
          <button className={`rail-btn ${showThumbs ? "active" : ""}`} title="Slide thumbnails" onClick={() => setShowThumbs(v => !v)}>
            <Icon name="addSlide" size={18} />
          </button>
          <div className="rail-spacer" />
          <button className="rail-btn" title="About" onClick={() => store.setStatus(`Presentation Editor v${typeof __APP_VERSION__ !== "undefined" ? __APP_VERSION__ : "dev"}`)}><InfoIcon /></button>
        </div>
        {showThumbs && <SlidePanel onPresentFrom={i => { store.setState({ presenting: true, presentIndex: i }); }} />}
        <div className="center-col">
          <EditorCanvas />
          <NotesBar />
        </div>
        <RightPanel />
      </div>
      <StatusBar onPresent={() => present(false)} />
      {state.presenting && <PresentMode />}
      {patternOpen && <PatternDialog onClose={() => setPatternOpen(false)} />}
      {findOpen && <FindDialog onClose={() => setFindOpen(false)} />}
      {chartShape?.kind === "chart" && (
        <ChartDataDialog shape={chartShape} onClose={() => store.setState({ chartEditId: null })} />
      )}
      <input
        ref={openInputRef}
        type="file"
        accept=".pptx,application/vnd.openxmlformats-officedocument.presentationml.presentation"
        style={{ display: "none" }}
        onChange={e => {
          const f = e.target.files?.[0];
          if (f) handleOpen(f);
          e.target.value = "";
        }}
      />
      <input
        id="image-file-input"
        type="file"
        accept="image/png,image/jpeg,image/gif,image/svg+xml"
        style={{ display: "none" }}
        onChange={e => {
          const f = e.target.files?.[0];
          if (f) handleImage(f);
          e.target.value = "";
        }}
      />
    </div>
  );
}

function SearchIcon() {
  return (
    <svg width="18" height="18" viewBox="0 0 20 20">
      <circle cx="8.5" cy="8.5" r="5" fill="none" stroke="currentColor" strokeWidth="1.5" />
      <path d="m12.5 12.5 4.5 4.5" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" />
    </svg>
  );
}

function InfoIcon() {
  return (
    <svg width="18" height="18" viewBox="0 0 20 20">
      <circle cx="10" cy="10" r="7" fill="none" stroke="currentColor" strokeWidth="1.4" />
      <path d="M10 9v4.5M10 6.2v.2" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" />
    </svg>
  );
}
