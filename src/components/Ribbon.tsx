import React, { useState } from "react";
import type { ChartKind, ColorRef, Paragraph, PresetGeom, Run, SpShape, TextAlign } from "../model/types";
import {
  CHART_NAMES, FONT_LIST, FONT_SIZES, LAYOUT_NAMES, MAJOR_FONT, MINOR_FONT,
  THEME_FONT_PAIRS, THEME_PRESETS, isThemeFont, makeChart, makeTable,
  resolveColor, type LayoutKind,
} from "../model/defaults";
import { SHAPE_CATEGORIES, SHAPE_GALLERY, geomLabel, presetPath } from "../render/geometry";
import { store } from "../state/store";
import { editorApi } from "../util/api";
import { permissions, uiConfig } from "../util/config";
import { editorBus, useEditorState } from "../state/useStore";
import {
  customFontNames, deleteCustomFontPair, deleteCustomTheme,
  loadCustomFontPairs, loadCustomThemes, saveCustomFontPair, saveCustomTheme,
} from "../util/custom";
import { ColorButton } from "./ColorPicker";
import { FontPairDialog, ThemeColorsDialog } from "./CustomizeDialogs";
import { Dropdown } from "./Dropdown";
import { Icon } from "./icons";

export type TabId = "file" | "home" | "insert" | "design" | "transitions" | "view";

const TABS: { id: TabId; label: string }[] = [
  { id: "file", label: "File" },
  { id: "home", label: "Home" },
  { id: "insert", label: "Insert" },
  { id: "design", label: "Design" },
  { id: "transitions", label: "Transitions" },
  { id: "view", label: "View" },
];

const CASE_FNS: Record<string, (s: string) => string> = {
  "Sentence case.": s => s.toLowerCase().replace(/(^\s*\p{L}|[.!?]\s+\p{L})/gu, m => m.toUpperCase()),
  "lowercase": s => s.toLowerCase(),
  "UPPERCASE": s => s.toUpperCase(),
  "Capitalize Each Word": s => s.toLowerCase().replace(/(^|\s)\p{L}/gu, m => m.toUpperCase()),
  "tOGGLE cASE": s => [...s].map(c => (c === c.toUpperCase() ? c.toLowerCase() : c.toUpperCase())).join(""),
};

export function Ribbon({ onOpenFile, onSave, onPresent, onImportPattern }: {
  onOpenFile: () => void;
  onSave: () => void;
  onPresent: (fromStart?: boolean) => void;
  onImportPattern: () => void;
}) {
  const state = useEditorState();
  // default to the first ribbon tab the host left enabled
  const firstTab = (["home", "insert", "design", "transitions", "view"] as const).find(t => uiConfig.tabs[t]) ?? "home";
  const [tab, setTab] = useState<TabId>(firstTab);
  const [fileOpen, setFileOpen] = useState(false);
  const { pres, selection, editingShapeId } = state;
  const theme = pres.theme;

  // user-defined palettes / font pairs (Design tab "Customize…")
  const [customThemes, setCustomThemes] = useState(loadCustomThemes);
  const [customPairs, setCustomPairs] = useState(loadCustomFontPairs);
  const [colorsDlg, setColorsDlg] = useState(false);
  const [fontsDlg, setFontsDlg] = useState(false);
  const customFonts = React.useMemo(customFontNames, [customPairs]);

  const selShapes = store.selectedShapes;
  const selSps = selShapes.filter((s): s is SpShape => s.kind === "sp");
  const hasTable = selShapes.some(s => s.kind === "table");
  const firstRun: Run | undefined = (() => {
    for (const s of selShapes) {
      if (s.kind === "sp") {
        for (const p of s.text?.paragraphs ?? []) for (const r of p.runs) return r;
      } else if (s.kind === "table") {
        for (const row of s.cells) for (const c of row) for (const p of c.text.paragraphs) for (const r of p.runs) return r;
      }
    }
    return undefined;
  })();
  const firstPara: Paragraph | undefined = selSps[0]?.text?.paragraphs[0];
  // chart text-element selection routes Home-tab font controls to that part
  const chartPart = state.chartPartSel;
  const chartPartStyle = chartPart ? store.chartPartStyle : null;
  const partDefaultSize = (): number => {
    if (!chartPart) return 18;
    if (chartPart.part === "title") return 16;
    const ch = selShapes.find(s => s.id === chartPart.id);
    // legend / axis labels / axis titles all default to the chart label size (12pt)
    return (ch?.kind === "chart" && ch.labelSizePt) || 12;
  };

  const hasTextTarget = editingShapeId !== null || selSps.length > 0 || hasTable || !!chartPart;
  const hasSel = selShapes.length > 0;

  // ---------- text formatting (live execCommand while editing, model ops otherwise) ----------
  const toggleRunProp = (prop: "bold" | "italic" | "underline" | "strike", cmd: string) => {
    if (chartPart && prop !== "strike") { store.formatChartPart({ [prop]: !(chartPartStyle?.[prop]) || undefined }); return; }
    if (editingShapeId) { editorBus.exec(cmd); return; }
    const cur = firstRun?.[prop] ?? false;
    store.formatRuns(r => ({ ...r, [prop]: !cur || undefined }));
  };

  const setFont = (font: string) => {
    if (chartPart) { store.formatChartPart({ font }); return; }
    if (isThemeFont(font)) {
      // theme refs are model-level — commit any open editor, then apply to the shape
      withCommitted(ids => store.formatRuns(r => ({ ...r, font }), ids));
      return;
    }
    if (editingShapeId) { editorBus.exec("fontName", font); return; }
    store.formatRuns(r => ({ ...r, font }));
  };

  const setSize = (sizePt: number) => {
    if (!sizePt || Number.isNaN(sizePt)) return;
    if (chartPart) { store.formatChartPart({ sizePt }); return; }
    if (editingShapeId) { execFontSize(sizePt); return; }
    store.formatRuns(r => ({ ...r, sizePt }));
  };

  const bumpSize = (dir: 1 | -1) => {
    if (chartPart) { setSize(Math.max(5, Math.min(96, partDefaultSize() + dir))); return; }
    if (editingShapeId) return; // keep simple while editing
    store.formatRuns(r => {
      const i = FONT_SIZES.findIndex(s => s >= r.sizePt);
      const at = dir === 1
        ? Math.min(FONT_SIZES.length - 1, (FONT_SIZES[i] === r.sizePt ? i : i < 0 ? FONT_SIZES.length - 1 : i) + (FONT_SIZES[i] === r.sizePt ? 1 : 0))
        : Math.max(0, (i < 0 ? FONT_SIZES.length : i) - 1);
      return { ...r, sizePt: FONT_SIZES[Math.max(0, Math.min(FONT_SIZES.length - 1, at))] };
    });
  };

  const setTextColor = (c: ColorRef) => {
    if (chartPart) { store.formatChartPart({ color: c }); return; }
    if (editingShapeId) { editorBus.exec("foreColor", resolveColor(c, theme)); return; }
    store.formatRuns(r => ({ ...r, color: c }));
  };

  const setAlign = (align: TextAlign) => {
    if (editingShapeId) {
      const cmd = align === "l" ? "justifyLeft" : align === "ctr" ? "justifyCenter" : align === "r" ? "justifyRight" : "justifyFull";
      editorBus.exec(cmd);
      return;
    }
    store.formatParagraphs(p => ({ ...p, align }));
  };

  const withCommitted = (fn: (ids: string[]) => void) => {
    const ids = editingShapeId ? [editingShapeId] : selection.shapeIds;
    editorBus.commitTextEdit?.();
    fn(ids);
  };

  const toggleBullet = (kind: "char" | "num") => withCommitted(ids => {
    const target = store.pres.slides[selection.slideIndex].shapes.find(s => s.id === ids[0]);
    const cur = target?.kind === "sp" ? target.text?.paragraphs[0]?.bullet : "none";
    store.formatParagraphs(p => ({ ...p, bullet: cur === kind ? "none" : kind }), ids);
  });

  const setAnchor = (anchor: "t" | "ctr" | "b") => withCommitted(ids => {
    store.updateShapes(ids, s => {
      if (s.kind === "sp" && s.text) return { ...s, text: { ...s.text, anchor } };
      if (s.kind === "table") return { ...s, cells: s.cells.map(row => row.map(c => ({ ...c, text: { ...c.text, anchor } }))) };
      return s;
    });
  });

  const setLineSpacing = (pct: number) => withCommitted(ids => {
    store.formatParagraphs(p => ({ ...p, lineSpacingPct: pct }), ids);
  });

  const toggleBaseline = (val: 30 | -25, cmd: string) => {
    if (editingShapeId) { editorBus.exec(cmd); return; }
    const cur = firstRun?.baseline ?? 0;
    store.formatRuns(r => ({ ...r, baseline: cur === val ? undefined : val }));
  };

  const setHighlight = (c: ColorRef | null) => {
    if (editingShapeId && c) { editorBus.exec("hiliteColor", resolveColor(c, theme)); return; }
    withCommitted(ids => store.formatRuns(r => ({ ...r, highlight: c ?? undefined }), ids));
  };

  const clearFormatting = () => withCommitted(ids => {
    store.formatRuns(r => ({ text: r.text, sizePt: r.sizePt, font: MINOR_FONT, color: { kind: "scheme", slot: "dk1" } }), ids);
  });

  const changeCase = (label: string) => withCommitted(ids => store.transformText(CASE_FNS[label], ids));

  const bumpIndent = (d: 1 | -1) => withCommitted(ids => {
    store.formatParagraphs(p => ({ ...p, level: Math.max(0, Math.min(8, p.level + d)) }), ids);
  });

  const setColumns = (n: number) => withCommitted(ids => {
    store.updateShapes(ids, s => s.kind === "sp" && s.text
      ? { ...s, text: { ...s.text, columns: n > 1 ? n : undefined, colSpacing: n > 1 ? (s.text.colSpacing ?? 360000) : undefined } }
      : s);
  });

  // ---------- shape formatting ----------
  const setShapeFill = (c: ColorRef | null) => {
    store.updateShapes(selection.shapeIds, s =>
      s.kind === "sp" ? { ...s, fill: c ? { kind: "solid", color: c } : { kind: "none" } } : s);
  };
  const setShapeLine = (c: ColorRef | null) => {
    store.updateShapes(selection.shapeIds, s =>
      s.kind === "sp" ? { ...s, line: { ...s.line, fill: c ? { kind: "solid", color: c } : { kind: "none" } } } : s);
  };

  const startShape = (geom: PresetGeom | "textbox") => {
    if (geom !== "textbox") pushRecentShape(geom);
    store.setState({ pendingShape: geom, editingShapeId: null });
  };

  const insertTable = (rows: number, cols: number) => {
    const W = pres.slideWidth, H = pres.slideHeight;
    const w = Math.round(W * 0.7);
    const h = Math.min(Math.round(H * 0.66), rows * 426720);
    store.addShape(makeTable(rows, cols, Math.round((W - w) / 2), Math.round(H * 0.2), w, h));
  };

  const insertChart = (v: ChartVariant) => {
    const W = pres.slideWidth, H = pres.slideHeight;
    const w = Math.round(W * 0.55), h = Math.round(H * 0.6);
    store.addShape(makeChart(v.kind, Math.round((W - w) / 2), Math.round((H - h) / 2), w, h, v.opts));
  };

  // ---------- toolbar groups ----------
  const StaticCluster = (
    <>
      <div className="tb-group tb-grid3">
        {uiConfig.save && <button className="tb-btn sm" title="Save as .pptx (Ctrl+S)" onClick={onSave}><Icon name="save" size={16} /></button>}
        <button className="tb-btn sm" title="Copy (Ctrl+C)" disabled={!hasSel} onClick={() => store.copySelection()}><Icon name="duplicate" size={16} /></button>
        <button className="tb-btn sm" title="Paste (Ctrl+V)" onClick={() => store.pasteClipboard()}><Icon name="doc" size={16} /></button>
        <button className="tb-btn sm" title="Undo (Ctrl+Z)" disabled={!state.canUndo} onClick={store.undo}><Icon name="undo" size={16} /></button>
        <button className="tb-btn sm" title="Redo (Ctrl+Y)" disabled={!state.canRedo} onClick={store.redo}><Icon name="redo" size={16} /></button>
        <button className="tb-btn sm" title="Delete (Del)" disabled={!hasSel} onClick={() => store.deleteSelectedShapes()}><Icon name="trash" size={16} /></button>
      </div>
      <div className="tb-sep" />
      <div className="tb-group">
        {/* split button: main face adds a Title-and-Content slide, chevron picks a layout */}
        <div className="split-btn">
          <button className="tb-btn xhuge split-main" title="New slide (Title and Content)" onClick={() => store.addSlide("titleContent")}>
            <span className="xhuge-face"><Icon name="addSlide" size={26} /><span>Add Slide</span></span>
          </button>
          <Dropdown className="split-chev" title="New slide with layout" button={<Icon name="chevDown" size={10} />}>
            {close => (
              <div className="menu">
                {(Object.keys(LAYOUT_NAMES) as LayoutKind[]).map(k => (
                  <button key={k} className="menu-item" onClick={() => { store.addSlide(k); close(); }}>{LAYOUT_NAMES[k]}</button>
                ))}
              </div>
            )}
          </Dropdown>
        </div>
        <button className="tb-btn xhuge" title="Start slideshow (F5)" onClick={() => onPresent(true)}>
          <span className="xhuge-face"><Icon name="present" size={26} /><span>Start<br />slideshow</span></span>
        </button>
      </div>
      <div className="tb-sep" />
    </>
  );

  // effective formatting source: chart part when one is selected, else the first run
  const fmtRef = chartPart ? chartPartStyle : firstRun;
  const fontFace = chartPart ? (chartPartStyle?.font ?? MINOR_FONT) : (firstRun?.font ?? "Arial");
  const fontSize = chartPart ? (chartPartStyle?.sizePt ?? partDefaultSize()) : (firstRun?.sizePt ?? 18);

  const FontGroup = (
    <div className="tb-group tb-col">
      <div className="tb-row">
        <select className="font-combo" title="Font" value={fontFace} disabled={!hasTextTarget} onChange={e => setFont(e.target.value)}>
          <optgroup label="Theme Fonts">
            <option value={MAJOR_FONT} style={{ fontFamily: theme.majorFont }}>{theme.majorFont} (Headings)</option>
            <option value={MINOR_FONT} style={{ fontFamily: theme.minorFont }}>{theme.minorFont} (Body)</option>
          </optgroup>
          {customFonts.length > 0 && (
            <optgroup label="Custom">
              {customFonts.map(f => <option key={f} value={f} style={{ fontFamily: `'${f}', sans-serif` }}>{f}</option>)}
            </optgroup>
          )}
          <optgroup label="All Fonts">
            {FONT_LIST.includes(fontFace) || isThemeFont(fontFace) || customFonts.includes(fontFace) ? null : <option value={fontFace}>{fontFace}</option>}
            {FONT_LIST.map(f => <option key={f} value={f} style={{ fontFamily: f }}>{f}</option>)}
          </optgroup>
        </select>
        <select className="size-combo" title="Font size" value={String(fontSize)} disabled={!hasTextTarget} onChange={e => setSize(parseFloat(e.target.value))}>
          {FONT_SIZES.includes(fontSize) ? null : <option value={String(fontSize)}>{fontSize}</option>}
          {FONT_SIZES.map(s => <option key={s} value={String(s)}>{s}</option>)}
        </select>
        <button className="tb-btn sm" title="Increase font size" disabled={!hasTextTarget} onClick={() => bumpSize(1)}><Icon name="fontUp" size={16} /></button>
        <button className="tb-btn sm" title="Decrease font size" disabled={!hasTextTarget} onClick={() => bumpSize(-1)}><Icon name="fontDown" size={16} /></button>
        <Dropdown title="Change case" disabled={!hasTextTarget} button={<span className="with-chev"><Icon name="changeCase" size={16} /><Icon name="chevDown" size={9} /></span>}>
          {close => (
            <div className="menu">
              {Object.keys(CASE_FNS).map(label => (
                <button key={label} className="menu-item" onClick={() => { changeCase(label); close(); }}>{label}</button>
              ))}
            </div>
          )}
        </Dropdown>
      </div>
      <div className="tb-row">
        <button className={`tb-btn sm ${fmtRef?.bold ? "active" : ""}`} title="Bold (Ctrl+B)" disabled={!hasTextTarget} onClick={() => toggleRunProp("bold", "bold")}><Icon name="bold" size={16} /></button>
        <button className={`tb-btn sm ${fmtRef?.italic ? "active" : ""}`} title="Italic (Ctrl+I)" disabled={!hasTextTarget} onClick={() => toggleRunProp("italic", "italic")}><Icon name="italic" size={16} /></button>
        <button className={`tb-btn sm ${fmtRef?.underline ? "active" : ""}`} title="Underline (Ctrl+U)" disabled={!hasTextTarget} onClick={() => toggleRunProp("underline", "underline")}><Icon name="underline" size={16} /></button>
        <button className={`tb-btn sm ${firstRun?.strike ? "active" : ""}`} title="Strikethrough" disabled={!hasTextTarget} onClick={() => toggleRunProp("strike", "strikeThrough")}><Icon name="strike" size={16} /></button>
        <button className={`tb-btn sm ${(firstRun?.baseline ?? 0) > 0 ? "active" : ""}`} title="Superscript" disabled={!hasTextTarget} onClick={() => toggleBaseline(30, "superscript")}><Icon name="superscript" size={16} /></button>
        <button className={`tb-btn sm ${(firstRun?.baseline ?? 0) < 0 ? "active" : ""}`} title="Subscript" disabled={!hasTextTarget} onClick={() => toggleBaseline(-25, "subscript")}><Icon name="subscript" size={16} /></button>
        <ColorButton
          theme={theme} icon="highlighter" title="Highlight color" disabled={!hasTextTarget}
          currentHex={firstRun?.highlight ? resolveColor(firstRun.highlight, theme) : "#FFFF00"}
          onPick={c => setHighlight(c)} allowNone noneLabel="No Highlight" onNone={() => setHighlight(null)}
        />
        <ColorButton
          theme={theme} icon="fontColor" title="Font color" disabled={!hasTextTarget}
          currentHex={fmtRef?.color ? resolveColor(fmtRef.color, theme) : "#000"}
          onPick={setTextColor}
        />
        <button className="tb-btn sm" title="Clear formatting" disabled={!hasTextTarget} onClick={clearFormatting}><Icon name="eraser" size={16} /></button>
      </div>
    </div>
  );

  const ParaGroup = (
    <div className="tb-group tb-col">
      <div className="tb-row">
        <button className={`tb-btn sm ${firstPara?.bullet === "char" ? "active" : ""}`} title="Bullets" disabled={!hasTextTarget} onClick={() => toggleBullet("char")}><Icon name="bullets" size={16} /></button>
        <button className={`tb-btn sm ${firstPara?.bullet === "num" ? "active" : ""}`} title="Numbering" disabled={!hasTextTarget} onClick={() => toggleBullet("num")}><Icon name="numbering" size={16} /></button>
        <button className="tb-btn sm" title="Decrease indent" disabled={!hasTextTarget} onClick={() => bumpIndent(-1)}><Icon name="indentDec" size={16} /></button>
        <button className="tb-btn sm" title="Increase indent" disabled={!hasTextTarget} onClick={() => bumpIndent(1)}><Icon name="indentInc" size={16} /></button>
        <Dropdown title="Line spacing" disabled={!hasTextTarget} button={<span className="with-chev"><Icon name="anchorM" size={16} /><Icon name="chevDown" size={9} /></span>}>
          {close => (
            <div className="menu">
              {[100, 115, 150, 200, 250, 300].map(v => (
                <button key={v} className="menu-item" onClick={() => { setLineSpacing(v); close(); }}>{(v / 100).toFixed(2).replace(/0+$/, "").replace(/\.$/, ".0")}</button>
              ))}
            </div>
          )}
        </Dropdown>
      </div>
      <div className="tb-row">
        <button className={`tb-btn sm ${firstPara?.align === "l" ? "active" : ""}`} title="Align left" disabled={!hasTextTarget} onClick={() => setAlign("l")}><Icon name="alignL" size={16} /></button>
        <button className={`tb-btn sm ${firstPara?.align === "ctr" ? "active" : ""}`} title="Align center" disabled={!hasTextTarget} onClick={() => setAlign("ctr")}><Icon name="alignC" size={16} /></button>
        <button className={`tb-btn sm ${firstPara?.align === "r" ? "active" : ""}`} title="Align right" disabled={!hasTextTarget} onClick={() => setAlign("r")}><Icon name="alignR" size={16} /></button>
        <button className={`tb-btn sm ${firstPara?.align === "just" ? "active" : ""}`} title="Justify" disabled={!hasTextTarget} onClick={() => setAlign("just")}><Icon name="alignJ" size={16} /></button>
        <Dropdown title="Vertical align" disabled={!hasTextTarget} button={<span className="with-chev"><Icon name="anchorT" size={16} /><Icon name="chevDown" size={9} /></span>}>
          {close => (
            <div className="menu">
              <button className="menu-item" onClick={() => { setAnchor("t"); close(); }}>Align text to top</button>
              <button className="menu-item" onClick={() => { setAnchor("ctr"); close(); }}>Align text to middle</button>
              <button className="menu-item" onClick={() => { setAnchor("b"); close(); }}>Align text to bottom</button>
            </div>
          )}
        </Dropdown>
        <Dropdown title="Text columns" disabled={!selSps.length} button={<span className="with-chev"><Icon name="columns" size={16} /><Icon name="chevDown" size={9} /></span>}>
          {close => (
            <div className="menu">
              <button className="menu-item" onClick={() => { setColumns(1); close(); }}><Icon name="alignJ" size={15} /> One Column</button>
              <button className="menu-item" onClick={() => { setColumns(2); close(); }}><Icon name="columns" size={15} /> Two Columns</button>
              <button className="menu-item" onClick={() => { setColumns(3); close(); }}><Icon name="columns" size={15} /> Three Columns</button>
              <div className="menu-div" />
              <button className="menu-item" onClick={() => {
                const v = parseInt(prompt("Number of columns (1-8):", "2") ?? "", 10);
                if (v >= 1 && v <= 8) setColumns(v);
                close();
              }}>Custom columns…</button>
            </div>
          )}
        </Dropdown>
      </div>
    </div>
  );

  const HomePanel = (
    <>
      {StaticCluster}
      {FontGroup}
      <div className="tb-sep" />
      {ParaGroup}
      <div className="tb-sep" />
      <div className="tb-group">
        <button className="tb-btn xhuge" title="Insert text box" onClick={() => startShape("textbox")}>
          <span className="xhuge-face"><Icon name="textbox" size={26} /><span>Text Box</span></span>
        </button>
        <button className="tb-btn xhuge" title="Insert image" onClick={onOpenImagePicker}>
          <span className="xhuge-face"><Icon name="image" size={26} /><span>Image</span></span>
        </button>
        <InlineShapeGallery onPick={startShape} />
      </div>
      <div className="tb-sep" />
      <div className="tb-group tb-col">
        <div className="tb-row">
          <ColorButton
            theme={theme} icon="fill" title="Shape fill" disabled={!selSps.length}
            currentHex={selSps[0] && selSps[0].fill.kind === "solid" ? resolveColor(selSps[0].fill.color, theme) : "#fff"}
            onPick={c => setShapeFill(c)} allowNone noneLabel="No Fill" onNone={() => setShapeFill(null)}
          />
          <ColorButton
            theme={theme} icon="outline" title="Shape outline" disabled={!selSps.length}
            currentHex={selSps[0] && selSps[0].line.fill.kind === "solid" ? resolveColor(selSps[0].line.fill.color, theme) : "#000"}
            onPick={c => setShapeLine(c)} allowNone noneLabel="No Outline" onNone={() => setShapeLine(null)}
          />
        </div>
        <div className="tb-row">
          <Dropdown title="Arrange" disabled={!hasSel} button={<span className="with-chev"><Icon name="arrange" size={16} /><Icon name="chevDown" size={9} /></span>}>
            {close => (
              <div className="menu">
                <button className="menu-item" onClick={() => { selection.shapeIds.forEach(id => store.reorderShape(id, "front")); close(); }}>Bring to foreground</button>
                <button className="menu-item" onClick={() => { selection.shapeIds.forEach(id => store.reorderShape(id, "back")); close(); }}>Send to background</button>
                <button className="menu-item" onClick={() => { selection.shapeIds.forEach(id => store.reorderShape(id, "forward")); close(); }}>Bring forward</button>
                <button className="menu-item" onClick={() => { selection.shapeIds.forEach(id => store.reorderShape(id, "backward")); close(); }}>Send backward</button>
                <div className="menu-div" />
                <button className="menu-item" onClick={() => { store.alignSelected("l"); close(); }}>Align Left</button>
                <button className="menu-item" onClick={() => { store.alignSelected("c"); close(); }}>Align Center</button>
                <button className="menu-item" onClick={() => { store.alignSelected("r"); close(); }}>Align Right</button>
                <button className="menu-item" onClick={() => { store.alignSelected("t"); close(); }}>Align Top</button>
                <button className="menu-item" onClick={() => { store.alignSelected("m"); close(); }}>Align Middle</button>
                <button className="menu-item" onClick={() => { store.alignSelected("b"); close(); }}>Align Bottom</button>
                <div className="menu-div" />
                <button className="menu-item" disabled={selection.shapeIds.length < 3} onClick={() => { store.distributeSelected("h"); close(); }}>Distribute Horizontally</button>
                <button className="menu-item" disabled={selection.shapeIds.length < 3} onClick={() => { store.distributeSelected("v"); close(); }}>Distribute Vertically</button>
              </div>
            )}
          </Dropdown>
        </div>
      </div>
    </>
  );

  const InsertPanel = (
    <>
      {StaticCluster}
      <div className="tb-group">
        <TableGridPicker onPick={insertTable} />
        <ChartMenu onPick={insertChart} />
      </div>
      <div className="tb-sep" />
      <div className="tb-group">
        <button className="tb-btn xhuge" title="Insert text box" onClick={() => startShape("textbox")}>
          <span className="xhuge-face"><Icon name="textbox" size={26} /><span>Text Box</span></span>
        </button>
        <ShapeGalleryButton onPick={startShape} />
        <button className="tb-btn xhuge" title="Insert image from file" onClick={onOpenImagePicker}>
          <span className="xhuge-face"><Icon name="image" size={26} /><span>Image</span></span>
        </button>
      </div>
    </>
  );

  const DesignPanel = (
    <>
      {StaticCluster}
      <div className="tb-group">
        <div className="theme-gallery">
          {[...THEME_PRESETS, ...customThemes].map((t, idx) => {
            const isCustom = idx >= THEME_PRESETS.length;
            return (
              <button
                key={t.name}
                className={`theme-chip ${pres.theme.name === t.name ? "active" : ""}`}
                title={`${t.name} theme${isCustom ? " — right-click to delete" : ""}`}
                onClick={() => store.commit({
                  ...pres,
                  // color scheme only — keep whatever heading/body fonts are active
                  theme: { ...t, majorFont: pres.theme.majorFont, minorFont: pres.theme.minorFont },
                })}
                onContextMenu={e => {
                  if (!isCustom) return;
                  e.preventDefault();
                  if (confirm(`Delete palette “${t.name}”?`)) {
                    deleteCustomTheme(t.name);
                    setCustomThemes(loadCustomThemes());
                  }
                }}
              >
                <span className="chip-name" style={{ color: "#" + t.dk2 }}>Aa</span>
                <span className="chip-colors">
                  {[t.accent1, t.accent2, t.accent3, t.accent4].map((c, i) => (
                    <span key={i} style={{ background: "#" + c }} />
                  ))}
                </span>
                <span className="chip-label">{t.name}</span>
              </button>
            );
          })}
          <button className="theme-chip add-chip" title="Create a custom color palette" onClick={() => setColorsDlg(true)}>
            <span className="chip-name">＋</span>
            <span className="chip-label">Customize…</span>
          </button>
        </div>
      </div>
      <div className="tb-sep" />
      <div className="tb-group">
        <Dropdown
          title="Theme fonts (Headings / Body)"
          className="xhuge"
          panelClassName="fonts-panel"
          button={<span className="xhuge-face"><Icon name="changeCase" size={26} /><span>Fonts <Icon name="chevDown" size={9} /></span></span>}
        >
          {close => (
            <div className="menu">
              {customPairs.length > 0 && <div className="color-section" style={{ padding: "2px 10px" }}>Custom</div>}
              {customPairs.map(p => (
                <button
                  key={p.name}
                  className={`menu-item font-pair ${theme.majorFont === p.major && theme.minorFont === p.minor ? "active" : ""}`}
                  title={`${p.name} — right-click to delete`}
                  onClick={() => {
                    store.commit({ ...pres, theme: { ...pres.theme, majorFont: p.major, minorFont: p.minor } });
                    close();
                  }}
                  onContextMenu={e => {
                    e.preventDefault();
                    if (confirm(`Delete font pair “${p.name}”?`)) {
                      deleteCustomFontPair(p.name);
                      setCustomPairs(loadCustomFontPairs());
                    }
                  }}
                >
                  <span className="font-pair-aa" style={{ fontFamily: `'${p.major}', sans-serif` }}>Aa</span>
                  <span className="font-pair-names">
                    <span style={{ fontFamily: `'${p.major}', sans-serif` }}>{p.major}</span>
                    <span style={{ fontFamily: `'${p.minor}', sans-serif` }}>{p.minor}</span>
                  </span>
                </button>
              ))}
              {customPairs.length > 0 && <div className="menu-div" />}
              {THEME_FONT_PAIRS.map(p => (
                <button
                  key={p.name}
                  className={`menu-item font-pair ${theme.majorFont === p.major && theme.minorFont === p.minor ? "active" : ""}`}
                  onClick={() => {
                    store.commit({ ...pres, theme: { ...pres.theme, majorFont: p.major, minorFont: p.minor } });
                    close();
                  }}
                >
                  <span className="font-pair-aa" style={{ fontFamily: p.major }}>Aa</span>
                  <span className="font-pair-names">
                    <span style={{ fontFamily: p.major }}>{p.major}</span>
                    <span style={{ fontFamily: p.minor }}>{p.minor}</span>
                  </span>
                </button>
              ))}
              <div className="menu-div" />
              <button className="menu-item" onClick={() => { setFontsDlg(true); close(); }}>Customize Fonts…</button>
            </div>
          )}
        </Dropdown>
      </div>
      <div className="tb-sep" />
      <div className="tb-group tb-col">
        <div className="tb-row">
          <ColorButton
            theme={theme} icon="fill" title="Slide background color"
            currentHex={slideBgHex(state)}
            onPick={c => store.updateSlideBg({ kind: "solid", color: c })}
            allowNone noneLabel="Reset to default" onNone={() => store.updateSlideBg(undefined)}
          />
        </div>
        <div className="tb-row">
          <button className="tb-btn label-btn" onClick={() => store.applyBgToAll()}>Apply to All Slides</button>
        </div>
      </div>
      <div className="tb-sep" />
      <div className="tb-group">
        <Dropdown title="Slide size" button={<span className="xhuge-face"><Icon name="fit" size={26} /><span>Slide Size <Icon name="chevDown" size={9} /></span></span>} className="xhuge">
          {close => (
            <div className="menu">
              <button className="menu-item" onClick={() => { store.setSlideSize(12192000, 6858000); close(); }}>Widescreen (16:9)</button>
              <button className="menu-item" onClick={() => { store.setSlideSize(9144000, 6858000); close(); }}>Standard (4:3)</button>
            </div>
          )}
        </Dropdown>
      </div>
    </>
  );

  const curTrans = pres.slides[selection.slideIndex]?.transition;
  const TransitionsPanel = (
    <>
      {StaticCluster}
      <div className="tb-group">
        {([
          { t: "none" as const, label: "None" },
          { t: "fade" as const, label: "Fade" },
          { t: "push" as const, label: "Push" },
        ]).map(({ t, label }) => (
          <button
            key={t}
            className={`tb-btn xhuge ${(curTrans?.type ?? "none") === t ? "active" : ""}`}
            onClick={() => store.setTransition(t === "none" ? undefined : { type: t, dir: "l", speed: curTrans?.speed ?? "med" })}
          >
            <span className="xhuge-face"><Icon name={t === "none" ? "fit" : t === "fade" ? "duplicate" : "redo"} size={26} /><span>{label}</span></span>
          </button>
        ))}
      </div>
      <div className="tb-sep" />
      <div className="tb-group tb-col">
        <div className="tb-row">
          <label className="tb-label">Direction</label>
          <select
            className="size-combo" disabled={curTrans?.type !== "push"}
            value={curTrans?.dir ?? "l"}
            onChange={e => curTrans && store.setTransition({ ...curTrans, dir: e.target.value as "l" | "r" | "u" | "d" })}
          >
            <option value="l">From right</option>
            <option value="r">From left</option>
            <option value="u">From bottom</option>
            <option value="d">From top</option>
          </select>
        </div>
        <div className="tb-row">
          <label className="tb-label">Speed</label>
          <select
            className="size-combo" disabled={!curTrans}
            value={curTrans?.speed ?? "med"}
            onChange={e => curTrans && store.setTransition({ ...curTrans, speed: e.target.value as "slow" | "med" | "fast" })}
          >
            <option value="slow">Slow</option>
            <option value="med">Medium</option>
            <option value="fast">Fast</option>
          </select>
        </div>
      </div>
      <div className="tb-sep" />
      <div className="tb-group tb-col">
        <div className="tb-row">
          <button className="tb-btn label-btn" disabled={!curTrans} onClick={() => store.applyTransitionToAll()}>Apply to All Slides</button>
        </div>
        <div className="tb-row">
          <button className="tb-btn label-btn" onClick={() => onPresent(false)}>Preview</button>
        </div>
      </div>
    </>
  );

  const ViewPanel = (
    <>
      {StaticCluster}
      <div className="tb-group">
        <button
          className={`tb-btn xhuge ${state.showRuler ? "active" : ""}`}
          title="Show or hide rulers"
          onClick={() => store.setState({ showRuler: !state.showRuler })}
        >
          <span className="xhuge-face"><Icon name="ruler" size={26} /><span>Ruler</span></span>
        </button>
      </div>
      <div className="tb-sep" />
      <div className="tb-group tb-col">
        <div className="tb-row">
          <button className={`tb-btn label-btn ${state.zoom === "fit" ? "active" : ""}`} onClick={() => store.setState({ zoom: "fit" })}>Fit to Slide</button>
        </div>
        <div className="tb-row">
          <button className={`tb-btn label-btn ${state.zoom === "fitw" ? "active" : ""}`} onClick={() => store.setState({ zoom: "fitw" })}>Fit to Width</button>
        </div>
      </div>
      <div className="tb-group tb-col">
        <div className="tb-row">
          <button className="tb-btn label-btn" onClick={() => store.setState({ zoom: 1 })}>Zoom 100%</button>
        </div>
        <div className="tb-row">
          <button className="tb-btn label-btn" onClick={() => store.setState({ zoom: 2 })}>Zoom 200%</button>
        </div>
      </div>
    </>
  );

  return (
    <div className="ribbon">
      <div className="header-bar">
        <div className="tab-strip">
          {TABS.filter(t => (t.id === "file" ? uiConfig.fileMenu : uiConfig.tabs[t.id as keyof typeof uiConfig.tabs])).map(t => (
            <button
              key={t.id}
              className={`tab ${tab === t.id && t.id !== "file" ? "active" : ""}`}
              onClick={() => (t.id === "file" ? setFileOpen(true) : setTab(t.id))}
            >
              {t.label}
            </button>
          ))}
        </div>
        <div className="tab-spacer" />
        {uiConfig.docTitle && (
          <input
            className="doc-title"
            value={pres.title}
            onChange={e => store.setState({ pres: { ...pres, title: e.target.value } })}
            spellCheck={false}
          />
        )}
        {uiConfig.present && (
          <button className="present-btn" title="Start slideshow (F5)" onClick={() => onPresent(true)}>
            <Icon name="present" size={12} /> Present
          </button>
        )}
      </div>
      <div className="toolbar">
        {uiConfig.tabs.home && tab === "home" && HomePanel}
        {uiConfig.tabs.insert && tab === "insert" && InsertPanel}
        {uiConfig.tabs.design && tab === "design" && DesignPanel}
        {uiConfig.tabs.transitions && tab === "transitions" && TransitionsPanel}
        {uiConfig.tabs.view && tab === "view" && ViewPanel}
      </div>
      {fileOpen && <FileMenu onClose={() => setFileOpen(false)} onOpenFile={onOpenFile} onSave={onSave} onImportPattern={onImportPattern} />}
      {colorsDlg && (
        <ThemeColorsDialog
          initial={pres.theme}
          takenNames={[...THEME_PRESETS, ...customThemes].map(t => t.name)}
          onClose={() => setColorsDlg(false)}
          onSave={t => {
            saveCustomTheme(t);
            setCustomThemes(loadCustomThemes());
            store.commit({ ...pres, theme: { ...t, majorFont: pres.theme.majorFont, minorFont: pres.theme.minorFont } });
            setColorsDlg(false);
          }}
        />
      )}
      {fontsDlg && (
        <FontPairDialog
          takenNames={customPairs.map(p => p.name)}
          initialMajor={theme.majorFont}
          initialMinor={theme.minorFont}
          onClose={() => setFontsDlg(false)}
          onSave={p => {
            saveCustomFontPair(p);
            setCustomPairs(loadCustomFontPairs());
            store.commit({ ...pres, theme: { ...pres.theme, majorFont: p.major, minorFont: p.minor } });
            setFontsDlg(false);
          }}
        />
      )}
    </div>
  );

  function onOpenImagePicker() {
    document.getElementById("image-file-input")?.click();
  }
}

function slideBgHex(state: ReturnType<typeof store.getState>): string {
  const slide = state.pres.slides[state.selection.slideIndex];
  if (slide?.background && slide.background.kind === "solid") return resolveColor(slide.background.color, state.pres.theme);
  return "#FFFFFF";
}

// ---------- shape gallery (full ECMA-376 preset library, categorized like PowerPoint) ----------
const RECENT_KEY = "recentShapes";
function getRecentShapes(): string[] {
  try {
    const v = JSON.parse(localStorage.getItem(RECENT_KEY) ?? "[]");
    return Array.isArray(v) ? v.filter(x => typeof x === "string") : [];
  } catch { return []; }
}
export function pushRecentShape(g: string) {
  const list = [g, ...getRecentShapes().filter(x => x !== g)].slice(0, 14);
  try { localStorage.setItem(RECENT_KEY, JSON.stringify(list)); } catch { /* private mode */ }
}

function ShapePreview({ geom, w = 18, h = 14 }: { geom: PresetGeom; w?: number; h?: number }) {
  return (
    <svg viewBox={`-1.5 -1.5 ${w + 3} ${h + 3}`} width={w + 3} height={h + 3}>
      <path d={presetPath(geom, w, h)} fill="none" fillRule="evenodd" stroke="#555" strokeWidth="1.1" strokeLinejoin="round" />
    </svg>
  );
}

function ShapeCatSection({ label, shapes, pick }: { label: string; shapes: PresetGeom[]; pick: (g: PresetGeom) => void }) {
  if (!shapes.length) return null;
  return (
    <>
      <div className="shape-cat-title">{label}</div>
      <div className="shape-cat-grid">
        {shapes.map(geom => (
          <button key={geom} title={geomLabel(geom)} className="shape-cell" onClick={() => pick(geom)}>
            <ShapePreview geom={geom} />
          </button>
        ))}
      </div>
    </>
  );
}

function CategorizedShapePanel({ onPick, close }: { onPick: (g: PresetGeom) => void; close: () => void }) {
  const pick = (g: PresetGeom) => { onPick(g); close(); };
  const recent = getRecentShapes();
  return (
    <div className="shape-cats">
      {recent.length > 0 && <ShapeCatSection label="Recently used" shapes={recent} pick={pick} />}
      {SHAPE_CATEGORIES.map(c => (
        <ShapeCatSection key={c.label} label={c.label} shapes={c.shapes} pick={pick} />
      ))}
    </div>
  );
}

/** Inline two-row mini gallery (kept right in the toolbar) + chevron for the full panel. */
function InlineShapeGallery({ onPick }: { onPick: (g: PresetGeom) => void }) {
  const recent = getRecentShapes();
  const inline = [...recent, ...SHAPE_GALLERY.map(s => s.geom).filter(g => !recent.includes(g))].slice(0, 14);
  return (
    <div className="inline-shapes">
      <div className="inline-shapes-grid">
        {inline.map(geom => (
          <button key={geom} title={geomLabel(geom)} className="shape-cell sm" onClick={() => onPick(geom)}>
            <svg viewBox="-2 -2 28 22" width="22" height="16">
              <path d={presetPath(geom, 24, 18)} fill="none" fillRule="evenodd" stroke="#555" strokeWidth="1.5" strokeLinejoin="round" />
            </svg>
          </button>
        ))}
      </div>
      <Dropdown title="All shapes" panelClassName="shape-panel" button={<Icon name="chevDown" size={11} />}>
        {close => <CategorizedShapePanel onPick={onPick} close={close} />}
      </Dropdown>
    </div>
  );
}

function ShapeGalleryButton({ onPick }: { onPick: (g: PresetGeom) => void }) {
  return (
    <Dropdown
      className="xhuge"
      title="Insert autoshape"
      panelClassName="shape-panel"
      button={<span className="xhuge-face"><Icon name="shapes" size={26} /><span>Shape <Icon name="chevDown" size={9} /></span></span>}
    >
      {close => <CategorizedShapePanel onPick={onPick} close={close} />}
    </Dropdown>
  );
}

/** PowerPoint-style N×M hover grid for table insertion. */
function TableGridPicker({ onPick }: { onPick: (rows: number, cols: number) => void }) {
  const [hover, setHover] = useState({ r: 3, c: 4 });
  const ROWS = 6, COLS = 8;
  return (
    <Dropdown
      className="xhuge"
      title="Insert table"
      panelClassName="tbl-picker-panel"
      button={<span className="xhuge-face"><TableIcon /><span>Table <Icon name="chevDown" size={9} /></span></span>}
    >
      {close => (
        <div className="tbl-picker">
          <div className="tbl-picker-grid" onMouseLeave={() => setHover({ r: 3, c: 4 })}>
            {Array.from({ length: ROWS }, (_, r) => (
              <div className="tbl-picker-row" key={r}>
                {Array.from({ length: COLS }, (_, c) => (
                  <button
                    key={c}
                    className={`tbl-picker-cell ${r <= hover.r && c <= hover.c ? "on" : ""}`}
                    onMouseEnter={() => setHover({ r, c })}
                    onClick={() => { onPick(r + 1, c + 1); close(); }}
                  />
                ))}
              </div>
            ))}
          </div>
          <div className="tbl-picker-label">{hover.c + 1} × {hover.r + 1} Table</div>
        </div>
      )}
    </Dropdown>
  );
}

export interface ChartVariant {
  id: string;
  label: string;
  kind: ChartKind;
  opts: Partial<import("../model/types").ChartShape>;
}

export const CHART_SECTIONS: { label: string; items: ChartVariant[] }[] = [
  {
    label: "Column",
    items: [
      { id: "col", label: "Clustered Column", kind: "column", opts: {} },
      { id: "colStacked", label: "Stacked Column", kind: "column", opts: { grouping: "stacked" } },
      { id: "colPercent", label: "100% Stacked Column", kind: "column", opts: { grouping: "percentStacked" } },
    ],
  },
  {
    label: "Line",
    items: [
      { id: "line", label: "Line", kind: "line", opts: { marker: false } },
      { id: "lineMarker", label: "Line with Markers", kind: "line", opts: { marker: true } },
      { id: "lineStacked", label: "Stacked Line", kind: "line", opts: { grouping: "stacked", marker: false } },
      { id: "lineSmooth", label: "Smooth Line", kind: "line", opts: { smooth: true, marker: true } },
    ],
  },
  {
    label: "Pie",
    items: [
      { id: "pie", label: "Pie", kind: "pie", opts: {} },
      { id: "doughnut", label: "Doughnut", kind: "doughnut", opts: {} },
    ],
  },
  {
    label: "Bar",
    items: [
      { id: "bar", label: "Clustered Bar", kind: "bar", opts: {} },
      { id: "barStacked", label: "Stacked Bar", kind: "bar", opts: { grouping: "stacked" } },
      { id: "barPercent", label: "100% Stacked Bar", kind: "bar", opts: { grouping: "percentStacked" } },
    ],
  },
  {
    label: "Area",
    items: [
      { id: "area", label: "Area", kind: "area", opts: {} },
      { id: "areaStacked", label: "Stacked Area", kind: "area", opts: { grouping: "stacked" } },
      { id: "areaPercent", label: "100% Stacked Area", kind: "area", opts: { grouping: "percentStacked" } },
    ],
  },
  {
    label: "XY (Scatter)",
    items: [
      { id: "scatter", label: "Scatter", kind: "scatter", opts: { marker: true } },
      { id: "scatterLines", label: "Scatter with Straight Lines", kind: "scatter", opts: { marker: true, smooth: false } },
      { id: "scatterSmooth", label: "Scatter with Smooth Lines", kind: "scatter", opts: { marker: true, smooth: true } },
    ],
  },
  {
    label: "Radar",
    items: [
      { id: "radar", label: "Radar", kind: "radar", opts: { radarStyle: "standard" } },
      { id: "radarMarker", label: "Radar with Markers", kind: "radar", opts: { radarStyle: "marker" } },
      { id: "radarFilled", label: "Filled Radar", kind: "radar", opts: { radarStyle: "filled" } },
    ],
  },
];

/** Miniature previews per variant — drawn, not iconfonts, like PowerPoint's gallery. */
function ChartVariantIcon({ id }: { id: string }) {
  const C = "#5A7FB8", C2 = "#B75B44", G = "#C9C9C9";
  const body = (() => {
    switch (id) {
      case "col": return <><rect x="4" y="11" width="4" height="9" fill={C} /><rect x="9" y="5" width="4" height="15" fill={C2} /><rect x="14" y="8" width="4" height="12" fill={C} /></>;
      case "colStacked": return <><rect x="5" y="10" width="5" height="6" fill={C} /><rect x="5" y="16" width="5" height="4" fill={C2} /><rect x="12" y="6" width="5" height="8" fill={C} /><rect x="12" y="14" width="5" height="6" fill={C2} /></>;
      case "colPercent": return <><rect x="5" y="4" width="5" height="9" fill={C} /><rect x="5" y="13" width="5" height="7" fill={C2} /><rect x="12" y="4" width="5" height="5" fill={C} /><rect x="12" y="9" width="5" height="11" fill={C2} /></>;
      case "line": return <path d="M3 16 8 9l4 3 6-8" fill="none" stroke={C} strokeWidth="1.8" strokeLinejoin="round" />;
      case "lineMarker": return <><path d="M3 16 8 9l4 3 6-8" fill="none" stroke={C} strokeWidth="1.6" strokeLinejoin="round" /><circle cx="3" cy="16" r="1.6" fill={C2} /><circle cx="8" cy="9" r="1.6" fill={C2} /><circle cx="12" cy="12" r="1.6" fill={C2} /><circle cx="18" cy="4" r="1.6" fill={C2} /></>;
      case "lineStacked": return <><path d="M3 16l5-4 4 2 6-5" fill="none" stroke={C} strokeWidth="1.6" /><path d="M3 11l5-5 4 3 6-6" fill="none" stroke={C2} strokeWidth="1.6" /></>;
      case "lineSmooth": return <path d="M3 15c3-7 5 1 8-3s4-7 7-7" fill="none" stroke={C} strokeWidth="1.8" />;
      case "pie": return <><circle cx="11" cy="12" r="8" fill={C} /><path d="M11 12V4a8 8 0 0 1 7.4 5z" fill={C2} /></>;
      case "doughnut": return <path d="M11 4a8 8 0 1 1-8 8h4a4 4 0 1 0 4-4z" fill={C} />;
      case "bar": return <><rect x="3" y="4" width="13" height="3.4" fill={C} /><rect x="3" y="9" width="8" height="3.4" fill={C2} /><rect x="3" y="14" width="15" height="3.4" fill={C} /></>;
      case "barStacked": return <><rect x="3" y="5" width="8" height="4.5" fill={C} /><rect x="11" y="5" width="5" height="4.5" fill={C2} /><rect x="3" y="12" width="6" height="4.5" fill={C} /><rect x="9" y="12" width="8" height="4.5" fill={C2} /></>;
      case "barPercent": return <><rect x="3" y="5" width="10" height="4.5" fill={C} /><rect x="13" y="5" width="6" height="4.5" fill={C2} /><rect x="3" y="12" width="6" height="4.5" fill={C} /><rect x="9" y="12" width="10" height="4.5" fill={C2} /></>;
      case "area": return <path d="M3 18V10l5-4 4 3 6-6v15z" fill={C} opacity="0.85" />;
      case "areaStacked": return <><path d="M3 18v-6l5-3 4 2 6-4v11z" fill={C2} opacity="0.9" /><path d="M3 18v-3l5-2 4 1 6-3v7z" fill={C} /></>;
      case "areaPercent": return <><rect x="3" y="4" width="16" height="14" fill={C2} opacity="0.55" /><path d="M3 18v-8l5 2 4-4 7 3v7z" fill={C} /></>;
      case "scatter": return <>{[[5, 14], [8, 9], [11, 13], [14, 6], [17, 10]].map(([x, y], i) => <circle key={i} cx={x} cy={y} r="1.7" fill={i % 2 ? C2 : C} />)}</>;
      case "scatterLines": return <><path d="M4 15 9 8l5 5 4-8" fill="none" stroke={G} strokeWidth="1.3" />{[[4, 15], [9, 8], [14, 13], [18, 5]].map(([x, y], i) => <circle key={i} cx={x} cy={y} r="1.7" fill={C} />)}</>;
      case "scatterSmooth": return <><path d="M4 15c2-8 4 0 6-4s4-7 8-6" fill="none" stroke={G} strokeWidth="1.3" />{[[4, 15], [10, 11], [18, 5]].map(([x, y], i) => <circle key={i} cx={x} cy={y} r="1.7" fill={C} />)}</>;
      case "radar": return <><polygon points="11,3 18,16 4,16" fill="none" stroke={G} strokeWidth="0.8" /><polygon points="11,6 15.5,14.5 6.5,14.5" fill="none" stroke={C} strokeWidth="1.6" /></>;
      case "radarMarker": return <><polygon points="11,3 18,16 4,16" fill="none" stroke={G} strokeWidth="0.8" /><polygon points="11,6 15.5,14.5 6.5,14.5" fill="none" stroke={C} strokeWidth="1.4" /><circle cx="11" cy="6" r="1.5" fill={C2} /><circle cx="15.5" cy="14.5" r="1.5" fill={C2} /><circle cx="6.5" cy="14.5" r="1.5" fill={C2} /></>;
      case "radarFilled": return <><polygon points="11,3 18,16 4,16" fill="none" stroke={G} strokeWidth="0.8" /><polygon points="11,6 15.5,14.5 6.5,14.5" fill={C} opacity="0.6" /></>;
      default: return <rect x="4" y="4" width="14" height="14" fill={G} />;
    }
  })();
  return <svg width="30" height="24" viewBox="0 0 22 22">{body}</svg>;
}

function ChartMenu({ onPick }: { onPick: (v: ChartVariant) => void }) {
  return (
    <Dropdown
      className="xhuge"
      title="Insert chart"
      panelClassName="chart-panel"
      button={<span className="xhuge-face"><ChartIcon /><span>Chart <Icon name="chevDown" size={9} /></span></span>}
    >
      {close => (
        <div className="shape-cats chart-cats">
          {CHART_SECTIONS.map(sec => (
            <React.Fragment key={sec.label}>
              <div className="shape-cat-title">{sec.label}</div>
              <div className="chart-cat-grid">
                {sec.items.map(v => (
                  <button key={v.id} title={v.label} className="shape-cell chart-cell" onClick={() => { onPick(v); close(); }}>
                    <ChartVariantIcon id={v.id} />
                  </button>
                ))}
              </div>
            </React.Fragment>
          ))}
        </div>
      )}
    </Dropdown>
  );
}

function TableIcon() {
  return (
    <svg width="26" height="26" viewBox="0 0 26 26">
      <rect x="3" y="5" width="20" height="16" rx="1" fill="none" stroke="currentColor" strokeWidth="1.5" />
      <path d="M3 10h20M3 15.5h20M10 5v16M16.5 5v16" stroke="currentColor" strokeWidth="1.2" />
    </svg>
  );
}

function ChartIcon() {
  return (
    <svg width="26" height="26" viewBox="0 0 26 26">
      <path d="M4 22V4" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round" />
      <path d="M4 22h18" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round" />
      <rect x="7" y="12" width="3.6" height="8" fill="currentColor" />
      <rect x="12.4" y="7" width="3.6" height="13" fill="currentColor" />
      <rect x="17.8" y="10" width="3.6" height="10" fill="currentColor" />
    </svg>
  );
}

/**
 * Download slides as PDF / PNG. Routes through the scripting API so the host
 * export permission (and any exportAuthUrl hook) gates the UI the same as a
 * programmatic call — one chokepoint for every path.
 */
async function runExport(kind: "pdf" | "png" | "pngzip" | "pptx", onClose: () => void) {
  onClose();
  try {
    store.setStatus(kind === "pptx" ? "Embedding fonts…" : "Exporting…");
    const base = (store.pres.title || "Presentation").replace(/[^\w.-]+/g, "_") || "Presentation";
    let blob: Blob, name: string;
    if (kind === "pptx") { blob = await editorApi.exportPPTX(); name = `${base}.pptx`; }
    else if (kind === "pdf") { blob = await editorApi.exportPDF(); name = `${base}.pdf`; }
    else if (kind === "pngzip") { blob = await editorApi.exportPNGZip(); name = `${base}-slides.zip`; }
    else { const i = store.getState().selection.slideIndex; blob = await editorApi.exportSlidePNG(i); name = `${base}-slide-${i + 1}.png`; }
    const { downloadBlob } = await import("../util/export");
    downloadBlob(blob, name);
    store.setStatus("Exported");
    setTimeout(() => store.setStatus(null), 2500);
  } catch (e) {
    store.setStatus("Export failed: " + (e as Error).message);
  }
}

function FileMenu({ onClose, onOpenFile, onSave, onImportPattern }: { onClose: () => void; onOpenFile: () => void; onSave: () => void; onImportPattern: () => void }) {
  return (
    <div className="file-overlay" onClick={onClose}>
      <div className="file-menu" onClick={e => e.stopPropagation()}>
        <button className="file-close" onClick={onClose}>‹ &nbsp;Close Menu</button>
        <div className="file-items">
          {uiConfig.save && (
            <button className="file-item" onClick={() => { onSave(); onClose(); }}>
              <Icon name="save" size={18} /> Download as… <span className="file-hint">.pptx</span>
            </button>
          )}
          {uiConfig.export && permissions.export && <>
            <button className="file-item" onClick={() => runExport("pptx", onClose)}>
              <Icon name="save" size={18} /> Export for PowerPoint <span className="file-hint">.pptx · fonts embedded</span>
            </button>
            <button className="file-item" onClick={() => runExport("pdf", onClose)}>
              <Icon name="doc" size={18} /> Export as PDF <span className="file-hint">.pdf</span>
            </button>
            <button className="file-item" onClick={() => runExport("png", onClose)}>
              <Icon name="image" size={18} /> Export current slide <span className="file-hint">.png</span>
            </button>
            <button className="file-item" onClick={() => runExport("pngzip", onClose)}>
              <Icon name="image" size={18} /> Export all slides <span className="file-hint">.png .zip</span>
            </button>
          </>}
          <div className="file-div" />
          {uiConfig.open && (
            <button className="file-item" onClick={() => { onOpenFile(); onClose(); }}>
              <Icon name="open" size={18} /> Open… <span className="file-hint">.pptx</span>
            </button>
          )}
          {uiConfig.importPattern && (
            <button className="file-item" onClick={() => { onImportPattern(); onClose(); }}>
              <Icon name="doc" size={18} /> Import Pattern… <span className="file-hint">JSON</span>
            </button>
          )}
          {uiConfig.newPresentation && (
            <button
              className="file-item"
              onClick={() => {
                if (store.getState().dirty && !confirm("Discard current presentation and create a new one?")) return;
                store.newPresentation();
                onClose();
              }}
            >
              <Icon name="doc" size={18} /> Create New
            </button>
          )}
          <div className="file-div" />
          <div className="file-about">
            <b>Presentation Editor</b>
            <p>A browser-based slide editor. Files open and save as standard .pptx presentations.</p>
          </div>
        </div>
      </div>
    </div>
  );
}

/** execCommand('fontSize') only supports 1-7 — set 7 then rewrite to exact px. */
function execFontSize(pt: number) {
  document.execCommand("fontSize", false, "7");
  document.querySelectorAll('.text-edit font[size="7"]').forEach(f => {
    const span = document.createElement("span");
    span.style.fontSize = `${(pt * 4) / 3}px`;
    while (f.firstChild) span.appendChild(f.firstChild);
    f.replaceWith(span);
  });
}
