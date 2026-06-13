import JSZip from "jszip";
import { DOMParser as XmldomParser, XMLSerializer as XmldomSerializer } from "@xmldom/xmldom";
import { buildPptx } from "../src/ooxml/write";
import { PARSE_LIMITS, parsePptx, setDOMParser, setXMLSerializer } from "../src/ooxml/parse";
import { store } from "../src/state/store";
import { importPatternSlide } from "../src/ooxml/pattern";
import { makeChart, makeShape, makeSlide, makeTable, makeTextBox, newPresentation } from "../src/model/defaults";
import { PRESET_NAMES, presetPaths } from "../src/render/presetGeom";
import type { ChartShape, MediaItem, SpShape, TableShape } from "../src/model/types";

/** 1x1 red PNG */
const PNG_BYTES = Uint8Array.from(atob(
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=="
), c => c.charCodeAt(0));

export async function runSmoke(patternJson?: string): Promise<{ zipBytes: Uint8Array; report: string[] }> {
  setDOMParser(XmldomParser as never);
  setXMLSerializer(XmldomSerializer as never);
  const report: string[] = [];
  const ok = (cond: boolean, label: string, detail = "") =>
    report.push(`${cond ? "ok  " : "FAIL"} ${label}${cond || !detail ? "" : " — " + detail}`);

  // ---------- build a representative deck ----------
  const pres = newPresentation();
  pres.title = "Smoke Deck";
  // slide 1 (title layout) — give the title some styled text
  const title = pres.slides[0].shapes[0] as SpShape;
  title.text!.paragraphs = [{
    runs: [
      { text: "Hello ", sizePt: 44, font: "Calibri Light", color: { kind: "scheme", slot: "dk1" } },
      { text: "OOXML", sizePt: 44, bold: true, italic: true, font: "Calibri Light", color: { kind: "scheme", slot: "accent1" } },
    ],
    align: "ctr", bullet: "none", level: 0,
  }];
  // slide 2: shapes, bullets, rotation, line, picture, background, transition
  const s2 = makeSlide("blank");
  s2.background = { kind: "solid", color: { kind: "srgb", hex: "FFF2CC" } };
  s2.transition = { type: "push", dir: "u", speed: "fast" };
  const star = makeShape("star5", 1000000, 1000000, 2000000, 2000000);
  star.rot = 30;
  star.fill = { kind: "solid", color: { kind: "scheme", slot: "accent4", lumMod: 60, lumOff: 40 } };
  const arrow = makeShape("rightArrow", 4000000, 1200000, 2500000, 1200000);
  arrow.flipH = true;
  const line = makeShape("line", 1000000, 4000000, 3000000, 800000);
  const box = makeTextBox(4000000, 3500000, 4000000, 2000000, "");
  box.text!.paragraphs = [
    { runs: [{ text: "First bullet", sizePt: 20, font: "Arial", color: { kind: "srgb", hex: "404040" } }], align: "l", bullet: "char", level: 0 },
    { runs: [{ text: "Nested & <escaped> \"text\"", sizePt: 18, font: "Arial", color: { kind: "srgb", hex: "C00000" } }], align: "l", bullet: "char", level: 1 },
    { runs: [{ text: "Numbered", sizePt: 18, font: "Georgia", color: { kind: "scheme", slot: "accent2" } }], align: "l", bullet: "num", level: 0 },
  ];
  const media = new Map<string, MediaItem>();
  media.set("m1", { id: "m1", mime: "image/png", bytes: PNG_BYTES, dataUrl: "data:image/png;base64," });
  s2.shapes.push(star, arrow, line, box, {
    kind: "pic", id: "p1", name: "Picture 1", mediaId: "m1",
    x: 8000000, y: 1000000, w: 1500000, h: 1500000, rot: 0,
  });
  pres.slides.push(s2);

  // ---------- write ----------
  const zip = await buildPptx(pres, media);
  const zipBytes = await zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
  ok(zipBytes.length > 1000, "pptx produced", `${zipBytes.length} bytes`);

  // ---------- structural validation: walk relationships like a consumer would ----------
  const names = Object.keys(zip.files).filter(n => !zip.files[n].dir);
  const required = [
    "[Content_Types].xml", "_rels/.rels", "ppt/presentation.xml", "ppt/_rels/presentation.xml.rels",
    "ppt/slideMasters/slideMaster1.xml", "ppt/slideMasters/_rels/slideMaster1.xml.rels",
    "ppt/slideLayouts/slideLayout1.xml", "ppt/theme/theme1.xml",
    "ppt/slides/slide1.xml", "ppt/slides/slide2.xml", "ppt/slides/_rels/slide2.xml.rels",
    "ppt/media/image1.png", "docProps/core.xml", "docProps/app.xml",
  ];
  for (const r of required) ok(names.includes(r), `part ${r}`);

  // XML well-formedness of every xml part (xmldom reports via errorHandler)
  for (const n of names.filter(n => n.endsWith(".xml") || n.endsWith(".rels"))) {
    const text = await zip.file(n)!.async("text");
    let bad = "";
    try {
      const doc = new XmldomParser({
        onError: (level: string, msg: string) => { if (level !== "warn") bad = msg; },
      } as never).parseFromString(text, "application/xml") as unknown as Document;
      if (!doc.documentElement) bad = "no root";
    } catch (e) { bad = String(e); }
    ok(!bad, `well-formed ${n}`, bad.slice(0, 120));
  }

  // content types covers every part
  const ct = await zip.file("[Content_Types].xml")!.async("text");
  for (const n of names) {
    if (n === "[Content_Types].xml") continue;
    const ext = n.split(".").pop()!;
    const covered = ct.includes(`PartName="/${n}"`) || ct.includes(`Extension="${ext}"`);
    ok(covered, `content-type covers ${n}`);
  }

  // every relationship target must exist
  for (const reln of names.filter(n => n.endsWith(".rels"))) {
    const text = await zip.file(reln)!.async("text");
    const dir = reln.replace(/_rels\/[^/]+$/, "").replace(/\/$/, "");
    for (const m of text.matchAll(/Target="([^"]+)"/g)) {
      const t = m[1];
      if (t.startsWith("http")) continue;
      const parts = (dir ? dir.split("/") : []).concat(t.split("/"));
      const out: string[] = [];
      for (const p of parts) {
        if (p === "." || p === "") continue;
        if (p === "..") out.pop(); else out.push(p);
      }
      const resolved = out.join("/");
      ok(names.includes(resolved), `rel target ${t} (from ${reln})`, `resolved: ${resolved}`);
    }
  }

  // ---------- re-parse through our own reader ----------
  const { pres: re, media: reMedia, warnings } = await parsePptx(zipBytes, "Smoke Deck.pptx");
  ok(re.slides.length === 2, "round-trip slide count", `${re.slides.length}`);
  ok(re.slideWidth === 12192000 && re.slideHeight === 6858000, "slide size");
  const reTitle = re.slides[0].shapes.find(s => s.kind === "sp" && s.text?.paragraphs.some(p => p.runs.some(r => r.text.includes("OOXML"))));
  ok(!!reTitle, "title text survived");
  const boldRun = re.slides[0].shapes.flatMap(s => s.kind === "sp" ? s.text?.paragraphs ?? [] : []).flatMap(p => p.runs).find(r => r.text === "OOXML");
  ok(!!boldRun?.bold && !!boldRun?.italic, "run bold/italic survived");
  ok(boldRun?.color.kind === "scheme" && boldRun.color.slot === "accent1", "scheme color survived", JSON.stringify(boldRun?.color));
  const reS2 = re.slides[1];
  ok(reS2.background?.kind === "solid" && (reS2.background.color as { hex?: string }).hex === "FFF2CC", "background survived");
  ok(reS2.transition?.type === "push" && reS2.transition.dir === "u" && reS2.transition.speed === "fast", "transition survived", JSON.stringify(reS2.transition));
  const reStar = reS2.shapes.find(s => s.kind === "sp" && s.geom === "star5") as SpShape | undefined;
  ok(!!reStar && Math.abs(reStar.rot - 30) < 0.01, "star5 + rotation survived", `rot=${reStar?.rot}`);
  ok(reStar?.fill.kind === "solid" && reStar.fill.color.kind === "scheme" && reStar.fill.color.lumMod === 60 && reStar.fill.color.lumOff === 40, "lumMod/lumOff survived", JSON.stringify(reStar?.fill));
  const reArrow = reS2.shapes.find(s => s.kind === "sp" && s.geom === "rightArrow");
  ok(!!reArrow?.flipH, "flipH survived");
  const bullets = reS2.shapes.flatMap(s => s.kind === "sp" ? s.text?.paragraphs ?? [] : []);
  ok(bullets.some(p => p.bullet === "char" && p.level === 1 && p.runs.some(r => r.text.includes("<escaped>"))), "nested bullet + escaping survived");
  ok(bullets.some(p => p.bullet === "num"), "numbered list survived");
  const rePic = reS2.shapes.find(s => s.kind === "pic");
  ok(!!rePic, "picture survived");
  ok(rePic ? reMedia.get((rePic as { mediaId: string }).mediaId)!.bytes.length === PNG_BYTES.length : false, "media bytes survived");
  ok(warnings.length === 0, "no import warnings", warnings.join("; "));

  // ---------- SlideBazaar pattern: import -> render model -> pptx round-trip ----------
  if (patternJson) {
    const { slide, warnings: pw, values } = await importPatternSlide(patternJson);
    ok(slide.shapes.length === 11, "pattern: shape count (1 bg + 3 group + 1 graphic + 6 text)", `${slide.shapes.length}`);
    ok(pw.length === 0, "pattern: no import warnings", pw.join("; "));
    const sps = slide.shapes.filter((s): s is SpShape => s.kind === "sp");
    const bg = sps[0];
    ok(bg.fill.kind === "gradient" && bg.fill.stops.length === 4 && Math.abs(bg.fill.angle - 320.03) < 0.1,
      "pattern: background gradient (4 stops, 320°)", JSON.stringify(bg.fill.kind === "gradient" ? { stops: bg.fill.stops.length, angle: bg.fill.angle } : {}));
    const custs = sps.filter(s => s.custPath);
    ok(custs.length === 4, "pattern: 4 freeform custGeom shapes parsed", `${custs.length}`);
    ok(custs.every(s => s.rawGeomXml?.includes("custGeom")), "pattern: raw custGeom preserved for export");
    const grpChild = custs.find(s => s.name === "Freeform: Shape 8");
    ok(grpChild?.fill.kind === "solid" && grpChild.fill.color.kind === "scheme" && grpChild.fill.color.slot === "lt1" && grpChild.fill.color.alpha === 14,
      "pattern: grpFill inherited (bg1 @ 14% alpha)", JSON.stringify(grpChild?.fill));
    const graphic = sps.find(s => s.name === "Graphic 69");
    ok(graphic?.fill.kind === "gradient" && graphic.fill.stops.some(st => st.color.alpha === 0) && graphic.fill.stops.some(st => st.color.alpha === 65),
      "pattern: alpha gradient stops on chart graphic");
    const texts = sps.filter(s => s.text?.paragraphs.some(p => p.runs.some(r => r.text)));
    const allText = texts.flatMap(s => s.text!.paragraphs.flatMap(p => p.runs.map(r => r.text))).join(" | ");
    ok(!allText.includes("{{"), "pattern: all {{placeholders}} substituted", allText);
    ok(allText.includes(values["coverpage_title"]), "pattern: title value present", values["coverpage_title"]);
    const title = sps.find(s => s.name === "content.coverpage_title");
    ok(title?.text?.paragraphs[0].runs[0]?.sizePt === 72, "pattern: 72pt title size");
    const year = sps.find(s => s.name === "content.coverpage_year");
    const yearColor = year?.text?.paragraphs[0].runs[0]?.color;
    ok(yearColor?.kind === "scheme" && yearColor.slot === "accent1", "pattern: year keeps accent1 scheme ref");
    // round-trip the imported slide through a full pptx
    const pres2 = newPresentation();
    pres2.slides = [slide];
    const zip2 = await buildPptx(pres2, new Map());
    const bytes2 = await zip2.generateAsync({ type: "uint8array" });
    const re2 = await parsePptx(bytes2, "pattern-roundtrip.pptx");
    const sps2 = re2.pres.slides[0].shapes.filter((s): s is SpShape => s.kind === "sp");
    ok(sps2.filter(s => s.custPath).length === 4, "pattern round-trip: custGeom survives pptx write/read");
    const bg2 = sps2[0];
    ok(bg2.fill.kind === "gradient" && bg2.fill.stops.length === 4, "pattern round-trip: gradient survives");
    const grp2 = sps2.find(s => s.name === "Freeform: Shape 8");
    ok(grp2?.fill.kind === "solid" && grp2.fill.color.kind === "scheme" && grp2.fill.color.alpha === 14, "pattern round-trip: alpha survives");
    ok(re2.warnings.length === 0, "pattern round-trip: no warnings", re2.warnings.join("; "));
  } else {
    report.push("skip pattern test (no pattern json provided)");
  }

  // ---------- tables, charts, notes round-trip ----------
  {
    const pres3 = newPresentation();
    const tbl = makeTable(3, 4, 914400, 914400, 7315200, 1828800);
    tbl.cells[0].forEach((c, i) => { c.text.paragraphs[0].runs[0].text = `Head ${i + 1}`; });
    tbl.cells[1][0].text.paragraphs[0].runs[0].text = "Cell A <&>";
    const col = makeChart("column", 914400, 3200400, 5486400, 2743200);
    col.title = "Revenue by quarter";
    const pie = makeChart("pie", 6858000, 3200400, 4114800, 2743200);
    pie.legend = true;
    pres3.slides = [{ ...makeSlide("blank"), shapes: [tbl, col, pie], notes: "Speaker line one\nSecond line" }];

    const zip3 = await buildPptx(pres3, new Map());
    const bytes3 = await zip3.generateAsync({ type: "uint8array" });
    ok(!!zip3.file("ppt/charts/chart1.xml") && !!zip3.file("ppt/charts/chart2.xml"), "charts: parts written");
    ok(!!zip3.file("ppt/notesSlides/notesSlide1.xml") && !!zip3.file("ppt/notesMasters/notesMaster1.xml"), "notes: parts written");

    const re3 = await parsePptx(bytes3, "graphics.pptx");
    const s3 = re3.pres.slides[0];
    const rtbl = s3.shapes.find(s => s.kind === "table") as TableShape | undefined;
    ok(!!rtbl, "table survived round-trip");
    ok(rtbl?.cells.length === 3 && rtbl.cells[0].length === 4, "table: 3×4 grid", `${rtbl?.cells.length}×${rtbl?.cells[0]?.length}`);
    ok(rtbl?.firstRow === true && rtbl.bandRow === true, "table: header/banded flags");
    ok(rtbl?.cells[0][1].text.paragraphs[0].runs[0]?.text === "Head 2", "table: header text");
    ok(rtbl?.cells[1][0].text.paragraphs[0].runs.map(r => r.text).join("") === "Cell A <&>", "table: cell text + escaping");
    const headFill = rtbl?.cells[0][0].fill;
    ok(headFill?.kind === "solid" && headFill.color.kind === "scheme" && headFill.color.slot === "accent1", "table: header fill accent1");

    // table design: style flags, dark family fills, border modes, merges
    const tblD = makeTable(4, 4, 914400, 914400, 7315200, 2438400);
    tblD.styleFamily = "dark"; tblD.accent = "accent3";
    tblD.totalRow = true; tblD.firstCol = true; tblD.lastCol = true; tblD.bandCol = true;
    tblD.borderMode = "none";
    tblD.cells[1][1].gridSpan = 2; tblD.cells[1][2].merged = "h";
    tblD.cells[2][0].rowSpan = 2; tblD.cells[3][0].merged = "v";
    const presT = newPresentation();
    presT.slides = [{ ...makeSlide("blank"), shapes: [tblD] }];
    const zipT = await buildPptx(presT, new Map());
    const slideTxt = await zipT.file("ppt/slides/slide1.xml")!.async("text");
    ok(slideTxt.includes('lastRow="1"') && slideTxt.includes('firstCol="1"') && slideTxt.includes('lastCol="1"') && slideTxt.includes('bandCol="1"'),
      "table design: tblPr flags written");
    ok(!/<a:lnL[^>]*><a:solidFill>/.test(slideTxt), "table design: borderMode none writes noFill edges");
    const reT = await parsePptx(await zipT.generateAsync({ type: "uint8array" }), "tbl-design.pptx");
    const rT = reT.pres.slides[0].shapes.find(s => s.kind === "table") as TableShape;
    ok(rT.totalRow === true && rT.firstCol === true && rT.lastCol === true && rT.bandCol === true,
      "table design: style-option flags survive round-trip");
    ok(rT.cells[1][1].gridSpan === 2 && rT.cells[1][2].merged === "h", "table design: column merge survives");
    ok(rT.cells[2][0].rowSpan === 2 && rT.cells[3][0].merged === "v", "table design: row merge survives");
    const darkBody = rT.cells[1][3].fill; // unmerged body cell — dark style baked as explicit fill
    ok(darkBody?.kind === "solid" && darkBody.color.kind === "scheme" && darkBody.color.slot === "dk1" && (darkBody.color.lumOff ?? 0) > 0,
      "table design: dark-family fill baked into cells", JSON.stringify(darkBody));
    const charts = s3.shapes.filter((s): s is ChartShape => s.kind === "chart");
    ok(charts.length === 2, "charts survived round-trip", `${charts.length}`);
    const rcol = charts.find(c => c.chart === "column");
    ok(!!rcol, "chart: column type");
    ok(rcol?.title === "Revenue by quarter", "chart: title", rcol?.title);
    ok(rcol?.series.length === 2 && rcol.series[0].values.join(",") === "4.3,2.5,3.5,4.5", "chart: series values", rcol?.series[0]?.values.join(","));
    ok(rcol?.series.map(x => x.name).join("|") === "Series 1|Series 2", "chart: series names via c:tx/c:v", rcol?.series.map(x => x.name).join("|"));
    const chartXml = await zip3.file("ppt/charts/chart1.xml")!.async("text");
    ok(!chartXml.includes("<c:tx><c:strLit>"), "chart: no invalid strLit in c:tx (PowerPoint repair trigger)");
    ok(chartXml.includes("<c:tx><c:v>"), "chart: series name as c:v literal");
    ok(rcol?.categories.join(",") === "Q1,Q2,Q3,Q4", "chart: categories");
    ok(rcol?.legend === true, "chart: legend");
    ok(charts.some(c => c.chart === "pie"), "chart: pie type");
    ok(s3.notes === "Speaker line one\nSecond line", "notes survived round-trip", JSON.stringify(s3.notes));
    ok(re3.warnings.length === 0, "graphics round-trip: no warnings", re3.warnings.join("; "));
  }

  // ---------- svg image: dual embed (png fallback + asvg:svgBlip) ----------
  {
    const svgText = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><circle cx="12" cy="12" r="10" fill="#4F9CFF"/></svg>`;
    const svgBytes = new TextEncoder().encode(svgText);
    const pres4 = newPresentation();
    const media4 = new Map<string, MediaItem>();
    media4.set("m-svg", {
      id: "m-svg", mime: "image/svg+xml", bytes: svgBytes,
      dataUrl: "data:image/svg+xml;base64,", pngFallback: PNG_BYTES,
    });
    pres4.slides = [{
      ...makeSlide("blank"),
      shapes: [{ kind: "pic", id: "p1", name: "icon.svg", mediaId: "m-svg", x: 914400, y: 914400, w: 914400, h: 914400, rot: 0 }],
    }];
    const zip4 = await buildPptx(pres4, media4);
    ok(!!zip4.file("ppt/media/image1.png") && !!zip4.file("ppt/media/image2.svg"), "svg: fallback png + svg files written");
    const slide4 = await zip4.file("ppt/slides/slide1.xml")!.async("text");
    ok(slide4.includes("asvg:svgBlip") && slide4.includes("{96DAC541-7B7A-43D3-8B79-37D633B846F1}"), "svg: asvg extension emitted");
    const bytes4 = await zip4.generateAsync({ type: "uint8array" });
    const re4 = await parsePptx(bytes4, "svg.pptx");
    const pic4 = re4.pres.slides[0].shapes.find(s => s.kind === "pic");
    ok(!!pic4, "svg: picture survived");
    const m4 = pic4 ? re4.media.get((pic4 as { mediaId: string }).mediaId) : undefined;
    ok(m4?.mime === "image/svg+xml", "svg: vector kept as primary media", m4?.mime);
    ok(m4 ? new TextDecoder().decode(m4.bytes).includes('width="24"') : false, "svg: width/height normalized for renderer");
    ok(m4?.pngFallback?.length === PNG_BYTES.length, "svg: png fallback preserved");
    ok(re4.warnings.length === 0, "svg round-trip: no warnings", re4.warnings.join("; "));
  }

  // ---------- full preset geometry library ----------
  {
    const bad: string[] = [];
    for (const name of PRESET_NAMES) {
      for (const [w, h] of [[200, 150], [120, 120]] as const) {
        const paths = presetPaths(name, w, h);
        if (!paths || !paths.length || paths.some(p => !p.d || /NaN|Infinity/i.test(p.d))) {
          bad.push(`${name}@${w}x${h}`);
        }
      }
    }
    ok(PRESET_NAMES.length === 187, "presets: 187 spec definitions loaded", String(PRESET_NAMES.length));
    ok(bad.length === 0, "presets: every shape evaluates to clean paths at two sizes", bad.slice(0, 6).join(", "));

    const pres5 = newPresentation();
    const rr = makeShape("roundRect", 914400, 914400, 1828800, 1219200);
    rr.adj = { adj: 40000 };
    const callout = makeShape("wedgeRectCallout", 3657600, 914400, 1828800, 1219200);
    callout.adj = { adj1: -30000, adj2: 70000 };
    const exotic = ["star7", "flowChartDecision", "actionButtonHome", "mathPlus", "bracketPair", "cloudCallout"]
      .map((g, i) => makeShape(g, 914400 + i * 1000000, 2743200, 914400, 914400));
    pres5.slides = [{ ...makeSlide("blank"), shapes: [rr, callout, ...exotic] }];
    const bytes5 = await (await buildPptx(pres5, new Map())).generateAsync({ type: "uint8array" });
    const re5 = await parsePptx(bytes5, "shapes.pptx");
    const sps5 = re5.pres.slides[0].shapes.filter((s): s is SpShape => s.kind === "sp");
    const wantGeoms = "roundRect,wedgeRectCallout,star7,flowChartDecision,actionButtonHome,mathPlus,bracketPair,cloudCallout";
    ok(sps5.map(s => s.geom).join(",") === wantGeoms, "presets: exotic geometry names survive round-trip", sps5.map(s => s.geom).join(","));
    ok(sps5[0].adj?.adj === 40000, "presets: roundRect adjust value survives");
    ok(sps5[1].adj?.adj1 === -30000 && sps5[1].adj?.adj2 === 70000, "presets: callout adj1/adj2 survive");
    ok(re5.warnings.length === 0, "presets round-trip: no warnings", re5.warnings.join("; "));
  }

  // ---------- line arrowheads + dash styles round-trip ----------
  {
    const presL = newPresentation();
    const ln = makeShape("line", 914400, 914400, 3000000, 0);
    ln.line = {
      fill: { kind: "solid", color: { kind: "srgb", hex: "C00000" } },
      widthPt: 3,
      dash: "lgDashDot",
      headEnd: { type: "oval", w: "lg", len: "lg" },
      tailEnd: { type: "triangle", w: "med", len: "lg" },
    };
    presL.slides = [{ ...makeSlide("blank"), shapes: [ln] }];
    const slideL = await (await buildPptx(presL, new Map())).file("ppt/slides/slide1.xml")!.async("text");
    ok(slideL.includes('<a:tailEnd type="triangle"') && slideL.includes('<a:headEnd type="oval"'), "arrows: head/tail ends written");
    ok(slideL.includes('<a:prstDash val="lgDashDot"/>'), "arrows: extended dash written");
    const reL = await parsePptx(await (await buildPptx(presL, new Map())).generateAsync({ type: "uint8array" }), "lines.pptx");
    const rln = reL.pres.slides[0].shapes.find(s => s.kind === "sp") as SpShape;
    ok(rln.line.headEnd?.type === "oval" && rln.line.headEnd.len === "lg" && rln.line.headEnd.w === "lg", "arrows: begin arrowhead survives", JSON.stringify(rln.line.headEnd));
    ok(rln.line.tailEnd?.type === "triangle" && rln.line.tailEnd.len === "lg", "arrows: end arrowhead survives", JSON.stringify(rln.line.tailEnd));
    ok(rln.line.dash === "lgDashDot", "arrows: extended dash survives round-trip", String(rln.line.dash));
    // a plain line gains no arrows
    const plain = makeShape("line", 0, 0, 1000000, 0);
    const reP = await parsePptx(await (await buildPptx({ ...newPresentation(), slides: [{ ...makeSlide("blank"), shapes: [plain] }] }, new Map())).generateAsync({ type: "uint8array" }), "plainline.pptx");
    const rp = reP.pres.slides[0].shapes.find(s => s.kind === "sp") as SpShape;
    ok(!rp.line.headEnd && !rp.line.tailEnd, "arrows: plain line stays arrowless");
    ok(reL.warnings.length === 0, "arrows round-trip: no warnings");
  }

  // ---------- text features (case/columns/super/highlight) + chart variants ----------
  {
    const pres6 = newPresentation();
    const tb = makeTextBox(914400, 914400, 5486400, 1828800, "E = mc2 highlighted");
    const runs = tb.text!.paragraphs[0].runs;
    runs[0].text = "E = mc";
    runs.push({ text: "2", sizePt: 18, font: "Arial", color: { kind: "scheme", slot: "dk1" }, baseline: 30 });
    runs.push({ text: " marked", sizePt: 18, font: "Arial", color: { kind: "scheme", slot: "dk1" }, highlight: { kind: "srgb", hex: "FFFF00" } });
    tb.text!.columns = 2;
    tb.text!.colSpacing = 360000;

    const colStacked = makeChart("column", 914400, 3200400, 4114800, 2286000, { grouping: "percentStacked" });
    const scatter = makeChart("scatter", 5486400, 3200400, 3200400, 2286000, { marker: true, smooth: true });
    const radar = makeChart("radar", 9144000, 3200400, 2743200, 2286000, { radarStyle: "filled" });

    pres6.slides = [{ ...makeSlide("blank"), shapes: [tb, colStacked, scatter, radar] }];
    const bytes6 = await (await buildPptx(pres6, new Map())).generateAsync({ type: "uint8array" });
    const re6 = await parsePptx(bytes6, "features.pptx");
    const s6 = re6.pres.slides[0];
    const rtb = s6.shapes.find(s => s.kind === "sp") as SpShape;
    ok(rtb.text?.columns === 2 && rtb.text.colSpacing === 360000, "text: columns round-trip", `${rtb.text?.columns}/${rtb.text?.colSpacing}`);
    const rRuns = rtb.text!.paragraphs[0].runs;
    ok(rRuns.some(r => r.baseline === 30), "text: superscript baseline round-trip");
    ok(rRuns.some(r => r.highlight && r.highlight.kind === "srgb" && r.highlight.hex === "FFFF00"), "text: highlight round-trip");
    const rCharts = s6.shapes.filter((s): s is ChartShape => s.kind === "chart");
    ok(rCharts.find(c => c.chart === "column")?.grouping === "percentStacked", "chart: percentStacked grouping survives");
    const rScatter = rCharts.find(c => c.chart === "scatter");
    ok(!!rScatter && rScatter.smooth === true && rScatter.marker === true, "chart: scatter smooth+markers survive");
    ok(rScatter?.categories.map(parseFloat).every(n => Number.isFinite(n)) === true, "chart: scatter x-values numeric");
    ok(rCharts.find(c => c.chart === "radar")?.radarStyle === "filled", "chart: radar filled style survives");
    ok(re6.warnings.length === 0, "features round-trip: no warnings", re6.warnings.join("; "));
  }

  // ---------- theme fonts: +mj-lt / +mn-lt stay symbolic ----------
  {
    const pres7 = newPresentation(); // title-layout slide: title seeds +mj-lt, subtitle +mn-lt
    pres7.theme = { ...pres7.theme, majorFont: "Georgia", minorFont: "Verdana" };
    const bytes7 = await (await buildPptx(pres7, new Map())).generateAsync({ type: "uint8array" });
    const re7 = await parsePptx(bytes7, "fonts.pptx");
    const sps7 = re7.pres.slides[0].shapes.filter((s): s is SpShape => s.kind === "sp");
    const title7 = sps7.find(s => s.name.includes("Title"));
    const sub7 = sps7.find(s => s.name.includes("Subtitle"));
    ok(title7?.text?.paragraphs[0].runs[0]?.font === "+mj-lt", "theme fonts: heading ref survives round-trip", title7?.text?.paragraphs[0].runs[0]?.font);
    ok(sub7?.text?.paragraphs[0].runs[0]?.font === "+mn-lt", "theme fonts: body ref survives round-trip", sub7?.text?.paragraphs[0].runs[0]?.font);
    ok(re7.pres.theme.majorFont === "Georgia" && re7.pres.theme.minorFont === "Verdana", "theme fonts: font pair in theme part", `${re7.pres.theme.majorFont}/${re7.pres.theme.minorFont}`);
  }

  // ---------- multi-image bug: AlternateContent wrappers + picture fills + crop ----------
  {
    const NS = `xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"`;
    const mkPic = (id: number, rid: string, x: number) =>
      `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="Pic${id}"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>` +
      `<p:blipFill><a:blip r:embed="${rid}"/><a:srcRect l="10000" t="5000" r="20000" b="0"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
      `<p:spPr><a:xfrm><a:off x="${x}" y="914400"/><a:ext cx="914400" cy="914400"/></a:xfrm>` +
      `<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 25000"/></a:avLst></a:prstGeom></p:spPr></p:pic>`;
    const dummySp = `<p:sp><p:nvSpPr><p:cNvPr id="9" name="Modern"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1" cy="1"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>`;
    const imgFillSp = `<p:sp><p:nvSpPr><p:cNvPr id="5" name="ImgFill"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="5486400" y="914400"/><a:ext cx="1828800" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:blipFill rotWithShape="1"><a:blip r:embed="rId2"/><a:stretch><a:fillRect/></a:stretch></a:blipFill><a:ln><a:noFill/></a:ln></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>`;
    const slideXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<p:sld ${NS}><p:cSld><p:spTree>` +
      `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>` +
      mkPic(2, "rId2", 914400) +
      `<mc:AlternateContent><mc:Choice Requires="we-do-not-support-this">${dummySp}</mc:Choice><mc:Fallback>${mkPic(3, "rId3", 2438400)}</mc:Fallback></mc:AlternateContent>` +
      `<mc:AlternateContent><mc:Choice>${mkPic(4, "rId4", 3962400)}</mc:Choice></mc:AlternateContent>` +
      imgFillSp +
      `</p:spTree></p:cSld></p:sld>`;

    const PNG2 = new Uint8Array([...PNG_BYTES, 1]);
    const PNG3 = new Uint8Array([...PNG_BYTES, 1, 2]);
    const zip8 = new JSZip();
    const relsNs = `xmlns="http://schemas.openxmlformats.org/package/2006/relationships"`;
    zip8.file("_rels/.rels", `<?xml version="1.0"?><Relationships ${relsNs}><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>`);
    zip8.file("ppt/presentation.xml", `<?xml version="1.0"?><p:presentation ${NS}><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/></p:presentation>`);
    zip8.file("ppt/_rels/presentation.xml.rels", `<?xml version="1.0"?><Relationships ${relsNs}><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/></Relationships>`);
    zip8.file("ppt/slides/slide1.xml", slideXml);
    zip8.file("ppt/slides/_rels/slide1.xml.rels", `<?xml version="1.0"?><Relationships ${relsNs}>` +
      `<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/img1.png"/>` +
      `<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/img2.png"/>` +
      `<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/img3.png"/>` +
      `</Relationships>`);
    zip8.file("ppt/media/img1.png", PNG_BYTES);
    zip8.file("ppt/media/img2.png", PNG2);
    zip8.file("ppt/media/img3.png", PNG3);

    const re8 = await parsePptx(await zip8.generateAsync({ type: "uint8array" }), "multi.pptx");
    const s8 = re8.pres.slides[0];
    const pics8 = s8.shapes.filter(s => s.kind === "pic");
    ok(pics8.length === 3, "multi-image: all 3 pictures parsed (incl. AlternateContent)", String(pics8.length));
    ok(new Set(pics8.map(p => (p as { mediaId: string }).mediaId)).size === 3, "multi-image: 3 distinct media items");
    const p0 = pics8[0] as import("../src/model/types").PicShape;
    ok(!!p0.srcRect && Math.abs(p0.srcRect.l - 0.1) < 1e-6 && Math.abs(p0.srcRect.t - 0.05) < 1e-6 && Math.abs(p0.srcRect.r - 0.2) < 1e-6, "multi-image: srcRect crop parsed", JSON.stringify(p0.srcRect));
    ok(p0.geom === "roundRect" && p0.adj?.adj === 25000, "multi-image: rounded frame parsed");
    const fillSp = s8.shapes.find(s => s.kind === "sp" && s.name === "ImgFill") as SpShape;
    ok(fillSp?.fill.kind === "image", "multi-image: shape picture-fill kept as image", fillSp?.fill.kind);
    ok(fillSp?.fill.kind === "image" && re8.media.get(fillSp.fill.mediaId)?.bytes.length === PNG_BYTES.length, "multi-image: fill media resolved");

    // full round-trip back through our writer
    const re9 = await parsePptx(await (await buildPptx(re8.pres, re8.media)).generateAsync({ type: "uint8array" }), "multi2.pptx");
    const pics9 = re9.pres.slides[0].shapes.filter(s => s.kind === "pic") as import("../src/model/types").PicShape[];
    ok(pics9.length === 3, "multi-image round-trip: 3 pictures survive");
    ok(pics9[0].srcRect !== undefined && pics9[0].geom === "roundRect", "multi-image round-trip: crop + corners survive");
    const fillSp9 = re9.pres.slides[0].shapes.find(s => s.kind === "sp" && s.name === "ImgFill") as SpShape;
    ok(fillSp9?.fill.kind === "image", "multi-image round-trip: picture fill survives");
  }

  // ---------- fill types + chart styling round-trip ----------
  {
    const presA = newPresentation();
    const patternSp = makeShape("rect", 914400, 914400, 1828800, 1219200);
    patternSp.fill = { kind: "pattern", prst: "diagCross", fg: { kind: "srgb", hex: "C00000" }, bg: { kind: "srgb", hex: "FFF2CC" } };
    const chartA = makeChart("column", 3657600, 914400, 4114800, 2743200);
    chartA.labelSizePt = 12;
    chartA.series[0].color = { kind: "srgb", hex: "FF0000" };
    presA.slides = [{ ...makeSlide("blank"), background: { kind: "image", mediaId: "m-bg" }, shapes: [patternSp, chartA] }];
    const mediaA = new Map<string, MediaItem>();
    mediaA.set("m-bg", { id: "m-bg", mime: "image/png", bytes: PNG_BYTES, dataUrl: "data:image/png;base64," });
    // gradient + pattern backgrounds on extra slides (slide-settings Background pane)
    const gradSlide = makeSlide("blank");
    gradSlide.background = {
      kind: "gradient", angle: 45,
      stops: [
        { pos: 0, color: { kind: "srgb", hex: "FFD7C2" } },
        { pos: 100, color: { kind: "scheme", slot: "accent1" } },
      ],
    };
    const patSlide = makeSlide("blank");
    patSlide.background = { kind: "pattern", prst: "horz", fg: { kind: "srgb", hex: "404040" }, bg: { kind: "srgb", hex: "F2F2F2" } };
    presA.slides.push(gradSlide, patSlide);
    const reA = await parsePptx(await (await buildPptx(presA, mediaA)).generateAsync({ type: "uint8array" }), "fills.pptx");
    const sA = reA.pres.slides[0];
    ok(sA.background?.kind === "image", "fills: image background survives", sA.background?.kind);
    const bgG = reA.pres.slides[1].background;
    ok(bgG?.kind === "gradient" && bgG.stops.length === 2 && bgG.angle === 45
      && bgG.stops[1].color.kind === "scheme" && bgG.stops[1].color.slot === "accent1",
      "fills: gradient background survives", JSON.stringify(bgG));
    const bgP = reA.pres.slides[2].background;
    ok(bgP?.kind === "pattern" && bgP.prst === "horz" && bgP.fg.kind === "srgb" && bgP.fg.hex === "404040",
      "fills: pattern background survives", JSON.stringify(bgP));
    const pat = sA.shapes.find(s => s.kind === "sp") as SpShape;
    ok(pat.fill.kind === "pattern" && pat.fill.prst === "diagCross" && pat.fill.fg.kind === "srgb" && pat.fill.fg.hex === "C00000", "fills: pattern fill survives", JSON.stringify(pat.fill.kind === "pattern" ? pat.fill.prst : ""));
    const chA = sA.shapes.find(s => s.kind === "chart") as ChartShape;
    ok(chA.labelSizePt === 12, "chart: label size survives (c:txPr)", String(chA.labelSizePt));
    ok(chA.series[0].color?.kind === "srgb" && chA.series[0].color.hex === "FF0000", "chart: series color survives");
    ok(reA.warnings.length === 0, "fills round-trip: no warnings", reA.warnings.join("; "));
  }

  // ---------- chart elements round-trip ----------
  {
    const presB = newPresentation();
    const chB = makeChart("column", 914400, 914400, 5486400, 3657600);
    chB.title = "Revenue by Quarter";
    chB.legendPos = "b";
    chB.dataLabels = true;
    chB.errorBarsPct = 5;
    chB.axisTitleX = "Quarter";
    chB.axisTitleY = "Revenue (M)";
    chB.hideAxisX = undefined;
    chB.chartFill = { kind: "solid", color: { kind: "srgb", hex: "F2F7FF" } };
    const chC = makeChart("line", 7315200, 914400, 4114800, 2743200);
    chC.hideAxisY = true;
    presB.slides = [{ ...makeSlide("blank"), shapes: [chB, chC] }];
    const reB = await parsePptx(await (await buildPptx(presB, new Map())).generateAsync({ type: "uint8array" }), "elements.pptx");
    const rB = reB.pres.slides[0].shapes.filter((s): s is ChartShape => s.kind === "chart");
    const c1 = rB.find(c => c.chart === "column")!;
    ok(c1.legendPos === "b", "chart elements: legend position survives", c1.legendPos);
    ok(c1.dataLabels === true, "chart elements: data labels survive");
    ok(c1.errorBarsPct === 5, "chart elements: error bars survive", String(c1.errorBarsPct));
    ok(c1.axisTitleX === "Quarter" && c1.axisTitleY === "Revenue (M)", "chart elements: axis titles survive", `${c1.axisTitleX}/${c1.axisTitleY}`);
    ok(c1.chartFill?.kind === "solid" && c1.chartFill.color.kind === "srgb" && c1.chartFill.color.hex === "F2F7FF", "chart elements: chart background fill survives");
    const c2 = rB.find(c => c.chart === "line")!;
    ok(c2.hideAxisY === true, "chart elements: hidden axis survives");
    ok(reB.warnings.length === 0, "chart elements round-trip: no warnings", reB.warnings.join("; "));
  }

  // ---------- text size inheritance: lstStyle / pPr defRPr / layout ph / master txStyles ----------
  {
    const NS = `xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"`;
    const relsNs = `xmlns="http://schemas.openxmlformats.org/package/2006/relationships"`;
    const grpHeader = `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>`;

    // text box: 9pt via its own lstStyle; paragraph 2 overrides via pPr defRPr; run 3 explicit
    const textBox =
      `<p:sp><p:nvSpPr><p:cNvPr id="2" name="Notes"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>` +
      `<p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="1828800"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>` +
      `<p:txBody><a:bodyPr/><a:lstStyle>` +
      `<a:lvl1pPr><a:defRPr sz="900" b="1"/></a:lvl1pPr>` +
      `<a:lvl2pPr><a:defRPr sz="800"/></a:lvl2pPr>` +
      `</a:lstStyle>` +
      `<a:p><a:r><a:rPr lang="en-US"/><a:t>nine point bold</a:t></a:r></a:p>` +
      `<a:p><a:pPr><a:defRPr sz="1200"/></a:pPr><a:r><a:rPr lang="en-US"/><a:t>twelve point</a:t></a:r></a:p>` +
      `<a:p><a:r><a:rPr lang="en-US" sz="2000"/><a:t>twenty point</a:t></a:r></a:p>` +
      `<a:p><a:pPr lvl="1"/><a:r><a:rPr lang="en-US"/><a:t>eight point</a:t></a:r></a:p>` +
      `</p:txBody></p:sp>`;
    // placeholders with no sizes anywhere on the slide
    const titlePh =
      `<p:sp><p:nvSpPr><p:cNvPr id="3" name="Title 1"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr/>` +
      `<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>Master sized title</a:t></a:r></a:p></p:txBody></p:sp>`;
    const bodyPh =
      `<p:sp><p:nvSpPr><p:cNvPr id="4" name="Content 2"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/>` +
      `<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>Layout sized body</a:t></a:r></a:p></p:txBody></p:sp>`;

    const zipC = new JSZip();
    zipC.file("_rels/.rels", `<?xml version="1.0"?><Relationships ${relsNs}><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>`);
    zipC.file("ppt/presentation.xml", `<?xml version="1.0"?><p:presentation ${NS}><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/></p:presentation>`);
    zipC.file("ppt/_rels/presentation.xml.rels", `<?xml version="1.0"?><Relationships ${relsNs}><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/></Relationships>`);
    zipC.file("ppt/slides/slide1.xml", `<?xml version="1.0"?><p:sld ${NS}><p:cSld><p:spTree>${grpHeader}${textBox}${titlePh}${bodyPh}</p:spTree></p:cSld></p:sld>`);
    zipC.file("ppt/slides/_rels/slide1.xml.rels", `<?xml version="1.0"?><Relationships ${relsNs}><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>`);
    // layout: body ph carries geometry + lstStyle sz=1100 (overrides master's 1500)
    zipC.file("ppt/slideLayouts/slideLayout1.xml", `<?xml version="1.0"?><p:sldLayout ${NS}><p:cSld><p:spTree>${grpHeader}` +
      `<p:sp><p:nvSpPr><p:cNvPr id="2" name="Body"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>` +
      `<p:spPr><a:xfrm><a:off x="914400" y="3200400"/><a:ext cx="10363200" cy="2743200"/></a:xfrm></p:spPr>` +
      `<p:txBody><a:bodyPr/><a:lstStyle><a:lvl1pPr><a:defRPr sz="1100"/></a:lvl1pPr></a:lstStyle><a:p/></p:txBody></p:sp>` +
      `<p:sp><p:nvSpPr><p:cNvPr id="3" name="Title"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>` +
      `<p:spPr><a:xfrm><a:off x="914400" y="457200"/><a:ext cx="10363200" cy="1371600"/></a:xfrm></p:spPr>` +
      `<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>` +
      `</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`);
    zipC.file("ppt/slideLayouts/_rels/slideLayout1.xml.rels", `<?xml version="1.0"?><Relationships ${relsNs}><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>`);
    // master: txStyles set title 30pt, body 15pt
    zipC.file("ppt/slideMasters/slideMaster1.xml", `<?xml version="1.0"?><p:sldMaster ${NS}><p:cSld><p:spTree>${grpHeader}</p:spTree></p:cSld>` +
      `<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>` +
      `<p:txStyles><p:titleStyle><a:lvl1pPr><a:defRPr sz="3000"/></a:lvl1pPr></p:titleStyle>` +
      `<p:bodyStyle><a:lvl1pPr><a:defRPr sz="1500"/></a:lvl1pPr></p:bodyStyle><p:otherStyle/></p:txStyles></p:sldMaster>`);

    const reC = await parsePptx(await zipC.generateAsync({ type: "uint8array" }), "sizes.pptx");
    const sC = reC.pres.slides[0];
    const tb = sC.shapes.find(s => s.kind === "sp" && s.name === "Notes") as SpShape;
    const sizes = tb.text!.paragraphs.map(p => p.runs[0]?.sizePt);
    ok(sizes[0] === 9, "size inheritance: shape lstStyle 9pt applied (user bug)", String(sizes[0]));
    ok(tb.text!.paragraphs[0].runs[0]?.bold === true, "size inheritance: lstStyle bold applied");
    ok(sizes[1] === 12, "size inheritance: paragraph defRPr overrides lstStyle", String(sizes[1]));
    ok(sizes[2] === 20, "size inheritance: explicit run size wins", String(sizes[2]));
    ok(sizes[3] === 8, "size inheritance: level-2 lstStyle applied", String(sizes[3]));
    const tPh = sC.shapes.find(s => s.kind === "sp" && s.name === "Title 1") as SpShape;
    ok(tPh.text!.paragraphs[0].runs[0]?.sizePt === 30, "size inheritance: master titleStyle 30pt", String(tPh.text!.paragraphs[0].runs[0]?.sizePt));
    const bPh = sC.shapes.find(s => s.kind === "sp" && s.name === "Content 2") as SpShape;
    ok(bPh.text!.paragraphs[0].runs[0]?.sizePt === 11, "size inheritance: layout ph lstStyle overrides master body 15pt", String(bPh.text!.paragraphs[0].runs[0]?.sizePt));
    ok(bPh.y > 0 && bPh.h > 0, "size inheritance: layout geometry still resolves");
  }

  // ---------- svg graphics: tint + fallback-less (pure graphic) round-trip ----------
  {
    // theme-bound Office-icon form: PowerPoint re-binds MsftOfcThm_* classes to
    // the live theme color, so a baked tint must also neutralize the class name
    const svgText = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" width="24" height="24"><style>.MsftOfcThm_Accent1_Fill_v2{fill:#4472C4;}</style><path class="MsftOfcThm_Accent1_Fill_v2" fill="#FFFFFF" d="M4 4h16v16H4z"/></svg>`;
    const presD = newPresentation();
    const mediaD = new Map<string, MediaItem>();
    mediaD.set("m-ico", {
      id: "m-ico", mime: "image/svg+xml",
      bytes: new TextEncoder().encode(svgText),
      dataUrl: "data:image/svg+xml;base64,", // no pngFallback: pure graphic, like PowerPoint icons
    });
    presD.slides = [{
      ...makeSlide("blank"),
      shapes: [
        { kind: "pic", id: "g1", name: "Graphic 1", mediaId: "m-ico", x: 914400, y: 914400, w: 457200, h: 457200, rot: 0, svgTint: { kind: "srgb", hex: "C00000" } },
        { kind: "pic", id: "g2", name: "Graphic 2", mediaId: "m-ico", x: 1828800, y: 914400, w: 457200, h: 457200, rot: 0 },
      ],
    }];
    const zipD = await buildPptx(presD, mediaD);
    const slideD = await zipD.file("ppt/slides/slide1.xml")!.async("text");
    ok(slideD.includes("<a:blip><a:extLst>"), "svg graphic: blip without r:embed (PowerPoint pure-graphic form)");
    ok((slideD.match(/asvg:svgBlip/g) ?? []).length === 2, "svg graphic: both pics use svgBlip");
    const files = Object.keys(zipD.files).filter(f => f.endsWith(".svg"));
    ok(files.length === 2, "svg graphic: tinted copy gets its own media file", files.join(","));
    const tinted = await Promise.all(files.map(f => zipD.file(f)!.async("text")));
    ok(tinted.some(t => t.includes("#C00000")) && tinted.some(t => t.includes("#FFFFFF")), "svg graphic: tint baked into one copy, original kept");
    const tintedCopy = tinted.find(t => t.includes("#C00000"))!;
    const plainCopy = tinted.find(t => !t.includes("#C00000"))!;
    ok(!tintedCopy.includes("MsftOfcThm_") && !/fill:\s*#4472C4/.test(tintedCopy),
      "svg graphic: tinted copy drops MsftOfcThm theme binding (white→accent1 bug)");
    ok(plainCopy.includes("MsftOfcThm_"), "svg graphic: untinted copy keeps theme binding");
    const reD = await parsePptx(await zipD.generateAsync({ type: "uint8array" }), "graphics.pptx");
    const picsD = reD.pres.slides[0].shapes.filter(s => s.kind === "pic");
    ok(picsD.length === 2, "svg graphic: both survive reparse");
    const mimes = picsD.map(p => reD.media.get((p as { mediaId: string }).mediaId)?.mime);
    ok(mimes.every(m => m === "image/svg+xml"), "svg graphic: stay vector after round-trip", mimes.join(","));
    ok(reD.warnings.length === 0, "svg graphic round-trip: no warnings", reD.warnings.join("; "));
  }

  // ---------- chart formatting: area/plot fill+border, series lines, marker size ----------
  {
    const presE = newPresentation();
    const chE = makeChart("line", 914400, 914400, 6400800, 3657600, { marker: true });
    chE.chartFill = { kind: "solid", color: { kind: "srgb", hex: "FBFBFB" } };
    chE.chartBorder = { fill: { kind: "solid", color: { kind: "srgb", hex: "C00000" } }, widthPt: 2, dash: "dash" };
    chE.plotFill = { kind: "solid", color: { kind: "srgb", hex: "EEF4FF" } };
    chE.plotBorder = { fill: { kind: "solid", color: { kind: "srgb", hex: "8FAADC" } }, widthPt: 1 };
    chE.markerSizePt = 9;
    chE.series = chE.series.map(s => ({ ...s, lineWidthPt: 3.5, dash: "dot" as const }));
    presE.slides = [{ ...makeSlide("blank"), shapes: [chE] }];
    const reE = await parsePptx(await (await buildPptx(presE, new Map())).generateAsync({ type: "uint8array" }), "chartfmt.pptx");
    const cE = reE.pres.slides[0].shapes.find(s => s.kind === "chart") as ChartShape;
    ok(cE.chartBorder?.fill.kind === "solid" && cE.chartBorder.fill.color.kind === "srgb" && cE.chartBorder.fill.color.hex === "C00000" && cE.chartBorder.widthPt === 2 && cE.chartBorder.dash === "dash",
      "chart fmt: chart area border survives", JSON.stringify(cE.chartBorder));
    ok(cE.plotFill?.kind === "solid" && cE.plotFill.color.kind === "srgb" && cE.plotFill.color.hex === "EEF4FF", "chart fmt: plot fill survives");
    ok(cE.plotBorder?.fill.kind === "solid" && cE.plotBorder.widthPt === 1, "chart fmt: plot border survives");
    ok(cE.series.every(s => s.lineWidthPt === 3.5 && s.dash === "dot"), "chart fmt: series line width + dash survive", JSON.stringify(cE.series[0].lineWidthPt));
    ok(cE.markerSizePt === 9, "chart fmt: marker size survives", String(cE.markerSizePt));
    ok(reE.warnings.length === 0, "chart fmt round-trip: no warnings", reE.warnings.join("; "));
  }

  // ---------- per-element chart text styling (title/axes/legend/labels) ----------
  {
    const presPS = newPresentation();
    const chPS = makeChart("column", 914400, 914400, 6400800, 3657600, {});
    chPS.title = "Quarterly Revenue";
    chPS.axisTitleX = "Quarter";
    chPS.axisTitleY = "USD (M)";
    chPS.labelSizePt = 10;
    chPS.partStyles = {
      title: { font: "Georgia", sizePt: 20, color: { kind: "srgb", hex: "C00000" }, bold: true, italic: true },
      axisTitleX: { sizePt: 12, color: { kind: "srgb", hex: "1F4E79" }, bold: true },
      axisTitleY: { underline: true, color: { kind: "scheme", slot: "accent2" } },
      legend: { sizePt: 9, color: { kind: "srgb", hex: "548235" } },
      axisLabels: { sizePt: 8, color: { kind: "srgb", hex: "404040" } },
    };
    presPS.slides = [{ ...makeSlide("blank"), shapes: [chPS] }];
    const rePS = await parsePptx(await (await buildPptx(presPS, new Map())).generateAsync({ type: "uint8array" }), "chartparts.pptx");
    const cPS = rePS.pres.slides[0].shapes.find(s => s.kind === "chart") as ChartShape;
    const t = cPS.partStyles?.title;
    ok(t?.font === "Georgia" && t.sizePt === 20 && t.bold === true && t.italic === true && t.color?.kind === "srgb" && t.color.hex === "C00000",
      "chart parts: title font/size/bold/italic/color survive", JSON.stringify(t));
    ok(cPS.partStyles?.axisTitleX?.sizePt === 12 && cPS.partStyles.axisTitleX.bold === true && cPS.partStyles.axisTitleX.color?.kind === "srgb" && cPS.partStyles.axisTitleX.color.hex === "1F4E79",
      "chart parts: X axis title style survives", JSON.stringify(cPS.partStyles?.axisTitleX));
    ok(cPS.partStyles?.axisTitleY?.underline === true && cPS.partStyles.axisTitleY.color?.kind === "scheme" && cPS.partStyles.axisTitleY.color.slot === "accent2",
      "chart parts: Y axis title underline + scheme color survive", JSON.stringify(cPS.partStyles?.axisTitleY));
    ok(cPS.partStyles?.legend?.sizePt === 9 && cPS.partStyles.legend.color?.kind === "srgb" && cPS.partStyles.legend.color.hex === "548235",
      "chart parts: legend style survives", JSON.stringify(cPS.partStyles?.legend));
    ok(cPS.partStyles?.axisLabels?.sizePt === 8 && cPS.partStyles.axisLabels.color?.kind === "srgb",
      "chart parts: axis label style survives", JSON.stringify(cPS.partStyles?.axisLabels));
    ok(cPS.labelSizePt === 10, "chart parts: global label size still distinct from per-part", String(cPS.labelSizePt));

    // an UNSTYLED chart must NOT gain partStyles on round-trip (no default pollution)
    const chPlain = makeChart("column", 0, 0, 4572000, 3200400, {});
    chPlain.title = "Plain";
    const rePlain = await parsePptx(await (await buildPptx({ ...newPresentation(), slides: [{ ...makeSlide("blank"), shapes: [chPlain] }] }, new Map())).generateAsync({ type: "uint8array" }), "plain.pptx");
    const cPlain = rePlain.pres.slides[0].shapes.find(s => s.kind === "chart") as ChartShape;
    ok(cPlain.partStyles === undefined, "chart parts: unstyled chart stays unstyled (no default pollution)", JSON.stringify(cPlain.partStyles));
    ok(rePS.warnings.length === 0 && rePlain.warnings.length === 0, "chart parts round-trip: no warnings");
  }

  // ---------- gridline color/visibility + pie per-slice colors (dPt) ----------
  {
    const presF = newPresentation();
    const chF = makeChart("column", 914400, 914400, 4572000, 3200400, {});
    chF.gridColor = { kind: "srgb", hex: "BFBFBF" };
    const chG = makeChart("line", 5486400, 914400, 4572000, 3200400, {});
    chG.hideGridlines = true;
    const chP = makeChart("pie", 914400, 4419600, 4572000, 3200400, {});
    chP.pointColors = [
      { kind: "srgb", hex: "ED7D31" },
      null,
      { kind: "srgb", hex: "262626" },
      { kind: "srgb", hex: "29ABE2" },
    ];
    presF.slides = [{ ...makeSlide("blank"), shapes: [chF, chG, chP] }];
    const reF = await parsePptx(await (await buildPptx(presF, new Map())).generateAsync({ type: "uint8array" }), "gridpie.pptx");
    const chartsF = reF.pres.slides[0].shapes.filter(s => s.kind === "chart") as ChartShape[];
    const colF = chartsF.find(c => c.chart === "column")!;
    const linF = chartsF.find(c => c.chart === "line")!;
    const pieF = chartsF.find(c => c.chart === "pie")!;
    ok(colF.gridColor?.kind === "srgb" && colF.gridColor.hex === "BFBFBF" && !colF.hideGridlines,
      "gridlines: color survives round-trip", JSON.stringify(colF.gridColor));
    ok(linF.hideGridlines === true, "gridlines: hidden survives round-trip (majorGridlines omitted)");
    ok(pieF.pointColors?.length === 4
      && pieF.pointColors[0]?.kind === "srgb" && pieF.pointColors[0].hex === "ED7D31"
      && pieF.pointColors[1] == null
      && pieF.pointColors[3]?.kind === "srgb" && pieF.pointColors[3].hex === "29ABE2",
      "pie slices: dPt colors survive round-trip", JSON.stringify(pieF.pointColors));
    ok(!colF.pointColors && !pieF.hideGridlines, "pie/grid flags stay scoped to their chart kinds");
    ok(reF.warnings.length === 0, "gridline/pie round-trip: no warnings", reF.warnings.join("; "));
  }

  // ---------- table cell-range merge / split (store-level, PowerPoint "Merge Cells") ----------
  {
    const presM = newPresentation();
    const tM = makeTable(4, 4, 914400, 914400, 7315200, 2438400);
    tM.cells[1][1].text.paragraphs[0].runs[0].text = "A";
    tM.cells[2][2].text.paragraphs[0].runs[0].text = "B";
    presM.slides = [{ ...makeSlide("blank"), shapes: [tM] }];
    store.loadPresentation(presM, new Map());
    store.mergeTableCells(tM.id, 1, 1, 2, 2); // merge a 2×2 block
    const merged = store.currentSlide.shapes.find(s => s.kind === "table") as TableShape;
    ok(merged.cells[1][1].gridSpan === 2 && merged.cells[1][1].rowSpan === 2,
      "merge cells: anchor spans 2×2", JSON.stringify({ g: merged.cells[1][1].gridSpan, r: merged.cells[1][1].rowSpan }));
    ok(merged.cells[1][2].merged === "h" && merged.cells[2][1].merged === "v" && merged.cells[2][2].merged === "v",
      "merge cells: covered cells flagged hMerge/vMerge");
    const mergedText = merged.cells[1][1].text.paragraphs.flatMap(p => p.runs.map(r => r.text)).join("");
    ok(mergedText.includes("A") && mergedText.includes("B"), "merge cells: text of merged cells concatenated", mergedText);
    const reM = await parsePptx(await (await buildPptx(store.pres, new Map())).generateAsync({ type: "uint8array" }), "merge.pptx");
    const rM = reM.pres.slides[0].shapes.find(s => s.kind === "table") as TableShape;
    ok(rM.cells[1][1].gridSpan === 2 && rM.cells[1][1].rowSpan === 2, "merge cells: 2×2 span survives round-trip");
    store.splitTableCell(tM.id, 1, 1);
    const split = store.currentSlide.shapes.find(s => s.kind === "table") as TableShape;
    ok(!split.cells[1][1].gridSpan && !split.cells[1][1].rowSpan && !split.cells[1][2].merged,
      "merge cells: split restores the grid");
  }

  // ---------- shape grouping: groupId clusters round-trip as real p:grpSp ----------
  {
    const presG = newPresentation();
    presG.slides = [{ ...makeSlide("blank"), shapes: [] }];
    store.loadPresentation(presG, new Map());
    const a = makeShape("rect", 914400, 914400, 1828800, 914400);
    const b = makeShape("ellipse", 3200400, 1371600, 1371600, 1371600);
    const lone = makeShape("star5", 6400800, 914400, 914400, 914400);
    store.addShape(a, false); store.addShape(b, false); store.addShape(lone, false);
    store.selectShapes([a.id, b.id]);
    store.groupSelection();
    const grouped = store.currentSlide.shapes;
    const gidA = grouped.find(s => s.id === a.id)?.groupId;
    ok(!!gidA && gidA === grouped.find(s => s.id === b.id)?.groupId, "group: members share a groupId");
    ok(grouped.find(s => s.id === lone.id)?.groupId === undefined, "group: unselected shape untouched");
    ok(store.expandToGroups([a.id]).sort().join() === [a.id, b.id].sort().join(), "group: expandToGroups returns whole group");

    const zipG = await buildPptx(store.pres, new Map());
    const slideG = await zipG.file("ppt/slides/slide1.xml")!.async("text");
    ok((slideG.match(/<p:grpSp>/g) ?? []).length === 1, "group: exactly one p:grpSp written");
    ok(/<a:chOff[^>]*\/><a:chExt/.test(slideG), "group: child-space mapping written");

    const reG = await parsePptx(await zipG.generateAsync({ type: "uint8array" }), "groups.pptx");
    const rs = reG.pres.slides[0].shapes;
    const gids = rs.map(s => s.groupId).filter(Boolean);
    ok(gids.length === 2 && gids[0] === gids[1], "group: membership survives round-trip", JSON.stringify(rs.map(s => !!s.groupId)));
    const rA = rs.find(s => s.kind === "sp" && s.geom === "rect")!;
    ok(Math.abs(rA.x - 914400) < 2 && Math.abs(rA.w - 1828800) < 2, "group: member geometry unchanged through grpSp identity mapping");
    ok(rs.length === 3, "group: flat shape count preserved");

    store.selectShapes([a.id, b.id]);
    store.ungroupSelection();
    ok(store.currentSlide.shapes.every(s => !s.groupId), "group: ungroup clears membership");

    // group rotate: members orbit the combined center and spin themselves
    const r1 = makeShape("rect", 0, 0, 914400, 914400);          // center (457200, 457200)
    const r2 = makeShape("rect", 1828800, 0, 914400, 914400);    // center (2286000, 457200)
    store.loadPresentation({ ...newPresentation(), slides: [{ ...makeSlide("blank"), shapes: [] }] }, new Map());
    store.addShape(r1, false); store.addShape(r2, false);
    store.selectShapes([r1.id, r2.id]);
    store.rotateSelected(90); // combined center (1371600, 457200)
    const g1 = store.currentSlide.shapes.find(s => s.id === r1.id)!;
    const g2 = store.currentSlide.shapes.find(s => s.id === r2.id)!;
    const c1 = { x: g1.x + g1.w / 2, y: g1.y + g1.h / 2 };
    ok(Math.abs(c1.x - 1371600) < 3 && Math.abs(c1.y - (457200 - 914400)) < 3,
      "group rotate: member orbits the combined center", JSON.stringify(c1));
    ok(g1.rot === 90 && g2.rot === 90, "group rotate: members spin with the group", `${g1.rot}/${g2.rot}`);
    store.flipSelected("h");
    const f1 = store.currentSlide.shapes.find(s => s.id === r1.id)!;
    ok(f1.flipH === true && f1.rot === 270, "group flip: mirrors flag and negates rotation", `${f1.flipH}/${f1.rot}`);
  }

  // ---------- embed config parsing (host integration) ----------
  {
    const { parseEmbedQuery } = await import("../src/util/embed");
    const base = "https://editor.example.com/";
    const c1 = parseEmbedQuery("?file=https%3A%2F%2Fcdn.x%2Fa.pptx&title=Q4&saveUrl=https%3A%2F%2Fput.x%2Fk&embed=1&parentOrigin=https%3A%2F%2Fapp.x", base);
    ok(c1.fileUrl === "https://cdn.x/a.pptx" && c1.saveUrl === "https://put.x/k", "embed: file/saveUrl parsed");
    ok(c1.title === "Q4" && c1.embed === true && c1.parentOrigin === "https://app.x", "embed: title/embed/parentOrigin parsed");
    const c2 = parseEmbedQuery("?file=javascript:alert(1)&saveUrl=ftp://x/y", base);
    ok(c2.fileUrl === undefined && c2.saveUrl === undefined, "embed: non-http(s) file/saveUrl rejected");
    const c3 = parseEmbedQuery("?file=/local/a.pptx", base);
    ok(c3.fileUrl === "https://editor.example.com/local/a.pptx", "embed: relative file resolved against page origin");
  }

  // ---------- security: hardening against untrusted .pptx input ----------
  {
    const NS = `xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"`;
    const relsNs = `xmlns="http://schemas.openxmlformats.org/package/2006/relationships"`;
    const grpHeader = `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>`;
    const miniPptx = (spTreeInner: string): JSZip => {
      const z = new JSZip();
      z.file("_rels/.rels", `<?xml version="1.0"?><Relationships ${relsNs}><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>`);
      z.file("ppt/presentation.xml", `<?xml version="1.0"?><p:presentation ${NS}><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/></p:presentation>`);
      z.file("ppt/_rels/presentation.xml.rels", `<?xml version="1.0"?><Relationships ${relsNs}><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/></Relationships>`);
      z.file("ppt/slides/slide1.xml", `<?xml version="1.0"?><p:sld ${NS}><p:cSld><p:spTree>${grpHeader}${spTreeInner}</p:spTree></p:cSld></p:sld>`);
      return z;
    };

    // 1) prototype pollution via a crafted avLst guide name
    const sentinel = {} as Record<string, unknown>;
    const evilSp = `<p:sp><p:nvSpPr><p:cNvPr id="2" name="Evil"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
      `<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm>` +
      `<a:prstGeom prst="roundRect"><a:avLst><a:gd name="__proto__" fmla="val 9999"/><a:gd name="constructor" fmla="val 1"/><a:gd name="adj" fmla="val 16667"/></a:avLst></a:prstGeom></p:spPr></p:sp>`;
    const reEvil = await parsePptx(await miniPptx(evilSp).generateAsync({ type: "uint8array" }), "evil.pptx");
    ok((sentinel as { polluted?: unknown }).polluted === undefined && (({} as Record<string, unknown>).polluted) === undefined,
      "security: avLst name '__proto__' does not pollute Object.prototype");
    const evilShape = reEvil.pres.slides[0].shapes.find(s => s.kind === "sp") as SpShape | undefined;
    const adjKeys = evilShape?.adj ? Object.keys(evilShape.adj) : [];
    ok(!adjKeys.includes("__proto__") && !adjKeys.includes("constructor") && adjKeys.includes("adj"),
      "security: dangerous adjust keys dropped, real ones kept", adjKeys.join(","));

    // 2) zip member-count cap (mutate the exported limit so the test stays fast)
    const savedEntries = PARSE_LIMITS.maxEntries;
    PARSE_LIMITS.maxEntries = 8;
    const bomb = new JSZip();
    for (let i = 0; i < 20; i++) bomb.file(`junk/file${i}.bin`, "x");
    let rejected = false;
    try {
      await parsePptx(await bomb.generateAsync({ type: "uint8array" }), "bomb.pptx");
    } catch (e) { rejected = /too many parts/i.test((e as Error).message); }
    PARSE_LIMITS.maxEntries = savedEntries;
    ok(rejected, "security: zip with too many entries is rejected (decompression-bomb guard)");

    // 3) deeply nested groups are truncated, not a stack overflow
    const savedDepth = PARSE_LIMITS.maxGroupDepth;
    PARSE_LIMITS.maxGroupDepth = 6;
    let nest = `<p:sp><p:nvSpPr><p:cNvPr id="99" name="leaf"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:sp>`;
    for (let i = 0; i < 40; i++) {
      nest = `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="${i + 10}" name="g"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
        `<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/><a:chOff x="0" y="0"/><a:chExt cx="100" cy="100"/></a:xfrm></p:grpSpPr>${nest}</p:grpSp>`;
    }
    let nestOk = false, nestWarned = false;
    try {
      const reNest = await parsePptx(await miniPptx(nest).generateAsync({ type: "uint8array" }), "nested.pptx");
      nestOk = true;
      nestWarned = reNest.warnings.some(w => /nested groups/i.test(w));
    } catch { nestOk = false; }
    PARSE_LIMITS.maxGroupDepth = savedDepth;
    ok(nestOk, "security: 40-deep nested groups parse without crashing");
    ok(nestWarned, "security: deep nesting is truncated with a warning");

    // 4) oversized input rejected up front
    const savedComp = PARSE_LIMITS.maxCompressedBytes;
    PARSE_LIMITS.maxCompressedBytes = 1024;
    let sizeRejected = false;
    try {
      await parsePptx(await miniPptx(evilSp).generateAsync({ type: "uint8array" }), "big.pptx");
    } catch (e) { sizeRejected = /too large/i.test((e as Error).message); }
    PARSE_LIMITS.maxCompressedBytes = savedComp;
    ok(sizeRejected, "security: oversized input is rejected before parsing");
  }

  return { zipBytes, report };
}
