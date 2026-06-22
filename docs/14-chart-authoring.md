# Charts — complete reference

Everything about charts in this editor: the data model, **inserting**, **reading data
back**, **updating**, and every styling dimension — **legend**, **data labels**, **series
colors/lines/markers**, **axes & gridlines**, **title**, **grouping**, **fills**, and
**per-element text formatting** — with, for each, *what you can set through the scripting
API* vs *what is editor-UI only* vs *what round-trips to the `.pptx`*.

> Companion to [13 — Scripting API](13-scripting-api.md). Examples were run against the
> live editor.

## Contents
1. [The two ways to call](#1-the-two-ways-to-call)
2. [The chart data model](#2-the-chart-data-model-chartshape)
3. [Capability matrix — set/read/round-trip](#3-capability-matrix)
4. [Inserting a chart](#4-inserting-a-chart)
5. [Per-type input (categories & series)](#5-per-type-input--categories--series)
6. [Reading chart data back (data access)](#6-reading-chart-data-back-data-access)
7. [Updating data](#7-updating-data) · [Changing appearance — `setChartOptions`](#75-changing-appearance--setchartoptions)
8. [Legend](#8-legend)
9. [Data labels](#9-data-labels)
10. [Series — colors, lines, markers](#10-series--colors-lines-markers)
11. [Axes, gridlines & title](#11-axes-gridlines--title)
12. [Grouping (stacked / 100 %)](#12-grouping-stacked--100)
13. [Chart-area & plot-area fill/border](#13-chart-area--plot-area-fillborder)
14. [Per-element text formatting](#14-per-element-text-formatting)
15. [Known limitations](#15-known-limitations)
16. [Source of truth](#16-source-of-truth)

---

## 1. The two ways to call

**In-page** (same JS context as the editor):
```js
const pe = window.presentationEditor;
const id = pe.insertChart("column", { categories: ["Q1","Q2"], series: [{ name:"R", values:[4,7] }] });
```

**Cross-origin** (host page → editor iframe, over `postMessage`): the `pe:invoke` bridge.
`args` is the **argument list as an array**:
```js
iframe.contentWindow.postMessage(
  { type: "pe:invoke", requestId, method: "insertChart", args: ["column", { categories:["Q1","Q2"], series:[{ name:"R", values:[4,7] }] }] },
  editorOrigin,
);
// reply: { type:"pe:result", requestId, ok:true, value:<id> }  | { ok:false, error }
```
See [§4](#4-inserting-a-chart) for the full bridge helper. **Only method *names* are
allow-listed; arguments are forwarded verbatim**, so all the input rules below apply
identically over the bridge.

---

## 2. The chart data model (`ChartShape`)

A chart is one shape on a slide. Its full model ([`src/model/types.ts`](../src/model/types.ts)):

```ts
interface ChartShape {
  kind: "chart";
  // geometry (shared by all shapes), EMU:
  id, name; x, y, w, h; rot; flipH?, flipV?; groupId?;

  // —— type & data ——
  chart: "column"|"bar"|"line"|"area"|"pie"|"doughnut"|"scatter"|"radar";
  categories: string[];                 // X labels / slice labels / scatter X-values
  series: {
    name: string;
    values: number[];
    color?: ColorRef;                   // undefined → theme accent cycle
    lineWidthPt?: number;               // line/scatter/radar stroke width
    dash?: "solid"|"dash"|"dot";        // line/scatter/radar stroke dash
  }[];

  // —— variants ——
  grouping?: "clustered"|"stacked"|"percentStacked";  // column/bar/area; default clustered
  marker?: boolean;                     // line/scatter point markers (default on)
  markerSizePt?: number;                // default ~5
  smooth?: boolean;                     // line/scatter curve smoothing
  radarStyle?: "standard"|"marker"|"filled";

  // —— title & legend ——
  title?: string;                       // undefined = no title
  legend: boolean;                      // show legend
  legendPos?: "r"|"b"|"t"|"l";          // default "r"

  // —— data labels ——
  dataLabels?: boolean;
  dataLabelPos?: "outEnd"|"ctr"|"inEnd";

  // —— axes & gridlines ——
  axisTitleX?: string;  axisTitleY?: string;
  hideAxisX?: boolean;  hideAxisY?: boolean;
  gridColor?: ColorRef; hideGridlines?: boolean;
  errorBarsPct?: number;                // % error bars, undefined = off

  // —— fills ——
  chartFill?: Fill;  chartBorder?: LineProps;   // chart area (outer)
  plotFill?: Fill;   plotBorder?: LineProps;    // plot area (inner)
  pointColors?: (ColorRef|null)[];      // per-slice colors (pie/doughnut); null = auto

  // —— text formatting ——
  labelSizePt?: number;                 // global label font size (default 12pt)
  partStyles?: Partial<Record<
    "title"|"axisTitleX"|"axisTitleY"|"legend"|"axisLabels"|"dataLabels",
    { font?, sizePt?, color?, bold?, italic?, underline? }
  >>;
}
```

---

## 3. Capability matrix

The API now covers **data** (`insertChart` / `setChartData`) **and most appearance**
(`setChartOptions` — see [§7.5](#75-changing-appearance--setchartoptions)). A few things
remain editor-UI only (per-element fonts, chart/plot **border**, per-series line width/dash).

| Aspect | Set via API | Read via API | In editor UI | Round-trips to `.pptx` |
|---|:--:|:--:|:--:|:--:|
| Chart **type** | ✅ `insertChart` | ✅ `chart.chartType` | ✅ | ✅ |
| **Position / size / rotation** | ✅ `insertChart` opts / `setElementProperties` | ✅ `x,y,w,h,rot…` | ✅ | ✅ |
| **Categories** | ✅ `insertChart` / `setChartData` | ✅ `chart.categories` | ✅ (Edit Data) | ✅ |
| **Series** name + values | ✅ `insertChart` / `setChartData` | ✅ `chart.series` | ✅ (Edit Data) | ✅ |
| **Title** (text) | ✅ `setChartOptions` | ✅ `chart.title` | ✅ | ✅ |
| **Legend** on/off | ✅ `setChartOptions` | ✅ `chart.legend` | ✅ | ✅ |
| **Legend position** | ✅ `setChartOptions` | ❌ | ✅ | ✅ |
| **Data labels** + position | ✅ `setChartOptions` | ❌ | ✅ | ✅ |
| **Series colors** | ✅ `setChartOptions` | ❌ | ✅ | ✅ |
| **Per-slice colors** (pie) | ✅ `setChartOptions` | ❌ | ✅ | ✅ |
| **Markers / smoothing / radar style** | ✅ `setChartOptions` | ❌ | ✅ | ✅ |
| **Grouping** (stacked / 100 %) | ✅ `setChartOptions` | ❌ | ✅ | ✅ |
| **Axis titles / hide axes / gridlines** | ✅ `setChartOptions` | ❌ | ✅ | ✅ |
| **Chart / plot fill** | ✅ `setChartOptions` | ❌ | ✅ | ✅ |
| Global **label font size** | ✅ `setChartOptions` (`labelSizePt`) | ❌ | ✅ | ✅ |
| **Per-element fonts** (`partStyles`) | ❌ | ❌ | ✅ | ✅ |
| Chart/plot **border**, per-series **line width/dash** | ❌ | ❌ | ✅ | ✅ |

> The ❌ "Set via API" rows are read from the model + `.pptx` but not yet projected into
> the public API — they're the natural next extension.

---

## 4. Inserting a chart

```ts
insertChart(chartType: ChartKind, opts?: {
  x?, y?, w?, h?: number;                 // EMU; omitted → centered 5,000,000 × 3,000,000
  categories?: string[];                  // coerced with String()
  series?: { name?: string; values: number[] }[];  // values coerced with Number()
}): string                                // returns the new chart's element id
```

⚠️ **`chartType` must be one of the 8 valid strings, passed FIRST.** `insertChart` now
**validates** it and **throws** on an unknown/typo'd type (over the bridge →
`pe:result { ok:false }`), instead of silently producing a chart that corrupts the
`.pptx`. An object passed as the first argument throws too. Optionally pass
`opts.options` (a [`ChartOptions`](#75-changing-appearance--setchartoptions)) to configure
appearance at insert time.

**Units (EMU):** `pe.EMU_PER_INCH` = 914400, `pe.EMU_PER_PX` = 9525, `pe.EMU_PER_PT` =
12700. A 16:9 slide is `12192000 × 6858000`.

```js
const id = pe.insertChart("column", {
  x: 914400, y: 914400, w: 6096000, h: 3429000,
  categories: ["Q1","Q2","Q3","Q4"],
  series: [{ name: "Revenue", values: [4.3, 2.5, 3.5, 4.5] }],
});
```

**Cross-origin bridge helper:**
```js
function invoke(iframe, origin, method, args) {
  return new Promise((resolve, reject) => {
    const requestId = crypto.randomUUID();
    function onMsg(e) {
      if (e.source !== iframe.contentWindow || e.data?.type !== "pe:result" || e.data.requestId !== requestId) return;
      window.removeEventListener("message", onMsg);
      e.data.ok ? resolve(e.data.value) : reject(new Error(e.data.error));
    }
    window.addEventListener("message", onMsg);
    iframe.contentWindow.postMessage({ type: "pe:invoke", requestId, method, args }, origin);
  });
}
const id = await invoke(iframe, editorOrigin, "insertChart",
  ["column", { categories:["Q1","Q2"], series:[{ name:"R", values:[4,7] }] }]);
```

---

## 5. Per-type input — `categories` & `series`

The same two fields feed every chart; their **meaning changes per type**.

| `chartType` | `categories` mean | `series[].values` mean | Series used | Multi-series |
|---|---|---|---|---|
| `column` | X-axis labels | bar heights | all | ✅ clustered |
| `bar` | Y-axis labels (horizontal) | bar lengths | all | ✅ clustered |
| `line` | X-axis labels | point heights | all | ✅ |
| `area` | X-axis labels | point heights | all | ✅ |
| `pie` | **slice labels** | slice sizes | **only `series[0]`** | ❌ extras ignored |
| `doughnut` | **slice labels** | slice sizes | **only `series[0]`** | ❌ extras ignored |
| `scatter` | **numeric X (as strings)** | **Y values** | all (shared X) | ✅ |
| `radar` | spoke (axis) labels | distance per spoke | all | ✅ |

```js
// column / bar / line / area — categories = X labels, one entry per series
pe.insertChart("column", { categories:["Q1","Q2","Q3","Q4"],
  series:[{name:"2024",values:[4.3,2.5,3.5,4.5]},{name:"2025",values:[2.4,4.4,1.8,2.8]}] });

// pie / doughnut — ONLY series[0]; categories label the slices; labels show %
pe.insertChart("pie", { categories:["Chrome","Safari","Edge","Other"],
  series:[{name:"Share",values:[63,19,5,13]}] });

// scatter — categories are numeric X (strings); values are Y; point i = (x_i, y_i)
pe.insertChart("scatter", { categories:["0.7","1.8","2.6","3.2","4.1"],
  series:[{name:"obs",values:[2.7,3.2,0.8,1.2,2.9]}] });

// radar — categories are spokes
pe.insertChart("radar", { categories:["Speed","Power","Range","Cost","UX"],
  series:[{name:"Model X",values:[8,6,9,4,7]}] });
```
> **Length rule:** keep each `series.values` the same length as `categories`.

---

## 6. Reading chart data back (data access)

Any method that returns an **element** exposes a chart's data under `.chart` as
**`ChartInfo`**:

```ts
interface ChartInfo {
  chartType: "column"|"bar"|…|"radar";
  title?: string;
  legend: boolean;
  categories: string[];
  series: { name: string; values: number[] }[];
}
```

Methods that return chart info:

| Method | Returns | Charts under |
|---|---|---|
| `getElement(id)` | one `ElementInfo` \| null | `.chart` |
| `getElements()` | `ElementInfo[]` (active slide) | `.chart` |
| `getActiveSlide()` | `SlideInfo` **with** `.elements` | `.elements[].chart` |
| `getSlide(i)` | `SlideInfo` **with** `.elements` | `.elements[].chart` |
| `getSelection()` | `{ ids, elements }` | `.elements[].chart` |
| `getSlides()` | `SlideInfo[]` (summaries, **no** elements) | — |

```js
const el = pe.getElement(id);
// el = { id, name, kind:"chart", x,y,w,h, rot, px:{…},
//        chart: { chartType:"column", title:undefined, legend:true,
//                 categories:["Q1","Q2","Q3","Q4"],
//                 series:[ {name:"2024",values:[4.3,2.5,3.5,4.5]},
//                          {name:"2025",values:[2.4,4.4,1.8,2.8]} ] } }

// every chart on the active slide:
const charts = pe.getElements().filter(e => e.kind === "chart").map(e => e.chart);
```

**What `ChartInfo` does *not* include:** legend *position*, grouping, series colors /
line width / dash, markers, data labels, axis titles, gridlines, fills, or fonts. Those
live in the model and the `.pptx`, but are **not** projected into the read API. If you
need them, read the raw model in-page via `window.__store` (dev builds only) or request
the `setChartOptions`/richer-`ChartInfo` extension.

---

## 7. Updating data

`setChartData(id, { categories?, series? })` → `boolean`. Replaces the data on an existing
chart; **per-series colors are preserved**; omit a field to leave it unchanged.

```js
pe.setChartData(id, {
  categories: ["X","Y","Z"],
  series: [{ name: "s", values: [5, 6, 7] }],
});   // → true
```
To change a chart's **type**, delete + re-insert (no API to mutate type in place).

Other geometry ops that apply to a chart shape: `setElementProperties(id, {x,y,w,h,rot,
flipH,flipV})`, `reorderElement(id, "front"|"back"|…)`, `deleteElement(id)`.

---

## 7.5 Changing appearance — `setChartOptions`

`setChartOptions(id, options)` → `boolean` sets the chart's **appearance** (legend,
grouping, data labels, titles, colors, …) on an existing chart. Pass only the keys you
want to change; **invalid values are ignored**, and **`null` clears** a value back to its
default/auto. Colors are **hex strings** (`"4472C4"` or `"#4472C4"`). The same `options`
object can also be passed to `insertChart` as `opts.options` to configure a chart at
creation time.

```ts
setChartOptions(id: string, options: ChartOptions): boolean

interface ChartOptions {
  title?: string | null;                                 // chart title ("" / null clears)
  legend?: boolean;                                      // show/hide legend
  legendPos?: "r" | "b" | "t" | "l";                     // right / bottom / top / left
  grouping?: "clustered" | "stacked" | "percentStacked"; // column / bar / area
  dataLabels?: boolean;                                  // value (or % for pie) labels
  dataLabelPos?: "outEnd" | "ctr" | "inEnd";             // outside / center / inside
  axisTitleX?: string | null;
  axisTitleY?: string | null;
  hideAxisX?: boolean;
  hideAxisY?: boolean;
  hideGridlines?: boolean;
  gridColor?: string | null;                             // hex
  markers?: boolean;                                     // line/scatter point markers
  markerSizePt?: number;
  smooth?: boolean;                                      // line/scatter curve smoothing
  radarStyle?: "standard" | "marker" | "filled";
  seriesColors?: (string | null)[];                      // per series; null = auto (accent cycle)
  sliceColors?: (string | null)[];                       // pie/doughnut per slice; null = auto
  chartFill?: string | null;                             // chart area: hex | "none" | null(clear)
  plotFill?: string | null;                              // plot area: hex | "none" | null(clear)
  labelSizePt?: number;                                  // global label font size
}
```

**Example — style a column chart in one call:**
```js
const id = pe.insertChart("column", {
  categories: ["Q1","Q2","Q3","Q4"],
  series: [{ name:"Plan", values:[3,5,4,6] }, { name:"Actual", values:[2,4,6,5] }],
});
pe.setChartOptions(id, {
  title: "Quarterly", legend: true, legendPos: "b",
  grouping: "stacked", dataLabels: true, dataLabelPos: "inEnd",
  axisTitleX: "Quarter", axisTitleY: "Units", hideGridlines: true,
  seriesColors: ["FF0000", "00CC44"], labelSizePt: 14,
});
```

**Example — configure at insert time + a pie with custom slice colors:**
```js
pe.insertChart("line", {
  categories: ["Jan","Feb","Mar","Apr"], series: [{ name:"Visits", values:[10,25,45,70] }],
  options: { legendPos: "b", dataLabels: true, smooth: true, seriesColors: ["7E57C2"] },
});

const pieId = pe.insertChart("pie", { categories:["A","B","C"], series:[{ name:"S", values:[5,3,2] }] });
pe.setChartOptions(pieId, { legendPos: "r", dataLabels: true, sliceColors: ["E53935","FB8C00","43A047"] });
```

**Clearing values:** `setChartOptions(id, { title: null, legendPos: "r", seriesColors: [null, null] })`
removes the title, returns the legend to the right, and reverts both series to automatic
(theme-accent) colors.

> Over the postMessage bridge: `invoke("setChartOptions", [id, { legendPos:"b", … }])`.

| | Value |
|---|---|
| Set via API | ✅ `setChartOptions` / `insertChart` `opts.options` |
| Editor UI | the chart panel & Add Chart Element menu |
| `.pptx` | ✅ every option round-trips |

---

## 8. Legend

**Model:** `legend: boolean` + `legendPos?: "r"|"b"|"t"|"l"` (default `"r"`). Text style via
`partStyles.legend`.

| | Value |
|---|---|
| Set via API | ✅ `setChartOptions` (`legend`, `legendPos`) |
| Read via API | ✅ `chart.legend` (boolean only — **not** position) |
| Editor UI | **Chart elements → Add Chart Element → Legend** → None / Right / Bottom / Top / Left |
| `.pptx` | ✅ `c:legend/c:legendPos` |

**Auto-fit width (no manual sizing needed):** a **left/right** legend column auto-sizes to
the longest entry (so names like "Series 1" aren't truncated), capped at 45 % of the chart
width; a **top/bottom** legend shows full names when they fit and truncates only when they
don't. This is derived from the labels at render time, so it needs no configuration and
re-fits on reopen.

---

## 9. Data labels

**Model:** `dataLabels?: boolean` + `dataLabelPos?: "outEnd"|"ctr"|"inEnd"`. Text style via
`partStyles.dataLabels`.

- **Content:** value for category/scatter charts; **percentage** for pie/doughnut.
- **Position** (`dataLabelPos`) maps per chart type:

| Token | column | bar | line / area / scatter | pie / doughnut |
|---|---|---|---|---|
| `outEnd` *(default cartesian)* | above bar | right of bar | above point | outside slice |
| `ctr` *(default pie)* | bar center | bar center | on the point | slice center |
| `inEnd` | inside top | inside end | below point | toward outer edge |

| | Value |
|---|---|
| Set via API | ✅ `setChartOptions` (`dataLabels`, `dataLabelPos`) |
| Read via API | ❌ |
| Editor UI | **Add Chart Element → Data Labels** → show + Outside End / Center / Inside End |
| `.pptx` | ✅ `c:dLbls` (`c:dLblPos` mapped to the per-type OOXML token; text style as `c:txPr`) |

---

## 10. Series — colors, lines, markers

Per-series fields on `ChartSeries`: `color?`, `lineWidthPt?`, `dash?` (`"solid"|"dash"|
"dot"`). Plus chart-level `marker?`, `markerSizePt?`, `smooth?`, and `pointColors?`
(pie/doughnut per-slice).

- **Default colors:** when `series.color` is unset, series cycle through the active theme's
  **accent1 → accent6** (e.g. Office `accent1` = `#4472C4`), then repeat. Pie/doughnut
  slices cycle the same accents unless overridden by `pointColors`.
- **Markers:** `line` and `scatter` show point markers by default (`marker ?? true`);
  `markerSizePt` defaults to ~5.
- **Smoothing:** `smooth: true` curves line/scatter. Scatter is tri-state:
  `smooth: undefined` → markers only (no line), `false` → straight line + markers,
  `true` → smooth line + markers.

| | Value |
|---|---|
| Set via API | ✅ colors/markers/smoothing via `setChartOptions` (`seriesColors`, `sliceColors`, `markers`, `markerSizePt`, `smooth`, `radarStyle`); name+values via `setChartData`. Per-series line width/dash: editor-UI only |
| Read via API | ❌ (color/line/dash not in `ChartInfo`) |
| Editor UI | series color pickers, **Series lines & markers** panel, per-slice colors |
| `.pptx` | ✅ series `c:spPr`/`a:ln`, `c:marker`, pie `c:dPt` |

---

## 11. Axes, gridlines & title

**Model:** `axisTitleX?`, `axisTitleY?` (strings); `hideAxisX?`, `hideAxisY?`;
`gridColor?`, `hideGridlines?`; `errorBarsPct?`; `title?`.

- Pie/doughnut/radar have no cartesian axes (titles/hide/gridlines are ignored).
- `title` is read via `chart.title` but **not** settable via the API.

| | Value |
|---|---|
| Set via API | ✅ `setChartOptions` (`title`, `axisTitleX/Y`, `hideAxisX/Y`, `hideGridlines`, `gridColor`) |
| Read via API | only `chart.title` |
| Editor UI | **Add Chart Element → Axis Titles / Axes / Chart Title / Error Bars**; gridline color & visibility in the chart panel |
| `.pptx` | ✅ axis `c:title`, `c:delete`, `c:majorGridlines`, `c:errBars` |

---

## 12. Grouping (stacked / 100 %)

**Model:** `grouping?: "clustered"|"stacked"|"percentStacked"` — applies to **column, bar,
area** (default `clustered`/standard). Multiple series stack instead of clustering;
`percentStacked` normalizes each category to 100 %.

| | Value |
|---|---|
| Set via API | ✅ `setChartOptions` (`grouping`) |
| Editor UI | **Type & options → Grouping** |
| `.pptx` | ✅ `c:grouping` + `c:overlap` |

---

## 13. Chart-area & plot-area fill/border

**Model:** `chartFill?`/`chartBorder?` (outer chart area) and `plotFill?`/`plotBorder?`
(inner plot rectangle). `Fill` supports solid / gradient / image / pattern / none.

| | Value |
|---|---|
| Set via API | ✅ `setChartOptions` (`chartFill`, `plotFill`); border is editor-UI only |
| Editor UI | **Chart area & gridlines** panel |
| `.pptx` | ✅ `c:chartSpace/c:spPr` and `c:plotArea/c:spPr` |

---

## 14. Per-element text formatting

Each chart text element can be styled independently via `partStyles` — keyed by part:
`"title" | "axisTitleX" | "axisTitleY" | "legend" | "axisLabels" | "dataLabels"`. Each value
is a `ChartTextStyle`:

```ts
{ font?: string; sizePt?: number; color?: ColorRef; bold?, italic?, underline?: boolean }
```
`labelSizePt` is the **global** label font-size fallback (default 12 pt) that per-part
styles override.

| | Value |
|---|---|
| Set via API | global `labelSizePt` via `setChartOptions`; per-part fonts ❌ |
| Editor UI | select the element (click it, or pick in the **Text elements** row) → format from the **Home tab** |
| `.pptx` | ✅ each part's `c:txPr` (data labels: `c:txPr` inside `c:dLbls`) |

---

## 15. Known limitations

1. **Pass a valid chart type.** `insertChart` now validates `chartType` and **throws** on
   an unknown value (and the writer has a column fallback), so a bad type can no longer
   silently corrupt the `.pptx`. Catch the throw if you forward arbitrary host input.
2. **Negative values clamp to zero** in the renderer today — `[-3, 5, -2]` draws the
   negatives flat on the baseline. Avoid negative data (or offset it) until fixed.
3. **A few options remain editor-UI only:** per-element fonts (`partStyles`), chart/plot
   **border**, and per-series **line width / dash**. Everything else — title, legend +
   position, grouping, data labels, series & slice colors, markers/smoothing, axes,
   gridlines, fills, global label size — is settable via
   [`setChartOptions`](#75-changing-appearance--setchartoptions).

---

## 16. Source of truth

- `insertChart` / `setChartData` / `describeElement` (→ `ChartInfo`) — [`src/util/api.ts`](../src/util/api.ts)
- Model (`ChartShape`, `ChartKind`, `ChartSeries`, `ChartTextStyle`, `ChartPart`) — [`src/model/types.ts`](../src/model/types.ts)
- Defaults (`makeChart`, `CHART_NAMES`, theme accents) — [`src/model/defaults.ts`](../src/model/defaults.ts)
- Rendering & layout (legend auto-fit, data-label position, axis/legend math) — [`src/render/GraphicViews.tsx`](../src/render/GraphicViews.tsx)
- `.pptx` read/write (round-trip of every field above) — [`src/ooxml/parse.ts`](../src/ooxml/parse.ts), [`src/ooxml/write.ts`](../src/ooxml/write.ts)
- `pe:invoke` bridge — [`src/util/embed.ts`](../src/util/embed.ts)
