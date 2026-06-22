# Inserting charts programmatically

How to add charts to the active slide through the scripting API, and exactly what the
**data input** must look like for each of the 8 chart types. Every example here was run
against the live editor and produces a valid chart.

> Companion to [13 — Scripting API](13-scripting-api.md). Chart insertion is part of that
> API surface (`window.presentationEditor`); this doc is the chart-specific deep dive.

---

## 1. The one rule that matters most

```js
window.presentationEditor.insertChart(chartType, opts)
//                                    ^^^^^^^^^  a STRING, passed FIRST
```

`chartType` is a **string** and must be **exactly** one of the 8 valid kinds:

```
"column" | "bar" | "line" | "area" | "pie" | "doughnut" | "scatter" | "radar"
```

⚠️ **There is no input validation today.** An unknown string (`"collumn"`), or an
**object passed as the first argument**, does *not* throw — it silently creates a broken
chart (`name: "undefined Chart"`) that renders blank in the editor and makes the saved
`.pptx` **unreadable in PowerPoint** ("repair" prompt). Always pass a valid kind. Use the
[safe wrapper](#9-recommended-safe-wrapper) until the API guards this itself.

```js
// ❌ WRONG — object first arg → corrupt chart
pe.insertChart({ type: "column", data: [...] });

// ❌ WRONG — typo'd kind → "undefined Chart", corrupts on save
pe.insertChart("collumn", { ... });

// ✅ RIGHT — kind string first, options second
pe.insertChart("column", { categories: ["Q1","Q2"], series: [{ name: "2025", values: [4,7] }] });
```

---

## 2. Signature & input shape

```ts
insertChart(chartType: ChartKind, opts?: InsertChartOpts): string   // returns the new element id

interface InsertChartOpts {
  // position & size, in EMU (English Metric Units). All optional.
  x?: number;  y?: number;  w?: number;  h?: number;
  // the data:
  categories?: string[];                       // axis/slice labels (see per-type rules)
  series?: { name?: string; values: number[] }[];
}
```

- **Returns** the new chart's element id (a string like `"chart_ab12_3"`). Keep it to
  update the chart later with [`setChartData`](#7-updating-an-existing-charts-data).
- **Coercion:** `categories` are coerced with `String(...)`, every `values` entry with
  `Number(...)`. So `categories: [1, 2, 3]` becomes `["1","2","3"]`, and a non-numeric
  value becomes `NaN` — pass real numbers.
- **`series[].name`** is optional; it defaults to `"Series 1"`, `"Series 2"`, … in order.
- **Omitting `categories`/`series`** fills the chart with built-in **sample data**
  (Q1–Q4 with two series; scatter gets sample XY points). Good for a quick placeholder,
  but for real output always pass your own.

### Units & positioning (EMU)

All geometry is in **EMU**. The API exposes the constants:

```js
pe.EMU_PER_INCH  // 914400
pe.EMU_PER_PX    // 9525   (96 px / inch)
pe.EMU_PER_PT    // 12700  (72 pt / inch)
```

| You want | EMU |
|---|---|
| 1 inch | `914400` |
| 1 cm | `360000` |
| 1 px (@96dpi) | `9525` |
| A 16:9 slide | `12192000 × 6858000` (13.33ʺ × 7.5ʺ) |

If you **omit** `x/y/w/h`, the chart is placed **centered** at a default size of
**5,000,000 × 3,000,000 EMU** (≈ 5.47ʺ × 3.28ʺ). To place explicitly:

```js
pe.insertChart("column", {
  x: 914400, y: 914400,        // 1ʺ from top-left
  w: 6_096_000, h: 3_429_000,  // half-slide-ish
  categories: ["Q1","Q2","Q3","Q4"],
  series: [{ name: "Revenue", values: [4.3, 2.5, 3.5, 4.5] }],
});
```

---

## 3. How `categories` and `series` map — by chart type

The **same two fields** (`categories`, `series`) feed every chart, but their meaning
changes per type. This is the heart of the doc.

### Category charts — `column`, `bar`, `line`, `area`

The default shape. `categories` are the **X-axis labels**; each entry in `series` is one
data series, and its `values[i]` lines up with `categories[i]`. Supports **multiple
series** (drawn side-by-side / layered).

```js
// column: vertical bars, grouped by category. Two series → clustered bars.
pe.insertChart("column", {
  categories: ["Q1", "Q2", "Q3", "Q4"],
  series: [
    { name: "2024", values: [4.3, 2.5, 3.5, 4.5] },
    { name: "2025", values: [2.4, 4.4, 1.8, 2.8] },
  ],
});
```

- **`bar`** — identical input to `column`, just drawn **horizontally** (categories run
  down the left, bars extend right).
- **`line`** — each series is a line across the categories. Markers/smoothing are
  styling, set in the editor (not via this API — see [§8](#8-what-you-can-and-cant-set-via-the-api-today)).
- **`area`** — like `line`, but filled to the baseline.

```js
pe.insertChart("line", {
  categories: ["Jan", "Feb", "Mar", "Apr"],
  series: [{ name: "Visits", values: [10, 25, 45, 70] }],
});

pe.insertChart("area", {
  categories: ["Jan", "Feb", "Mar"],
  series: [
    { name: "Organic", values: [3, 5, 4] },
    { name: "Paid",    values: [2, 3, 5] },
  ],
});
```

> **Length rule:** make every `series.values` the same length as `categories`. A short
> series leaves gaps; a long one is clipped to the category count.

### Part-of-whole — `pie`, `doughnut`

These show **one** series as slices of a whole.

- **Only `series[0]` is used.** Any additional series are **ignored**.
- Each **value** is one slice; **`categories`** are the **slice labels** (and the legend).
- Data labels render as **percentages** (not raw values).

```js
// pie: categories = slice labels, series[0].values = slice sizes
pe.insertChart("pie", {
  categories: ["Chrome", "Safari", "Edge", "Other"],
  series: [{ name: "Browser share", values: [63, 19, 5, 13] }],
});

// doughnut: same input, drawn with a center hole
pe.insertChart("doughnut", {
  categories: ["Done", "Remaining"],
  series: [{ name: "Progress", values: [72, 28] }],
});
```

> Values don't need to sum to 100 — slices are sized by proportion. Use whatever raw
> counts you have.

### Correlation — `scatter` (XY)

The one type where **`categories` are not labels** — they hold the **numeric X values**
(as strings), and `series[].values` are the **Y values**. Point *i* is at
`(categories[i], series.values[i])`.

```js
// scatter: categories = X (numeric strings), values = Y. Point i = (x_i, y_i).
pe.insertChart("scatter", {
  categories: ["0.7", "1.8", "2.6", "3.2", "4.1"],   // X axis (numbers as strings)
  series: [{ name: "Observations", values: [2.7, 3.2, 0.8, 1.2, 2.9] }], // Y axis
});
```

- Pass X values as **numeric strings** (`"0.7"`); non-numeric X falls back to its index
  (1, 2, 3…).
- Multiple series are allowed; each shares the **same X values** (`categories`) but has
  its own Y `values`.
- Markers are **on by default** for scatter.

### Comparison across axes — `radar`

`categories` are the **spokes** (the axes radiating from the center); each series is one
closed ring. Multiple series overlay for comparison.

```js
pe.insertChart("radar", {
  categories: ["Speed", "Power", "Range", "Cost", "UX"],
  series: [
    { name: "Model X", values: [8, 6, 9, 4, 7] },
    { name: "Model Y", values: [6, 8, 5, 7, 6] },
  ],
});
```

---

## 4. Quick reference table

| `chartType` | Orientation / shape | `categories` mean | `series[].values` mean | Series used | Multi-series? |
|---|---|---|---|---|---|
| `"column"` | vertical bars | X-axis labels | bar heights | all | ✅ clustered |
| `"bar"` | horizontal bars | Y-axis labels | bar lengths | all | ✅ clustered |
| `"line"` | lines | X-axis labels | point heights | all | ✅ |
| `"area"` | filled lines | X-axis labels | point heights | all | ✅ |
| `"pie"` | slices | **slice labels** | slice sizes | **only `series[0]`** | ❌ (extras ignored) |
| `"doughnut"` | slices + hole | **slice labels** | slice sizes | **only `series[0]`** | ❌ (extras ignored) |
| `"scatter"` | XY points | **X values (numeric strings)** | **Y values** | all (shared X) | ✅ |
| `"radar"` | spokes/rings | spoke (axis) labels | distance per spoke | all | ✅ |

---

## 5. A complete example (one chart per type)

```js
const pe = window.presentationEditor;

pe.insertChart("column",   { categories:["Q1","Q2","Q3","Q4"], series:[{name:"2024",values:[4.3,2.5,3.5,4.5]},{name:"2025",values:[2.4,4.4,1.8,2.8]}] });
pe.insertChart("bar",      { categories:["North","South","East"], series:[{name:"Units",values:[120,90,150]}] });
pe.insertChart("line",     { categories:["Jan","Feb","Mar","Apr"], series:[{name:"Visits",values:[10,25,45,70]}] });
pe.insertChart("area",     { categories:["Jan","Feb","Mar"], series:[{name:"A",values:[3,5,4]},{name:"B",values:[2,3,5]}] });
pe.insertChart("pie",      { categories:["Chrome","Safari","Edge","Other"], series:[{name:"Share",values:[63,19,5,13]}] });
pe.insertChart("doughnut", { categories:["Done","Left"], series:[{name:"Progress",values:[72,28]}] });
pe.insertChart("scatter",  { categories:["0.7","1.8","2.6","3.2","4.1"], series:[{name:"obs",values:[2.7,3.2,0.8,1.2,2.9]}] });
pe.insertChart("radar",    { categories:["Speed","Power","Range","Cost","UX"], series:[{name:"Model X",values:[8,6,9,4,7]}] });
```

---

## 6. Inserting from a host page (cross-origin `postMessage`)

When the editor runs in an iframe, call the same method over the `pe:invoke` bridge.
**`args` is the argument list as an array** — so chart insertion is
`args: [chartType, opts]`.

```js
const iframe = document.getElementById("editor");          // the <iframe> running the editor
const editorOrigin = "https://editor.example.com";         // its origin

function invoke(method, args) {
  return new Promise((resolve, reject) => {
    const requestId = crypto.randomUUID();
    function onMsg(e) {
      if (e.source !== iframe.contentWindow) return;
      const d = e.data;
      if (d?.type !== "pe:result" || d.requestId !== requestId) return;
      window.removeEventListener("message", onMsg);
      d.ok ? resolve(d.value) : reject(new Error(d.error));
    }
    window.addEventListener("message", onMsg);
    iframe.contentWindow.postMessage({ type: "pe:invoke", requestId, method, args }, editorOrigin);
  });
}

// insertChart(chartType, opts)  →  args = [chartType, opts]
const id = await invoke("insertChart", [
  "column",
  { categories: ["Q1","Q2","Q3","Q4"], series: [{ name: "Revenue", values: [4.3,2.5,3.5,4.5] }] },
]);
```

The reply is `{ type:"pe:result", requestId, ok:true, value:<id> }`, or
`{ ok:false, error:"…" }`. Only allow-listed method **names** are checked — **arguments
are forwarded verbatim**, so the "valid type string first" rule applies here too. Note:
if you pass `args` as a bare object instead of an array, the bridge wraps it as
`[object]`, i.e. `insertChart(object)` → corrupt chart.

---

## 7. Updating an existing chart's data

`setChartData(id, { categories?, series? })` replaces the data on an existing chart and
returns `true` on success (`false` if the id isn't a chart on the active slide). Per-series
**colors are preserved**; omit a field to leave it unchanged.

```js
const id = pe.insertChart("column", { categories:["A","B"], series:[{ name:"s", values:[1,2] }] });

pe.setChartData(id, {
  categories: ["X", "Y", "Z"],
  series: [{ name: "s", values: [5, 6, 7] }],
});
```

To change the **chart type** of an existing chart, there is no API yet — delete and
re-insert, or change it in the editor UI.

---

## 8. What you can — and can't — set via the API today

| Set programmatically | How |
|---|---|
| Chart **type** | first arg of `insertChart` |
| **Position / size** | `opts.x/y/w/h` (or `setElementProperties(id, {x,y,w,h,rot})`) |
| **Categories & series** (names + values) | `opts.categories/series`, or `setChartData` |

**Not exposed through the API yet** (configure these in the editor UI after inserting):
chart **title**, **legend** on/off + position, **grouping** (stacked / 100 % stacked for
column/bar/area), **per-series & per-slice colors**, **data labels** + position, **axis
titles**, hide axes, line **markers / smoothing**, **radar style**, and fonts. The model
supports all of these (`ChartShape` in [`src/model/types.ts`](../src/model/types.ts)); they
just aren't wired into the public API surface.

### Two known limitations (from the code audit)

1. **No type validation** — an invalid `chartType` corrupts the `.pptx` (see [§1](#1-the-one-rule-that-matters-most)). Guard it yourself for now ([§9](#9-recommended-safe-wrapper)).
2. **Negative values currently clamp to zero** in the renderer — a series like
   `[-3, 5, -2]` draws the negatives flat on the baseline. Avoid charts with negative data
   until this is fixed, or offset your data.

---

## 9. Recommended safe wrapper

Until the API validates input itself, wrap `insertChart` so a bad type can never reach it:

```js
const CHART_TYPES = ["column","bar","line","area","pie","doughnut","scatter","radar"];

function insertChartSafe(pe, chartType, opts = {}) {
  if (!CHART_TYPES.includes(chartType)) {
    throw new Error(`insertChart: invalid chartType "${chartType}". Use one of: ${CHART_TYPES.join(", ")}`);
  }
  if (opts.series) {
    for (const s of opts.series) {
      if (!Array.isArray(s.values) || s.values.some(v => typeof v !== "number" || Number.isNaN(v))) {
        throw new Error(`insertChart: series "${s.name ?? "?"}" has non-numeric values`);
      }
    }
  }
  return pe.insertChart(chartType, opts);
}
```

---

## Source of truth

- `insertChart` / `setChartData` — [`src/util/api.ts`](../src/util/api.ts)
- Chart model (`ChartShape`, `ChartKind`, `ChartSeries`) — [`src/model/types.ts`](../src/model/types.ts)
- Chart defaults (`makeChart`, `CHART_NAMES`) — [`src/model/defaults.ts`](../src/model/defaults.ts)
- `pe:invoke` bridge — [`src/util/embed.ts`](../src/util/embed.ts)
