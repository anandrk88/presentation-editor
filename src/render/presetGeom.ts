import rawDefs from "./presetDefs.json";

/**
 * ECMA-376 preset geometry evaluator. Each preset is a list of adjust values
 * (avLst), guide formulas (gdLst) and path commands whose arguments reference
 * guides — the same tables PowerPoint evaluates. Evaluating them
 * at the shape's current size yields correct outlines for all 187 presets at
 * any aspect ratio, honoring per-shape adjust values from imported files.
 */

interface RawPath {
  w?: number;
  h?: number;
  fill?: string;   // "none" | "lighten" | ... (absent = norm)
  stroke?: 0;      // 0 = no stroke (absent = stroke)
  cmds: (string | number)[][];
}
interface RawDef {
  av?: [string, string][];
  gd?: [string, string][];
  rect?: [string, string, string, string];
  paths: RawPath[];
}

const DEFS = rawDefs as unknown as Record<string, RawDef>;

export const PRESET_NAMES = Object.keys(DEFS);

export function hasPreset(name: string): boolean {
  return Object.prototype.hasOwnProperty.call(DEFS, name);
}

/** Default adjust values (name -> raw value) from the preset's avLst. */
export function defaultAdj(name: string): Record<string, number> {
  const def = DEFS[name];
  const out: Record<string, number> = {};
  if (!def?.av) return out;
  for (const [n, fmla] of def.av) {
    const m = /^val (-?[\d.]+)$/.exec(fmla);
    out[n] = m ? parseFloat(m[1]) : 0;
  }
  return out;
}

// 60000ths of a degree -> radians
const A = Math.PI / 10800000;

function builtins(w: number, h: number): Record<string, number> {
  const ss = Math.min(w, h);
  return {
    w, h, ss, ls: Math.max(w, h),
    t: 0, l: 0, b: h, r: w,
    hc: w / 2, vc: h / 2,
    cd2: 10800000, cd4: 5400000, cd8: 2700000,
    "3cd4": 16200000, "3cd8": 8100000, "5cd8": 13500000, "7cd8": 18900000,
    wd2: w / 2, wd3: w / 3, wd4: w / 4, wd5: w / 5, wd6: w / 6, wd8: w / 8,
    wd10: w / 10, wd12: w / 12, wd16: w / 16, wd32: w / 32,
    hd2: h / 2, hd3: h / 3, hd4: h / 4, hd5: h / 5, hd6: h / 6, hd8: h / 8,
    hd10: h / 10, hd12: h / 12, hd16: h / 16, hd32: h / 32,
    ssd2: ss / 2, ssd4: ss / 4, ssd6: ss / 6, ssd8: ss / 8, ssd16: ss / 16, ssd32: ss / 32,
  };
}

const NUM_RE = /^[+-]?(\d+\.?\d*|\.\d+)$/;

function resolve(tok: string | number, env: Record<string, number>): number {
  if (typeof tok === "number") return tok;
  if (NUM_RE.test(tok)) return parseFloat(tok);
  const v = env[tok];
  return v === undefined || Number.isNaN(v) ? 0 : v;
}

function evalFormula(fmla: string, env: Record<string, number>): number {
  const parts = fmla.trim().split(/\s+/);
  const op = parts[0];
  const a = () => resolve(parts[1], env);
  const b = () => resolve(parts[2], env);
  const c = () => resolve(parts[3], env);
  switch (op) {
    case "val": return a();
    case "*/": { const z = c(); return z === 0 ? 0 : (a() * b()) / z; }
    case "+-": return a() + b() - c();
    case "+/": { const z = c(); return z === 0 ? 0 : (a() + b()) / z; }
    case "?:": return a() > 0 ? b() : c();
    case "abs": return Math.abs(a());
    case "min": return Math.min(a(), b());
    case "max": return Math.max(a(), b());
    case "sqrt": return Math.sqrt(Math.max(0, a()));
    case "mod": return Math.sqrt(a() ** 2 + b() ** 2 + c() ** 2);
    case "pin": return Math.min(Math.max(b(), a()), c());
    case "at2": return Math.atan2(b(), a()) / A;
    case "cat2": return a() * Math.cos(Math.atan2(c(), b()));
    case "sat2": return a() * Math.sin(Math.atan2(c(), b()));
    case "cos": return a() * Math.cos(b() * A);
    case "sin": return a() * Math.sin(b() * A);
    case "tan": {
      const t = Math.tan(b() * A);
      return Number.isFinite(t) ? a() * t : 0;
    }
    default: return 0;
  }
}

function buildEnv(def: RawDef, w: number, h: number, adj?: Record<string, number>): Record<string, number> {
  const env = builtins(w, h);
  for (const [name, fmla] of def.av ?? []) {
    env[name] = adj?.[name] !== undefined ? adj[name] : evalFormula(fmla, env);
  }
  for (const [name, fmla] of def.gd ?? []) {
    env[name] = evalFormula(fmla, env);
  }
  return env;
}

export interface EvaluatedPath {
  d: string;
  noFill: boolean;   // path is stroke-only decoration
  noStroke: boolean; // path is a fill-only layer
  darken?: boolean;  // fill mode lighten/darken hints (rendered as opacity tweaks)
  lighten?: boolean;
}

const fmt = (v: number) => {
  const r = Math.round(v * 100) / 100;
  return Object.is(r, -0) ? "0" : String(r);
};

/** Evaluate a preset to renderable SVG paths at the given size. */
export function presetPaths(name: string, w: number, h: number, adj?: Record<string, number>): EvaluatedPath[] | null {
  const def = DEFS[name];
  if (!def) return null;
  const env = buildEnv(def, Math.max(w, 0.01), Math.max(h, 0.01), adj);

  return def.paths.map(path => {
    const sx = path.w ? w / path.w : 1;
    const sy = path.h ? h / path.h : 1;
    let d = "";
    let cx = 0, cy = 0;
    for (const cmd of path.cmds) {
      const op = cmd[0];
      const arg = (i: number) => resolve(cmd[i] as string, env);
      switch (op) {
        case "M": cx = arg(1) * sx; cy = arg(2) * sy; d += `M${fmt(cx)} ${fmt(cy)}`; break;
        case "L": cx = arg(1) * sx; cy = arg(2) * sy; d += `L${fmt(cx)} ${fmt(cy)}`; break;
        case "C": {
          const pts = [1, 2, 3, 4, 5, 6].map(i => arg(i) * (i % 2 ? sx : sy));
          d += `C${fmt(pts[0])} ${fmt(pts[1])} ${fmt(pts[2])} ${fmt(pts[3])} ${fmt(pts[4])} ${fmt(pts[5])}`;
          cx = pts[4]; cy = pts[5];
          break;
        }
        case "Q": {
          const pts = [1, 2, 3, 4].map(i => arg(i) * (i % 2 ? sx : sy));
          d += `Q${fmt(pts[0])} ${fmt(pts[1])} ${fmt(pts[2])} ${fmt(pts[3])}`;
          cx = pts[2]; cy = pts[3];
          break;
        }
        case "A": {
          // arcTo: ellipse radii + start/swing angles; current point lies ON the arc at stAng
          const wR = arg(1) * sx, hR = arg(2) * sy;
          const st = arg(3) * A, sw = arg(4) * A;
          const ecx = cx - wR * Math.cos(st), ecy = cy - hR * Math.sin(st);
          const ex = ecx + wR * Math.cos(st + sw), ey = ecy + hR * Math.sin(st + sw);
          const large = Math.abs(sw) > Math.PI ? 1 : 0;
          const sweep = sw > 0 ? 1 : 0;
          // full-circle arcs collapse in SVG — split into two halves
          if (Math.abs(sw) >= Math.PI * 2 - 1e-6) {
            const mx = ecx + wR * Math.cos(st + sw / 2), my = ecy + hR * Math.sin(st + sw / 2);
            d += `A${fmt(wR)} ${fmt(hR)} 0 0 ${sweep} ${fmt(mx)} ${fmt(my)}`;
            d += `A${fmt(wR)} ${fmt(hR)} 0 0 ${sweep} ${fmt(ex)} ${fmt(ey)}`;
          } else {
            d += `A${fmt(wR)} ${fmt(hR)} 0 ${large} ${sweep} ${fmt(ex)} ${fmt(ey)}`;
          }
          cx = ex; cy = ey;
          break;
        }
        case "Z": d += "Z"; break;
      }
    }
    return {
      d,
      noFill: path.fill === "none",
      noStroke: path.stroke === 0,
      darken: path.fill === "darken" || path.fill === "darkenLess",
      lighten: path.fill === "lighten" || path.fill === "lightenLess",
    };
  });
}

/** Single concatenated outline (ghost previews, hit areas, gallery icons). */
export function presetOutline(name: string, w: number, h: number, adj?: Record<string, number>): string {
  const paths = presetPaths(name, w, h, adj);
  if (!paths) return `M0 0H${w}V${h}H0Z`;
  return paths.map(p => p.d).join("");
}
