import React, { useState } from "react";
import type { ChartKind, ChartShape } from "../model/types";
import { CHART_NAMES } from "../model/defaults";
import { store } from "../state/store";

/** Spreadsheet-lite editor for a chart's categories/series — the stand-in for OnlyOffice's embedded sheet. */
export function ChartDataDialog({ shape, onClose }: { shape: ChartShape; onClose: () => void }) {
  const [kind, setKind] = useState<ChartKind>(shape.chart);
  const [title, setTitle] = useState(shape.title ?? "");
  const [legend, setLegend] = useState(shape.legend);
  const [cats, setCats] = useState<string[]>([...shape.categories]);
  const [series, setSeries] = useState(shape.series.map(s => ({ name: s.name, values: [...s.values] })));

  const setVal = (si: number, ci: number, v: string) => {
    setSeries(prev => prev.map((s, i) => i !== si ? s : { ...s, values: s.values.map((x, j) => j !== ci ? x : parseFloat(v) || 0) }));
  };

  const addCategory = () => {
    setCats(p => [...p, `Category ${p.length + 1}`]);
    setSeries(p => p.map(s => ({ ...s, values: [...s.values, 0] })));
  };
  const removeCategory = (ci: number) => {
    if (cats.length <= 1) return;
    setCats(p => p.filter((_, i) => i !== ci));
    setSeries(p => p.map(s => ({ ...s, values: s.values.filter((_, i) => i !== ci) })));
  };
  const addSeries = () => setSeries(p => [...p, { name: `Series ${p.length + 1}`, values: cats.map(() => 0) }]);
  const removeSeries = (si: number) => {
    if (series.length <= 1) return;
    setSeries(p => p.filter((_, i) => i !== si));
  };

  const apply = () => {
    store.updateShapes([shape.id], s => s.kind !== "chart" ? s : {
      ...s,
      chart: kind,
      title: title.trim() || undefined,
      legend,
      categories: cats,
      series: series.map((sr, i) => ({ ...shape.series[i], name: sr.name, values: sr.values, color: shape.series[i]?.color })),
    });
    onClose();
  };

  return (
    <div className="modal-overlay" onClick={onClose}>
      <div className="modal chart-modal" onClick={e => e.stopPropagation()}>
        <div className="modal-title">Chart Data</div>

        <div className="chart-form-row">
          <label>Type
            <select value={kind} onChange={e => setKind(e.target.value as ChartKind)}>
              {(Object.keys(CHART_NAMES) as ChartKind[]).map(k => (
                <option key={k} value={k}>{CHART_NAMES[k]}</option>
              ))}
            </select>
          </label>
          <label>Title
            <input value={title} placeholder="(none)" onChange={e => setTitle(e.target.value)} />
          </label>
          <label className="chk">
            <input type="checkbox" checked={legend} onChange={e => setLegend(e.target.checked)} /> Legend
          </label>
        </div>

        <div className="chart-grid-wrap">
          <table className="chart-grid">
            <thead>
              <tr>
                <th></th>
                {series.map((s, si) => (
                  <th key={si}>
                    <input value={s.name} onChange={e => setSeries(p => p.map((x, i) => i !== si ? x : { ...x, name: e.target.value }))} />
                    <button className="mini-x" title="Remove series" onClick={() => removeSeries(si)}>×</button>
                  </th>
                ))}
                <th className="add-col"><button className="link-btn" onClick={addSeries}>+ Series</button></th>
              </tr>
            </thead>
            <tbody>
              {cats.map((cat, ci) => (
                <tr key={ci}>
                  <td>
                    <input value={cat} onChange={e => setCats(p => p.map((x, i) => i !== ci ? x : e.target.value))} />
                    <button className="mini-x" title="Remove category" onClick={() => removeCategory(ci)}>×</button>
                  </td>
                  {series.map((s, si) => (
                    <td key={si}>
                      <input
                        type="number"
                        value={Number.isFinite(s.values[ci]) ? s.values[ci] : 0}
                        onChange={e => setVal(si, ci, e.target.value)}
                      />
                    </td>
                  ))}
                  <td className="add-col"></td>
                </tr>
              ))}
              <tr>
                <td className="add-col"><button className="link-btn" onClick={addCategory}>+ Category</button></td>
                <td colSpan={series.length + 1}></td>
              </tr>
            </tbody>
          </table>
        </div>

        <div className="modal-actions">
          <button className="pane-btn primary" onClick={apply}>OK</button>
          <button className="pane-btn" onClick={onClose}>Cancel</button>
        </div>
      </div>
    </div>
  );
}
