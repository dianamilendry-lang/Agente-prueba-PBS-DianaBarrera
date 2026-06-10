import { useState, useRef, useEffect, useCallback } from "react";
import { loadProjectIndex, loadProject, saveProject, deleteProject, storageMode } from "./lib/storage";
import * as XLSX from "xlsx";
import {
  BarChart, Bar, LineChart, Line, XAxis, YAxis,
  CartesianGrid, Tooltip, ResponsiveContainer, Cell, Legend,
} from "recharts";
import {
  Send, BarChart3, TrendingUp, Calculator, Layers,
  FileSpreadsheet, ChevronRight, Menu, X, ShieldCheck,
  Upload, FolderPlus, Folder, MessageSquare, Download,
  Boxes, Trash2, Check, Target, Disc, ShoppingBag,
  Award, AlertTriangle, FileCheck, Lightbulb,
} from "lucide-react";

// ─── Constants ────────────────────────────────────────────────────────────────
const NAVY = "#0a2a4a";
const NAVY2 = "#103a63";
const BLUE = "#1668b3";

const STRATEGY_PROMPT = `Con base en el análisis ejecutivo previo de esta conversación y en mis DATOS CARGADOS (cumplimiento, ventas, inventario), diseña ESTRATEGIAS DE VENTAS ACCIONABLES PARA EL PRÓXIMO MES. Usa este formato:
**1. Resumen ejecutivo** — diagnóstico breve del mes analizado.
**2. Estrategias por prioridad** — cada una con: objetivo, acción concreta, línea/cliente foco, responsable sugerido e impacto esperado en cifras.
**3. Metas y proyección del próximo mes** — con cifras base reales de mis datos.
**4. Riesgos y supuestos.**
Cierra con conclusión y recomendaciones. Incluye al final el bloque de gráficas (proyección/metas) en el formato indicado.`;

const SYSTEM_PROMPT = `Eres "GPT Ingeniera Comercialización", Asistente Senior para Comercialización y Finanzas Industriales de POLYTEC — GRUPO SIGMA Q (Polímeros y Tecnología, S.A.), especializado en empaque plástico flexible (bobinas impresas, bolsas, películas y laminados para los sectores agrícola, comercio, alimentos, limpieza, bebidas y textiles).

OBJETIVO: Apoyar el análisis financiero, comercial y estratégico para facilitar decisiones, SIN sustituir el criterio ejecutivo del usuario.

ESPECIALIDADES: Comercialización (presupuesto de ventas, proyecciones, márgenes, portafolio, rentabilidad, cumplimiento, estrategias de venta); Finanzas industriales (flujo de caja, PNL, costeo, escenarios, sensibilidad); Empaque flexible (producción–ventas–inventario, mezcla comercial, desempeño).

REGLAS (estrictas):
1. NO inventar cifras. Usa SOLO los datos provistos.
2. NO asumir información faltante; solicítala.
3. Separar SIEMPRE: DATOS, CÁLCULOS y CONCLUSIONES.
4. Explicar impacto financiero y comercial.
5. Mostrar riesgos y supuestos.
6. NO decidir por el usuario; presenta opciones.

FORMATO (cuando aplique):
**1. Resumen ejecutivo**
**2. Análisis detallado**
**3. Escenarios comparativos**
**4. Observaciones y riesgos**

TONO: Visión estratégica y autoridad ejecutiva, como directora de negocio, con empatía. Dirígete al usuario como "Jefa". Cierra con conclusión y recomendaciones. Responde en español. Usa las cifras reales de los DATOS CARGADOS. Si hay reporte de cumplimiento o presupuesto, calcula cumplimiento (%) y desviaciones por línea.

GRÁFICAS: Cuando entregues un análisis ejecutivo y existan DATOS CARGADOS, además del texto agrega AL FINAL un bloque para graficar los resultados, SOLO con cifras reales o derivadas de ellas, máximo 3 gráficas. Formato EXACTO, sin explicarlo:
\`\`\`charts
[{"title":"Top líneas por venta","type":"bar","data":[{"name":"Línea A","value":12500}]},{"title":"Real vs Presupuesto","type":"grouped","series":["Real","Presupuesto"],"data":[{"name":"Abril","Real":98000,"Presupuesto":110000}]},{"title":"Tendencia","type":"line","data":[{"name":"Ene","value":80000}]}]
\`\`\`
Usa "grouped" para comparativos (real vs meta), "line" para series de tiempo y "bar" para rankings. No inventes valores que no se deriven de los datos.`;

const QUICK = [
  { icon: FileSpreadsheet, label: "Analizar cumplimiento del mes", prompt: "Analiza el reporte de cumplimiento mensual cargado: lectura ejecutiva del desempeño, cumplimiento vs meta y principales desviaciones. Grafica los resultados." },
  { icon: Lightbulb, label: "Estrategias de venta · próximo mes", prompt: STRATEGY_PROMPT },
  { icon: Target, label: "Ventas reales vs presupuesto", prompt: "Con mis datos cargados, compara ventas reales contra presupuesto: cumplimiento (%), desviaciones y dónde están las brechas. Grafica los resultados." },
  { icon: Boxes, label: "Inventario vs demanda", prompt: "Evalúa mi inventario cargado frente al comportamiento de ventas. Señala excesos, faltantes y riesgos de rotación. Grafica los resultados." },
  { icon: Calculator, label: "Impacto de alza de resina", prompt: "Si la resina sube de precio, ¿qué impacto tendría en mis márgenes según mis datos cargados? Dime qué datos adicionales necesitas." },
  { icon: Layers, label: "Escenarios 2026", prompt: "Construye escenario optimista, base y pesimista para 2026 a partir de mis datos cargados y supuestos razonables que propongas. Grafica los escenarios." },
];

const SPECIALTIES = [
  { t: "Comercialización", i: ["Presupuesto de ventas", "Cumplimiento", "Estrategias de venta", "Márgenes", "Rentabilidad"] },
  { t: "Finanzas industriales", i: ["Flujo de caja", "PNL", "Costeo", "Escenarios", "Sensibilidad"] },
  { t: "Empaque flexible", i: ["Producción–ventas–inventario", "Mezcla comercial", "Desempeño"] },
];

const SLOT = {
  sales:     { t: "Reporte de ventas / cumplimiento", ic: TrendingUp, icon: "text-blue-700",  ctx: "REPORTE DE VENTAS / CUMPLIMIENTO" },
  budget:    { t: "Presupuesto de ventas",            ic: Target,     icon: "text-sky-600",   ctx: "PRESUPUESTO DE VENTAS" },
  inventory: { t: "Reporte de inventario",            ic: Boxes,      icon: "text-cyan-700",  ctx: "REPORTE DE INVENTARIO" },
};
const SLOT_KEYS = ["sales", "budget", "inventory"];

const HERO_TOP = [
  { ic: Disc,      l: "Bobinas impresas" },
  { ic: ShoppingBag, l: "Bolsas personalizadas" },
  { ic: ShieldCheck, l: "Alta barrera y protección" },
  { ic: Award,     l: "Calidad e innovación" },
];
const SECTORS = ["Agrícola", "Comercio", "Alimentos", "Limpieza", "Bebidas", "Textiles"];

const COLORS = [
  "#0a2a4a","#1668b3","#2b8fd6","#5aa9e6","#0d4e8b",
  "#7bbde8","#103a63","#3f97d8","#8fc7ec","#185a96","#4ea3df","#a9d4f0",
];

// ─── Helpers ──────────────────────────────────────────────────────────────────
const fmt = (n) =>
  n == null || isNaN(n) ? "—" : Number(n).toLocaleString("es-GT", { maximumFractionDigits: 0 });

function num(v) {
  if (typeof v === "number") return v;
  if (typeof v === "string") { const n = parseFloat(v.replace(/[^0-9.\-]/g, "")); return isNaN(n) ? null : n; }
  return null;
}

function extractCharts(text) {
  const m = text.match(/```charts\s*([\s\S]*?)```/i) || text.match(/```json\s*([\s\S]*?)```/i);
  if (!m) return { clean: text, charts: [] };
  let charts = [];
  try { const j = JSON.parse(m[1].trim()); if (Array.isArray(j)) charts = j; } catch {}
  charts = charts.filter((c) => c && Array.isArray(c.data) && c.data.length);
  return { clean: charts.length ? text.replace(m[0], "").trim() : text, charts };
}


// ─── Data processing ──────────────────────────────────────────────────────────
function cleanAndAnalyze(wb, fileName) {
  const ws = wb.Sheets[wb.SheetNames[0]];
  const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, blankrows: false });
  if (!aoa.length) return null;

  let headerIdx = 0, best = -1;
  for (let i = 0; i < Math.min(aoa.length, 12); i++) {
    const row = aoa[i] || [];
    const nonEmpty = row.filter((c) => c != null && String(c).trim() !== "").length;
    const txt = row.filter((c) => typeof c === "string" && String(c).trim() !== "" && num(c) == null).length;
    const sc = nonEmpty + txt * 1.5;
    if (sc > best) { best = sc; headerIdx = i; }
  }

  let header = (aoa[headerIdx] || []).map((c, j) =>
    c == null || String(c).trim() === "" ? "Columna " + (j + 1) : String(c).trim().replace(/\s+/g, " ")
  );
  const seen = {};
  header = header.map((h) => { if (seen[h] != null) { seen[h] += 1; return h + " " + seen[h]; } seen[h] = 0; return h; });

  const raw = aoa.slice(headerIdx + 1);
  const totalRe = /^(total|subtotal|gran total|totales|suma|promedio)/i;
  let removedEmpty = 0, removedTotals = 0, removedDup = 0;
  const dup = new Set();
  let objs = [];

  raw.forEach((r) => {
    const cells = header.map((h, j) => {
      let v = r[j];
      if (typeof v === "string") { v = v.trim().replace(/\s+/g, " "); if (v === "") v = null; }
      return v;
    });
    if (cells.every((v) => v == null)) { removedEmpty++; return; }
    const first = cells.find((v) => v != null);
    if (typeof first === "string" && totalRe.test(first)) { removedTotals++; return; }
    const key = JSON.stringify(cells);
    if (dup.has(key)) { removedDup++; return; }
    dup.add(key);
    const o = {};
    header.forEach((h, j) => (o[h] = cells[j]));
    objs.push(o);
  });

  const keepCols = header.filter((h) => objs.some((o) => o[h] != null));
  const droppedCols = header.length - keepCols.length;
  objs = objs.map((o) => { const n = {}; keepCols.forEach((h) => (n[h] = o[h])); return n; });
  const cols = keepCols;

  const numericCols = cols.filter((c) => objs.filter((o) => num(o[c]) != null).length > objs.length * 0.5);
  objs.forEach((o) => numericCols.forEach((c) => (o[c] = num(o[c]))));

  const reVal = /cumpl|venta|real|total|monto|importe|valor|ingreso|presupuest|meta|cantidad|existencia|stock|inventar|unidad|kg|peso/i;
  const reLab = /producto|cliente|sku|material|item|categor|descrip|nombre|art|presentaci|linea|línea|familia|vendedor|zona|regi/i;
  const reTime = /fecha|mes|periodo|año|year|month|date|trimestre|semana/i;

  let valueCol =
    numericCols.find((c) => reVal.test(c) && !/%|porc|cumpl/i.test(c)) ||
    numericCols.find((c) => reVal.test(c));
  if (!valueCol && numericCols.length)
    valueCol = numericCols.reduce((a, b) =>
      objs.reduce((s, o) => s + Math.abs(o[b] || 0), 0) > objs.reduce((s, o) => s + Math.abs(o[a] || 0), 0) ? b : a
    );

  const labelCol = cols.find((c) => reLab.test(c) && c !== valueCol) || cols.find((c) => !numericCols.includes(c)) || cols[0];
  const timeCol = cols.find((c) => reTime.test(c) && c !== valueCol && c !== labelCol);

  const group = (key) => {
    const m = new Map();
    objs.forEach((o) => {
      const k = o[key] == null ? "—" : String(o[key]).trim();
      const v = valueCol ? (o[valueCol] || 0) : 1;
      m.set(k, (m.get(k) || 0) + v);
    });
    return [...m.entries()].map(([name, value]) => ({
      name: name.length > 22 ? name.slice(0, 20) + "…" : name,
      value,
    }));
  };

  const byCategory = group(labelCol).sort((a, b) => b.value - a.value).slice(0, 12);
  const byTime = timeCol ? group(timeCol).slice(0, 24) : null;

  const vals = valueCol ? objs.map((o) => o[valueCol] || 0) : [];
  const total = vals.reduce((s, v) => s + v, 0);
  const kpis = valueCol
    ? { total, count: objs.length, avg: total / (objs.length || 1), max: vals.length ? Math.max(...vals) : 0 }
    : { total: null, count: objs.length, avg: null, max: null };

  const numericSummary = numericCols.map((c) => ({ col: c, sum: objs.reduce((s, o) => s + (o[c] || 0), 0) }));

  const warnings = [];
  if (headerIdx > 0) warnings.push("Encabezados en la fila " + (headerIdx + 1) + " (se ignoraron filas previas).");
  if (!numericCols.length) warnings.push("No se detectaron columnas numéricas; revise el formato.");
  if (valueCol && vals.some((v) => v < 0)) warnings.push(`Valores negativos en "${valueCol}".`);
  if (valueCol) {
    const b = objs.filter((o) => o[valueCol] == null).length;
    if (b) warnings.push(`${b} registro(s) sin valor en "${valueCol}".`);
  }

  const report = { rowsRead: raw.length, headerRow: headerIdx + 1, kept: objs.length, removedEmpty, removedTotals, removedDup, droppedCols, numericCols: numericCols.length, warnings };
  return { fileName, cols, valueCol, labelCol, timeCol, byCategory, byTime, kpis, rowCount: objs.length, numericSummary, report };
}

function dataContext(p) {
  if (!p) return "";
  let s = "\n\n=== DATOS CARGADOS (cifras reales, ya validadas y limpias; no inventes otras) ===\n";
  SLOT_KEYS.forEach((k) => {
    const d = p[k];
    if (!d) return;
    s += `\n[${SLOT[k].ctx}] archivo: ${d.fileName} · ${d.rowCount} registros limpios · columnas: ${d.cols.join(", ")}.\n`;
    if (d.numericSummary?.length)
      s += `Totales por columna numérica: ` + d.numericSummary.slice(0, 8).map((x) => `${x.col}=${fmt(x.sum)}`).join("; ") + ".\n";
    s += `Top por "${d.labelCol}" (métrica ${d.valueCol || "conteo"}): ` + d.byCategory.slice(0, 8).map((x) => `${x.name}=${fmt(x.value)}`).join("; ") + ".\n";
    if (d.byTime)
      s += `Serie por "${d.timeCol}": ` + d.byTime.map((x) => `${x.name}=${fmt(x.value)}`).join("; ") + ".\n";
  });
  return s;
}

// ─── Markdown renderer ────────────────────────────────────────────────────────
function renderInline(text, k) {
  return text.split(/(\*\*[^*]+\*\*)/g).map((p, i) =>
    p.startsWith("**") && p.endsWith("**")
      ? <strong key={k + i} className="font-semibold text-slate-900">{p.slice(2, -2)}</strong>
      : <span key={k + i}>{p}</span>
  );
}

function Markdown({ text }) {
  const lines = text.split("\n");
  const out = [];
  let bul = [];

  const flush = (key) => {
    if (bul.length) {
      out.push(
        <ul key={"u" + key} className="my-2 space-y-1.5">
          {bul.map((b, i) => (
            <li key={i} className="flex gap-2 leading-relaxed text-slate-700">
              <span className="mt-2 h-1.5 w-1.5 shrink-0 rounded-full bg-sky-500" />
              <span>{renderInline(b, "b" + key + i)}</span>
            </li>
          ))}
        </ul>
      );
      bul = [];
    }
  };

  lines.forEach((raw, idx) => {
    const line = raw.trimEnd();
    if (/^\s*[-*]\s+/.test(line)) { bul.push(line.replace(/^\s*[-*]\s+/, "")); return; }
    flush(idx);
    if (line.trim() === "") { out.push(<div key={"s" + idx} className="h-2" />); return; }
    const h = line.match(/^\*\*(\d+\.\s.*?)\*\*\s*(.*)$/);
    if (h) {
      out.push(
        <h4 key={"h" + idx} className="mb-1 mt-4 flex items-center gap-2 text-xs font-bold uppercase tracking-wide text-blue-800">
          <span className="h-4 w-1 rounded bg-blue-700" />{h[1]}
        </h4>
      );
      if (h[2]) out.push(<p key={"hp" + idx} className="leading-relaxed text-slate-700">{renderInline(h[2], "hp" + idx)}</p>);
      return;
    }
    out.push(<p key={"p" + idx} className="leading-relaxed text-slate-700">{renderInline(line, "p" + idx)}</p>);
  });

  flush("e");
  return <div className="text-sm">{out}</div>;
}

function mdToHtml(t) {
  return t.split("\n").map((l) => {
    l = l.replace(/\*\*(.+?)\*\*/g, "<strong>$1</strong>");
    if (/^\s*[-*]\s+/.test(l)) return "<li>" + l.replace(/^\s*[-*]\s+/, "") + "</li>";
    if (l.trim() === "") return "<br/>";
    return "<p>" + l + "</p>";
  }).join("");
}

// ─── Sub-components ───────────────────────────────────────────────────────────
function ChartCard({ c }) {
  if (!c || !Array.isArray(c.data) || !c.data.length) return null;
  return (
    <div className="rounded-lg border border-slate-200 bg-slate-50 p-2.5">
      <p className="mb-1 flex items-center gap-1.5 text-xs font-bold text-blue-800">
        <BarChart3 className="h-3.5 w-3.5" />{c.title || "Resultado"}
      </p>
      <div style={{ height: 200 }}>
        <ResponsiveContainer width="100%" height="100%">
          {c.type === "line" ? (
            <LineChart data={c.data} margin={{ top: 6, right: 10, bottom: 4, left: 0 }}>
              <CartesianGrid strokeDasharray="3 3" stroke="#e2e8f0" />
              <XAxis dataKey="name" tick={{ fontSize: 9, fill: "#64748b" }} />
              <YAxis tick={{ fontSize: 9, fill: "#64748b" }} tickFormatter={fmt} width={44} />
              <Tooltip formatter={(v) => fmt(v)} contentStyle={{ fontSize: 12, borderRadius: 8 }} />
              <Line type="monotone" dataKey="value" stroke="#1668b3" strokeWidth={2.5} dot={{ r: 3 }} />
            </LineChart>
          ) : c.type === "grouped" && Array.isArray(c.series) ? (
            <BarChart data={c.data} margin={{ top: 6, right: 10, bottom: 4, left: 0 }}>
              <CartesianGrid strokeDasharray="3 3" stroke="#e2e8f0" />
              <XAxis dataKey="name" tick={{ fontSize: 9, fill: "#64748b" }} interval={0} angle={-20} textAnchor="end" height={44} />
              <YAxis tick={{ fontSize: 9, fill: "#64748b" }} tickFormatter={fmt} width={44} />
              <Tooltip formatter={(v) => fmt(v)} contentStyle={{ fontSize: 12, borderRadius: 8 }} />
              <Legend wrapperStyle={{ fontSize: 11 }} />
              {c.series.map((s, i) => <Bar key={s} dataKey={s} fill={COLORS[i % COLORS.length]} radius={[3, 3, 0, 0]} />)}
            </BarChart>
          ) : (
            <BarChart data={c.data} margin={{ top: 6, right: 10, bottom: 4, left: 0 }}>
              <CartesianGrid strokeDasharray="3 3" stroke="#e2e8f0" />
              <XAxis dataKey="name" tick={{ fontSize: 9, fill: "#64748b" }} interval={0} angle={-20} textAnchor="end" height={48} />
              <YAxis tick={{ fontSize: 9, fill: "#64748b" }} tickFormatter={fmt} width={44} />
              <Tooltip formatter={(v) => fmt(v)} contentStyle={{ fontSize: 12, borderRadius: 8 }} />
              <Bar dataKey="value" radius={[3, 3, 0, 0]}>
                {c.data.map((e, i) => <Cell key={i} fill={COLORS[i % COLORS.length]} />)}
              </Bar>
            </BarChart>
          )}
        </ResponsiveContainer>
      </div>
    </div>
  );
}

const KPI = ({ label, value }) => (
  <div className="rounded-xl border border-slate-200 bg-white p-3">
    <p className="text-xs font-medium text-slate-500">{label}</p>
    <p className="mt-0.5 text-xl font-bold text-slate-900">{value}</p>
  </div>
);

function ReportCard({ d }) {
  const r = d.report;
  if (!r) return null;
  const chips = [["vacías", r.removedEmpty], ["totales", r.removedTotals], ["duplicadas", r.removedDup], ["columnas vacías", r.droppedCols]]
    .filter(([, n]) => n > 0);
  return (
    <div className="mt-2 rounded-lg border border-slate-200 bg-slate-50 p-2.5 text-xs">
      <div className="flex items-center gap-1.5 font-semibold text-emerald-700">
        <FileCheck className="h-3.5 w-3.5" />Datos validados: {r.kept} filas limpias
      </div>
      <p className="mt-0.5 text-slate-500">
        De {r.rowsRead} leídas · {r.numericCols} columna(s) numérica(s)
        {r.headerRow > 1 ? ` · encabezado fila ${r.headerRow}` : ""}
      </p>
      {chips.length > 0 && (
        <div className="mt-1 flex flex-wrap gap-1">
          {chips.map(([t, n]) => <span key={t} className="rounded bg-slate-200 px-1.5 py-0.5 text-slate-600">−{n} {t}</span>)}
        </div>
      )}
      {d.valueCol && <p className="mt-1 text-slate-600">Métrica: <b>{d.valueCol}</b> · Agrupación: <b>{d.labelCol}</b></p>}
      {r.warnings.map((w, i) => (
        <p key={i} className="mt-1 flex items-start gap-1 text-amber-700">
          <AlertTriangle className="mt-0.5 h-3 w-3 shrink-0" />{w}
        </p>
      ))}
    </div>
  );
}

function Logo({ size = 36 }) {
  return (
    <svg width={size} height={size} viewBox="0 0 40 40" fill="none">
      <defs>
        <linearGradient id="pl" x1="0" y1="0" x2="40" y2="40">
          <stop stopColor="#103a63" />
          <stop offset="1" stopColor="#1668b3" />
        </linearGradient>
      </defs>
      <rect width="40" height="40" rx="9" fill="url(#pl)" />
      <ellipse cx="16" cy="20" rx="7.5" ry="9.5" fill="none" stroke="#fff" strokeWidth="2.2" />
      <ellipse cx="16" cy="20" rx="2.4" ry="3.2" fill="#fff" />
      <path d="M22 12 C30 13 31 27 23.5 28.5" fill="none" stroke="#7bbde8" strokeWidth="2.2" strokeLinecap="round" />
    </svg>
  );
}

// ─── Main App ─────────────────────────────────────────────────────────────────
export default function App() {
  const [projects, setProjects] = useState([]);
  const [activeId, setActiveId] = useState(null);
  const [data, setData] = useState(null);
  const [input, setInput] = useState("");
  const [loading, setLoading] = useState(false);
  const [tab, setTab] = useState("chat");
  const [sidebar, setSidebar] = useState(false);
  const [newName, setNewName] = useState("");
  const [ready, setReady] = useState(false);
  const [busy, setBusy] = useState("");
  const scrollRef = useRef(null);

  // Load project index on mount
  useEffect(() => {
    loadProjectIndex()
      .then(setProjects)
      .catch(() => setProjects([]))
      .finally(() => setReady(true));
  }, []);

  useEffect(() => {
    if (scrollRef.current) scrollRef.current.scrollTop = scrollRef.current.scrollHeight;
  }, [data?.messages, loading]);

  const createProject = useCallback(async () => {
    const name = newName.trim() || "Proyecto " + (projects.length + 1);
    const id = "p" + Date.now();
    const payload = { name, messages: [], sales: null, budget: null, inventory: null };
    await saveProject(id, payload);
    const list = [...projects, { id, name }];
    setProjects(list);
    setNewName("");
    setActiveId(id);
    setData(payload);
    setTab("data");
    setSidebar(false);
  }, [newName, projects]);

  const openProject = useCallback(async (id) => {
    setActiveId(id);
    setSidebar(false);
    try {
      const p = await loadProject(id);
      setData(p || { name: projects.find((x) => x.id === id)?.name || "Proyecto", messages: [], sales: null, budget: null, inventory: null });
      setTab("chat");
    } catch {
      setData({ name: projects.find((x) => x.id === id)?.name || "Proyecto", messages: [], sales: null, budget: null, inventory: null });
    }
  }, [projects]);

  const removeProject = useCallback(async (id, e) => {
    e.stopPropagation();
    await deleteProject(id);
    const list = projects.filter((p) => p.id !== id);
    setProjects(list);
    if (activeId === id) { setActiveId(null); setData(null); }
  }, [projects, activeId]);

  const handleFile = useCallback((file, kind) => {
    if (!file || !data) return;
    setBusy(kind);
    const reader = new FileReader();
    reader.onload = async (e) => {
      try {
        const wb = XLSX.read(new Uint8Array(e.target.result), { type: "array" });
        const parsed = cleanAndAnalyze(wb, file.name);
        const upd = { ...data, [kind]: parsed };
        setData(upd);
        await saveProject(activeId, upd);
      } catch {
        alert("No pude leer el archivo. Verifica que sea Excel (.xlsx/.xls) válido, Jefa.");
      }
      setBusy("");
    };
    reader.readAsArrayBuffer(file);
  }, [data, activeId]);

  const send = async (text) => {
    const content = (text ?? input).trim();
    if (!content || loading || !data) return;
    const next = [...data.messages, { role: "user", content }];
    setData({ ...data, messages: next });
    setInput("");
    setLoading(true);
    setTab("chat");

    // En producción (Vercel) usa el proxy /api/chat.
    // En dev local usa Anthropic directo si VITE_ANTHROPIC_API_KEY está en .env.
    const isDev = import.meta.env.DEV;
    const devKey = import.meta.env.VITE_ANTHROPIC_API_KEY;

    if (isDev && !devKey) {
      const final = { ...data, messages: [...next, { role: "assistant", content: "⚠️ Configura `VITE_ANTHROPIC_API_KEY` en tu `.env` para usar el asistente en local, Jefa." }] };
      setData(final);
      await saveProject(activeId, final);
      setLoading(false);
      return;
    }

    const endpoint = isDev ? "https://api.anthropic.com/v1/messages" : "/api/chat";
    const headers = { "Content-Type": "application/json" };
    if (isDev) {
      headers["x-api-key"] = devKey;
      headers["anthropic-version"] = "2023-06-01";
      headers["anthropic-dangerous-direct-browser-access"] = "true";
    }

    try {
      const res = await fetch(endpoint, {
        method: "POST",
        headers,
        body: JSON.stringify({
          model: "claude-sonnet-4-20250514",
          max_tokens: 4096,
          system: SYSTEM_PROMPT + dataContext(data),
          messages: next.map((m) => ({ role: m.role, content: m.content })),
        }),
      });
      const d = await res.json();
      const reply = (d.content || []).filter((b) => b.type === "text").map((b) => b.text).join("\n").trim()
        || "No recibí respuesta. Intentemos de nuevo, Jefa.";
      const { clean, charts } = extractCharts(reply);
      const final = { ...data, messages: [...next, { role: "assistant", content: clean, charts }] };
      setData(final);
      await saveProject(activeId, final);
    } catch {
      const final = { ...data, messages: [...next, { role: "assistant", content: "Hubo un problema de conexión con el motor de análisis. Intentemos nuevamente, Jefa." }] };
      setData(final);
      await saveProject(activeId, final);
    } finally {
      setLoading(false);
    }
  };

  const exportReport = () => {
    if (!data) return;
    const last = [...data.messages].reverse().find((m) => m.role === "assistant");
    const date = new Date().toLocaleDateString("es-GT", { year: "numeric", month: "long", day: "numeric" });

    const block = (k) => {
      const d = data[k];
      if (!d) return "";
      return `<h3>${SLOT[k].ctx}</h3><p><b>Archivo:</b> ${d.fileName} · ${d.rowCount} registros limpios</p>
        ${d.numericSummary?.length ? `<table><tr>${d.numericSummary.slice(0, 6).map((x) => `<th>${x.col}</th>`).join("")}</tr><tr>${d.numericSummary.slice(0, 6).map((x) => `<td>${fmt(x.sum)}</td>`).join("")}</tr></table>` : ""}
        <table><tr><th>${d.labelCol}</th><th>${d.valueCol || "Conteo"}</th></tr>${d.byCategory.slice(0, 8).map((x) => `<tr><td>${x.name}</td><td>${fmt(x.value)}</td></tr>`).join("")}</table>`;
    };

    const chartHtml = (c) => {
      if (!c || !Array.isArray(c.data) || !c.data.length) return "";
      if (c.type === "grouped" && Array.isArray(c.series))
        return `<h4>${c.title || ""}</h4><table><tr><th>Concepto</th>${c.series.map((s) => `<th>${s}</th>`).join("")}</tr>${c.data.slice(0, 12).map((d) => `<tr><td>${d.name}</td>${c.series.map((s) => `<td>${fmt(d[s])}</td>`).join("")}</tr>`).join("")}</table>`;
      const max = Math.max(...c.data.map((d) => Math.abs(d.value || 0)), 1);
      return `<h4>${c.title || ""}</h4>` + c.data.slice(0, 12).map((d) => {
        const w = Math.round((Math.abs(d.value || 0) / max) * 240);
        return `<div style="display:flex;align-items:center;gap:8px;margin:3px 0;font-size:11px;font-family:Arial"><span style="width:150px;color:#475569;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${d.name}</span><span style="height:12px;width:${w}px;background:#1668b3;border-radius:2px;display:inline-block"></span><span style="color:#0f172a">${fmt(d.value)}</span></div>`;
      }).join("");
    };

    const charts = last?.charts?.length ? last.charts.map(chartHtml).join("") : "";

    const html = `<!doctype html><html lang="es"><head><meta charset="utf-8"><title>Análisis — ${data.name}</title>
<style>body{font-family:Georgia,serif;color:#1e293b;max-width:760px;margin:0 auto;padding:40px}
.head{display:flex;align-items:center;gap:14px;border-bottom:3px solid #0a2a4a;padding-bottom:14px;margin-bottom:8px}
.logo{width:46px;height:46px;border-radius:11px;background:linear-gradient(135deg,#103a63,#1668b3)}
h1{font-size:20px;margin:0;color:#0a2a4a}.sub{color:#64748b;font-size:12px;letter-spacing:1px;text-transform:uppercase}
h2{font-size:16px;color:#0a2a4a;border-left:4px solid #1668b3;padding-left:8px;margin-top:26px}
h3{font-size:13px;color:#103a63;margin-top:18px}h4{color:#0a2a4a;font-size:12px;text-transform:uppercase;margin:16px 0 4px}
table{border-collapse:collapse;width:100%;margin:8px 0;font-size:12px;font-family:Arial}
th,td{border:1px solid #cbd5e1;padding:5px 8px;text-align:left}th{background:#e8f1fb}
p,li{font-size:13px;line-height:1.6}.foot{margin-top:34px;border-top:1px solid #e2e8f0;padding-top:10px;color:#94a3b8;font-size:11px}</style></head>
<body><div class="head"><div class="logo"></div><div><div class="sub">POLYTEC — Grupo Sigma Q</div><h1>Análisis Ejecutivo — ${data.name}</h1></div></div>
<p class="sub">Generado el ${date}</p>
<h2>Datos analizados</h2>${block("sales") || "<p>Sin datos de ventas cargados.</p>"}${block("budget")}${block("inventory")}
<h2>Lectura del asistente</h2>${last ? mdToHtml(last.content) : "<p>Aún no hay análisis generado.</p>"}
${charts ? `<h2>Gráficas del análisis</h2>${charts}` : ""}
<div class="foot">Generado por GPT Ingeniera Comercialización — POLYTEC · Grupo Sigma Q (Polímeros y Tecnología, S.A.). Asistente analítico; no sustituye el criterio ejecutivo. Verifique siempre las cifras.</div></body></html>`;

    const blob = new Blob([html], { type: "text/html" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `Analisis_${data.name.replace(/\s+/g, "_")}.html`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  };

  const D = data;
  const hasAnalysis = D && D.messages.some((m) => m.role === "assistant");

  return (
    <div className="flex h-screen w-full overflow-hidden bg-slate-50 font-sans">
      {/* Sidebar overlay on mobile */}
      {sidebar && <div className="fixed inset-0 z-20 bg-black/40 md:hidden" onClick={() => setSidebar(false)} />}

      {/* Sidebar */}
      <aside
        className={`${sidebar ? "fixed inset-y-0 left-0 z-30 w-72" : "hidden"} flex flex-col text-slate-100 md:relative md:flex md:w-72 md:shrink-0`}
        style={{ backgroundImage: `linear-gradient(180deg, ${NAVY}, ${NAVY2})` }}
      >
        <div className="flex items-center justify-between border-b border-white/10 p-4">
          <div className="flex items-center gap-2.5">
            <Logo size={36} />
            <div>
              <div className="text-sm font-extrabold leading-none tracking-tight">POLYTEC</div>
              <div className="text-[11px] font-medium tracking-wider text-sky-300">GRUPO SIGMA Q</div>
            </div>
          </div>
          <button onClick={() => setSidebar(false)} className="md:hidden"><X className="h-5 w-5 text-slate-300" /></button>
        </div>

        <div className="border-b border-white/10 p-3">
          <div className="mb-1.5 flex items-center gap-1.5 text-xs font-bold uppercase tracking-wider text-sky-300">
            <FolderPlus className="h-3.5 w-3.5" />Nuevo proyecto
          </div>
          <div className="flex gap-1.5">
            <input
              value={newName}
              onChange={(e) => setNewName(e.target.value)}
              onKeyDown={(e) => e.key === "Enter" && createProject()}
              placeholder="Ej. Cumplimiento Abr-26"
              className="min-w-0 flex-1 rounded-lg bg-white/10 px-2.5 py-1.5 text-xs text-white outline-none placeholder:text-slate-400 focus:ring-2 focus:ring-sky-400"
            />
            <button onClick={createProject} className="flex h-8 w-8 shrink-0 items-center justify-center rounded-lg bg-sky-500 hover:bg-sky-400">
              <Check className="h-4 w-4" />
            </button>
          </div>
        </div>

        <div className="flex-1 overflow-y-auto p-3">
          <div className="mb-1.5 text-xs font-bold uppercase tracking-wider text-slate-400">Proyectos</div>
          {projects.length === 0 && <p className="text-xs text-slate-500">Cree un proyecto para empezar.</p>}
          <div className="space-y-1">
            {projects.map((p) => (
              <div
                key={p.id}
                onClick={() => openProject(p.id)}
                className={`group flex cursor-pointer items-center justify-between rounded-lg px-2.5 py-2 text-sm transition-colors ${activeId === p.id ? "bg-sky-500/20 text-white" : "text-slate-300 hover:bg-white/5"}`}
              >
                <span className="flex items-center gap-2 truncate">
                  <Folder className="h-4 w-4 shrink-0 text-sky-400" />{p.name}
                </span>
                <button onClick={(e) => removeProject(p.id, e)} className="opacity-0 transition-opacity group-hover:opacity-100">
                  <Trash2 className="h-3.5 w-3.5 text-slate-400 hover:text-red-400" />
                </button>
              </div>
            ))}
          </div>

          <div className="mt-5 space-y-3">
            {SPECIALTIES.map((s) => (
              <div key={s.t}>
                <h3 className="mb-1 text-xs font-bold uppercase tracking-wider text-sky-400">{s.t}</h3>
                <ul className="space-y-0.5">
                  {s.i.map((it) => (
                    <li key={it} className="flex items-center gap-1.5 text-xs text-slate-400">
                      <ChevronRight className="h-3 w-3 shrink-0 text-sky-500" />{it}
                    </li>
                  ))}
                </ul>
              </div>
            ))}
          </div>
        </div>

        <div className="border-t border-white/10 p-3">
          <div className="mb-1 flex items-center gap-1.5 text-xs font-bold uppercase tracking-wider text-amber-400">
            <ShieldCheck className="h-3.5 w-3.5" />Reglas
          </div>
          <p className="text-xs leading-relaxed text-slate-400">
            No inventa cifras · Valida y limpia datos · Separa Datos / Cálculos / Conclusiones · No decide por usted.
          </p>
          <div className={`mt-2 inline-flex items-center gap-1 rounded px-1.5 py-0.5 text-[10px] font-semibold ${storageMode === "supabase" ? "bg-emerald-500/20 text-emerald-300" : "bg-slate-600 text-slate-400"}`}>
            {storageMode === "supabase" ? "● Supabase" : "○ Local"}
          </div>
        </div>
      </aside>

      {/* Main */}
      <div className="flex flex-1 flex-col overflow-hidden">
        {/* Header */}
        <header className="flex items-center justify-between border-b border-slate-200 bg-white px-4 py-2.5 shrink-0">
          <div className="flex items-center gap-3">
            <button onClick={() => setSidebar(true)} className="md:hidden">
              <Menu className="h-5 w-5 text-slate-600" />
            </button>
            <div className="flex items-center gap-2">
              <span className="md:hidden"><Logo size={28} /></span>
              <div>
                <h2 className="text-sm font-bold text-slate-900">{D ? D.name : "Ingeniera Comercialización"}</h2>
                <p className="text-xs text-slate-500">POLYTEC · Grupo Sigma Q — Empaque flexible</p>
              </div>
            </div>
          </div>
          {D && (
            <button onClick={exportReport} className="flex items-center gap-1.5 rounded-lg px-3 py-1.5 text-xs font-semibold text-white hover:opacity-90" style={{ backgroundColor: BLUE }}>
              <Download className="h-3.5 w-3.5" />Exportar
            </button>
          )}
        </header>

        {/* Tabs */}
        {D && (
          <div className="flex border-b border-slate-200 bg-white px-3 shrink-0">
            {[{ k: "chat", l: "Asistente", ic: MessageSquare }, { k: "dash", l: "Tablero", ic: BarChart3 }, { k: "data", l: "Datos", ic: Upload }].map((t) => (
              <button
                key={t.k}
                onClick={() => setTab(t.k)}
                className={`flex items-center gap-1.5 border-b-2 px-3 py-2 text-sm font-medium transition-colors ${tab === t.k ? "border-blue-700 text-blue-800" : "border-transparent text-slate-500 hover:text-slate-700"}`}
              >
                <t.ic className="h-4 w-4" />{t.l}
              </button>
            ))}
          </div>
        )}

        {/* Welcome screen */}
        {!D && (
          <div className="flex-1 overflow-y-auto">
            <div className="px-6 py-8 text-white" style={{ backgroundImage: `linear-gradient(135deg, ${NAVY} 60%, ${NAVY2})` }}>
              <div className="mx-auto max-w-2xl">
                <div className="mb-4 flex items-center gap-2.5">
                  <Logo size={40} />
                  <div>
                    <div className="text-base font-extrabold tracking-tight">POLYTEC</div>
                    <div className="text-xs font-medium tracking-widest text-sky-300">GRUPO SIGMA Q · EMPAQUE FLEXIBLE</div>
                  </div>
                </div>
                <h1 className="text-2xl font-extrabold leading-tight sm:text-3xl">EMPAQUE PLÁSTICO FLEXIBLE</h1>
                <p className="text-sm font-semibold tracking-wide text-sky-300">EN BOBINAS Y BOLSAS</p>
                <p className="mt-2 max-w-md text-sm text-slate-300">Soluciones de empaque flexibles para todo tipo de industria. Inteligencia comercial y financiera para su negocio.</p>
                <div className="mt-5 grid grid-cols-2 gap-4 sm:grid-cols-4">
                  {HERO_TOP.map((f) => (
                    <div key={f.l} className="flex flex-col items-center text-center">
                      <div className="flex h-12 w-12 items-center justify-center rounded-full" style={{ backgroundColor: BLUE }}>
                        <f.ic className="h-5 w-5 text-white" />
                      </div>
                      <span className="mt-1.5 text-xs font-medium leading-tight text-slate-200">{f.l}</span>
                    </div>
                  ))}
                </div>
                <div className="mt-6 flex flex-col gap-2 sm:flex-row">
                  <button
                    onClick={() => setSidebar(true)}
                    className="rounded-lg bg-white px-4 py-2 text-sm font-bold hover:bg-slate-100"
                    style={{ color: NAVY }}
                  >
                    Crear proyecto →
                  </button>
                </div>
              </div>
            </div>
            <div className="bg-white px-6 py-5">
              <div className="mx-auto max-w-2xl">
                <div className="mb-3 flex flex-wrap items-center gap-2">
                  <span className="text-xs font-bold uppercase tracking-wider text-slate-500">Sectores que atendemos</span>
                  {SECTORS.map((s) => (
                    <span key={s} className="rounded-full border border-slate-200 bg-slate-50 px-2.5 py-0.5 text-xs font-medium text-slate-700">{s}</span>
                  ))}
                </div>
                <div className="flex items-center gap-1.5 text-xs text-slate-500">
                  <ShieldCheck className="h-4 w-4 text-emerald-600" />
                  Primera planta en Centroamérica con Certificación FSSC 22000 en empaque flexible para alimentos.
                </div>
                {!ready && <p className="mt-3 text-center text-xs text-slate-400">Cargando proyectos…</p>}
              </div>
            </div>
          </div>
        )}

        {/* Data tab */}
        {D && tab === "data" && (
          <div className="flex-1 overflow-y-auto p-4">
            <div className="mx-auto max-w-3xl">
              <div className="grid gap-4 sm:grid-cols-3">
                {SLOT_KEYS.map((k) => {
                  const meta = SLOT[k];
                  const d = D[k];
                  return (
                    <div key={k} className="rounded-xl border border-slate-200 bg-white p-4">
                      <div className="mb-2 flex items-center gap-2">
                        <meta.ic className={`h-5 w-5 ${meta.icon}`} />
                        <h4 className="text-sm font-bold text-slate-800">{meta.t}</h4>
                      </div>
                      <label className={`flex cursor-pointer flex-col items-center justify-center rounded-lg border-2 border-dashed py-6 text-center transition-colors ${d ? "border-sky-300 bg-sky-50" : "border-slate-300 hover:border-blue-400 hover:bg-blue-50"}`}>
                        <input type="file" accept=".xlsx,.xls,.csv" className="hidden" onChange={(e) => handleFile(e.target.files[0], k)} />
                        {busy === k ? (
                          <span className="text-xs text-slate-500">Procesando…</span>
                        ) : d ? (
                          <>
                            <Check className="h-6 w-6 text-sky-600" />
                            <span className="mt-1 text-xs font-medium text-sky-700">{d.fileName}</span>
                            <span className="text-xs text-slate-500">{d.rowCount} reg. · reemplazar</span>
                          </>
                        ) : (
                          <>
                            <Upload className="h-6 w-6 text-slate-400" />
                            <span className="mt-1 text-xs font-medium text-slate-600">Cargar Excel</span>
                            <span className="text-xs text-slate-400">.xlsx · .xls · .csv</span>
                          </>
                        )}
                      </label>
                      {d && <ReportCard d={d} />}
                    </div>
                  );
                })}
              </div>
              <p className="mt-4 text-center text-xs text-slate-400">
                El motor detecta encabezados, descarta filas de totales/vacías/duplicadas, recorta columnas vacías y convierte texto a número antes de analizar.
              </p>
            </div>
          </div>
        )}

        {/* Dashboard tab */}
        {D && tab === "dash" && (
          <div className="flex-1 overflow-y-auto p-4">
            <div className="mx-auto max-w-3xl space-y-5">
              {!D.sales && !D.budget && !D.inventory && (
                <div className="rounded-xl border border-dashed border-slate-300 bg-white p-8 text-center text-sm text-slate-500">
                  Cargue un reporte en la pestaña <b>Datos</b> para ver el tablero.
                </div>
              )}
              {SLOT_KEYS.map((k) => {
                const d = D[k];
                if (!d) return null;
                const meta = SLOT[k];
                return (
                  <div key={k} className="rounded-xl border border-slate-200 bg-white p-4">
                    <h4 className="mb-3 flex items-center gap-2 text-sm font-bold text-slate-800">
                      <meta.ic className={`h-4 w-4 ${meta.icon}`} />{meta.t} · {d.fileName}
                    </h4>
                    <div className="mb-4 grid grid-cols-2 gap-2 sm:grid-cols-4">
                      <KPI label={`Total ${d.valueCol || ""}`} value={fmt(d.kpis.total)} />
                      <KPI label="Registros" value={fmt(d.kpis.count)} />
                      <KPI label="Promedio" value={fmt(d.kpis.avg)} />
                      <KPI label="Máximo" value={fmt(d.kpis.max)} />
                    </div>
                    <p className="mb-1 text-xs font-semibold text-slate-500">Top por {d.labelCol}</p>
                    <div style={{ height: 220 }}>
                      <ResponsiveContainer width="100%" height="100%">
                        <BarChart data={d.byCategory} margin={{ top: 4, right: 8, bottom: 4, left: 0 }}>
                          <CartesianGrid strokeDasharray="3 3" stroke="#f1f5f9" />
                          <XAxis dataKey="name" tick={{ fontSize: 9, fill: "#64748b" }} interval={0} angle={-30} textAnchor="end" height={56} />
                          <YAxis tick={{ fontSize: 10, fill: "#64748b" }} tickFormatter={fmt} width={48} />
                          <Tooltip formatter={(v) => fmt(v)} contentStyle={{ fontSize: 12, borderRadius: 8 }} />
                          <Bar dataKey="value" radius={[4, 4, 0, 0]}>
                            {d.byCategory.map((e, i) => <Cell key={i} fill={COLORS[i % COLORS.length]} />)}
                          </Bar>
                        </BarChart>
                      </ResponsiveContainer>
                    </div>
                    {d.byTime && (
                      <>
                        <p className="mb-1 mt-3 text-xs font-semibold text-slate-500">Serie por {d.timeCol}</p>
                        <div style={{ height: 180 }}>
                          <ResponsiveContainer width="100%" height="100%">
                            <LineChart data={d.byTime} margin={{ top: 4, right: 8, bottom: 4, left: 0 }}>
                              <CartesianGrid strokeDasharray="3 3" stroke="#f1f5f9" />
                              <XAxis dataKey="name" tick={{ fontSize: 9, fill: "#64748b" }} />
                              <YAxis tick={{ fontSize: 10, fill: "#64748b" }} tickFormatter={fmt} width={48} />
                              <Tooltip formatter={(v) => fmt(v)} contentStyle={{ fontSize: 12, borderRadius: 8 }} />
                              <Line type="monotone" dataKey="value" stroke="#1668b3" strokeWidth={2.5} dot={{ r: 3 }} />
                            </LineChart>
                          </ResponsiveContainer>
                        </div>
                      </>
                    )}
                  </div>
                );
              })}
            </div>
          </div>
        )}

        {/* Chat tab */}
        {D && tab === "chat" && (
          <>
            <div ref={scrollRef} className="flex-1 overflow-y-auto px-4 py-5">
              {D.messages.length === 0 ? (
                <div className="mx-auto max-w-2xl">
                  <div className="mb-5 text-center">
                    <div className="mx-auto mb-2 flex h-12 w-12 items-center justify-center">
                      <Logo size={48} />
                    </div>
                    <h3 className="text-base font-bold text-slate-900">¿En qué analizamos hoy, Jefa?</h3>
                    <p className="mt-1 text-sm text-slate-500">
                      {D.sales || D.budget || D.inventory
                        ? "Sus datos están validados y cargados. Elija un análisis o escríbame."
                        : "Sugiero cargar primero sus reportes en la pestaña Datos."}
                    </p>
                  </div>
                  <div className="grid gap-2.5 sm:grid-cols-2">
                    {QUICK.map((a) => (
                      <button
                        key={a.label}
                        onClick={() => send(a.prompt)}
                        className="group flex items-start gap-3 rounded-lg border border-slate-200 bg-white p-3 text-left transition-all hover:border-blue-400 hover:shadow-md"
                      >
                        <div className="flex h-9 w-9 shrink-0 items-center justify-center rounded-lg bg-blue-50 group-hover:bg-blue-100">
                          <a.icon className="h-4 w-4 text-blue-700" />
                        </div>
                        <span className="text-xs font-medium leading-snug text-slate-700">{a.label}</span>
                      </button>
                    ))}
                  </div>
                </div>
              ) : (
                <div className="mx-auto max-w-2xl space-y-4">
                  {D.messages.map((m, i) => (
                    <div key={i} className={`flex ${m.role === "user" ? "justify-end" : "justify-start"}`}>
                      {m.role === "assistant" && <div className="mr-2.5 mt-0.5 shrink-0"><Logo size={32} /></div>}
                      <div
                        className={`max-w-[88%] rounded-2xl px-4 py-3 ${m.role === "user" ? "text-white" : "border border-slate-200 bg-white shadow-sm"}`}
                        style={m.role === "user" ? { backgroundColor: NAVY } : {}}
                      >
                        {m.role === "user"
                          ? <p className="whitespace-pre-wrap text-sm leading-relaxed">{m.content}</p>
                          : (
                            <>
                              <Markdown text={m.content} />
                              {m.charts?.length > 0 && (
                                <div className="mt-3 space-y-3">
                                  {m.charts.map((c, ci) => <ChartCard key={ci} c={c} />)}
                                </div>
                              )}
                            </>
                          )}
                      </div>
                    </div>
                  ))}
                  {loading && (
                    <div className="flex justify-start">
                      <div className="mr-2.5 mt-0.5 shrink-0"><Logo size={32} /></div>
                      <div className="rounded-2xl border border-slate-200 bg-white px-4 py-3.5 shadow-sm">
                        <div className="flex gap-1.5">
                          {[0, 150, 300].map((d) => <span key={d} className="h-2 w-2 animate-bounce rounded-full bg-blue-500" style={{ animationDelay: d + "ms" }} />)}
                        </div>
                      </div>
                    </div>
                  )}
                </div>
              )}
            </div>

            <div className="border-t border-slate-200 bg-white px-4 py-3 shrink-0">
              {hasAnalysis && !loading && (
                <div className="mx-auto mb-2 flex max-w-2xl flex-wrap gap-2">
                  <button
                    onClick={() => send(STRATEGY_PROMPT)}
                    className="flex items-center gap-1.5 rounded-lg border border-amber-300 bg-amber-50 px-3 py-1.5 text-xs font-semibold text-amber-800 hover:bg-amber-100"
                  >
                    <Lightbulb className="h-3.5 w-3.5" />Estrategias de ventas · próximo mes
                  </button>
                </div>
              )}
              <div className="mx-auto flex max-w-2xl items-end gap-2">
                <textarea
                  value={input}
                  onChange={(e) => setInput(e.target.value)}
                  onKeyDown={(e) => { if (e.key === "Enter" && !e.shiftKey) { e.preventDefault(); send(); } }}
                  placeholder="Escriba su consulta de análisis…"
                  rows={1}
                  className="max-h-32 flex-1 resize-none rounded-xl border border-slate-300 px-4 py-2.5 text-sm text-slate-800 outline-none focus:border-blue-500 focus:ring-2 focus:ring-blue-100"
                />
                <button
                  onClick={() => send()}
                  disabled={loading || !input.trim()}
                  className="flex h-11 w-11 shrink-0 items-center justify-center rounded-xl text-white disabled:opacity-40"
                  style={{ backgroundColor: BLUE }}
                >
                  <Send className="h-5 w-5" />
                </button>
              </div>
              <p className="mx-auto mt-2 max-w-2xl text-center text-xs text-slate-400">
                Asistente analítico · No sustituye el criterio ejecutivo · Verifique siempre las cifras.
              </p>
            </div>
          </>
        )}
      </div>
    </div>
  );
}
