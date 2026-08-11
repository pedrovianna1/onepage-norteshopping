import React, { useState, useEffect, useCallback } from "react";
import * as XLSX from "xlsx";
import {
  LineChart, Line, BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip,
  ResponsiveContainer, Legend, ComposedChart, PieChart, Pie, Cell, LabelList
} from "recharts";
import { Upload, Download, TrendingUp, TrendingDown, ChevronDown, ArrowUp, ArrowDown, FileDown, Save, RotateCcw, User } from "lucide-react";

// ---------- Shim: window.storage -> localStorage (fora do ambiente de artifacts) ----------
if (typeof window !== "undefined" && !window.storage) {
  const PREFIX = "onepage_ns::";
  window.storage = {
    async get(key) {
      const raw = window.localStorage.getItem(PREFIX + key);
      if (raw === null) throw new Error("key not found");
      return { key, value: raw, shared: false };
    },
    async set(key, value) {
      window.localStorage.setItem(PREFIX + key, value);
      return { key, value, shared: false };
    },
    async delete(key) {
      window.localStorage.removeItem(PREFIX + key);
      return { key, deleted: true, shared: false };
    },
    async list(prefix = "") {
      const keys = Object.keys(window.localStorage)
        .filter(k => k.startsWith(PREFIX + prefix))
        .map(k => k.slice(PREFIX.length));
      return { keys, prefix, shared: false };
    },
  };
}


// ---------- Design tokens (extraídos do PPT + imagem de referência NorteShopping) ----------
const C = {
  page: "#F4EFE4",
  headerFrom: "#0D454A",
  headerTo: "#0F6470",
  card: "#12474D",
  cardAlt: "#0F5259",
  teal: "#3E8489",
  gold: "#EDA754",
  sand: "#B9CFC9",
  green: "#8FD19E",
  red: "#E8846B",
  ink: "#FFFFFF",
  inkDim: "#AFC9C5",
  textDark: "#1D3A3C",
};

const COLUMNS = [
  "competencia",
  "noi_atual", "noi_meta", "noi_ano_anterior",
  "sss_atual_pct", "sss_meta_pct", "sss_ano_pct", "cpi1_pct", "cpi2_pct", "cpi3_pct",
  "vt_atual", "vt_ano_anterior_pct", "vt_ytd_pct",
  "inad_atual_pct", "inad_meta_pct", "inad_ano_pct",
  "ocup_atual_pct", "ocup_meta_pct", "ocup_ano_pct",
  "aluguel_min_atual", "aluguel_min_meta", "aluguel_min_ano",
  "mall_atual", "mall_meta", "mall_ano",
  "midia_atual", "midia_meta", "midia_ano",
  "fluxo_veiculos_atual", "fluxo_veiculos_meta", "fluxo_veiculos_ano",
  "fluxo_pagantes_atual", "fluxo_pagantes_meta", "fluxo_pagantes_ano",
  "fluxo_pessoas_atual", "fluxo_pessoas_meta", "fluxo_pessoas_ano",
  "tm_atual", "tm_meta", "tm_ano",
  "estac_atual", "estac_meta", "estac_ano",
  "abl_total", "abl_vago", "vagas", "vagas_delta_ano",
  "pipeline_andamento", "pipeline_assinados_mes", "pipeline_assinados_ytd",
  "pipeline_distratos_mes", "pipeline_distratos_ytd",
  "top_ofensores", "maiores_altas", "maiores_baixas", "top_segmentos",
];

const EXAMPLE_ROW = {
  competencia: "2026-06",
  noi_atual: 13600000, noi_meta: 13260000, noi_ano_anterior: 13300000,
  sss_atual_pct: 1.6, sss_meta_pct: 1.3, sss_ano_pct: -0.6, cpi1_pct: 1.9, cpi2_pct: 4.7, cpi3_pct: -4.9,
  vt_atual: 125700000, vt_ano_anterior_pct: 5.7, vt_ytd_pct: 6.2,
  inad_atual_pct: -1.8, inad_meta_pct: -2.1, inad_ano_pct: -7.4,
  ocup_atual_pct: 98.1, ocup_meta_pct: 96.1, ocup_ano_pct: 91.2,
  aluguel_min_atual: 8700000, aluguel_min_meta: 9666000, aluguel_min_ano: 8878000,
  mall_atual: 1500000, mall_meta: 1442000, mall_ano: 1282000,
  midia_atual: 864000, midia_meta: 726000, midia_ano: 557000,
  fluxo_veiculos_atual: 301059, fluxo_veiculos_meta: 295000, fluxo_veiculos_ano: 295500,
  fluxo_pagantes_atual: 117271, fluxo_pagantes_meta: 112000, fluxo_pagantes_ano: 107300,
  fluxo_pessoas_atual: 1818361, fluxo_pessoas_meta: 1780000, fluxo_pessoas_ano: 1664000,
  tm_atual: 25.86, tm_meta: 24.68, tm_ano: 24.86,
  estac_atual: 2800000, estac_meta: 2593000, estac_ano: 2617000,
  abl_total: 71655, abl_vago: 1245.66, vagas: 22, vagas_delta_ano: -5,
  pipeline_andamento: 2, pipeline_assinados_mes: 2, pipeline_assinados_ytd: 20,
  pipeline_distratos_mes: 0, pipeline_distratos_ytd: 25,
  top_ofensores: "BODY TECH:348000|SOUTH:102000|MAMMA JAMMA:88000|BADALADO:61000",
  maiores_altas: "CLARO:729000|MOVIDA:669000|CENTAURO:451000",
  maiores_baixas: "FAST SHOP:-1200000|DAISO JAPAN:-334000|ESTACIO:-320000",
  top_segmentos: "Lazer:10|Mall:7|Eletrodomésticos e Eletrônicos:6",
};

function parsePairs(str) {
  if (!str) return [];
  return String(str).split("|").filter(Boolean).map(pair => {
    const [nome, valor] = pair.split(":");
    return { nome: (nome || "").trim(), valor: Number(valor) };
  });
}

const MESES_PT = ["JANEIRO", "FEVEREIRO", "MARÇO", "ABRIL", "MAIO", "JUNHO", "JULHO", "AGOSTO", "SETEMBRO", "OUTUBRO", "NOVEMBRO", "DEZEMBRO"];
function competenciaLabel(comp) {
  if (!comp) return { mes: "—", ano: "" };
  const [ano, mesNum] = comp.split("-");
  const idx = parseInt(mesNum, 10) - 1;
  return { mes: MESES_PT[idx] || mesNum, ano };
}

// ---------- Formatação segundo o Guia de Comunicação Executiva ----------
function fmtMoney(v) {
  if (v === undefined || v === null || isNaN(v)) return "—";
  const abs = Math.abs(v);
  if (abs >= 1e6) return (v / 1e6).toFixed(2).replace(".", ",") + " MM";
  if (abs >= 1e3) return Math.round(v / 1e3) + "k";
  return String(Math.round(v));
}
function fmtCount(v) {
  if (v === undefined || v === null || isNaN(v)) return "—";
  return Math.round(v).toLocaleString("pt-BR");
}
function fmtPct(v, digits = 1) {
  if (v === undefined || v === null || isNaN(v)) return "—";
  const s = v.toFixed(digits).replace(".", ",");
  return (v > 0 ? "+" : "") + s + "%";
}
function fmtPP(v, digits = 1) {
  if (v === undefined || v === null || isNaN(v)) return "—";
  const s = v.toFixed(digits).replace(".", ",");
  return (v > 0 ? "+" : "") + s + " p.p.";
}
// delta percentual entre dois valores absolutos (NOI, receitas, fluxos)
function deltaPct(atual, base) {
  if (!base) return null;
  return ((atual - base) / Math.abs(base)) * 100;
}
// delta em pontos percentuais entre dois indicadores que já são % (SSS, Inad, Ocupação)
function deltaPP(atual, base) {
  if (atual === undefined || base === undefined) return null;
  return atual - base;
}

function Delta({ value, isPP, label, small }) {
  if (value === null || value === undefined || isNaN(value)) return null;
  const positive = value >= 0;
  const color = positive ? C.green : C.red;
  const Icon = positive ? ArrowUp : ArrowDown;
  return (
    <div style={{ display: "flex", alignItems: "center", gap: 4, fontSize: small ? 10.5 : 11.5, color, marginTop: 2 }}>
      <Icon size={10} strokeWidth={3} />
      <span style={{ fontWeight: 700 }}>{isPP ? fmtPP(value) : fmtPct(value)}</span>
      {label && <span style={{ color: C.inkDim, fontWeight: 400 }}>{label}</span>}
    </div>
  );
}

function SectionHeader({ dot, children }) {
  return (
    <div style={{ display: "flex", alignItems: "center", gap: 8, margin: "26px 0 12px" }}>
      <span style={{ width: 8, height: 8, borderRadius: 999, background: dot || C.gold, flexShrink: 0 }} />
      <span style={{ fontSize: 12, fontWeight: 800, color: C.textDark, letterSpacing: "0.14em" }}>{children}</span>
      <span style={{ flex: 1, height: 1, background: "#D8CFBE" }} />
    </div>
  );
}

function Badge({ children }) {
  return (
    <span style={{
      background: "rgba(237,167,84,0.16)", color: C.gold, border: `1px solid rgba(237,167,84,0.5)`,
      borderRadius: 999, padding: "3px 12px", fontSize: 10, fontWeight: 800, letterSpacing: "0.06em",
    }}>{children}</span>
  );
}

function KpiCard({ title, value, deltaMeta, deltaAno, isPP }) {
  return (
    <div style={{
      background: C.card, borderRadius: 12, padding: "16px 18px",
      display: "flex", flexDirection: "column", gap: 4, minWidth: 0,
      boxShadow: "0 6px 16px rgba(13,69,74,0.18)", borderTop: `3px solid ${C.gold}`,
    }}>
      <div style={{ fontSize: 10.5, letterSpacing: "0.1em", color: C.sand, fontWeight: 700, textTransform: "uppercase" }}>
        {title}
      </div>
      <div style={{ fontSize: 26, fontWeight: 800, color: C.ink, lineHeight: 1.1 }}>
        {value}
      </div>
      <div style={{ display: "flex", flexDirection: "column", gap: 1, marginTop: 2 }}>
        <Delta value={deltaAno} isPP={isPP} label="vs. mesmo mês ano anterior" small />
        <Delta value={deltaMeta} isPP={isPP} label="vs. meta" small />
      </div>
    </div>
  );
}

const Card = ({ children, style }) => (
  <div style={{
    background: C.card, borderRadius: 14, padding: 18,
    boxShadow: "0 6px 18px rgba(13,69,74,0.16)", ...style,
  }}>
    {children}
  </div>
);

const CardTitle = ({ children, badge }) => (
  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14 }}>
    <span style={{ fontSize: 13, fontWeight: 800, color: C.ink, letterSpacing: "0.03em" }}>{children}</span>
    {badge && <Badge>{badge}</Badge>}
  </div>
);

function ReceitaLine({ title, atual, meta, ano }) {
  const dMeta = deltaPct(atual, meta);
  const dAno = deltaPct(atual, ano);
  return (
    <div style={{ marginBottom: 14 }}>
      <div style={{ fontSize: 10, letterSpacing: "0.06em", color: C.inkDim, fontWeight: 700, textTransform: "uppercase" }}>{title}</div>
      <div style={{ fontSize: 22, fontWeight: 800, color: C.ink, marginTop: 2 }}>{fmtMoney(atual)}</div>
      <div style={{ display: "flex", gap: 10 }}>
        <Delta value={dMeta} label="vs. meta" small />
        <Delta value={dAno} label="vs. A-1" small />
      </div>
    </div>
  );
}

function MoverList({ title, items, color }) {
  return (
    <div>
      <div style={{ fontSize: 9.5, letterSpacing: "0.05em", color: C.inkDim, fontWeight: 700, marginBottom: 6 }}>{title}</div>
      {items.map((it, i) => (
        <div key={i} style={{ fontSize: 12, color, fontWeight: 700, marginBottom: 3 }}>
          {it.nome} · {it.valor >= 0 ? "+" : ""}{fmtMoney(it.valor)}
        </div>
      ))}
    </div>
  );
}

function SegmentBar({ nome, valor, max, color }) {
  const w = Math.max(6, Math.min(100, (valor / max) * 100));
  return (
    <div style={{ marginBottom: 12 }}>
      <div style={{ fontSize: 12, fontWeight: 700, color: C.ink, marginBottom: 4 }}>{nome}</div>
      <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
        <div style={{ flex: 1, height: 8, background: "rgba(0,0,0,0.22)", borderRadius: 999 }}>
          <div style={{ width: w + "%", height: "100%", background: color, borderRadius: 999 }} />
        </div>
        <span style={{ fontSize: 11, color: C.ink, fontWeight: 700, minWidth: 30 }}>{fmtPct(valor, 0)}</span>
      </div>
    </div>
  );
}
const SEGMENT_COLORS = ["#5FC9BE", "#EDA754", "#8FD19E", "#E8846B"];

function PipelineBar({ label, value, sub }) {
  const w = Math.min(100, value * 10 + 10);
  return (
    <div style={{ marginBottom: 12 }}>
      <div style={{ fontSize: 11, color: C.inkDim, fontWeight: 700, marginBottom: 4 }}>{label}</div>
      <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
        <div style={{ flex: 1, height: 7, background: "rgba(255,255,255,0.12)", borderRadius: 999 }}>
          <div style={{ width: w + "%", height: "100%", background: C.gold, borderRadius: 999 }} />
        </div>
        <span style={{ fontSize: 11, color: C.ink, fontWeight: 700, minWidth: 60, textAlign: "right" }}>{sub}</span>
      </div>
    </div>
  );
}

function hasVal(v) {
  return v !== null && v !== undefined && v !== "" && !(typeof v === "number" && isNaN(v));
}

function AnnotationBox({ storageKey, placeholder }) {
  const [value, setValue] = useState("");
  const [editing, setEditing] = useState(false);
  const [saving, setSaving] = useState(false);
  const [loaded, setLoaded] = useState(false);

  useEffect(() => {
    let cancelled = false;
    setLoaded(false);
    (async () => {
      let v = "";
      try {
        const r = await window.storage.get(storageKey, false);
        v = r?.value || "";
      } catch (e) { /* no note yet */ }
      if (!cancelled) { setValue(v); setEditing(false); setLoaded(true); }
    })();
    return () => { cancelled = true; };
  }, [storageKey]);

  const save = async (v) => {
    setSaving(true);
    try {
      if (v.trim() === "") await window.storage.delete(storageKey, false);
      else await window.storage.set(storageKey, v, false);
    } catch (e) { /* ignore */ }
    setSaving(false);
    setEditing(false);
  };

  if (!loaded) return null;

  if (!editing && value.trim() === "") {
    return (
      <button onClick={() => setEditing(true)} className="no-print" style={{
        background: "none", border: "none", color: C.inkDim, fontSize: 10.5, fontWeight: 600,
        cursor: "pointer", padding: "6px 0", textDecoration: "underline", textUnderlineOffset: 2,
      }}>
        + adicionar observação
      </button>
    );
  }

  if (!editing) {
    return (
      <div style={{ marginTop: 10 }} onClick={() => setEditing(true)} title="Clique para editar">
        <div style={{ fontSize: 9.5, letterSpacing: "0.05em", color: C.inkDim, fontWeight: 700, marginBottom: 4 }}>OBSERVAÇÕES</div>
        <div style={{
          fontSize: 11.5, color: C.ink, background: "rgba(255,255,255,0.06)", borderRadius: 8,
          padding: "8px 10px", cursor: "text", whiteSpace: "pre-wrap",
        }}>{value}</div>
      </div>
    );
  }

  return (
    <div style={{ marginTop: 10 }}>
      <div style={{ fontSize: 9.5, letterSpacing: "0.05em", color: C.inkDim, fontWeight: 700, marginBottom: 4 }}>
        OBSERVAÇÕES
      </div>
      <textarea
        autoFocus
        value={value}
        placeholder={placeholder || "Adicionar observação para este mês…"}
        onChange={e => setValue(e.target.value)}
        onBlur={e => save(e.target.value)}
        rows={2}
        style={{
          width: "100%", background: "rgba(255,255,255,0.08)", border: "1px solid rgba(255,255,255,0.18)",
          borderRadius: 8, padding: "8px 10px", fontSize: 11.5, color: C.ink, resize: "vertical",
          fontFamily: "inherit", boxSizing: "border-box", outline: "none",
        }}
      />
      {saving && <div style={{ fontSize: 9, color: C.inkDim, marginTop: 2 }}>salvando…</div>}
    </div>
  );
}

function ContactFooter() {
  const [photo, setPhoto] = useState(null);

  useEffect(() => {
    (async () => {
      try {
        const r = await window.storage.get("profile:photo", false);
        if (r?.value) setPhoto(r.value);
      } catch (e) { /* sem foto ainda */ }
    })();
  }, []);

  const onUpload = (file) => {
    const reader = new FileReader();
    reader.onload = async () => {
      const dataUrl = reader.result;
      setPhoto(dataUrl);
      try { await window.storage.set("profile:photo", dataUrl, false); } catch (e) { /* ignore */ }
    };
    reader.readAsDataURL(file);
  };

  return (
    <div style={{ marginTop: 30, borderRadius: 14, overflow: "hidden", boxShadow: "0 6px 18px rgba(13,69,74,0.16)" }}>
      <div style={{
        background: "linear-gradient(115deg, #0A3236 0%, #0D454A 40%, #1B7A72 100%)",
        padding: "16px 20px", display: "flex", alignItems: "center", gap: 14,
      }}>
        <label style={{ cursor: "pointer", flexShrink: 0 }} className="no-print">
          <div style={{
            width: 48, height: 48, borderRadius: "50%", overflow: "hidden", background: "rgba(255,255,255,0.15)",
            display: "flex", alignItems: "center", justifyContent: "center", border: "2px solid rgba(255,255,255,0.3)",
          }}>
            {photo ? <img src={photo} alt="Foto" style={{ width: "100%", height: "100%", objectFit: "cover" }} /> : <User size={22} color="#fff" />}
          </div>
          <input type="file" accept="image/*" style={{ display: "none" }}
            onChange={e => e.target.files[0] && onUpload(e.target.files[0])} />
        </label>
        <div>
          <div style={{ fontSize: 10, letterSpacing: "0.1em", color: C.gold, fontWeight: 700 }}>CONTATO</div>
          <div style={{ fontSize: 15, fontWeight: 800, color: "#fff" }}>João Vianna</div>
          <div style={{ fontSize: 12, color: C.inkDim }}>joao.vianna@allos.com.br</div>
        </div>
      </div>
      <div style={{ background: "#fff", padding: "8px 20px", fontSize: 11, fontWeight: 800, letterSpacing: "0.08em", color: C.headerFrom }}>
        COMERCIAL NORTESHOPPING
      </div>
    </div>
  );
}

export default function OnePageDashboard() {
  const [months, setMonths] = useState({}); // { "2026-06": {...row} }
  const [selected, setSelected] = useState(null);
  const [loading, setLoading] = useState(true);
  const [uploadMsg, setUploadMsg] = useState("");

  const loadAll = useCallback(async () => {
    setLoading(true);
    try {
      const list = await window.storage.list("onepage:", false);
      const keys = list?.keys || [];
      const data = {};
      for (const k of keys) {
        try {
          const r = await window.storage.get(k, false);
          if (r?.value) {
            const parsed = JSON.parse(r.value);
            data[parsed.competencia] = parsed;
          }
        } catch (e) { /* skip */ }
      }
      setMonths(data);
      const comps = Object.keys(data).sort();
      if (comps.length) setSelected(comps[comps.length - 1]);
    } catch (e) {
      setMonths({});
    }
    setLoading(false);
  }, []);

  useEffect(() => { loadAll(); }, [loadAll]);

  const handleFile = async (file) => {
    setUploadMsg("Lendo planilha…");
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array" });
      const sheet = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(sheet, { defval: null });
      if (!rows.length) { setUploadMsg("Planilha vazia."); return; }

      let saved = 0;
      for (const row of rows) {
        if (!row.competencia) continue;
        const comp = String(row.competencia).slice(0, 7);
        const clean = { competencia: comp };
        for (const col of COLUMNS) {
          if (col === "competencia") continue;
          const v = row[col];
          clean[col] = v === null || v === "" || v === undefined ? null : Number(v);
        }
        await window.storage.set("onepage:" + comp, JSON.stringify(clean), false);
        saved++;
      }
      setUploadMsg(`${saved} mês(es) importado(s) com sucesso.`);
      await loadAll();
    } catch (e) {
      setUploadMsg("Erro ao ler o arquivo: " + e.message);
    }
  };

  const downloadTemplate = () => {
    const ws = XLSX.utils.json_to_sheet([EXAMPLE_ROW], { header: COLUMNS });
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "OnePage");
    XLSX.writeFile(wb, "modelo_onepage_norteshopping.xlsx");
  };

  const exportBackup = async () => {
    try {
      const list = await window.storage.list("", false);
      const keys = list?.keys || [];
      const payload = {};
      for (const k of keys) {
        try {
          const r = await window.storage.get(k, false);
          if (r?.value !== undefined) payload[k] = r.value;
        } catch (e) { /* skip */ }
      }
      const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `backup_onepage_${new Date().toISOString().slice(0, 10)}.json`;
      a.click();
      URL.revokeObjectURL(url);
    } catch (e) {
      setUploadMsg("Erro ao gerar backup: " + e.message);
    }
  };

  const restoreBackup = async (file) => {
    try {
      const text = await file.text();
      const payload = JSON.parse(text);
      for (const [k, v] of Object.entries(payload)) {
        await window.storage.set(k, v, false);
      }
      setUploadMsg("Backup restaurado com sucesso.");
      await loadAll();
    } catch (e) {
      setUploadMsg("Erro ao restaurar backup: " + e.message);
    }
  };

  const comps = Object.keys(months).sort();
  const d = selected ? months[selected] : null;

  const noiSeries = comps.map(c => ({
    mes: c.slice(2).replace("-", "/"),
    "Real 2026": months[c].noi_atual,
    "Meta": months[c].noi_meta,
    "Real A-1": months[c].noi_ano_anterior,
  }));

  const inadSeries = comps.map(c => ({
    mes: c.slice(2).replace("-", "/"),
    "Meta 2026": months[c].inad_meta_pct,
    "Real 2026": months[c].inad_atual_pct,
    "Real 2025": months[c].inad_ano_pct,
  }));

  const fluxoCombo = d ? [
    { cat: "A-1", "Fluxo Total": d.fluxo_veiculos_ano, "Fluxo Pagante": d.fluxo_pagantes_ano, TM: d.tm_ano },
    { cat: String(competenciaLabel(selected).ano || "Atual"), "Fluxo Total": d.fluxo_veiculos_atual, "Fluxo Pagante": d.fluxo_pagantes_atual, TM: d.tm_atual },
    { cat: "Meta / Orç", "Fluxo Total": d.fluxo_veiculos_meta, "Fluxo Pagante": d.fluxo_pagantes_meta, TM: d.tm_meta },
  ] : [];

  const ofensores = d ? parsePairs(d.top_ofensores) : [];
  const altas = d ? parsePairs(d.maiores_altas) : [];
  const baixas = d ? parsePairs(d.maiores_baixas) : [];
  const segmentos = d ? parsePairs(d.top_segmentos) : [];
  const maxSeg = segmentos.length ? Math.max(...segmentos.map(s => s.valor)) : 1;

  const vtSeries = comps.map(c => ({
    mes: c.slice(2).replace("-", "/"),
    "Vendas Totais": months[c].vt_atual,
  }));
  const sssSeries = comps.map(c => ({
    mes: c.slice(2).replace("-", "/"),
    "SSS": months[c].sss_atual_pct,
  }));
  const vtHasSeries = vtSeries.filter(p => hasVal(p["Vendas Totais"])).length >= 2;
  const sssHasSeries = sssSeries.filter(p => hasVal(p["SSS"])).length >= 2;
  const hasCpi = hasVal(d?.cpi1_pct) || hasVal(d?.cpi2_pct) || hasVal(d?.cpi3_pct);
  const hasVtDelta = hasVal(d?.vt_ano_anterior_pct) || hasVal(d?.vt_ytd_pct);

  const ocupData = d ? [
    { name: "ocupado", value: d.ocup_atual_pct },
    { name: "vago", value: 100 - d.ocup_atual_pct },
  ] : [];

  return (
    <div style={{
      background: C.page, minHeight: "100%", fontFamily: "'Segoe UI', 'Helvetica Neue', Arial, sans-serif",
      color: C.textDark,
    }}>
      {/* Header */}
      <div style={{
        background: "linear-gradient(115deg, #0A3236 0%, #0D454A 35%, #1B7A72 100%)",
        padding: "26px 28px 22px", position: "relative",
      }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", flexWrap: "wrap", gap: 20 }}>
          <div>
            <div style={{ fontSize: 11, letterSpacing: "0.2em", color: C.gold, fontWeight: 700 }}>RESULTADO MENSAL · ONE PAGE</div>
            <div style={{ position: "relative", display: "inline-block", marginTop: 2 }}>
              <div style={{ fontSize: 32, fontWeight: 800, color: C.ink, letterSpacing: "0.01em" }}>NorteShopping</div>
              <div style={{
                position: "absolute", left: 0, right: -8, bottom: -4, height: 0,
                borderBottom: "2px dashed rgba(143,209,190,0.6)",
              }} />
            </div>
          </div>

          <div style={{ display: "flex", alignItems: "center", gap: 18 }}>
            <div style={{ textAlign: "right", fontFamily: "Calibri, 'Segoe UI', Arial, sans-serif" }}>
              <div style={{ fontSize: 32, fontWeight: 800, color: C.ink, letterSpacing: "0.01em" }}>
                {competenciaLabel(selected).mes}
              </div>
              <div style={{ fontSize: 11, color: C.inkDim, letterSpacing: "0.16em", fontWeight: 600 }}>{competenciaLabel(selected).ano}</div>
            </div>
            <div style={{ width: 1, height: 40, background: "rgba(255,255,255,0.25)" }} />
            <div style={{ display: "flex", flexDirection: "column", alignItems: "center", gap: 2, fontFamily: "Calibri, 'Segoe UI', Arial, sans-serif" }}>
              <div style={{ fontSize: 13, fontWeight: 700, color: C.ink }}>NorteShopping</div>
              <svg width="30" height="26" viewBox="0 0 60 52" fill="none">
                <path d="M6 48 C6 48 17 28 23 15 C26 8.5 28 4 30 4 C32 4 34 8.5 37 15 C43 28 54 48 54 48"
                  stroke="#6CC680" strokeWidth="4" strokeLinecap="round" fill="none" />
              </svg>
            </div>
          </div>
        </div>
      </div>
      <div style={{ height: 6, background: C.gold }} />

      {/* Toolbar funcional */}
      <div className="no-print" style={{
        display: "flex", justifyContent: "flex-end", alignItems: "center", gap: 10,
        padding: "10px 28px", background: "#EAE3D3", flexWrap: "wrap",
      }}>
        {comps.length > 0 && (
          <div style={{ position: "relative" }}>
            <select
              value={selected || ""}
              onChange={e => setSelected(e.target.value)}
              style={{
                background: "#fff", color: C.textDark, border: `1px solid #C9BEA5`,
                borderRadius: 8, padding: "7px 30px 7px 12px", fontSize: 12.5, fontWeight: 600,
                appearance: "none", cursor: "pointer",
              }}
            >
              {comps.map(c => <option key={c} value={c}>{c}</option>)}
            </select>
            <ChevronDown size={13} style={{ position: "absolute", right: 9, top: 9, pointerEvents: "none", color: C.textDark }} />
          </div>
        )}
        <button onClick={downloadTemplate} style={{
          display: "flex", alignItems: "center", gap: 6, background: "#fff",
          border: `1px solid #C9BEA5`, color: C.textDark, borderRadius: 8, padding: "7px 12px",
          fontSize: 12.5, fontWeight: 600, cursor: "pointer",
        }}>
          <Download size={13} /> Modelo
        </button>
        <label style={{
          display: "flex", alignItems: "center", gap: 6, background: C.gold, color: C.headerFrom,
          borderRadius: 8, padding: "7px 12px", fontSize: 12.5, fontWeight: 700, cursor: "pointer",
        }}>
          <Upload size={13} /> Importar planilha
          <input type="file" accept=".xlsx,.xls,.csv" style={{ display: "none" }}
            onChange={e => e.target.files[0] && handleFile(e.target.files[0])} />
        </label>
        <button onClick={() => window.print()} style={{
          display: "flex", alignItems: "center", gap: 6, background: C.headerFrom, color: "#fff",
          border: "none", borderRadius: 8, padding: "7px 12px", fontSize: 12.5, fontWeight: 700, cursor: "pointer",
        }}>
          <FileDown size={13} /> Exportar PDF
        </button>
        <button onClick={exportBackup} style={{
          display: "flex", alignItems: "center", gap: 6, background: "#fff",
          border: `1px solid #C9BEA5`, color: C.textDark, borderRadius: 8, padding: "7px 12px",
          fontSize: 12.5, fontWeight: 600, cursor: "pointer",
        }}>
          <Save size={13} /> Backup
        </button>
        <label style={{
          display: "flex", alignItems: "center", gap: 6, background: "#fff",
          border: `1px solid #C9BEA5`, color: C.textDark, borderRadius: 8, padding: "7px 12px",
          fontSize: 12.5, fontWeight: 600, cursor: "pointer",
        }}>
          <RotateCcw size={13} /> Restaurar
          <input type="file" accept=".json" style={{ display: "none" }}
            onChange={e => e.target.files[0] && restoreBackup(e.target.files[0])} />
        </label>
      </div>

      <div style={{ padding: "20px 28px 40px" }}>
        {uploadMsg && <div style={{ fontSize: 12, color: C.headerFrom, fontWeight: 700, marginBottom: 12 }}>{uploadMsg}</div>}
        {loading && <div style={{ color: C.textDark, fontSize: 13 }}>Carregando…</div>}

        {!loading && comps.length === 0 && (
          <div style={{ border: `1px dashed ${C.teal}`, borderRadius: 16, padding: 40, textAlign: "center", color: C.textDark }}>
            <div style={{ fontSize: 14, marginBottom: 8, fontWeight: 700 }}>Nenhum dado importado ainda.</div>
            <div style={{ fontSize: 12 }}>Baixe o modelo, preencha um mês por linha e importe a planilha.</div>
          </div>
        )}

        {!loading && d && (
          <>
            <SectionHeader dot={C.gold}>DESTAQUES DO MÊS</SectionHeader>
            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(190px, 1fr))", gap: 14 }}>
              <KpiCard title="NOI Total" value={fmtMoney(d.noi_atual)}
                deltaMeta={deltaPct(d.noi_atual, d.noi_meta)} deltaAno={deltaPct(d.noi_atual, d.noi_ano_anterior)} />
              <KpiCard title="Vendas (SSS)" value={fmtPct(d.sss_atual_pct)}
                isPP deltaMeta={deltaPP(d.sss_atual_pct, d.sss_meta_pct)} deltaAno={deltaPP(d.sss_atual_pct, d.sss_ano_pct)} />
              <KpiCard title="Inadimplência" value={fmtPct(d.inad_atual_pct)}
                isPP deltaMeta={deltaPP(d.inad_atual_pct, d.inad_meta_pct)} deltaAno={deltaPP(d.inad_atual_pct, d.inad_ano_pct)} />
              <KpiCard title="Taxa de Ocupação" value={fmtPct(d.ocup_atual_pct)}
                isPP deltaMeta={deltaPP(d.ocup_atual_pct, d.ocup_meta_pct)} deltaAno={deltaPP(d.ocup_atual_pct, d.ocup_ano_pct)} />
            </div>

            <SectionHeader dot={C.gold}>NOI</SectionHeader>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16 }}>
              <Card>
                <CardTitle badge="RECEITA">NOI — Evolução linhas</CardTitle>
                <div style={{ display: "grid", gridTemplateColumns: "1fr 1.3fr", gap: 14 }}>
                  <div>
                    <ReceitaLine title="Receita Aluguel Mínimo Líquida" atual={d.aluguel_min_atual} meta={d.aluguel_min_meta} ano={d.aluguel_min_ano} />
                    <ReceitaLine title="Receita de Mall" atual={d.mall_atual} meta={d.mall_meta} ano={d.mall_ano} />
                    <ReceitaLine title="Receita de Mídia" atual={d.midia_atual} meta={d.midia_meta} ano={d.midia_ano} />
                  </div>
                  <ResponsiveContainer width="100%" height={190}>
                    <LineChart data={noiSeries} margin={{ top: 4, right: 4, left: 0, bottom: 0 }}>
                      <CartesianGrid stroke="rgba(255,255,255,0.08)" vertical={false} />
                      <XAxis dataKey="mes" tick={{ fill: C.sand, fontSize: 10 }} axisLine={{ stroke: C.teal }} tickLine={false} />
                      <YAxis tickFormatter={v => fmtMoney(v)} tick={{ fill: C.sand, fontSize: 9 }} axisLine={false} tickLine={false} width={44} />
                      <Tooltip contentStyle={{ background: C.cardAlt, border: `1px solid ${C.teal}`, borderRadius: 8, fontSize: 11 }} formatter={v => fmtMoney(v)} />
                      <Legend wrapperStyle={{ fontSize: 10, color: C.sand }} />
                      <Line type="monotone" dataKey="Real 2026" stroke={C.gold} strokeWidth={2.5} dot={{ r: 3 }} />
                      <Line type="monotone" dataKey="Meta" stroke={C.sand} strokeWidth={1.5} strokeDasharray="4 3" dot={false} />
                      <Line type="monotone" dataKey="Real A-1" stroke={C.green} strokeWidth={1.5} dot={false} />
                    </LineChart>
                  </ResponsiveContainer>
                </div>
              </Card>

              <Card>
                <CardTitle badge="INAD">Inadimplência AMM + COND</CardTitle>
                <ResponsiveContainer width="100%" height={150}>
                  <LineChart data={inadSeries} margin={{ top: 4, right: 4, left: 0, bottom: 0 }}>
                    <CartesianGrid stroke="rgba(255,255,255,0.08)" vertical={false} />
                    <XAxis dataKey="mes" tick={{ fill: C.sand, fontSize: 10 }} axisLine={{ stroke: C.teal }} tickLine={false} />
                    <YAxis tickFormatter={v => v + "%"} tick={{ fill: C.sand, fontSize: 9 }} axisLine={false} tickLine={false} width={36} />
                    <Tooltip contentStyle={{ background: C.cardAlt, border: `1px solid ${C.teal}`, borderRadius: 8, fontSize: 11 }} formatter={v => fmtPct(v)} />
                    <Legend wrapperStyle={{ fontSize: 10, color: C.sand }} />
                    <Line type="monotone" dataKey="Meta 2026" stroke={C.sand} strokeWidth={1.5} strokeDasharray="4 3" dot={false} />
                    <Line type="monotone" dataKey="Real 2026" stroke={C.gold} strokeWidth={2.5} dot={{ r: 3 }} />
                    <Line type="monotone" dataKey="Real 2025" stroke={C.green} strokeWidth={1.5} dot={false} />
                  </LineChart>
                </ResponsiveContainer>
                <MoverList title="TOP OFENSORES (MÊS EM VALOR ABSOLUTO)" items={ofensores} color={C.red} />
                <AnnotationBox storageKey={`note:${selected}:inad-ofensores`} placeholder="Observação sobre os ofensores…" />
              </Card>
            </div>

            <SectionHeader dot={C.teal}>VT · SSS · LEASING</SectionHeader>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 16 }}>
              <Card>
                <CardTitle badge="VT">Vendas Totais</CardTitle>
                <div style={{ fontSize: 30, fontWeight: 800, color: C.ink }}>{fmtMoney(d.vt_atual)}</div>
                {hasVtDelta ? (
                  <div style={{ display: "flex", gap: 10, marginBottom: 10 }}>
                    <Delta value={d.vt_ano_anterior_pct} label="vs. mesmo mês A-1" small />
                    <Delta value={d.vt_ytd_pct} label="YTD" small />
                  </div>
                ) : (
                  <AnnotationBox storageKey={`note:${selected}:vt-delta`} placeholder="Observação sobre a variação de vendas…" />
                )}
                {vtHasSeries && (
                  <ResponsiveContainer width="100%" height={70}>
                    <LineChart data={vtSeries} margin={{ top: 2, right: 4, left: 4, bottom: 0 }}>
                      <Line type="monotone" dataKey="Vendas Totais" stroke="#5FC9BE" strokeWidth={2} dot={false} />
                      <Tooltip contentStyle={{ background: C.cardAlt, border: `1px solid ${C.teal}`, borderRadius: 8, fontSize: 11 }} formatter={v => fmtMoney(v)} />
                    </LineChart>
                  </ResponsiveContainer>
                )}
                {segmentos.length > 0 ? (
                  <>
                    <div style={{ fontSize: 9.5, letterSpacing: "0.05em", color: C.inkDim, fontWeight: 700, margin: "10px 0 8px" }}>
                      TOP SEGMENTOS — {selected}
                    </div>
                    {segmentos.map((s, i) => <SegmentBar key={i} nome={s.nome} valor={s.valor} max={maxSeg} color={SEGMENT_COLORS[i % SEGMENT_COLORS.length]} />)}
                  </>
                ) : (
                  <AnnotationBox storageKey={`note:${selected}:vt-segmentos`} placeholder="Observação sobre os segmentos…" />
                )}
              </Card>

              <Card>
                <CardTitle badge="SSS">SSS — Same Store Sale</CardTitle>
                <div style={{ fontSize: 30, fontWeight: 800, color: C.ink }}>{fmtPct(d.sss_atual_pct)}</div>
                {hasCpi && (
                  <div style={{ display: "flex", gap: 16, fontSize: 11, color: C.inkDim, margin: "8px 0 14px" }}>
                    <span>CPI 1 · {fmtPct(d.cpi1_pct)}</span>
                    <span>CPI 2 · {fmtPct(d.cpi2_pct)}</span>
                    <span>CPI 3 · {fmtPct(d.cpi3_pct)}</span>
                  </div>
                )}
                {sssHasSeries ? (
                  <ResponsiveContainer width="100%" height={70}>
                    <LineChart data={sssSeries} margin={{ top: 2, right: 4, left: 4, bottom: 0 }}>
                      <Line type="monotone" dataKey="SSS" stroke={C.gold} strokeWidth={2} dot={false} />
                      <Tooltip contentStyle={{ background: C.cardAlt, border: `1px solid ${C.teal}`, borderRadius: 8, fontSize: 11 }} formatter={v => fmtPct(v)} />
                    </LineChart>
                  </ResponsiveContainer>
                ) : (
                  <AnnotationBox storageKey={`note:${selected}:sss-chart`} placeholder="Observação sobre a tendência de SSS…" />
                )}
                {(altas.length > 0 || baixas.length > 0) ? (
                  <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginTop: 10 }}>
                    <MoverList title="↑ MAIORES ALTAS" items={altas} color={C.green} />
                    <MoverList title="↓ MAIORES BAIXAS" items={baixas} color={C.red} />
                  </div>
                ) : (
                  <AnnotationBox storageKey={`note:${selected}:sss-movers`} placeholder="Observação sobre altas e baixas…" />
                )}
              </Card>

              <Card>
                <CardTitle badge="ABL">Leasing & Ocupação</CardTitle>
                <div style={{ display: "flex", alignItems: "center", gap: 14, marginBottom: 14 }}>
                  <div style={{ width: 84, height: 84, position: "relative" }}>
                    <ResponsiveContainer width="100%" height="100%">
                      <PieChart>
                        <Pie data={ocupData} dataKey="value" innerRadius={28} outerRadius={40} startAngle={90} endAngle={-270} stroke="none">
                          <Cell fill={C.gold} />
                          <Cell fill="rgba(255,255,255,0.12)" />
                        </Pie>
                      </PieChart>
                    </ResponsiveContainer>
                    <div style={{ position: "absolute", inset: 0, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 15, fontWeight: 800, color: C.ink }}>
                      {fmtPct(d.ocup_atual_pct, 1)}
                    </div>
                  </div>
                  <div style={{ fontSize: 11.5, color: C.inkDim, lineHeight: 1.7 }}>
                    <div>ABL Total: <b style={{ color: C.ink }}>{fmtCount(d.abl_total)} m²</b></div>
                    <div style={{ color: C.red }}>Vagas: {d.vagas} un. ({d.vagas_delta_ano} vs A-1)</div>
                    <div>ABL Vago: <b style={{ color: C.ink }}>{d.abl_vago} m²</b></div>
                  </div>
                </div>
                <AnnotationBox storageKey={`note:${selected}:leasing-vagas`} placeholder="Observação sobre vagas/ABL…" />
                <div style={{ fontSize: 9.5, letterSpacing: "0.05em", color: C.inkDim, fontWeight: 700, margin: "12px 0 8px" }}>PIPELINE CONTRATOS</div>
                <PipelineBar label="Em andamento" value={d.pipeline_andamento} sub={`${fmtCount(d.pipeline_andamento)} un.`} />
                <PipelineBar label="Assinados/mês" value={d.pipeline_assinados_mes} sub={`${fmtCount(d.pipeline_assinados_mes)} un. / ${fmtCount(d.pipeline_assinados_ytd)} YTD`} />
                <PipelineBar label="Distratos" value={d.pipeline_distratos_mes} sub={`${fmtCount(d.pipeline_distratos_mes)} un. / ${fmtCount(d.pipeline_distratos_ytd)} YTD`} />
                <AnnotationBox storageKey={`note:${selected}:pipeline-ytd`} placeholder="Observação sobre o YTD do pipeline…" />
              </Card>
            </div>

            <SectionHeader dot={C.gold}>FLUXOS</SectionHeader>
            <div style={{ display: "grid", gridTemplateColumns: "1.4fr 1fr", gap: 16 }}>
              <Card>
                <CardTitle>Fluxo de Veículos</CardTitle>
                <ResponsiveContainer width="100%" height={210}>
                  <ComposedChart data={fluxoCombo} margin={{ top: 20, right: 8, left: 0, bottom: 0 }}>
                    <CartesianGrid stroke="rgba(255,255,255,0.08)" vertical={false} />
                    <XAxis dataKey="cat" tick={{ fill: C.sand, fontSize: 11 }} axisLine={{ stroke: C.teal }} tickLine={false} />
                    <YAxis yAxisId="left" tickFormatter={v => fmtCount(v)} tick={{ fill: C.sand, fontSize: 10 }} axisLine={false} tickLine={false} width={50} />
                    <YAxis yAxisId="right" orientation="right" tickFormatter={v => "R$" + v} tick={{ fill: C.sand, fontSize: 10 }} axisLine={false} tickLine={false} width={38} />
                    <Tooltip contentStyle={{ background: C.cardAlt, border: `1px solid ${C.teal}`, borderRadius: 8, fontSize: 12 }} />
                    <Legend wrapperStyle={{ fontSize: 11, color: C.sand }} />
                    <Bar yAxisId="left" dataKey="Fluxo Total" fill={C.teal} radius={[4, 4, 0, 0]}>
                      <LabelList dataKey="Fluxo Total" position="top" formatter={v => fmtCount(v)} fill={C.sand} fontSize={10} />
                    </Bar>
                    <Bar yAxisId="left" dataKey="Fluxo Pagante" fill="#5FC9BE" radius={[4, 4, 0, 0]}>
                      <LabelList dataKey="Fluxo Pagante" position="top" formatter={v => fmtCount(v)} fill={C.sand} fontSize={10} />
                    </Bar>
                    <Line yAxisId="right" type="monotone" dataKey="TM" name="TM (Ticket Médio)" stroke={C.gold} strokeWidth={2.5} dot={{ r: 4 }}>
                      <LabelList dataKey="TM" position="top" formatter={v => "R$" + (v ?? "").toString().replace(".", ",")} fill={C.gold} fontSize={10} />
                    </Line>
                  </ComposedChart>
                </ResponsiveContainer>
              </Card>
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14 }}>
                <Card><ReceitaLine title="Fluxo de Veículos" atual={d.fluxo_veiculos_atual} meta={d.fluxo_veiculos_meta} ano={d.fluxo_veiculos_ano} /></Card>
                <Card><ReceitaLine title="Receita de Estacionamento" atual={d.estac_atual} meta={d.estac_meta} ano={d.estac_ano} /></Card>
                <Card><ReceitaLine title="Fluxo de Pagantes" atual={d.fluxo_pagantes_atual} meta={d.fluxo_pagantes_meta} ano={d.fluxo_pagantes_ano} /></Card>
                <Card><ReceitaLine title="Fluxo de Pessoas" atual={d.fluxo_pessoas_atual} meta={d.fluxo_pessoas_meta} ano={d.fluxo_pessoas_ano} /></Card>
              </div>
            </div>

            <ContactFooter />
          </>
        )}
      </div>
      <style>{`
        @media print {
          .no-print { display: none !important; }
          body { background: #fff !important; }
        }
      `}</style>
    </div>
  );
}
