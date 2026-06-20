import { useState, useEffect, useCallback, useRef } from "react";
import * as XLSX from "xlsx";
import { createClient } from "@supabase/supabase-js";

const SUPABASE_URL = import.meta.env.VITE_SUPABASE_URL;
const SUPABASE_KEY = import.meta.env.VITE_SUPABASE_KEY;
const supabase = createClient(SUPABASE_URL, SUPABASE_KEY);

const COLORS = {
  bg: "#F0F4F1", card: "#FFFFFF",
  green: "#1B4332", greenMid: "#2D6A4F", greenLight: "#52B788",
  accent: "#D4A017", success: "#2D9C6A", warn: "#C77B2E",
  danger: "#B5302A", muted: "#6B8070", border: "#C8DDD1", lightBg: "#EAF2ED",
};

const UTI_OPTIONS = [
  { id: "uti7", label: "UTI HER — 7º Andar", icon: "🏥", cor: COLORS.green },
  { id: "uti8", label: "UTI HER — 8º Andar", icon: "🏥", cor: COLORS.greenMid },
];

const CVC_OPTS  = ["VJID","VJIE","VSCD","VSCE","VFD","VFE"];
const ART_OPTS  = ["ARD","ARE","AFD","AFE","ASCD","ASCE"];
const PICC_OPTS = ["MSD","MSE"];

const INITIAL_DISPOSITIVOS = {
  cvc: [], arterial: [], picc: [],
  hd: false, svd: false, sne: false, sng: false, drenos: false,
  desinvadir: { cvc: {}, arterial: {}, picc: {}, hd: null, svd: null, sne: null, sng: null, drenos: null },
};

const INITIAL_ROUND = {
  // Cuidados gerais
  contato: null, riscoBroncoaspiracao: null, riscoQueda: null,
  // Neurológico
  rass: null, dor: null, delirium: null, contencao: null,
  // Cardio
  dva: null, pam: "",
  // Resp
  suporteResp: null, vmProtetora: null, planoDesmame: [], preocResp: [], ims: "", progredirFuncional: null,
  // Nutrição
  viaAlimentar: null, riscoNutricional: null, terapiaNutricional: null,
  aceitacao: null, metaCalorica: null, fono: null, glicemia: null, evacuacao: null,
  // Renal / Infeccioso
  funcaoRenal: null, metaBH: null, pioraInfec: null, atb: null,
  // Profilaxias
  tev: null, lamg: null, cornea: null, higieneOral: null, decubito: null,
  bundles: null, bundlesPendente: "", mudancaDecubito: null, lesaoPressao: null, avalEspecializada: null,
  // Dispositivos / Planejamento
  dispositivos: INITIAL_DISPOSITIVOS,
  diretivas: [], pendenciaExame: null, descPendencia: "", previsaoAlta: null,
  diagnostico: "", medicoAssistente: "",
};

function normalizarRound(r) {
  if (!r) return null;
  const out = { ...INITIAL_ROUND, ...r };
  if (!out.dispositivos) out.dispositivos = { ...INITIAL_DISPOSITIVOS };
  else {
    out.dispositivos = { ...INITIAL_DISPOSITIVOS, ...out.dispositivos };
    if (!Array.isArray(out.dispositivos.cvc))     out.dispositivos.cvc = [];
    if (!Array.isArray(out.dispositivos.arterial)) out.dispositivos.arterial = [];
    if (!Array.isArray(out.dispositivos.picc))    out.dispositivos.picc = [];
    if (!out.dispositivos.desinvadir) out.dispositivos.desinvadir = { ...INITIAL_DISPOSITIVOS.desinvadir };
    else {
      out.dispositivos.desinvadir = { ...INITIAL_DISPOSITIVOS.desinvadir, ...out.dispositivos.desinvadir };
      if (!out.dispositivos.desinvadir.cvc     || typeof out.dispositivos.desinvadir.cvc !== "object")     out.dispositivos.desinvadir.cvc = {};
      if (!out.dispositivos.desinvadir.arterial || typeof out.dispositivos.desinvadir.arterial !== "object") out.dispositivos.desinvadir.arterial = {};
      if (!out.dispositivos.desinvadir.picc    || typeof out.dispositivos.desinvadir.picc !== "object")    out.dispositivos.desinvadir.picc = {};
    }
  }
  if (!Array.isArray(out.planoDesmame)) out.planoDesmame = [];
  if (!Array.isArray(out.preocResp))    out.preocResp = [];
  if (!Array.isArray(out.diretivas))    out.diretivas = [];
  return out;
}

function useIsMobile() {
  const [isMobile, setIsMobile] = useState(typeof window !== "undefined" ? window.innerWidth < 700 : false);
  useEffect(() => {
    function onResize() { setIsMobile(window.innerWidth < 700); }
    window.addEventListener("resize", onResize);
    return () => window.removeEventListener("resize", onResize);
  }, []);
  return isMobile;
}

function hojeStr() { return new Date().toISOString().slice(0, 10); }

function calcDias(d) {
  if (!d) return 0;
  const adm = new Date(d);
  if (isNaN(adm.getTime())) return 0;
  const diff = Math.floor((new Date() - adm) / 86400000);
  return diff >= 0 ? diff : null;
}

const gravBadge = (g) => ({
  alta:  [COLORS.danger,  "⚠ Alta"],
  media: [COLORS.warn,    "◈ Média"],
  baixa: [COLORS.success, "✓ Baixa"],
  livre: [COLORS.muted,   "Livre"],
}[g] || [COLORS.muted, g]);

function computeAlerts(r, p) {
  const a = [];
  if (!r || !p.iniciais) return a;
  const roundIniciado = Object.entries(r).some(([k, v]) =>
    k !== "diagnostico" && k !== "medicoAssistente" && k !== "dispositivos" &&
    v !== null && v !== "" && !(Array.isArray(v) && v.length === 0)
  );
  if (r.tev === "Não") a.push("Sem profilaxia TEV");
  else if (r.tev === null && roundIniciado) a.push("TEV não avaliado");
  if (r.riscoNutricional === "Sim" && r.terapiaNutricional !== "Sim") a.push("Risco nutricional sem terapia");
  if (r.suporteResp === "VM invasiva" && !(r.planoDesmame?.length)) a.push("VM sem plano de desmame");
  if (r.dva === "Sim" && !r.pam) a.push("DVA sem meta PAM");
  if (r.pendenciaExame === "Sim" && r.descPendencia) a.push(r.descPendencia);
  if (r.lesaoPressao === "Sim") a.push("LPP presente");
  if (r.bundles === "Não" && r.bundlesPendente) a.push(`Bundle pendente: ${r.bundlesPendente}`);
  const dias = p.dataAdm ? calcDias(p.dataAdm) : null;
  if (dias !== null && dias > 7) a.push(`${dias} dias internado`);
  return a;
}

function listarDispositivos(dev) {
  if (!dev) return "Nenhum";
  const partes = [];
  if (dev.cvc?.length)      partes.push(`CVC (${dev.cvc.join(", ")})`);
  if (dev.arterial?.length) partes.push(`Art (${dev.arterial.join(", ")})`);
  if (dev.picc?.length)     partes.push(`PICC (${dev.picc.join(", ")})`);
  if (dev.hd)     partes.push("HD");
  if (dev.svd)    partes.push("SVD");
  if (dev.sne)    partes.push("SNE");
  if (dev.sng)    partes.push("SNG");
  if (dev.drenos) partes.push("Drenos");
  return partes.length ? partes.join(" | ") : "Nenhum";
}

function simNaoEmoji(v) { return v === "Sim" ? "✅" : v === "Não" ? "❌" : "—"; }

function gerarRelatorioGeral(patients, rounds, utiLabel) {
  const ocupados    = patients.filter(p => p.iniciais);
  const vagos       = patients.filter(p => !p.iniciais);
  const altaHoje    = ocupados.filter(p => rounds[p.id]?.previsaoAlta === "Hoje");
  const graves      = ocupados.filter(p => p.gravidade === "alta");
  const intub       = ocupados.filter(p => rounds[p.id]?.suporteResp === "VM invasiva");
  const dva         = ocupados.filter(p => rounds[p.id]?.dva === "Sim");
  const hd          = ocupados.filter(p => rounds[p.id]?.funcaoRenal === "Em HD");
  const totalLeitos = patients.length;

  const alertasPorLeito = ocupados
    .map(p => ({ leito: p.leito, items: computeAlerts(rounds[p.id], p) }))
    .filter(g => g.items.length > 0);

  const alertasTexto = alertasPorLeito.length > 0
    ? alertasPorLeito.map(g => `🛏 *${g.leito}*\n${g.items.map(a => `  ⚠️ ${a}`).join("\n")}`).join("\n\n")
    : "✅ Sem alertas no momento";

  const dt = new Date().toLocaleDateString("pt-BR") + " — " + new Date().toLocaleTimeString("pt-BR", { hour: "2-digit", minute: "2-digit" });

  return `🏥 *ROUND ${utiLabel.toUpperCase()}*
📅 *${dt}*

━━━━━━━━━━━━━━━━━
📊 *RESUMO GERAL*
━━━━━━━━━━━━━━━━━

🛏 *Leitos ocupados:* ${ocupados.length}/${totalLeitos}
🔓 *Leitos vagos:* ${vagos.length > 0 ? vagos.map(p => p.leito).join(", ") : "Nenhum"}

🏠 *Previsão de alta hoje:* ${altaHoje.length > 0 ? altaHoje.map(p => p.leito).join(", ") : "Nenhum"}
⚠️ *Graves:* ${graves.length > 0 ? graves.map(p => p.leito).join(", ") : "Nenhum"}
🫁 *Intubados (VMI):* ${intub.length > 0 ? intub.map(p => p.leito).join(", ") : "Nenhum"}
💊 *Em DVA:* ${dva.length > 0 ? dva.map(p => p.leito).join(", ") : "Nenhum"}
🩸 *Em HD:* ${hd.length > 0 ? hd.map(p => p.leito).join(", ") : "Nenhum"}

━━━━━━━━━━━━━━━━━
🚨 *ALERTAS E PENDÊNCIAS*
━━━━━━━━━━━━━━━━━

${alertasTexto}

━━━━━━━━━━━━━━━━━
📊 _Gerado pelo UTI Round — HER_`;
}

function gerarRelatorioEspecifico(patients, rounds, utiLabel) {
  const ocupados = patients.filter(p => p.iniciais);
  const dt = new Date().toLocaleDateString("pt-BR") + " — " + new Date().toLocaleTimeString("pt-BR", { hour: "2-digit", minute: "2-digit" });
  if (ocupados.length === 0) return `🏥 *ROUND ${utiLabel.toUpperCase()}*\n📅 *${dt}*\n\nNenhum paciente admitido.`;

  const blocos = ocupados.map(p => {
    const r = rounds[p.id];
    const dias = calcDias(p.dataAdm);
    return `🛏 *${p.leito} — ${p.iniciais}*
👤 ${p.idade ? p.idade + "a" : "?a"} | 🕐 ${dias !== null ? dias : "?"}d internado
👨‍⚕️ *Médico:* ${r?.medicoAssistente || p.medicoAssistente || "—"}
📋 *Diagnóstico:* ${r?.diagnostico || p.diagnostico || "—"}
⚠️ *Gravidade:* ${p.gravidade}

🧠 *Neuro:* RASS ${r?.rass || "—"} | Dor ${simNaoEmoji(r?.dor)} | Delirium ${simNaoEmoji(r?.delirium)}
❤️ *Cardio:* DVA ${simNaoEmoji(r?.dva)} | PAM: ${r?.pam || "—"} mmHg
🫁 *Resp:* ${r?.suporteResp || "—"}${r?.planoDesmame?.length ? " | " + r.planoDesmame.join(", ") : ""}
🍽️ *Nutrição:* ${r?.viaAlimentar || "—"} | Risco nutricional ${simNaoEmoji(r?.riscoNutricional)}${r?.riscoNutricional === "Sim" ? " | Terapia " + simNaoEmoji(r?.terapiaNutricional) : ""}
🫘 *Renal:* ${r?.funcaoRenal || "—"} | BH ${r?.metaBH || "—"}
🦠 *Infeccioso:* ATB ${simNaoEmoji(r?.atb)} | Piora ${simNaoEmoji(r?.pioraInfec)}
💉 *Profilaxias:* TEV ${r?.tev || "—"} | LAMG ${r?.lamg || "—"} | HO ${r?.higieneOral || "—"}
🔌 *Dispositivos:* ${listarDispositivos(r?.dispositivos)}
${r?.pendenciaExame === "Sim" && r?.descPendencia ? `📌 *Pendência:* ${r.descPendencia}\n` : ""}🏠 *Alta:* ${r?.previsaoAlta || "—"}`;
  });

  return `🏥 *ROUND ${utiLabel.toUpperCase()}*
📅 *${dt}*

━━━━━━━━━━━━━━━━━
${blocos.join("\n\n━━━━━━━━━━━━━━━━━\n")}

━━━━━━━━━━━━━━━━━
📊 _Gerado pelo UTI Round — HER_`;
}

function exportExcel(patients, rounds, utiLabel) {
  const rows = patients.filter(p => p.iniciais);
  if (!rows.length) { alert("Nenhum paciente para exportar."); return; }
  const data = rows.map(p => {
    const r = rounds[p.id] || {};
    const dev = r.dispositivos || INITIAL_DISPOSITIVOS;
    return {
      "Leito": p.leito, "Iniciais": p.iniciais, "Idade": p.idade || "",
      "Data Admissão": p.dataAdm || "", "Dias Internado": calcDias(p.dataAdm) ?? "",
      "Médico Assistente": r.medicoAssistente || p.medicoAssistente || "",
      "Gravidade": p.gravidade, "Diagnóstico": r.diagnostico || p.diagnostico || "",
      "Precaução Contato": r.contato || "",
      "Risco Broncoaspiração": r.riscoBroncoaspiracao || "",
      "Risco de Queda": r.riscoQueda || "",
      "RASS": r.rass || "", "Dor": r.dor || "", "Delirium": r.delirium || "", "Contenção": r.contencao || "",
      "DVA": r.dva || "", "Meta PAM": r.pam || "",
      "Suporte Resp": r.suporteResp || "", "VM Protetora": r.vmProtetora || "",
      "Plano Desmame": (r.planoDesmame || []).join(" | "),
      "Preocupações Resp": (r.preocResp || []).join(" | "),
      "IMS": r.ims || "", "Progredir Funcional": r.progredirFuncional || "",
      "Via Alimentar": r.viaAlimentar || "",
      "Risco Nutricional": r.riscoNutricional || "", "Terapia Nutricional": r.terapiaNutricional || "",
      "Aceitação": r.aceitacao || "", "Meta Calórica": r.metaCalorica || "",
      "Fono": r.fono || "", "Glicemia": r.glicemia || "", "Evacuação": r.evacuacao || "",
      "Função Renal": r.funcaoRenal || "", "BH Meta": r.metaBH || "",
      "Piora Infecciosa": r.pioraInfec || "", "ATB": r.atb || "",
      "TEV": r.tev || "", "LAMG": r.lamg || "", "Córnea": r.cornea || "",
      "Higiene Oral": r.higieneOral || "", "Decúbito": r.decubito || "",
      "Bundles OK": r.bundles || "", "Bundle Pendente": r.bundlesPendente || "",
      "Mudança Decúbito": r.mudancaDecubito || "", "LPP": r.lesaoPressao || "",
      "Aval Especializada": r.avalEspecializada || "",
      "CVC": (dev.cvc || []).join(" | "),
      "CVC Desinvadir": Object.entries(dev.desinvadir?.cvc || {}).filter(([_, v]) => v === "Sim").map(([k]) => k).join(", "),
      "Cateter Arterial": (dev.arterial || []).join(" | "),
      "Art Desinvadir": Object.entries(dev.desinvadir?.arterial || {}).filter(([_, v]) => v === "Sim").map(([k]) => k).join(", "),
      "PICC": (dev.picc || []).join(" | "),
      "PICC Desinvadir": Object.entries(dev.desinvadir?.picc || {}).filter(([_, v]) => v === "Sim").map(([k]) => k).join(", "),
      "HD": dev.hd ? "Sim" : "Não", "HD Desinvadir": dev.desinvadir?.hd ?? "",
      "SVD": dev.svd ? "Sim" : "Não", "SVD Desinvadir": dev.desinvadir?.svd ?? "",
      "SNE": dev.sne ? "Sim" : "Não", "SNE Desinvadir": dev.desinvadir?.sne ?? "",
      "SNG": dev.sng ? "Sim" : "Não", "SNG Desinvadir": dev.desinvadir?.sng ?? "",
      "Drenos": dev.drenos ? "Sim" : "Não", "Drenos Desinvadir": dev.desinvadir?.drenos ?? "",
      "Diretivas": (r.diretivas || []).join(" | "),
      "Pendência": r.pendenciaExame || "", "Desc Pendência": r.descPendencia || "",
      "Previsão Alta": r.previsaoAlta || "",
      "Alertas": computeAlerts(r, p).join(" | "),
    };
  });
  const ws = XLSX.utils.json_to_sheet(data);
  ws["!cols"] = Object.keys(data[0]).map(() => ({ wch: 20 }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Round UTI");
  XLSX.writeFile(wb, `round-her-${hojeStr()}.xlsx`);
}

// ── Atoms ─────────────────────────────────────────────────────────────────────
function Pill({ label, selected, onClick, color }) {
  return (
    <button type="button" onClick={onClick} style={{
      padding: "7px 14px", borderRadius: 20, cursor: "pointer", whiteSpace: "nowrap",
      border: `1.5px solid ${selected ? (color || COLORS.greenMid) : COLORS.border}`,
      background: selected ? (color || COLORS.greenMid) : "#fff",
      color: selected ? "#fff" : COLORS.green,
      fontSize: 13, fontWeight: selected ? 600 : 400, transition: "all .15s",
      minHeight: 34, touchAction: "manipulation",
    }}>{label}</button>
  );
}
function MultiPill({ label, checked, onChange, color }) {
  return (
    <button type="button" onClick={() => onChange(!checked)} style={{
      padding: "7px 14px", borderRadius: 20, cursor: "pointer",
      border: `1.5px solid ${checked ? (color || COLORS.greenLight) : COLORS.border}`,
      background: checked ? (color || COLORS.greenLight) : "#fff",
      color: checked ? "#fff" : COLORS.green,
      fontSize: 13, fontWeight: checked ? 600 : 400, transition: "all .15s",
      minHeight: 34, touchAction: "manipulation",
    }}>{label}</button>
  );
}
function SecHdr({ title, icon }) {
  return (
    <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 14, marginTop: 4 }}>
      <span style={{ fontSize: 17 }}>{icon}</span>
      <span style={{ fontWeight: 700, fontSize: 12, color: COLORS.green, letterSpacing: ".6px", textTransform: "uppercase" }}>{title}</span>
      <div style={{ flex: 1, height: 1, background: COLORS.border }} />
    </div>
  );
}
function Field({ label, children }) {
  return (
    <div style={{ marginBottom: 14 }}>
      <div style={{ fontSize: 11, color: COLORS.muted, fontWeight: 700, marginBottom: 6, textTransform: "uppercase", letterSpacing: ".4px" }}>{label}</div>
      <div style={{ display: "flex", flexWrap: "wrap", gap: 6 }}>{children}</div>
    </div>
  );
}
function TInput({ value, onChange, placeholder, w, type = "text" }) {
  return (
    <input type={type} value={value || ""} onChange={e => onChange(e.target.value)} placeholder={placeholder}
      style={{
        border: `1.5px solid ${COLORS.border}`, borderRadius: 8, padding: "8px 12px",
        fontSize: 14, color: COLORS.green, outline: "none", background: "#fff",
        width: w || "100%", maxWidth: w || "100%", boxSizing: "border-box", minHeight: 36,
      }} />
  );
}
function TArea({ value, onChange, placeholder, rows = 2 }) {
  return (
    <textarea value={value || ""} onChange={e => onChange(e.target.value)} placeholder={placeholder} rows={rows}
      style={{
        border: `1.5px solid ${COLORS.border}`, borderRadius: 8, padding: "8px 12px",
        fontSize: 14, color: COLORS.green, outline: "none", width: "100%",
        resize: "vertical", background: "#fff", fontFamily: "inherit", boxSizing: "border-box",
      }} />
  );
}
function Grid({ cols = 2, children, isMobile, gap = 14 }) {
  return (
    <div style={{ display: "grid", gridTemplateColumns: isMobile ? "1fr" : `repeat(${cols}, 1fr)`, gap }}>
      {children}
    </div>
  );
}
function Modal({ children, onClose, maxWidth = 460 }) {
  return (
    <div style={{
      position: "fixed", inset: 0, background: "rgba(27,67,50,.65)", zIndex: 200,
      display: "flex", alignItems: "center", justifyContent: "center", padding: 12,
    }} onClick={onClose}>
      <div onClick={e => e.stopPropagation()} style={{
        background: "#fff", borderRadius: 16, padding: "22px",
        width: "100%", maxWidth, boxShadow: "0 20px 60px rgba(0,0,0,.25)",
        maxHeight: "92vh", overflowY: "auto",
      }}>{children}</div>
    </div>
  );
}

// ── Dispositivos ──────────────────────────────────────────────────────────────
function DispositivosSection({ dev, onChange }) {
  const d = dev || INITIAL_DISPOSITIVOS;

  const toggleArr = (key, opt) => {
    const cur = d[key] || [];
    const novo = cur.includes(opt) ? cur.filter(x => x !== opt) : [...cur, opt];
    const desKey = { ...(d.desinvadir?.[key] || {}) };
    if (!novo.includes(opt)) delete desKey[opt];
    onChange({ ...d, [key]: novo, desinvadir: { ...d.desinvadir, [key]: desKey } });
  };
  const toggleSimples = (key) => {
    const novoValor = !d[key];
    onChange({ ...d, [key]: novoValor, desinvadir: { ...d.desinvadir, [key]: null } });
  };
  const setDesinvArr  = (key, opt, val) =>
    onChange({ ...d, desinvadir: { ...d.desinvadir, [key]: { ...(d.desinvadir?.[key] || {}), [opt]: val } } });
  const setDesinvSimp = (key, val) =>
    onChange({ ...d, desinvadir: { ...d.desinvadir, [key]: val } });

  const boxStyle = (active) => ({
    background: active ? COLORS.greenMid + "12" : COLORS.lightBg,
    borderRadius: 10, padding: "12px 14px",
    border: `1.5px solid ${active ? COLORS.greenMid : COLORS.border}`,
    marginBottom: 8,
  });

  const renderMultiplo = (key, opts, label, icon, color = COLORS.greenMid) => {
    const lista = d[key] || [];
    const active = lista.length > 0;
    return (
      <div style={boxStyle(active)}>
        <div style={{ fontSize: 13, fontWeight: 700, color: COLORS.green, marginBottom: 8 }}>{icon} {label}</div>
        <div style={{ display: "flex", flexWrap: "wrap", gap: 6, marginBottom: lista.length ? 8 : 0 }}>
          {opts.map(opt => (
            <button type="button" key={opt} onClick={() => toggleArr(key, opt)} style={{
              padding: "6px 12px", borderRadius: 16, cursor: "pointer", fontSize: 12, fontWeight: 600,
              border: `1.5px solid ${lista.includes(opt) ? color : COLORS.border}`,
              background: lista.includes(opt) ? color : "#fff",
              color: lista.includes(opt) ? "#fff" : COLORS.green, minHeight: 30,
            }}>{opt}</button>
          ))}
        </div>
        {lista.map(opt => {
          const desVal = d.desinvadir?.[key]?.[opt];
          return (
            <div key={opt} style={{ display: "flex", alignItems: "center", gap: 8, marginTop: 6, flexWrap: "wrap" }}>
              <span style={{ fontSize: 12, color, fontWeight: 700, minWidth: 44 }}>{opt}</span>
              <span style={{ fontSize: 12, color: COLORS.muted }}>Desinvadir?</span>
              {["Sim","Não"].map(v => (
                <button type="button" key={v} onClick={() => setDesinvArr(key, opt, v)} style={{
                  padding: "4px 12px", borderRadius: 12, cursor: "pointer", fontSize: 12, fontWeight: 600,
                  border: `1.5px solid ${desVal === v ? (v === "Sim" ? COLORS.success : COLORS.danger) : COLORS.border}`,
                  background: desVal === v ? (v === "Sim" ? COLORS.success : COLORS.danger) : "#fff",
                  color: desVal === v ? "#fff" : COLORS.green, minHeight: 28,
                }}>{v}</button>
              ))}
            </div>
          );
        })}
      </div>
    );
  };

  const renderSimples = (key, label, icon) => {
    const active = d[key];
    const desVal = d.desinvadir?.[key];
    return (
      <div style={boxStyle(active)}>
        <div style={{ display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap" }}>
          <button type="button" onClick={() => toggleSimples(key)} style={{
            padding: "6px 14px", borderRadius: 16, cursor: "pointer", fontSize: 12, fontWeight: 700,
            border: `1.5px solid ${active ? COLORS.greenMid : COLORS.border}`,
            background: active ? COLORS.greenMid : "#fff",
            color: active ? "#fff" : COLORS.green, minHeight: 30,
          }}>{active ? "✓ Presente" : "Ausente"}</button>
          <span style={{ fontSize: 13, fontWeight: 700, color: COLORS.green }}>{icon} {label}</span>
        </div>
        {active && (
          <div style={{ display: "flex", alignItems: "center", gap: 8, marginTop: 8, flexWrap: "wrap" }}>
            <span style={{ fontSize: 12, color: COLORS.muted }}>Desinvadir?</span>
            {["Sim","Não"].map(v => (
              <button type="button" key={v} onClick={() => setDesinvSimp(key, v)} style={{
                padding: "4px 12px", borderRadius: 12, cursor: "pointer", fontSize: 12, fontWeight: 600,
                border: `1.5px solid ${desVal === v ? (v === "Sim" ? COLORS.success : COLORS.danger) : COLORS.border}`,
                background: desVal === v ? (v === "Sim" ? COLORS.success : COLORS.danger) : "#fff",
                color: desVal === v ? "#fff" : COLORS.green, minHeight: 28,
              }}>{v}</button>
            ))}
          </div>
        )}
      </div>
    );
  };

  return (
    <div>
      {renderMultiplo("cvc",      CVC_OPTS,  "Cateter Venoso Central", "🔵", COLORS.greenMid)}
      {renderMultiplo("arterial", ART_OPTS,  "Cateter Arterial",       "🔴", COLORS.danger)}
      {renderMultiplo("picc",     PICC_OPTS, "PICC",                   "🟣", "#7B2D8B")}
      {renderSimples("hd",     "Cateter de HD", "⚫")}
      {renderSimples("svd",    "SVD",           "🟡")}
      {renderSimples("sne",    "SNE",           "🟠")}
      {renderSimples("sng",    "SNG",           "⚪")}
      {renderSimples("drenos", "Drenos",        "🟤")}
    </div>
  );
}

// ── Modal Trocar Leito ────────────────────────────────────────────────────────
function TrocarLeitoModal({ pat, patients, onTrocar, onClose }) {
  const [destino, setDestino] = useState(null);
  const disponiveis = patients.filter(p => p.id !== pat.id && !p.iniciais);
  const ocupados    = patients.filter(p => p.id !== pat.id && p.iniciais);

  return (
    <Modal onClose={onClose} maxWidth={420}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 20 }}>
        <div>
          <div style={{ fontSize: 11, fontWeight: 700, color: COLORS.muted, textTransform: "uppercase" }}>Trocar de leito</div>
          <div style={{ fontSize: 17, fontWeight: 800, color: COLORS.green }}>{pat.iniciais} — {pat.leito}</div>
        </div>
        <button type="button" onClick={onClose} style={{ background: "none", border: "none", fontSize: 26, cursor: "pointer", color: COLORS.muted, lineHeight: 1, padding: 0 }}>×</button>
      </div>

      {disponiveis.length === 0 && ocupados.length === 0 ? (
        <div style={{ textAlign: "center", color: COLORS.muted, padding: 20 }}>Nenhum outro leito disponível.</div>
      ) : (
        <>
          {disponiveis.length > 0 && (
            <>
              <div style={{ fontSize: 11, fontWeight: 700, color: COLORS.muted, textTransform: "uppercase", marginBottom: 8 }}>Leitos vagos</div>
              <div style={{ display: "flex", flexWrap: "wrap", gap: 8, marginBottom: 16 }}>
                {disponiveis.map(p => (
                  <button type="button" key={p.id} onClick={() => setDestino(p)} style={{
                    padding: "8px 16px", borderRadius: 10, cursor: "pointer", fontSize: 13, fontWeight: 600,
                    border: `1.5px solid ${destino?.id === p.id ? COLORS.greenMid : COLORS.border}`,
                    background: destino?.id === p.id ? COLORS.greenMid : "#fff",
                    color: destino?.id === p.id ? "#fff" : COLORS.green, minHeight: 36,
                  }}>{p.leito}</button>
                ))}
              </div>
            </>
          )}
          {ocupados.length > 0 && (
            <>
              <div style={{ fontSize: 11, fontWeight: 700, color: COLORS.muted, textTransform: "uppercase", marginBottom: 8 }}>Trocar com outro paciente</div>
              <div style={{ display: "flex", flexWrap: "wrap", gap: 8, marginBottom: 16 }}>
                {ocupados.map(p => (
                  <button type="button" key={p.id} onClick={() => setDestino(p)} style={{
                    padding: "8px 16px", borderRadius: 10, cursor: "pointer", fontSize: 13, fontWeight: 600,
                    border: `1.5px solid ${destino?.id === p.id ? COLORS.warn : COLORS.border}`,
                    background: destino?.id === p.id ? COLORS.warn : "#fff",
                    color: destino?.id === p.id ? "#fff" : COLORS.green, minHeight: 36,
                  }}>{p.leito} — {p.iniciais}</button>
                ))}
              </div>
            </>
          )}

          {destino && (
            <div style={{ background: COLORS.lightBg, borderRadius: 10, padding: 12, marginBottom: 14 }}>
              <div style={{ fontSize: 13, color: COLORS.green, fontWeight: 600 }}>
                {destino.iniciais
                  ? `Trocar ${pat.iniciais} (${pat.leito}) ↔ ${destino.iniciais} (${destino.leito})`
                  : `Mover ${pat.iniciais} de ${pat.leito} → ${destino.leito}`}
              </div>
            </div>
          )}

          <div style={{ display: "flex", gap: 10 }}>
            <button type="button" onClick={onClose} style={{
              flex: 1, padding: "11px", borderRadius: 10, border: `1.5px solid ${COLORS.border}`,
              background: "#fff", color: COLORS.green, fontSize: 14, fontWeight: 600, cursor: "pointer", minHeight: 42,
            }}>Cancelar</button>
            <button type="button" disabled={!destino} onClick={() => { if (destino) { onTrocar(pat, destino); onClose(); }}} style={{
              flex: 2, padding: "11px", borderRadius: 10, border: "none",
              background: destino ? COLORS.greenMid : COLORS.border,
              color: "#fff", fontSize: 14, fontWeight: 700,
              cursor: destino ? "pointer" : "not-allowed", minHeight: 42,
            }}>↔ Confirmar troca</button>
          </div>
        </>
      )}
    </Modal>
  );
}

// ── Modais ────────────────────────────────────────────────────────────────────
function EditPatientModal({ pat, onSave, onClear, onClose, onTrocarLeito }) {
  const [iniciais, setIniciais] = useState(pat.iniciais || "");
  const [idade,    setIdade]    = useState(pat.idade || "");
  const [dataAdm,  setAdm]      = useState(pat.dataAdm || "");
  const [diag,     setDiag]     = useState(pat.diagnostico || "");
  const [medico,   setMedico]   = useState(pat.medicoAssistente || "");
  const [grav,     setGrav]     = useState(pat.gravidade || "livre");
  const [confirm,  setConfirm]  = useState(false);
  const canSave = iniciais.trim().length > 0;
  const diasAdm = dataAdm ? calcDias(dataAdm) : null;
  const admFutura = diasAdm === null && dataAdm;

  return (
    <Modal onClose={onClose}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 20 }}>
        <div>
          <div style={{ fontSize: 11, fontWeight: 700, color: COLORS.muted, textTransform: "uppercase" }}>{pat.leito}</div>
          <div style={{ fontSize: 17, fontWeight: 800, color: COLORS.green }}>{pat.iniciais ? "Editar Paciente" : "Admitir Paciente"}</div>
        </div>
        <button type="button" onClick={onClose} style={{ background: "none", border: "none", fontSize: 26, cursor: "pointer", color: COLORS.muted, lineHeight: 1, padding: 0 }}>×</button>
      </div>

      <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
        <div>
          <div style={{ fontSize: 11, color: COLORS.muted, fontWeight: 700, marginBottom: 5, textTransform: "uppercase" }}>Iniciais *</div>
          <input value={iniciais} onChange={e => setIniciais(e.target.value.toUpperCase())} placeholder="Ex: M.S.O."
            style={{ border: `1.5px solid ${COLORS.border}`, borderRadius: 8, padding: "10px 12px", fontSize: 15, color: COLORS.green, outline: "none", width: "100%", background: "#fff", boxSizing: "border-box" }} />
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
          <div>
            <div style={{ fontSize: 11, color: COLORS.muted, fontWeight: 700, marginBottom: 5, textTransform: "uppercase" }}>Idade (anos)</div>
            <input type="number" value={idade} onChange={e => setIdade(e.target.value)} placeholder="Ex: 65" min="0" max="130"
              style={{ border: `1.5px solid ${COLORS.border}`, borderRadius: 8, padding: "10px 12px", fontSize: 14, color: COLORS.green, outline: "none", width: "100%", background: "#fff", boxSizing: "border-box" }} />
          </div>
          <div>
            <div style={{ fontSize: 11, color: COLORS.muted, fontWeight: 700, marginBottom: 5, textTransform: "uppercase" }}>Admissão UTI</div>
            <input type="date" value={dataAdm} onChange={e => setAdm(e.target.value)}
              style={{ border: `1.5px solid ${admFutura ? COLORS.danger : COLORS.border}`, borderRadius: 8, padding: "10px 12px", fontSize: 14, color: COLORS.green, outline: "none", width: "100%", background: "#fff", boxSizing: "border-box" }} />
            {admFutura
              ? <div style={{ fontSize: 12, color: COLORS.danger, marginTop: 4, fontWeight: 600 }}>⚠ Data futura</div>
              : dataAdm && <div style={{ fontSize: 12, color: COLORS.greenMid, marginTop: 4, fontWeight: 600 }}>→ {diasAdm}d</div>
            }
          </div>
        </div>
        <div>
          <div style={{ fontSize: 11, color: COLORS.muted, fontWeight: 700, marginBottom: 5, textTransform: "uppercase" }}>Médico Assistente</div>
          <input value={medico} onChange={e => setMedico(e.target.value)} placeholder="Nome do médico"
            style={{ border: `1.5px solid ${COLORS.border}`, borderRadius: 8, padding: "10px 12px", fontSize: 15, color: COLORS.green, outline: "none", width: "100%", background: "#fff", boxSizing: "border-box" }} />
        </div>
        <div>
          <div style={{ fontSize: 11, color: COLORS.muted, fontWeight: 700, marginBottom: 5, textTransform: "uppercase" }}>Diagnóstico</div>
          <input value={diag} onChange={e => setDiag(e.target.value)} placeholder="Ex: Sepse pulmonar"
            style={{ border: `1.5px solid ${COLORS.border}`, borderRadius: 8, padding: "10px 12px", fontSize: 15, color: COLORS.green, outline: "none", width: "100%", background: "#fff", boxSizing: "border-box" }} />
        </div>
        <div>
          <div style={{ fontSize: 11, color: COLORS.muted, fontWeight: 700, marginBottom: 8, textTransform: "uppercase" }}>Gravidade</div>
          <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
            {[["alta",COLORS.danger,"⚠ Alta"],["media",COLORS.warn,"◈ Média"],["baixa",COLORS.success,"✓ Baixa"],["livre",COLORS.muted,"Livre"]].map(([v,c,l]) => (
              <button type="button" key={v} onClick={() => setGrav(v)} style={{
                padding: "8px 14px", borderRadius: 8, cursor: "pointer",
                border: `1.5px solid ${grav===v?c:COLORS.border}`,
                background: grav===v?c+"1A":"#fff", color: grav===v?c:COLORS.green,
                fontSize: 13, fontWeight: 700, flex: "1 1 auto", minHeight: 36,
              }}>{l}</button>
            ))}
          </div>
        </div>
      </div>

      {pat.iniciais && (
        <div style={{ marginTop: 20, paddingTop: 16, borderTop: `1px solid ${COLORS.border}`, display: "flex", flexDirection: "column", gap: 10 }}>
          <button type="button" onClick={() => { onClose(); onTrocarLeito(pat); }} style={{
            width: "100%", padding: "10px", borderRadius: 8,
            border: `1.5px solid ${COLORS.greenMid}`, background: "#fff", color: COLORS.greenMid,
            fontSize: 13, fontWeight: 700, cursor: "pointer", minHeight: 40,
          }}>↔ Trocar de leito</button>
          {!confirm ? (
            <button type="button" onClick={() => setConfirm(true)} style={{
              width: "100%", padding: "10px", borderRadius: 8,
              border: `1.5px solid ${COLORS.danger}`, background: "#fff", color: COLORS.danger,
              fontSize: 13, fontWeight: 700, cursor: "pointer", minHeight: 40,
            }}>🗑 Limpar dados e desocupar leito</button>
          ) : (
            <div style={{ background: COLORS.danger+"12", borderRadius: 10, padding: 12 }}>
              <div style={{ fontSize: 13, color: COLORS.danger, fontWeight: 700, marginBottom: 10 }}>Tem certeza? Todos os dados serão apagados.</div>
              <div style={{ display: "flex", gap: 8 }}>
                <button type="button" onClick={() => { onClear(pat.id); onClose(); }} style={{ flex: 1, padding: "10px", borderRadius: 8, border: "none", background: COLORS.danger, color: "#fff", fontSize: 13, fontWeight: 700, cursor: "pointer", minHeight: 40 }}>Sim, limpar</button>
                <button type="button" onClick={() => setConfirm(false)} style={{ flex: 1, padding: "10px", borderRadius: 8, border: `1.5px solid ${COLORS.border}`, background: "#fff", color: COLORS.green, fontSize: 13, fontWeight: 600, cursor: "pointer", minHeight: 40 }}>Cancelar</button>
              </div>
            </div>
          )}
        </div>
      )}

      <div style={{ display: "flex", gap: 10, marginTop: 20 }}>
        <button type="button" onClick={onClose} style={{ flex: 1, padding: "11px", borderRadius: 10, border: `1.5px solid ${COLORS.border}`, background: "#fff", color: COLORS.green, fontSize: 14, fontWeight: 600, cursor: "pointer", minHeight: 42 }}>Cancelar</button>
        <button type="button" onClick={() => { if (canSave) { onSave(pat.id, { iniciais: iniciais.trim(), idade: idade ? Number(idade) : null, dataAdm, diagnostico: diag, medicoAssistente: medico, gravidade: grav }); onClose(); }}}
          style={{ flex: 2, padding: "11px", borderRadius: 10, border: "none", background: canSave ? COLORS.greenMid : COLORS.border, color: "#fff", fontSize: 14, fontWeight: 700, cursor: canSave ? "pointer" : "not-allowed", minHeight: 42 }}>✓ Salvar</button>
      </div>
    </Modal>
  );
}

function ConfirmModal({ icon, title, message, confirmLabel, confirmColor, onConfirm, onClose }) {
  return (
    <Modal onClose={onClose} maxWidth={400}>
      <div style={{ textAlign: "center" }}>
        <div style={{ fontSize: 44, marginBottom: 12 }}>{icon}</div>
        <div style={{ fontSize: 19, fontWeight: 800, color: COLORS.green, marginBottom: 8 }}>{title}</div>
        <div style={{ fontSize: 14, color: COLORS.muted, marginBottom: 22 }}>{message}</div>
        <div style={{ display: "flex", gap: 10 }}>
          <button type="button" onClick={onClose} style={{ flex: 1, padding: "12px", borderRadius: 10, border: `1.5px solid ${COLORS.border}`, background: "#fff", color: COLORS.green, fontSize: 14, fontWeight: 600, cursor: "pointer", minHeight: 42 }}>Cancelar</button>
          <button type="button" onClick={onConfirm} style={{ flex: 1, padding: "12px", borderRadius: 10, border: "none", background: confirmColor || COLORS.danger, color: "#fff", fontSize: 14, fontWeight: 700, cursor: "pointer", minHeight: 42 }}>{confirmLabel}</button>
        </div>
      </div>
    </Modal>
  );
}

function RelatorioModal({ patients, rounds, onClose, onSave, utiLabel }) {
  const [tipo, setTipo] = useState("geral");
  const [copiado, setCopiado] = useState(false);
  const texto = tipo === "geral"
    ? gerarRelatorioGeral(patients, rounds, utiLabel)
    : gerarRelatorioEspecifico(patients, rounds, utiLabel);

  const copiar = async () => {
    try { await navigator.clipboard.writeText(texto); setCopiado(true); setTimeout(() => setCopiado(false), 2000); }
    catch { alert("Não foi possível copiar. Selecione o texto manualmente."); }
  };

  return (
    <Modal onClose={onClose} maxWidth={640}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14 }}>
        <div style={{ fontWeight: 800, fontSize: 17, color: COLORS.green }}>📋 Relatório</div>
        <button type="button" onClick={onClose} style={{ background: "none", border: "none", fontSize: 26, cursor: "pointer", color: COLORS.muted, lineHeight: 1, padding: 0 }}>×</button>
      </div>
      <div style={{ display: "flex", gap: 6, marginBottom: 14, flexWrap: "wrap" }}>
        {[["geral","📊 Geral"],["especifico","🛏 Por Paciente"]].map(([v,l]) => (
          <button type="button" key={v} onClick={() => setTipo(v)} style={{
            padding: "8px 16px", borderRadius: 10,
            border: `1.5px solid ${tipo === v ? COLORS.greenMid : COLORS.border}`,
            background: tipo === v ? COLORS.greenMid : "#fff",
            color: tipo === v ? "#fff" : COLORS.green,
            fontSize: 13, fontWeight: 600, cursor: "pointer", flex: "1 1 auto", minHeight: 38,
          }}>{l}</button>
        ))}
      </div>
      <textarea readOnly value={texto} style={{
        width: "100%", border: `1px solid ${COLORS.border}`, borderRadius: 10, padding: 12,
        fontSize: 12, fontFamily: "monospace", resize: "vertical",
        background: COLORS.lightBg, color: COLORS.green, minHeight: 240, maxHeight: "40vh", boxSizing: "border-box",
      }} />
      <div style={{ display: "flex", gap: 8, marginTop: 14, flexWrap: "wrap" }}>
        <button type="button" onClick={copiar} style={{
          flex: "2 1 160px", padding: "11px", borderRadius: 10, border: "none",
          background: copiado ? COLORS.success : "#25D366", color: "#fff",
          fontSize: 14, fontWeight: 700, cursor: "pointer", minHeight: 42,
        }}>{copiado ? "✓ Copiado!" : "📲 Copiar p/ WhatsApp"}</button>
        <button type="button" onClick={() => exportExcel(patients, rounds, utiLabel)} style={{
          flex: "1 1 120px", padding: "11px", borderRadius: 10, border: "none",
          background: COLORS.greenMid, color: "#fff", fontSize: 14, fontWeight: 700, cursor: "pointer", minHeight: 42,
        }}>📊 Excel</button>
        <button type="button" onClick={() => onSave(tipo, texto)} style={{
          flex: "1 1 100px", padding: "11px", borderRadius: 10, border: "none",
          background: COLORS.accent, color: "#fff", fontSize: 14, fontWeight: 700, cursor: "pointer", minHeight: 42,
        }}>💾 Salvar</button>
      </div>
    </Modal>
  );
}

function HistoricoModal({ relatorios, onDelete, onClose, isMobile }) {
  const [sel, setSel] = useState(null);
  const [confirm, setConfirm] = useState(false);
  const [copiado, setCopiado] = useState(false);

  const handleDelete = () => { onDelete(sel); setSel(null); setConfirm(false); };
  const handleCopy = async () => {
    try { await navigator.clipboard.writeText(sel.texto); setCopiado(true); setTimeout(() => setCopiado(false), 2000); }
    catch { alert("Erro ao copiar."); }
  };

  return (
    <Modal onClose={onClose} maxWidth={720}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14 }}>
        <div style={{ fontWeight: 800, fontSize: 17, color: COLORS.green }}>🗂 Histórico (últimos 7 dias)</div>
        <button type="button" onClick={onClose} style={{ background: "none", border: "none", fontSize: 26, cursor: "pointer", color: COLORS.muted, lineHeight: 1, padding: 0 }}>×</button>
      </div>
      {relatorios.length === 0 ? (
        <div style={{ textAlign: "center", color: COLORS.muted, padding: 40 }}>Nenhum relatório salvo.</div>
      ) : (
        <div style={{ display: "flex", flexDirection: isMobile ? "column" : "row", gap: 12 }}>
          <div style={{
            width: isMobile ? "100%" : 200, maxHeight: isMobile && sel ? 140 : 360, overflowY: "auto",
            borderRight: isMobile ? "none" : `1px solid ${COLORS.border}`,
            borderBottom: isMobile ? `1px solid ${COLORS.border}` : "none",
            paddingRight: isMobile ? 0 : 12, paddingBottom: isMobile ? 12 : 0,
          }}>
            {relatorios.map((r, i) => (
              <div key={r.id || i} onClick={() => { setSel(r); setConfirm(false); }} style={{
                padding: "8px 10px", borderRadius: 8, cursor: "pointer", marginBottom: 6,
                background: sel === r ? COLORS.greenMid + "18" : COLORS.lightBg,
                border: `1px solid ${sel === r ? COLORS.greenMid : COLORS.border}`,
              }}>
                <div style={{ fontSize: 12, fontWeight: 700, color: COLORS.green }}>{r.tipo === "geral" ? "📊 Geral" : "🛏 Por paciente"}</div>
                <div style={{ fontSize: 11, color: COLORS.muted }}>{r.data}</div>
              </div>
            ))}
          </div>
          <div style={{ flex: 1, display: "flex", flexDirection: "column" }}>
            {sel ? (
              <>
                <textarea readOnly value={sel.texto} style={{
                  width: "100%", border: `1px solid ${COLORS.border}`, borderRadius: 10, padding: 12,
                  fontSize: 12, fontFamily: "monospace", resize: "vertical",
                  background: COLORS.lightBg, color: COLORS.green, minHeight: 200, maxHeight: "40vh", boxSizing: "border-box",
                }} />
                <div style={{ display: "flex", gap: 8, marginTop: 10, flexWrap: "wrap" }}>
                  <button type="button" onClick={handleCopy} style={{
                    flex: "2 1 140px", padding: "10px", borderRadius: 10, border: "none",
                    background: copiado ? COLORS.success : "#25D366", color: "#fff",
                    fontSize: 13, fontWeight: 700, cursor: "pointer", minHeight: 40,
                  }}>{copiado ? "✓ Copiado!" : "📲 Copiar"}</button>
                  <button type="button" onClick={() => setConfirm(true)} style={{
                    flex: "1 1 100px", padding: "10px", borderRadius: 10,
                    border: `1.5px solid ${COLORS.danger}`, background: "#fff", color: COLORS.danger,
                    fontSize: 13, fontWeight: 700, cursor: "pointer", minHeight: 40,
                  }}>🗑 Apagar</button>
                </div>
                {confirm && (
                  <div style={{ marginTop: 10, background: COLORS.danger + "12", borderRadius: 10, padding: 12 }}>
                    <div style={{ fontSize: 13, color: COLORS.danger, fontWeight: 700, marginBottom: 8 }}>Apagar este relatório?</div>
                    <div style={{ display: "flex", gap: 8 }}>
                      <button type="button" onClick={handleDelete} style={{ flex: 1, padding: "9px", borderRadius: 8, border: "none", background: COLORS.danger, color: "#fff", fontSize: 13, fontWeight: 700, cursor: "pointer", minHeight: 38 }}>Sim, apagar</button>
                      <button type="button" onClick={() => setConfirm(false)} style={{ flex: 1, padding: "9px", borderRadius: 8, border: `1.5px solid ${COLORS.border}`, background: "#fff", color: COLORS.green, fontSize: 13, fontWeight: 600, cursor: "pointer", minHeight: 38 }}>Cancelar</button>
                    </div>
                  </div>
                )}
              </>
            ) : (
              <div style={{ display: "flex", alignItems: "center", justifyContent: "center", minHeight: 180, color: COLORS.muted, fontSize: 13 }}>Selecione um relatório</div>
            )}
          </div>
        </div>
      )}
    </Modal>
  );
}

// ── Patient Card ──────────────────────────────────────────────────────────────
function PatientCard({ pat, round, onSelect, onEdit }) {
  const [color, label] = gravBadge(pat.gravidade);
  const alerts  = computeAlerts(round, pat);
  const isEmpty = !pat.iniciais;
  const dias    = pat.dataAdm ? calcDias(pat.dataAdm) : null;

  return (
    <div style={{
      background: COLORS.card, borderRadius: 14,
      border: `1.5px solid ${isEmpty ? COLORS.border : color+"44"}`,
      padding: "14px 16px", position: "relative", overflow: "hidden",
      boxShadow: "0 2px 8px rgba(27,67,50,.07)",
    }}>
      {!isEmpty && <div style={{ position: "absolute", top: 0, left: 0, width: 4, height: "100%", background: color, borderRadius: "14px 0 0 14px" }} />}
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginLeft: isEmpty ? 0 : 8 }}>
        <span style={{ fontSize: 11, fontWeight: 700, color: COLORS.muted }}>{pat.leito}</span>
        <div style={{ display: "flex", gap: 6, alignItems: "center" }}>
          {!isEmpty && <span style={{ fontSize: 11, color, fontWeight: 600, background: color+"18", padding: "2px 8px", borderRadius: 10 }}>{label}</span>}
          <button type="button" onClick={e => { e.stopPropagation(); onEdit(pat.id); }}
            style={{ background: "none", border: `1px solid ${COLORS.border}`, borderRadius: 6, padding: "4px 9px", cursor: "pointer", fontSize: 13, color: COLORS.muted, minHeight: 28 }}>
            {isEmpty ? "＋" : "✏️"}
          </button>
        </div>
      </div>
      {isEmpty ? (
        <div onClick={() => onEdit(pat.id)} style={{ marginTop: 10, cursor: "pointer" }}>
          <div style={{ color: COLORS.muted, fontSize: 13, marginBottom: 4 }}>Leito disponível</div>
          <div style={{ fontSize: 12, color: COLORS.greenMid, fontWeight: 700 }}>+ Admitir paciente</div>
        </div>
      ) : (
        <div onClick={() => onSelect(pat.id)} style={{ cursor: "pointer" }}>
          <div style={{ marginTop: 6, fontWeight: 700, fontSize: 15, color: COLORS.green, marginLeft: 8 }}>{pat.iniciais}</div>
          <div style={{ fontSize: 12, color: COLORS.muted, marginTop: 2, marginLeft: 8 }}>{pat.diagnostico || "—"}</div>
          <div style={{ display: "flex", gap: 12, marginTop: 8, marginLeft: 8, flexWrap: "wrap" }}>
            {pat.idade && <span style={{ fontSize: 12, color: COLORS.muted }}>👤 {pat.idade}a</span>}
            {pat.dataAdm && (
              dias === null
                ? <span style={{ fontSize: 12, color: COLORS.danger }}>⚠ Data adm. futura</span>
                : <span style={{ fontSize: 12, color: COLORS.muted }}>🕐 {dias}d</span>
            )}
          </div>
          {alerts.length > 0 && (
            <div style={{ marginTop: 8, marginLeft: 8, display: "flex", flexWrap: "wrap", gap: 4 }}>
              {alerts.map((a, i) => <span key={i} style={{ fontSize: 11, background: COLORS.danger+"15", color: COLORS.danger, padding: "2px 7px", borderRadius: 8, fontWeight: 600 }}>⚠ {a}</span>)}
            </div>
          )}
          <div style={{ marginTop: 8, marginLeft: 8 }}>
            {round
              ? <span style={{ fontSize: 11, background: COLORS.success+"20", color: COLORS.success, padding: "2px 8px", borderRadius: 8, fontWeight: 600 }}>✓ Round feito</span>
              : <span style={{ fontSize: 11, background: COLORS.warn+"18", color: COLORS.warn, padding: "2px 8px", borderRadius: 8, fontWeight: 600 }}>⏳ Pendente</span>}
          </div>
        </div>
      )}
    </div>
  );
}

// ── Round Form ────────────────────────────────────────────────────────────────
function RoundForm({ pat, round, onChange, onBack, onNovoRound, isMobile, saveStatus }) {
  const [confirmNovo, setConfirmNovo] = useState(false);
  const r = round || { ...INITIAL_ROUND };
  const set = (k, v) => onChange({ ...r, [k]: v });
  const tog = (k, v) => {
    const a = r[k] || [];
    onChange({ ...r, [k]: a.includes(v) ? a.filter(x => x !== v) : [...a, v] });
  };
  const s2 = ["Sim","Não"], s3 = ["Sim","Não","Sem indicação"];

  const cardStyle = {
    background: COLORS.card, borderRadius: 14,
    padding: isMobile ? "14px 14px" : "18px 20px",
    marginBottom: 12, border: `1px solid ${COLORS.border}`,
  };
  const diasAdm = pat.dataAdm ? calcDias(pat.dataAdm) : null;

  return (
    <div style={{ maxWidth: 820, margin: "0 auto", paddingBottom: 80 }}>
      <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 20, flexWrap: "wrap" }}>
        <button type="button" onClick={onBack} style={{
          background: "none", border: `1.5px solid ${COLORS.border}`, borderRadius: 8,
          padding: "8px 14px", cursor: "pointer", fontSize: 13, color: COLORS.green, fontWeight: 600, minHeight: 38,
        }}>← Voltar</button>
        <div style={{ flex: 1, minWidth: 200 }}>
          <div style={{ fontSize: 11, color: COLORS.muted, fontWeight: 700, textTransform: "uppercase" }}>{pat.leito} · Round — {new Date().toLocaleDateString("pt-BR")}</div>
          <div style={{ fontSize: isMobile ? 17 : 20, fontWeight: 800, color: COLORS.green }}>{pat.iniciais}</div>
          <div style={{ fontSize: 12, color: COLORS.muted, marginTop: 2 }}>
            {pat.idade && `${pat.idade} anos`}
            {pat.idade && pat.dataAdm && " · "}
            {pat.dataAdm && diasAdm !== null && `${diasAdm} dia(s) internado`}
          </div>
        </div>
        <button type="button" onClick={() => setConfirmNovo(true)} style={{
          padding: "8px 14px", borderRadius: 8, border: `1.5px solid ${COLORS.warn}`,
          background: "#fff", color: COLORS.warn, fontSize: 13, fontWeight: 700, cursor: "pointer", minHeight: 38,
        }}>🔄 Novo Round</button>
      </div>

      {confirmNovo && (
        <div style={{ background: COLORS.warn+"18", border: `1.5px solid ${COLORS.warn}`, borderRadius: 12, padding: 14, marginBottom: 14 }}>
          <div style={{ fontSize: 14, fontWeight: 700, color: COLORS.warn, marginBottom: 10 }}>
            ⚠️ Iniciar novo round? Todos os campos do checklist serão zerados (o paciente continua admitido).
          </div>
          <div style={{ display: "flex", gap: 8 }}>
            <button type="button" onClick={() => { onNovoRound(); setConfirmNovo(false); }} style={{
              flex: 1, padding: "10px", borderRadius: 8, border: "none",
              background: COLORS.warn, color: "#fff", fontSize: 13, fontWeight: 700, cursor: "pointer", minHeight: 40,
            }}>Sim, novo round</button>
            <button type="button" onClick={() => setConfirmNovo(false)} style={{
              flex: 1, padding: "10px", borderRadius: 8, border: `1.5px solid ${COLORS.border}`,
              background: "#fff", color: COLORS.green, fontSize: 13, fontWeight: 600, cursor: "pointer", minHeight: 40,
            }}>Cancelar</button>
          </div>
        </div>
      )}

      {/* Identificação */}
      <div style={cardStyle}>
        <SecHdr title="Identificação" icon="🏥" />
        <Grid cols={2} isMobile={isMobile}>
          <Field label="Diagnóstico"><TInput value={r.diagnostico} onChange={v => set("diagnostico", v)} placeholder="Ex: Sepse pulmonar" /></Field>
          <Field label="Médico Assistente"><TInput value={r.medicoAssistente} onChange={v => set("medicoAssistente", v)} placeholder="Nome do médico" /></Field>
        </Grid>
      </div>

      {/* Cuidados Gerais — sem visita flexibilizada, com broncoaspiração e risco de queda */}
      <div style={cardStyle}>
        <SecHdr title="Cuidados Gerais" icon="🛡️" />
        <Grid cols={3} isMobile={isMobile}>
          <Field label="Precaução de contato?">{s2.map(o => <Pill key={o} label={o} selected={r.contato===o} onClick={() => set("contato",o)} />)}</Field>
          <Field label="Risco de broncoaspiração?">{s2.map(o => <Pill key={o} label={o} selected={r.riscoBroncoaspiracao===o} onClick={() => set("riscoBroncoaspiracao",o)} color={o==="Sim"?COLORS.danger:COLORS.greenMid} />)}</Field>
          <Field label="Risco de queda?">{s2.map(o => <Pill key={o} label={o} selected={r.riscoQueda===o} onClick={() => set("riscoQueda",o)} color={o==="Sim"?COLORS.warn:COLORS.greenMid} />)}</Field>
        </Grid>
      </div>

      {/* Neurológico */}
      <div style={cardStyle}>
        <SecHdr title="Neurológico" icon="🧠" />
        <Field label="Meta de sedação (RASS)">{["-5 a -4","-2 a 0","Não se aplica"].map(o => <Pill key={o} label={o} selected={r.rass===o} onClick={() => set("rass",o)} />)}</Field>
        <Grid cols={3} isMobile={isMobile}>
          <Field label="Controle de dor?">{s2.map(o => <Pill key={o} label={o} selected={r.dor===o} onClick={() => set("dor",o)} />)}</Field>
          <Field label="Delirium?">{s2.map(o => <Pill key={o} label={o} selected={r.delirium===o} onClick={() => set("delirium",o)} />)}</Field>
          <Field label="Contenção mecânica?">{s2.map(o => <Pill key={o} label={o} selected={r.contencao===o} onClick={() => set("contencao",o)} />)}</Field>
        </Grid>
      </div>

      {/* Cardiovascular */}
      <div style={cardStyle}>
        <SecHdr title="Cardiovascular / Hemodinâmica" icon="❤️" />
        <div style={{ display: "flex", gap: 20, flexWrap: "wrap", alignItems: "flex-end" }}>
          <Field label="Suporte hemodinâmico / DVA?">{s2.map(o => <Pill key={o} label={o} selected={r.dva===o} onClick={() => set("dva",o)} />)}</Field>
          <div style={{ minWidth: 130 }}>
            <Field label="Meta de PAM (mmHg)"><TInput value={r.pam} onChange={v => set("pam",v)} placeholder="Ex: 65" w="100px" type="number" /></Field>
          </div>
        </div>
      </div>

      {/* Respiratório — sem broncoaspiração (movida para Cuidados Gerais) */}
      <div style={cardStyle}>
        <SecHdr title="Respiratório / Reabilitação" icon="🫁" />
        <Field label="Suporte respiratório">{["Sem suporte","O2 suplementar","VNI","VM invasiva"].map(o => <Pill key={o} label={o} selected={r.suporteResp===o} onClick={() => set("suporteResp",o)} />)}</Field>
        {(r.suporteResp==="VM invasiva"||r.suporteResp==="VNI") && <Field label="VM protetora?">{s2.map(o => <Pill key={o} label={o} selected={r.vmProtetora===o} onClick={() => set("vmProtetora",o)} />)}</Field>}
        <Field label="Plano de desmame">{["Redução de parâmetros","TRE hoje","Ex-TOT","Sem proposta"].map(o => <MultiPill key={o} label={o} checked={(r.planoDesmame||[]).includes(o)} onChange={() => tog("planoDesmame",o)} />)}</Field>
        <Field label="Preocupações respiratórias">{["Piora respiratória","Piora de secreção"].map(o => <MultiPill key={o} label={o} checked={(r.preocResp||[]).includes(o)} onChange={() => tog("preocResp",o)} color={COLORS.danger} />)}</Field>
        <div style={{ display: "flex", gap: 20, flexWrap: "wrap" }}>
          <div style={{ minWidth: 110 }}><Field label="Escala IMS"><TInput value={r.ims} onChange={v => set("ims",v)} placeholder="0-10" w="90px" type="number" /></Field></div>
          <Field label="Progredir nível funcional?">{s2.map(o => <Pill key={o} label={o} selected={r.progredirFuncional===o} onClick={() => set("progredirFuncional",o)} />)}</Field>
        </div>
      </div>

      {/* Nutrição */}
      <div style={cardStyle}>
        <SecHdr title="Gastrointestinal / Nutrição" icon="🍽️" />
        <Field label="Via alimentar">{["VO","SNE/GTT","NPT","Zero"].map(o => <Pill key={o} label={o} selected={r.viaAlimentar===o} onClick={() => set("viaAlimentar",o)} />)}</Field>
        <Grid cols={2} isMobile={isMobile}>
          <div>
            <Field label="Risco nutricional?">{s2.map(o => <Pill key={o} label={o} selected={r.riscoNutricional===o} onClick={() => set("riscoNutricional",o)} color={o==="Sim"?COLORS.warn:COLORS.greenMid} />)}</Field>
            {r.riscoNutricional === "Sim" && (
              <Field label="Terapia nutricional?">{s2.map(o => <Pill key={o} label={o} selected={r.terapiaNutricional===o} onClick={() => set("terapiaNutricional",o)} color={o==="Não"?COLORS.danger:COLORS.greenMid} />)}</Field>
            )}
          </div>
          <Field label="Meta calórica?">{s2.map(o => <Pill key={o} label={o} selected={r.metaCalorica===o} onClick={() => set("metaCalorica",o)} />)}</Field>
          <Field label="Aceitação">{["Normal","Baixa","Intolerância"].map(o => <Pill key={o} label={o} selected={r.aceitacao===o} onClick={() => set("aceitacao",o)} />)}</Field>
          <Field label="Avaliação fono?">{s2.map(o => <Pill key={o} label={o} selected={r.fono===o} onClick={() => set("fono",o)} />)}</Field>
          <Field label="Glicemia adequada?">{s2.map(o => <Pill key={o} label={o} selected={r.glicemia===o} onClick={() => set("glicemia",o)} />)}</Field>
          <Field label="Evacuação < 3 dias?">{s2.map(o => <Pill key={o} label={o} selected={r.evacuacao===o} onClick={() => set("evacuacao",o)} />)}</Field>
        </Grid>
      </div>

      {/* Renal / Infeccioso */}
      <Grid cols={2} isMobile={isMobile} gap={12}>
        <div style={cardStyle}>
          <SecHdr title="Renal" icon="🫘" />
          <Field label="Função renal em piora?">{["Sim","Não","Em HD"].map(o => <Pill key={o} label={o} selected={r.funcaoRenal===o} onClick={() => set("funcaoRenal",o)} />)}</Field>
          <Field label="Meta de balanço hídrico">{["Positivo","Negativo","Neutro"].map(o => <Pill key={o} label={o} selected={r.metaBH===o} onClick={() => set("metaBH",o)} />)}</Field>
        </div>
        <div style={cardStyle}>
          <SecHdr title="Infeccioso" icon="🦠" />
          <Field label="Piora infecciosa?">{s2.map(o => <Pill key={o} label={o} selected={r.pioraInfec===o} onClick={() => set("pioraInfec",o)} />)}</Field>
          <Field label="Em uso de ATB?">{s2.map(o => <Pill key={o} label={o} selected={r.atb===o} onClick={() => set("atb",o)} />)}</Field>
        </div>
      </Grid>

      {/* Profilaxias */}
      <div style={cardStyle}>
        <SecHdr title="Profilaxias" icon="💉" />
        <Grid cols={2} isMobile={isMobile}>
          <Field label="Profilaxia TEV">{s3.map(o => <Pill key={o} label={o} selected={r.tev===o} onClick={() => set("tev",o)} color={o==="Não"?COLORS.danger:COLORS.greenMid} />)}</Field>
          <Field label="Profilaxia LAMG">{s3.map(o => <Pill key={o} label={o} selected={r.lamg===o} onClick={() => set("lamg",o)} />)}</Field>
          <Field label="Úlcera de córnea">{s3.map(o => <Pill key={o} label={o} selected={r.cornea===o} onClick={() => set("cornea",o)} />)}</Field>
          <Field label="Higiene oral">{s3.map(o => <Pill key={o} label={o} selected={r.higieneOral===o} onClick={() => set("higieneOral",o)} />)}</Field>
          <Field label="Decúbito elevado">{s3.map(o => <Pill key={o} label={o} selected={r.decubito===o} onClick={() => set("decubito",o)} />)}</Field>
          <Field label="Bundles OK?">{s2.map(o => <Pill key={o} label={o} selected={r.bundles===o} onClick={() => set("bundles",o)} color={o==="Não"?COLORS.danger:COLORS.greenMid} />)}</Field>
        </Grid>
        {r.bundles==="Não" && <Field label="Bundle pendente"><TInput value={r.bundlesPendente} onChange={v => set("bundlesPendente",v)} placeholder="Qual bundle?" /></Field>}
        <Grid cols={2} isMobile={isMobile}>
          <Field label="Mudança de decúbito?">{s2.map(o => <Pill key={o} label={o} selected={r.mudancaDecubito===o} onClick={() => set("mudancaDecubito",o)} />)}</Field>
          <div>
            <Field label="Lesão por pressão?">{s2.map(o => <Pill key={o} label={o} selected={r.lesaoPressao===o} onClick={() => set("lesaoPressao",o)} color={o==="Sim"?COLORS.danger:COLORS.greenMid} />)}</Field>
            {r.lesaoPressao==="Sim" && (
              <Field label="Avaliação especializada?">{["Comissão curativo","Cirurgia plástica","Não necessário"].map(o => <Pill key={o} label={o} selected={r.avalEspecializada===o} onClick={() => set("avalEspecializada",o)} />)}</Field>
            )}
          </div>
        </Grid>
      </div>

      {/* Dispositivos */}
      <div style={cardStyle}>
        <SecHdr title="Dispositivos Invasivos" icon="🔌" />
        <DispositivosSection dev={r.dispositivos} onChange={v => set("dispositivos", v)} />
      </div>

      {/* Planejamento */}
      <div style={cardStyle}>
        <SecHdr title="Objetivos de Cuidado e Planejamento" icon="📋" />
        <Field label="Diretivas de cuidado">{["Não RCP","Não HD","Não IOT","Não DVA","Não HGT","Não coletar exames","Sem diretivas"].map(o => <MultiPill key={o} label={o} checked={(r.diretivas||[]).includes(o)} onChange={() => tog("diretivas",o)} color={o==="Sem diretivas"?COLORS.success:COLORS.danger} />)}</Field>
        <Grid cols={2} isMobile={isMobile}>
          <div>
            <Field label="Pendência de exame/procedimento?">{s2.map(o => <Pill key={o} label={o} selected={r.pendenciaExame===o} onClick={() => set("pendenciaExame",o)} />)}</Field>
            {r.pendenciaExame==="Sim" && <TArea value={r.descPendencia} onChange={v => set("descPendencia",v)} placeholder="Descreva a pendência..." />}
          </div>
          <Field label="Previsão de alta">{["Hoje","24–48h","> 48h"].map(o => <Pill key={o} label={o} selected={r.previsaoAlta===o} onClick={() => set("previsaoAlta",o)} />)}</Field>
        </Grid>
      </div>

      {/* Status de salvamento */}
      <div style={{
        background: saveStatus === "saved" ? COLORS.success + "18" : saveStatus === "error" ? COLORS.danger + "18" : COLORS.lightBg,
        border: `1.5px solid ${saveStatus === "saved" ? COLORS.success : saveStatus === "error" ? COLORS.danger : COLORS.border}`,
        borderRadius: 12, padding: "12px 16px", marginTop: 4,
        display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 10,
      }}>
        <div style={{ fontSize: 13, fontWeight: 700, color: saveStatus === "saved" ? COLORS.success : saveStatus === "error" ? COLORS.danger : COLORS.muted }}>
          {saveStatus === "saving" && "💾 Salvando automaticamente..."}
          {saveStatus === "saved"  && "✓ Salvo com sucesso"}
          {saveStatus === "error"  && "⚠ Falha ao salvar — verifique sua conexão"}
          {!saveStatus && "💡 Suas alterações são salvas automaticamente"}
        </div>
        <button type="button" onClick={onBack} style={{
          padding: "10px 24px", borderRadius: 10, border: "none",
          background: COLORS.greenMid, color: "#fff", fontSize: 14, fontWeight: 700, cursor: "pointer", minHeight: 42,
        }}>← Voltar ao dashboard</button>
      </div>
    </div>
  );
}

// ── Tela seleção de UTI ───────────────────────────────────────────────────────
function UtiSelector({ onSelect }) {
  return (
    <div style={{
      minHeight: "100vh", background: COLORS.bg,
      display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center",
      fontFamily: "'DM Sans','Segoe UI',sans-serif", padding: 24,
    }}>
      <div style={{ textAlign: "center", marginBottom: 40 }}>
        <div style={{ fontSize: 52, marginBottom: 12 }}>🏥</div>
        <div style={{ fontSize: 26, fontWeight: 800, color: COLORS.green }}>UTI HER</div>
        <div style={{ fontSize: 15, color: COLORS.muted, marginTop: 6 }}>Hospital Esperança Recife</div>
        <div style={{ fontSize: 14, color: COLORS.muted, marginTop: 4 }}>Round Multidisciplinar</div>
      </div>
      <div style={{ fontSize: 13, fontWeight: 700, color: COLORS.muted, textTransform: "uppercase", letterSpacing: ".6px", marginBottom: 16 }}>
        Selecione a UTI
      </div>
      <div style={{ display: "flex", flexDirection: "column", gap: 14, width: "100%", maxWidth: 380 }}>
        {UTI_OPTIONS.map(u => (
          <button key={u.id} type="button" onClick={() => onSelect(u)} style={{
            padding: "22px 28px", borderRadius: 16, border: `2px solid ${u.cor}`,
            background: "#fff", cursor: "pointer", textAlign: "left",
            boxShadow: "0 4px 16px rgba(27,67,50,.10)", transition: "all .15s",
          }}
            onMouseEnter={e => { e.currentTarget.style.background = u.cor; e.currentTarget.querySelectorAll("*").forEach(el => { if (el.dataset.role === "label") el.style.color = "#fff"; if (el.dataset.role === "sub") el.style.color = "rgba(255,255,255,.75)"; }); }}
            onMouseLeave={e => { e.currentTarget.style.background = "#fff"; e.currentTarget.querySelectorAll("*").forEach(el => { if (el.dataset.role === "label") el.style.color = u.cor; if (el.dataset.role === "sub") el.style.color = COLORS.muted; }); }}
          >
            <div data-role="label" style={{ fontSize: 18, fontWeight: 800, color: u.cor, marginBottom: 4 }}>{u.icon} {u.label}</div>
            <div data-role="sub" style={{ fontSize: 13, color: COLORS.muted }}>14 leitos · Round do dia</div>
          </button>
        ))}
      </div>
      <div style={{ fontSize: 12, color: COLORS.muted, marginTop: 32 }}>
        {new Date().toLocaleDateString("pt-BR", { weekday: "long", day: "2-digit", month: "long", year: "numeric" })}
      </div>
    </div>
  );
}

// ── App Principal ─────────────────────────────────────────────────────────────
export default function App() {
  const isMobile = useIsMobile();

  const [utiSelecionada,  setUtiSelecionada]  = useState(null);
  const [patients,        setPatients]        = useState([]);
  const [rounds,          setRounds]          = useState({});
  const [selected,        setSelected]        = useState(null);
  const [editingId,       setEditingId]       = useState(null);
  const [trocarLeitoId,   setTrocarLeitoId]   = useState(null);
  const [showRelatorio,   setShowRelatorio]   = useState(false);
  const [showHistorico,   setShowHistorico]   = useState(false);
  const [showClearAll,    setShowClearAll]    = useState(false);
  const [search,          setSearch]          = useState("");
  const [filter,          setFilter]          = useState("todos");
  const [loading,         setLoading]         = useState(false);
  const [saveStatus,      setSaveStatus]      = useState("");
  const [error,           setError]           = useState(null);
  const [relatorios,      setRelatorios]      = useState([]);

  const pendingChanges = useRef({});
  const savingPatients = useRef(new Set());
  const selectedRef    = useRef(null);
  const saveStatusRef  = useRef("");

  useEffect(() => { selectedRef.current = selected; }, [selected]);
  useEffect(() => { saveStatusRef.current = saveStatus; }, [saveStatus]);

  const saveLoop = useCallback(async (patientId) => {
    if (savingPatients.current.has(patientId)) return;
    savingPatients.current.add(patientId);
    try {
      while (pendingChanges.current[patientId] !== undefined) {
        const dados = pendingChanges.current[patientId];
        delete pendingChanges.current[patientId];
        setSaveStatus("saving"); saveStatusRef.current = "saving";
        let tentativas = 0, sucesso = false;
        while (tentativas < 5 && !sucesso) {
          try {
            const { error: err } = await supabase.from("rounds").upsert({
              patient_id: patientId, data: hojeStr(), round_data: dados,
              updated_at: new Date().toISOString(),
            }, { onConflict: "patient_id,data" });
            if (err) throw err;
            sucesso = true;
          } catch {
            tentativas++;
            if (tentativas < 5) await new Promise(r => setTimeout(r, 1000 * tentativas));
          }
        }
        if (!sucesso) {
          pendingChanges.current[patientId] = dados;
          setSaveStatus("error"); saveStatusRef.current = "error";
          break;
        }
      }
      if (Object.keys(pendingChanges.current).length === 0 && saveStatusRef.current !== "error") {
        setSaveStatus("saved"); saveStatusRef.current = "saved";
        setTimeout(() => { setSaveStatus(prev => prev === "saved" ? "" : prev); if (saveStatusRef.current === "saved") saveStatusRef.current = ""; }, 2000);
      }
    } finally { savingPatients.current.delete(patientId); }
  }, []);

  const saveTimers = useRef({});
  const handleChange = useCallback((novoRound) => {
    const id = selectedRef.current;
    if (!id) return;
    setRounds(prev => ({ ...prev, [id]: novoRound }));
    pendingChanges.current[id] = novoRound;
    if (saveTimers.current[id]) clearTimeout(saveTimers.current[id]);
    saveTimers.current[id] = setTimeout(() => saveLoop(id), 400);
  }, [saveLoop]);

  const handleBackFromRound = useCallback(async () => {
    const id = selectedRef.current;
    if (id && saveTimers.current[id]) { clearTimeout(saveTimers.current[id]); delete saveTimers.current[id]; }
    if (id && pendingChanges.current[id] !== undefined) await saveLoop(id);
    setSelected(null);
  }, [saveLoop]);

  useEffect(() => {
    const i = setInterval(() => {
      Object.keys(pendingChanges.current).forEach(idStr => {
        const id = Number(idStr);
        if (!savingPatients.current.has(id)) saveLoop(id);
      });
    }, 10000);
    return () => clearInterval(i);
  }, [saveLoop]);

  const handleNovoRound = useCallback(async () => {
    const id = selectedRef.current;
    if (!id) return;
    const pat = patients.find(p => p.id === id);
    const zerado = { ...INITIAL_ROUND, diagnostico: pat?.diagnostico || "", medicoAssistente: pat?.medicoAssistente || "" };
    setRounds(prev => ({ ...prev, [id]: zerado }));
    pendingChanges.current[id] = zerado;
    if (saveTimers.current[id]) { clearTimeout(saveTimers.current[id]); delete saveTimers.current[id]; }
    await saveLoop(id);
  }, [patients, saveLoop]);

  // Carrega dados quando UTI é selecionada
  useEffect(() => {
    if (!utiSelecionada) return;
    async function load() {
      setLoading(true); setError(null);
      try {
        const dataHoje = hojeStr();
        const { data: pats, error: e1 } = await supabase.from("patients").select("*").eq("uti_id", utiSelecionada.id).order("id");
        if (e1) throw e1;
        const patIds = (pats || []).map(p => p.id);
        const { data: rds, error: e2 } = patIds.length > 0
          ? await supabase.from("rounds").select("*").eq("data", dataHoje).in("patient_id", patIds)
          : { data: [], error: null };
        if (e2) throw e2;
        setPatients((pats || []).map(p => ({
          ...p, iniciais: p.iniciais || "", diagnostico: p.diagnostico || "",
          medicoAssistente: p.medicoAssistente || "",
          gravidade: p.gravidade || "livre", dataAdm: p.dataAdm || null, idade: p.idade || null,
        })));
        const map = {};
        (rds || []).forEach(r => { map[r.patient_id] = normalizarRound(r.round_data); });
        setRounds(map);
        const seteDiasAtras = new Date(); seteDiasAtras.setDate(seteDiasAtras.getDate() - 7);
        const { data: rels } = await supabase.from("relatorios").select("*")
          .eq("uti_id", utiSelecionada.id)
          .gte("created_at", seteDiasAtras.toISOString())
          .order("created_at", { ascending: false });
        if (rels) setRelatorios(rels.map(r => ({ id: r.id, tipo: r.tipo, texto: r.texto, data: new Date(r.created_at).toLocaleString("pt-BR") })));
      } catch (err) { setError("Erro ao conectar ao banco de dados."); console.error(err); }
      setLoading(false);
    }
    load();
  }, [utiSelecionada]);

  // Auto-refresh
  useEffect(() => {
    if (!utiSelecionada || selected || editingId || showRelatorio || showHistorico) return;
    const i = setInterval(async () => {
      if (Object.keys(pendingChanges.current).length > 0) return;
      if (savingPatients.current.size > 0) return;
      try {
        const dataHoje = hojeStr();
        const { data: pats } = await supabase.from("patients").select("*").eq("uti_id", utiSelecionada.id).order("id");
        const patIds = (pats || []).map(p => p.id);
        const { data: rds } = patIds.length > 0
          ? await supabase.from("rounds").select("*").eq("data", dataHoje).in("patient_id", patIds)
          : { data: [] };
        if (pats) setPatients(pats.map(p => ({
          ...p, iniciais: p.iniciais||"", diagnostico: p.diagnostico||"",
          medicoAssistente: p.medicoAssistente||"",
          gravidade: p.gravidade||"livre", dataAdm: p.dataAdm||null, idade: p.idade||null,
        })));
        if (rds) { const map = {}; rds.forEach(r => { map[r.patient_id] = normalizarRound(r.round_data); }); setRounds(map); }
      } catch { /* silencioso */ }
    }, 30000);
    return () => clearInterval(i);
  }, [utiSelecionada, selected, editingId, showRelatorio, showHistorico]);

  const savePatient = async (id, data) => {
    setPatients(prev => prev.map(p => p.id===id ? {...p,...data} : p));
    await supabase.from("patients").update({
      iniciais: data.iniciais, idade: data.idade||null, dataAdm: data.dataAdm||null,
      diagnostico: data.diagnostico, medicoAssistente: data.medicoAssistente||"",
      gravidade: data.gravidade, updated_at: new Date().toISOString(),
    }).eq("id", id);
  };

  const clearPatient = async (id) => {
    setPatients(prev => prev.map(p => p.id===id ? {...p, iniciais:"", idade:null, dataAdm:null, diagnostico:"", medicoAssistente:"", gravidade:"livre"} : p));
    setRounds(prev => { const n={...prev}; delete n[id]; return n; });
    delete pendingChanges.current[id];
    if (saveTimers.current[id]) { clearTimeout(saveTimers.current[id]); delete saveTimers.current[id]; }
    await supabase.from("patients").update({ iniciais:"", idade:null, dataAdm:null, diagnostico:"", medicoAssistente:"", gravidade:"livre" }).eq("id", id);
    await supabase.from("rounds").delete().eq("patient_id", id).eq("data", hojeStr());
  };

  // Troca de leito: se destino vago, move; se ocupado, faz swap completo (dados + rounds)
  const handleTrocarLeito = async (origem, destino) => {
    const dadosOrigem  = { iniciais: origem.iniciais, idade: origem.idade||null, dataAdm: origem.dataAdm||null, diagnostico: origem.diagnostico, medicoAssistente: origem.medicoAssistente||"", gravidade: origem.gravidade };
    const dadosDestino = destino.iniciais
      ? { iniciais: destino.iniciais, idade: destino.idade||null, dataAdm: destino.dataAdm||null, diagnostico: destino.diagnostico, medicoAssistente: destino.medicoAssistente||"", gravidade: destino.gravidade }
      : { iniciais: "", idade: null, dataAdm: null, diagnostico: "", medicoAssistente: "", gravidade: "livre" };

    // Atualiza estado local imediatamente
    setPatients(prev => prev.map(p => {
      if (p.id === origem.id)  return { ...p, ...dadosDestino };
      if (p.id === destino.id) return { ...p, ...dadosOrigem };
      return p;
    }));

    // Troca rounds em memória
    setRounds(prev => {
      const n = { ...prev };
      const rOrigem  = n[origem.id];
      const rDestino = n[destino.id];
      if (rOrigem)  n[destino.id] = rOrigem;  else delete n[destino.id];
      if (rDestino) n[origem.id]  = rDestino; else delete n[origem.id];
      return n;
    });

    // Persiste no banco: dados do paciente
    await supabase.from("patients").update({ ...dadosDestino, updated_at: new Date().toISOString() }).eq("id", origem.id);
    await supabase.from("patients").update({ ...dadosOrigem,  updated_at: new Date().toISOString() }).eq("id", destino.id);

    // Troca rounds do dia no banco usando patient_id
    const dataHoje = hojeStr();
    const { data: rdsHoje } = await supabase.from("rounds").select("id,patient_id,round_data").eq("data", dataHoje).in("patient_id", [origem.id, destino.id]);

    if (rdsHoje && rdsHoje.length > 0) {
      const rOrigem  = rdsHoje.find(r => r.patient_id === origem.id);
      const rDestino = rdsHoje.find(r => r.patient_id === destino.id);
      if (rOrigem && rDestino) {
        // Swap: atualiza cada um com o round_data do outro
        await supabase.from("rounds").update({ round_data: rDestino.round_data, updated_at: new Date().toISOString() }).eq("id", rOrigem.id);
        await supabase.from("rounds").update({ round_data: rOrigem.round_data,  updated_at: new Date().toISOString() }).eq("id", rDestino.id);
      } else if (rOrigem) {
        // Apenas origem tinha round — move para destino
        await supabase.from("rounds").update({ patient_id: destino.id, updated_at: new Date().toISOString() }).eq("id", rOrigem.id);
      } else if (rDestino) {
        // Apenas destino tinha round — move para origem
        await supabase.from("rounds").update({ patient_id: origem.id, updated_at: new Date().toISOString() }).eq("id", rDestino.id);
      }
    }
  };

  const clearAll = async () => {
    const ids = patients.map(p => p.id);
    setPatients(prev => prev.map(p => ({...p, iniciais:"", idade:null, dataAdm:null, diagnostico:"", medicoAssistente:"", gravidade:"livre"})));
    setRounds({}); setShowClearAll(false);
    pendingChanges.current = {};
    Object.keys(saveTimers.current).forEach(k => clearTimeout(saveTimers.current[k]));
    saveTimers.current = {};
    if (ids.length > 0) {
      await supabase.from("patients").update({ iniciais:"", idade:null, dataAdm:null, diagnostico:"", medicoAssistente:"", gravidade:"livre" }).in("id", ids);
      await supabase.from("rounds").delete().eq("data", hojeStr()).in("patient_id", ids);
    }
  };

  const salvarRelatorio = async (tipo, texto) => {
    try {
      const { data, error: err } = await supabase.from("relatorios")
        .insert({ tipo, texto, uti_id: utiSelecionada.id, created_at: new Date().toISOString() })
        .select().single();
      if (err) throw err;
      setRelatorios(prev => [{ id: data.id, tipo, texto, data: new Date().toLocaleString("pt-BR") }, ...prev]);
    } catch (err) { console.error("Erro ao salvar relatório:", err); alert("Erro ao salvar relatório. Verifique sua conexão."); }
  };

  const deletarRelatorio = async (rel) => {
    setRelatorios(prev => prev.filter(r => r !== rel));
    if (rel.id) await supabase.from("relatorios").delete().eq("id", rel.id);
  };

  if (!utiSelecionada) return <UtiSelector onSelect={(u) => { setUtiSelecionada(u); setPatients([]); setRounds({}); setRelatorios([]); }} />;

  const selPat      = patients.find(p => p.id === selected);
  const editPat     = patients.find(p => p.id === editingId);
  const trocarPat   = patients.find(p => p.id === trocarLeitoId);
  const date = new Date().toLocaleDateString("pt-BR", isMobile
    ? { day: "2-digit", month: "2-digit", year: "numeric" }
    : { weekday: "long", day: "2-digit", month: "long", year: "numeric" });

  const withPats     = patients.filter(p => p.iniciais);
  const roundsDone   = withPats.filter(p => rounds[p.id]).length;
  const total        = withPats.length;
  const totalAlertas = patients.filter(p => computeAlerts(rounds[p.id], p).length > 0).length;

  const filtered = patients.filter(p => {
    if (search) {
      const q = search.toLowerCase();
      if (!p.iniciais?.toLowerCase().includes(q) && !p.leito?.toLowerCase().includes(q)) return false;
    }
    if (filter==="pendente" && (!p.iniciais || rounds[p.id])) return false;
    if (filter==="alta"     && p.gravidade!=="alta") return false;
    return true;
  });

  if (loading) return (
    <div style={{ background: COLORS.bg, minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center", fontFamily: "'DM Sans','Segoe UI',sans-serif", padding: 20 }}>
      <div style={{ textAlign: "center" }}>
        <div style={{ fontSize: 48, marginBottom: 16 }}>🏥</div>
        <div style={{ fontSize: 17, fontWeight: 700, color: COLORS.green }}>Carregando dados da UTI...</div>
        <div style={{ fontSize: 13, color: COLORS.muted, marginTop: 8 }}>Conectando ao banco de dados</div>
      </div>
    </div>
  );

  if (error) return (
    <div style={{ background: COLORS.bg, minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center", fontFamily: "'DM Sans','Segoe UI',sans-serif", padding: 20 }}>
      <div style={{ textAlign: "center", maxWidth: 400 }}>
        <div style={{ fontSize: 48, marginBottom: 16 }}>⚠️</div>
        <div style={{ fontSize: 17, fontWeight: 700, color: COLORS.danger, marginBottom: 8 }}>{error}</div>
        <button type="button" onClick={() => window.location.reload()} style={{ padding: "11px 22px", borderRadius: 10, border: "none", background: COLORS.greenMid, color: "#fff", fontSize: 14, fontWeight: 700, cursor: "pointer" }}>Tentar novamente</button>
      </div>
    </div>
  );

  if (selected && selPat) return (
    <div style={{ background: COLORS.bg, minHeight: "100vh", fontFamily: "'DM Sans','Segoe UI',sans-serif" }}>
      <div style={{ background: utiSelecionada.cor, padding: "10px 16px", display: "flex", alignItems: "center", justifyContent: "space-between", gap: 10, position: "sticky", top: 0, zIndex: 50 }}>
        <div style={{ color: "#fff", fontWeight: 800, fontSize: isMobile ? 13 : 15 }}>🏥 {utiSelecionada.label}</div>
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          {saveStatus === "saving" && <span style={{ color: "#FFD54F", fontSize: 11, fontWeight: 600 }}>💾 Salvando...</span>}
          {saveStatus === "saved"  && <span style={{ color: "#7FEFA7", fontSize: 11, fontWeight: 600 }}>✓ Salvo</span>}
          {saveStatus === "error"  && <span style={{ color: "#FF6B6B", fontSize: 11, fontWeight: 600 }}>⚠ Erro</span>}
          <div style={{ color: "rgba(255,255,255,.7)", fontSize: 12 }}>{date}</div>
        </div>
      </div>
      <div style={{ padding: isMobile ? 14 : 24 }}>
        <RoundForm
          pat={selPat}
          round={rounds[selected] || { ...INITIAL_ROUND, diagnostico: selPat.diagnostico || "", medicoAssistente: selPat.medicoAssistente || "" }}
          onChange={handleChange}
          onBack={handleBackFromRound}
          onNovoRound={handleNovoRound}
          isMobile={isMobile}
          saveStatus={saveStatus}
        />
      </div>
    </div>
  );

  return (
    <div style={{ background: COLORS.bg, minHeight: "100vh", fontFamily: "'DM Sans','Segoe UI',sans-serif" }}>
      <div style={{ background: utiSelecionada.cor, padding: isMobile ? "10px 14px" : "13px 28px", position: "sticky", top: 0, zIndex: 50 }}>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 8, flexWrap: "wrap" }}>
          <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
            <span style={{ fontSize: isMobile ? 20 : 22 }}>🏥</span>
            <div>
              <div style={{ color: "#fff", fontWeight: 800, fontSize: isMobile ? 14 : 17 }}>{utiSelecionada.label}</div>
              <div style={{ color: "rgba(255,255,255,.7)", fontSize: 11 }}>{isMobile ? date : `Round Multidisciplinar · ${patients.length} leitos`}</div>
            </div>
          </div>
          <div style={{ display: "flex", alignItems: "center", gap: 6, flexWrap: "wrap" }}>
            {!isMobile && <div style={{ color: "rgba(255,255,255,.7)", fontSize: 13, marginRight: 4 }}>{date}</div>}
            <button type="button" onClick={() => { setUtiSelecionada(null); setSelected(null); setEditingId(null); }} style={{
              padding: isMobile ? "7px 10px" : "8px 14px", borderRadius: 10,
              border: "1.5px solid rgba(255,255,255,.3)", background: "transparent", color: "#fff",
              fontSize: 12, fontWeight: 600, cursor: "pointer", minHeight: 34,
            }}>🔄 {!isMobile && "Trocar UTI"}</button>
            <button type="button" onClick={() => setShowHistorico(true)} style={{
              padding: isMobile ? "7px 10px" : "8px 14px", borderRadius: 10,
              border: "1.5px solid rgba(255,255,255,.3)", background: "transparent", color: "#fff",
              fontSize: 12, fontWeight: 600, cursor: "pointer", minHeight: 34,
            }}>🗂 {!isMobile && "Histórico"}</button>
            <button type="button" onClick={() => setShowClearAll(true)} style={{
              padding: isMobile ? "7px 10px" : "8px 14px", borderRadius: 10,
              border: `1.5px solid ${COLORS.danger}88`, background: "transparent", color: "#ffaaaa",
              fontSize: 12, fontWeight: 700, cursor: "pointer", minHeight: 34,
            }}>🗑 {!isMobile && "Limpar"}</button>
            <button type="button" onClick={() => setShowRelatorio(true)} style={{
              padding: isMobile ? "7px 12px" : "8px 18px", borderRadius: 10,
              border: "none", background: COLORS.accent, color: "#fff",
              fontSize: 12, fontWeight: 700, cursor: "pointer", minHeight: 34,
            }}>📋 Relatório</button>
          </div>
        </div>
      </div>

      <div style={{ padding: isMobile ? "14px 14px" : "24px 28px" }}>
        <div style={{ display: "grid", gridTemplateColumns: isMobile ? "1fr 1fr" : "repeat(4,1fr)", gap: isMobile ? 10 : 14, marginBottom: 16 }}>
          {[
            ["Leitos ocupados", total,             COLORS.green,   "🛏"],
            ["Rounds feitos",   roundsDone,         COLORS.success, "✅"],
            ["Pendentes",       total - roundsDone, COLORS.warn,    "⏳"],
            ["Alertas",         totalAlertas,        COLORS.danger,  "⚠️"],
          ].map(([lbl, val, color, icon]) => (
            <div key={lbl} style={{ background: "#fff", borderRadius: 12, padding: isMobile ? "12px 14px" : "16px 20px", border: `1px solid ${COLORS.border}` }}>
              <div style={{ fontSize: isMobile ? 18 : 22, marginBottom: 4 }}>{icon}</div>
              <div style={{ fontSize: isMobile ? 22 : 26, fontWeight: 800, color }}>{val}</div>
              <div style={{ fontSize: 11, color: COLORS.muted, fontWeight: 600 }}>{lbl}</div>
            </div>
          ))}
        </div>

        <div style={{ background: "#fff", borderRadius: 12, padding: "14px 16px", border: `1px solid ${COLORS.border}`, marginBottom: 16 }}>
          <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 8 }}>
            <span style={{ fontSize: 13, fontWeight: 700, color: COLORS.green }}>Progresso do round</span>
            <span style={{ fontSize: 13, color: COLORS.muted }}>{roundsDone}/{total} leitos</span>
          </div>
          <div style={{ background: COLORS.border, borderRadius: 8, height: 8 }}>
            <div style={{ width: `${total ? (roundsDone / total) * 100 : 0}%`, height: "100%", background: COLORS.greenMid, borderRadius: 8, transition: "width .4s" }} />
          </div>
        </div>

        <div style={{ display: "flex", gap: 8, marginBottom: 16, flexWrap: "wrap" }}>
          <input value={search} onChange={e => setSearch(e.target.value)} placeholder="🔍 Buscar paciente ou leito..."
            style={{
              flex: "1 1 200px", minWidth: 140,
              border: `1.5px solid ${COLORS.border}`, borderRadius: 10, padding: "10px 14px",
              fontSize: 14, color: COLORS.green, outline: "none", background: "#fff",
              minHeight: 40, boxSizing: "border-box",
            }} />
          {[["todos","Todos"],["pendente","Pendentes"],["alta","Grav. Alta"]].map(([v,l]) => (
            <button type="button" key={v} onClick={() => setFilter(v)} style={{
              padding: "8px 14px", borderRadius: 10,
              border: `1.5px solid ${filter===v ? COLORS.greenMid : COLORS.border}`,
              background: filter===v ? COLORS.greenMid : "#fff",
              color: filter===v ? "#fff" : COLORS.green,
              fontSize: 13, fontWeight: 600, cursor: "pointer", minHeight: 40,
            }}>{l}</button>
          ))}
        </div>

        <div style={{
          display: "grid",
          gridTemplateColumns: isMobile ? "repeat(auto-fill, minmax(160px, 1fr))" : "repeat(auto-fill, minmax(220px, 1fr))",
          gap: isMobile ? 10 : 14,
        }}>
          {filtered.length === 0 ? (
            <div style={{ gridColumn: "1 / -1", textAlign: "center", padding: 40, color: COLORS.muted, fontSize: 14 }}>Nenhum leito encontrado.</div>
          ) : (
            filtered.map(p => <PatientCard key={p.id} pat={p} round={rounds[p.id]} onSelect={setSelected} onEdit={setEditingId} />)
          )}
        </div>
      </div>

      {editingId && editPat && (
        <EditPatientModal
          pat={editPat}
          onSave={savePatient}
          onClear={clearPatient}
          onClose={() => setEditingId(null)}
          onTrocarLeito={(p) => { setEditingId(null); setTrocarLeitoId(p.id); }}
        />
      )}
      {trocarLeitoId && trocarPat && (
        <TrocarLeitoModal
          pat={trocarPat}
          patients={patients}
          onTrocar={handleTrocarLeito}
          onClose={() => setTrocarLeitoId(null)}
        />
      )}
      {showRelatorio  && <RelatorioModal patients={patients} rounds={rounds} onClose={() => setShowRelatorio(false)} onSave={salvarRelatorio} utiLabel={utiSelecionada.label} />}
      {showHistorico  && <HistoricoModal relatorios={relatorios} onDelete={deletarRelatorio} onClose={() => setShowHistorico(false)} isMobile={isMobile} />}
      {showClearAll   && <ConfirmModal
        icon="🗑️"
        title="Limpar todos os pacientes?"
        message={`Todos os pacientes e rounds de hoje da ${utiSelecionada.label} serão apagados.`}
        confirmLabel="Sim, limpar"
        onConfirm={clearAll}
        onClose={() => setShowClearAll(false)} />}
    </div>
  );
}
