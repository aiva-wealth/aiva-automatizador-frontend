import React, { useState, useEffect } from "react";
import { supabase } from "./supabaseClient";

const NAVY = "#1D2E49";
const CREAM = "#F5F1EC";
const TEAL = "#557787";
const LIGHTBLUE = "#D1DFEA";
const GREY = "#9B9B93";

const BACKEND_URL = import.meta.env.VITE_BACKEND_URL;

const USUARIOS = ["Carme", "Fede", "Juli", "Mati"];

const RISK_TEXT = {
  Conservador:
    "Inversor preocupado mayormente por la preservación del capital y menos preocupado por los rendimientos de la inversión. Prefiere inversiones de bajo riesgo, retornos estables y baja volatilidad.",
  Balanceado:
    "El inversor reconoce que para obtener mayores retornos es necesario correr algunos riesgos; está dispuesto entonces a tomar mayor exposición en activos de riesgo, tolerando mayores fluctuaciones, pero evitando en la medida de los posible, cualquier cambio brusco, drástico o muy frecuente.",
  Agresivo:
    "El inversor presenta una alta tolerancia al riesgo y está dispuesto a asumir mayores riesgos y tolerar alta volatilidad. Esta dispuesto a tolerar fluctuaciones frecuentes e incluso profundas en su portafolio, disminución de valor, a cambio de obtener rendimientos potencialmente mayores en el mediano plazo.",
};

const STEPS = [
  "Portada",
  "Equipo",
  "Propuesta de valor",
  "Estrategia",
  "Portafolio actual",
  "Portafolio propuesto",
  "Comentarios",
];

const DEFAULT_TEAM = [
  { id: 1, nombre: "Carmela Hernández", puesto: "Investment Specialist", educacion: "Lic. Economía, Candidate CFA III", incluido: true },
  { id: 2, nombre: "Gregorio Anza", puesto: "Dealing Desk Specialist", educacion: "Lic. Administración de Empresas", incluido: true },
  { id: 3, nombre: "Julieta Broggi", puesto: "Investment Analyst", educacion: "Contadora Pública, Bsc. in Finance, University of London", incluido: true },
  { id: 4, nombre: "Federico Cañette", puesto: "Investment Support", educacion: "Lic. Economía", incluido: true },
  { id: 5, nombre: "Belén Rodriguez", puesto: "Investment Execution Analyst", educacion: "Estudiante Lic. Economía", incluido: true },
];

const CATEGORIAS = ["Renta Fija", "Multi Activo", "Renta Variable", "Alternativos Líquidos"];

function Section({ title, children }) {
  return (
    <div style={{ marginBottom: 28 }}>
      <h3 style={{ color: NAVY, fontSize: 13, fontWeight: 700, letterSpacing: 1, textTransform: "uppercase", marginBottom: 14 }}>
        {title}
      </h3>
      {children}
    </div>
  );
}

function Field({ label, children }) {
  return (
    <div style={{ marginBottom: 14 }}>
      <label style={{ display: "block", fontSize: 12.5, color: "#5b5b55", marginBottom: 5 }}>{label}</label>
      {children}
    </div>
  );
}

const inputStyle = {
  width: "100%", padding: "9px 11px", borderRadius: 6, border: "1px solid #d8d5cc",
  fontSize: 14, fontFamily: "inherit", boxSizing: "border-box", background: "#fff",
};

export default function App() {
  const [usuario, setUsuario] = useState("");
  const [repcode, setRepcode] = useState("");
  const [asesorQuery, setAsesorQuery] = useState("");
  const [asesorResultados, setAsesorResultados] = useState([]);
  const [asesorSel, setAsesorSel] = useState(null);
  const [step, setStep] = useState(0);

  const [tipo, setTipo] = useState("Propuesta");
  const [cliente, setCliente] = useState("");
  const [nroCuenta, setNroCuenta] = useState("");
  const [incluirPagina2, setIncluirPagina2] = useState(false);
  const [team, setTeam] = useState(DEFAULT_TEAM);
  const [incluirValueProp, setIncluirValueProp] = useState(true);
  const [perfil, setPerfil] = useState("Balanceado");

  const [currentAssets, setCurrentAssets] = useState([]);
  const [montoInvertir, setMontoInvertir] = useState(500000);
  const [fondoQuery, setFondoQuery] = useState("");
  const [fondoResultados, setFondoResultados] = useState([]);
  const [proposedAssets, setProposedAssets] = useState([]);
  const [comentarios, setComentarios] = useState("");

  const [generando, setGenerando] = useState(false);
  const [resultado, setResultado] = useState(null);
  const [error, setError] = useState("");

  // --- Búsqueda de asesor en Supabase (por repcode o nombre) ---
  useEffect(() => {
    if (asesorQuery.trim().length < 2) { setAsesorResultados([]); return; }
    const t = setTimeout(async () => {
      const { data } = await supabase
        .from("asesores")
        .select("repcode, nombre, email")
        .or(`repcode.ilike.%${asesorQuery}%,nombre.ilike.%${asesorQuery}%`)
        .limit(8);
      setAsesorResultados(data || []);
    }, 250);
    return () => clearTimeout(t);
  }, [asesorQuery]);

  // --- Búsqueda de fondos en Supabase (por ISIN o nombre) ---
  useEffect(() => {
    if (fondoQuery.trim().length < 2) { setFondoResultados([]); return; }
    const t = setTimeout(async () => {
      const { data } = await supabase
        .from("fondos")
        .select("isin, nombre, sector, categoria, ter, uso_frecuente")
        .or(`isin.ilike.%${fondoQuery}%,nombre.ilike.%${fondoQuery}%`)
        .order("uso_frecuente", { ascending: false })
        .limit(8);
      setFondoResultados(data || []);
    }, 250);
    return () => clearTimeout(t);
  }, [fondoQuery]);

  function addProposedAsset(fondo) {
    setProposedAssets((prev) => [...prev, { ...fondo, categoria: fondo.categoria || "Renta Variable", pct: 0, monto: 0, ytd: 0, y1: 0, y3: 0, y5: 0 }]);
    setFondoQuery("");
    setFondoResultados([]);
  }

  function updateProposedField(idx, field, value) {
    setProposedAssets((prev) => prev.map((a, i) => {
      if (i !== idx) return a;
      const next = { ...a, [field]: value };
      if (field === "pct") next.monto = Math.round((value / 100) * montoInvertir);
      if (field === "monto") next.pct = montoInvertir ? +(100 * value / montoInvertir).toFixed(1) : 0;
      return next;
    }));
  }

  function buildConfig() {
    const categorias = CATEGORIAS.map((label) => ({
      label: label === "Renta Fija" ? "Fondos Renta Fija" : label === "Multi Activo" ? "Fondo Multi Activo" : label === "Renta Variable" ? "Fondo Renta Variable" : "Fondos Alternativos Líquidos",
      fondos: proposedAssets.filter((a) => a.categoria === label).map((a) => ({
        isin: a.isin, nombre: a.nombre, sector: a.sector || "", ytd: a.ytd || 0, y1: a.y1 || 0, y3: a.y3 || 0, y5: a.y5 || 0, pct: a.pct, monto: a.monto, ter: a.ter || 0,
      })),
    })).filter((c) => c.fondos.length > 0);

    const cashMonto = Math.max(0, montoInvertir - proposedAssets.reduce((s, a) => s + (a.monto || 0), 0));

    const byCat = {};
    proposedAssets.forEach((a) => { byCat[a.categoria] = (byCat[a.categoria] || 0) + (a.pct || 0) / 100; });
    const donut1 = {
      "Fondos Renta Fija": byCat["Renta Fija"] || 0,
      "Fondos Renta Variable": byCat["Renta Variable"] || 0,
      "Fondos Multi Asset": byCat["Multi Activo"] || 0,
      "Cash": montoInvertir ? cashMonto / montoInvertir : 0,
      "Fondos Alternativos Liquidos": byCat["Alternativos Líquidos"] || 0,
    };
    const donut2 = {
      "Fondos Renta Fija + Cash": (donut1["Fondos Renta Fija"] || 0) + (donut1["Cash"] || 0),
      "Fondos Renta Variable": (donut1["Fondos Renta Variable"] || 0) + (donut1["Fondos Multi Asset"] || 0) + (donut1["Fondos Alternativos Liquidos"] || 0),
    };

    const fondosPorCategoria = {
      "Renta Fija & Multi Activo": proposedAssets.filter((a) => a.categoria === "Renta Fija" || a.categoria === "Multi Activo").map((a) => ({ nombre: a.nombre, descripcion: a.descripcion || "", factsheet_url: a.factsheet_url || "" })),
      "Renta Variable": proposedAssets.filter((a) => a.categoria === "Renta Variable").map((a) => ({ nombre: a.nombre, descripcion: a.descripcion || "", factsheet_url: a.factsheet_url || "" })),
      "Alternativos Líquidos": proposedAssets.filter((a) => a.categoria === "Alternativos Líquidos").map((a) => ({ nombre: a.nombre, descripcion: a.descripcion || "", factsheet_url: a.factsheet_url || "" })),
    };

    return {
      cliente, nro_cuenta: nroCuenta,
      incluir_pagina2: incluirPagina2, incluir_valor: incluirValueProp,
      equipo: team.filter((m) => m.incluido).map((m) => ({ nombre: m.nombre, puesto: m.puesto, educacion: m.educacion })),
      perfil_riesgo: perfil,
      portafolio_actual: currentAssets.map((a) => ({ nombre: a.nombre, pct: a.pct, importe: a.importe })),
      categorias_propuesto: categorias,
      cash_monto: cashMonto,
      monto_total: montoInvertir,
      asset_allocation_donut1: donut1,
      asset_allocation_donut2: donut2,
      fondos_por_categoria: fondosPorCategoria,
      comentarios,
    };
  }

  async function handleGenerar() {
    setGenerando(true); setError(""); setResultado(null);
    try {
      const config = buildConfig();
      const res = await fetch(`${BACKEND_URL}/generar`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ tipo, creado_por: usuario, repcode, cliente, config }),
      });
      if (!res.ok) throw new Error(`El backend respondió ${res.status}`);
      const data = await res.json();
      setResultado(data);
    } catch (e) {
      setError(e.message || "Error generando la propuesta");
    } finally {
      setGenerando(false);
    }
  }

  if (!usuario || !asesorSel) {
    return (
      <div style={{ fontFamily: "Montserrat, sans-serif", background: CREAM, minHeight: "100vh", padding: 40, display: "flex", alignItems: "center", justifyContent: "center" }}>
        <div style={{ background: "#fff", borderRadius: 12, padding: "40px 44px", width: 440, boxShadow: "0 1px 3px rgba(0,0,0,0.08)" }}>
          <h2 style={{ color: NAVY, margin: "0 0 6px", fontSize: 20 }}>Automatizador de propuestas</h2>
          <p style={{ color: "#78776f", fontSize: 13.5, margin: "0 0 26px" }}>Elegí quién sos y a qué asesor corresponde esta propuesta.</p>

          <Field label="¿Quién sos?">
            <select style={inputStyle} value={usuario} onChange={(e) => setUsuario(e.target.value)}>
              <option value="">Seleccioná tu nombre</option>
              {USUARIOS.map((u) => <option key={u} value={u}>{u}</option>)}
            </select>
          </Field>

          <Field label="Asesor / RepCode">
            <input style={inputStyle} value={asesorQuery} onChange={(e) => { setAsesorQuery(e.target.value); setAsesorSel(null); setRepcode(""); }} placeholder="Buscar por RepCode o nombre" />
          </Field>
          {asesorResultados.length > 0 && !asesorSel && (
            <div style={{ border: "1px solid #eae7dc", borderRadius: 6, marginTop: -8, marginBottom: 14, maxHeight: 180, overflowY: "auto" }}>
              {asesorResultados.map((a) => (
                <div key={a.repcode} onClick={() => { setAsesorSel(a); setRepcode(a.repcode); setAsesorQuery(`${a.repcode} — ${a.nombre}`); setAsesorResultados([]); }}
                  style={{ padding: "8px 10px", fontSize: 13, cursor: "pointer", borderBottom: "1px solid #f2f0e9" }}>
                  <b>{a.repcode}</b> — {a.nombre}
                </div>
              ))}
            </div>
          )}
          {asesorSel && (
            <div style={{ fontSize: 12.5, color: "#78776f", background: CREAM, borderRadius: 6, padding: "8px 10px", marginBottom: 18 }}>
              {asesorSel.email}
            </div>
          )}
        </div>
      </div>
    );
  }

  return (
    <div style={{ fontFamily: "Montserrat, sans-serif", background: CREAM, minHeight: "100vh" }}>
      <div style={{ background: NAVY, color: "#fff", padding: "14px 24px", display: "flex", justifyContent: "space-between" }}>
        <span style={{ fontSize: 14, fontWeight: 600 }}>Automatizador de propuestas</span>
        <span style={{ fontSize: 12.5, opacity: 0.85 }}>{usuario} · {repcode} — {asesorSel.nombre}</span>
      </div>

      <div style={{ display: "flex" }}>
        <div style={{ width: 210, padding: "24px 0", borderRight: "1px solid #e4e1d6" }}>
          {STEPS.map((s, i) => (
            <div key={s} onClick={() => setStep(i)} style={{
              padding: "10px 22px", fontSize: 13.5, cursor: "pointer",
              color: i === step ? NAVY : "#8c8b83", fontWeight: i === step ? 600 : 400,
              borderLeft: i === step ? `3px solid ${NAVY}` : "3px solid transparent",
              background: i === step ? "#fff" : "transparent",
            }}>{i + 1}. {s}</div>
          ))}
        </div>

        <div style={{ flex: 1, padding: "28px 36px", maxWidth: 720 }}>
          {step === 0 && (
            <Section title="Portada">
              <Field label="Tipo de documento">
                <div style={{ display: "flex", gap: 10 }}>
                  {["Propuesta", "Revision"].map((t) => (
                    <button key={t} onClick={() => setTipo(t)} style={{ padding: "8px 16px", borderRadius: 6, border: `1px solid ${tipo === t ? NAVY : "#d8d5cc"}`, background: tipo === t ? NAVY : "#fff", color: tipo === t ? "#fff" : "#333", cursor: "pointer" }}>{t === "Revision" ? "Revisión" : t}</button>
                  ))}
                </div>
              </Field>
              <Field label="Nombre del cliente">
                <input style={inputStyle} value={cliente} onChange={(e) => setCliente(e.target.value)} />
              </Field>
              {tipo === "Revision" && (
                <Field label="Número de cuenta StoneX">
                  <input style={inputStyle} value={nroCuenta} onChange={(e) => setNroCuenta(e.target.value)} />
                </Field>
              )}
              {tipo === "Propuesta" && (
                <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                  <input type="checkbox" checked={incluirPagina2} onChange={(e) => setIncluirPagina2(e.target.checked)} />
                  <span style={{ fontSize: 13.5 }}>Incluir página 2</span>
                </div>
              )}
            </Section>
          )}

          {step === 1 && (
            <Section title="Equipo">
              {team.map((m) => (
                <div key={m.id} style={{ display: "flex", alignItems: "center", gap: 10, padding: "8px 0" }}>
                  <input type="checkbox" checked={m.incluido} onChange={() => setTeam((prev) => prev.map((x) => x.id === m.id ? { ...x, incluido: !x.incluido } : x))} />
                  <div style={{ fontSize: 13.5 }}>{m.nombre} — <span style={{ color: "#78776f" }}>{m.puesto}</span></div>
                </div>
              ))}
            </Section>
          )}

          {step === 2 && (
            <Section title="Propuesta de valor">
              <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                <input type="checkbox" checked={incluirValueProp} onChange={(e) => setIncluirValueProp(e.target.checked)} />
                <span style={{ fontSize: 13.5 }}>Incluir esta página</span>
              </div>
            </Section>
          )}

          {step === 3 && (
            <Section title="Estrategia">
              <Field label="Perfil de riesgo">
                <div style={{ display: "flex", gap: 10 }}>
                  {Object.keys(RISK_TEXT).map((p) => (
                    <button key={p} onClick={() => setPerfil(p)} style={{ padding: "8px 16px", borderRadius: 6, border: `1px solid ${perfil === p ? NAVY : "#d8d5cc"}`, background: perfil === p ? NAVY : "#fff", color: perfil === p ? "#fff" : "#333", cursor: "pointer" }}>{p}</button>
                  ))}
                </div>
              </Field>
              <div style={{ background: "#fff", border: "1px solid #eae7dc", borderRadius: 8, padding: 16, fontSize: 13.5 }}>{RISK_TEXT[perfil]}</div>
            </Section>
          )}

          {step === 4 && (
            <Section title="Portafolio actual">
              <button onClick={() => setCurrentAssets((prev) => [...prev, { nombre: "", pct: 0, importe: 0 }])} style={{ marginBottom: 12, padding: "6px 12px", borderRadius: 6, border: "1px dashed #b8b5a9", background: "none", cursor: "pointer", fontSize: 12.5 }}>+ Agregar activo</button>
              {currentAssets.map((a, i) => (
                <div key={i} style={{ display: "flex", gap: 8, marginBottom: 6 }}>
                  <input style={{ ...inputStyle, flex: 2 }} placeholder="Nombre" value={a.nombre} onChange={(e) => setCurrentAssets((prev) => prev.map((x, j) => j === i ? { ...x, nombre: e.target.value } : x))} />
                  <input style={{ ...inputStyle, width: 70 }} type="number" placeholder="%" value={a.pct} onChange={(e) => setCurrentAssets((prev) => prev.map((x, j) => j === i ? { ...x, pct: +e.target.value } : x))} />
                  <input style={{ ...inputStyle, width: 120 }} type="number" placeholder="Importe" value={a.importe} onChange={(e) => setCurrentAssets((prev) => prev.map((x, j) => j === i ? { ...x, importe: +e.target.value } : x))} />
                </div>
              ))}
            </Section>
          )}

          {step === 5 && (
            <Section title="Portafolio propuesto">
              <Field label="Monto total a invertir (USD)">
                <input type="number" style={{ ...inputStyle, maxWidth: 220 }} value={montoInvertir} onChange={(e) => setMontoInvertir(+e.target.value)} />
              </Field>
              <Field label="Buscar fondo (ISIN o nombre) en la biblioteca de Supabase">
                <input style={inputStyle} value={fondoQuery} onChange={(e) => setFondoQuery(e.target.value)} placeholder="Ej: IE00B3XXRP09 o Vanguard" />
              </Field>
              {fondoResultados.length > 0 && (
                <div style={{ border: "1px solid #eae7dc", borderRadius: 6, marginBottom: 14, maxHeight: 200, overflowY: "auto" }}>
                  {fondoResultados.map((f) => (
                    <div key={f.isin} onClick={() => addProposedAsset(f)} style={{ padding: "8px 10px", fontSize: 12.5, cursor: "pointer", borderBottom: "1px solid #f2f0e9" }}>
                      <b>{f.isin}</b> — {f.nombre} {f.uso_frecuente && <span style={{ color: TEAL }}>★ frecuente</span>}
                    </div>
                  ))}
                </div>
              )}
              {proposedAssets.map((a, i) => (
                <div key={i} style={{ display: "flex", gap: 8, alignItems: "center", padding: "6px 0", borderBottom: "1px solid #eae7dc" }}>
                  <div style={{ flex: 1, fontSize: 13 }}>{a.nombre}</div>
                  <select style={{ ...inputStyle, width: 150 }} value={a.categoria} onChange={(e) => updateProposedField(i, "categoria", e.target.value)}>
                    {CATEGORIAS.map((c) => <option key={c} value={c}>{c}</option>)}
                  </select>
                  <input type="number" style={{ ...inputStyle, width: 60 }} value={a.pct} onChange={(e) => updateProposedField(i, "pct", +e.target.value)} />
                  <input type="number" style={{ ...inputStyle, width: 100 }} value={a.monto} onChange={(e) => updateProposedField(i, "monto", +e.target.value)} />
                </div>
              ))}
            </Section>
          )}

          {step === 6 && (
            <Section title="Comentarios">
              <textarea value={comentarios} onChange={(e) => setComentarios(e.target.value)} rows={8} style={{ ...inputStyle, resize: "vertical" }} />
            </Section>
          )}

          <div style={{ display: "flex", justifyContent: "space-between", marginTop: 24, paddingTop: 18, borderTop: "1px solid #eae7dc" }}>
            <button disabled={step === 0} onClick={() => setStep((s) => s - 1)} style={{ padding: "9px 18px", borderRadius: 6, border: "1px solid #d8d5cc", background: "#fff", opacity: step === 0 ? 0.4 : 1 }}>← Atrás</button>
            {step < STEPS.length - 1 ? (
              <button onClick={() => setStep((s) => s + 1)} style={{ padding: "9px 18px", borderRadius: 6, border: "none", background: NAVY, color: "#fff", fontWeight: 600 }}>Siguiente →</button>
            ) : (
              <button onClick={handleGenerar} disabled={generando} style={{ padding: "9px 20px", borderRadius: 6, border: "none", background: TEAL, color: "#fff", fontWeight: 600 }}>
                {generando ? "Generando…" : "Generar PPT + PDF"}
              </button>
            )}
          </div>

          {error && <div style={{ marginTop: 14, color: "#b23b3b", fontSize: 13 }}>{error}</div>}
          {resultado && (
            <div style={{ marginTop: 14, background: "#fff", border: "1px solid #eae7dc", borderRadius: 8, padding: 16 }}>
              <div style={{ fontSize: 13.5, fontWeight: 600, marginBottom: 8 }}>¡Listo!</div>
              <a href={resultado.pptx_url} target="_blank" rel="noreferrer" style={{ marginRight: 16, color: NAVY }}>Descargar PPTX</a>
              {resultado.pdf_url && <a href={resultado.pdf_url} target="_blank" rel="noreferrer" style={{ color: NAVY }}>Descargar PDF</a>}
            </div>
          )}
        </div>
      </div>
    </div>
  );
}
