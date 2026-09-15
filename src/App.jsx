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

// Cada tipo de documento tiene su propio recorrido de pasos — Revisión no
// arma un portafolio propuesto (es sobre una cuenta que ya existe), pero sí
// tiene "Evolución de la cuenta", que Propuesta no tiene.
const STEPS_PROPUESTA = ["Portada", "Equipo", "Propuesta de valor", "Estrategia", "Portafolio actual", "Portafolio propuesto", "Comentarios"];
const STEPS_REVISION = ["Portada", "Equipo", "Propuesta de valor", "Estrategia", "Evolución de la cuenta", "Portafolio actual", "Comentarios"];

const DEFAULT_TEAM = [
  { id: 1, nombre: "Carmela Hernández", puesto: "Investment Specialist", educacion: "Lic. Economía, Candidate CFA III", incluido: true },
  { id: 2, nombre: "Gregorio Anza", puesto: "Dealing Desk Specialist", educacion: "Lic. Administración de Empresas", incluido: true },
  { id: 3, nombre: "Julieta Broggi", puesto: "Investment Analyst", educacion: "Contadora Pública, Bsc. in Finance, University of London", incluido: true },
  { id: 4, nombre: "Federico Cañette", puesto: "Investment Support", educacion: "Lic. Economía", incluido: true },
  { id: 5, nombre: "Belén Rodriguez", puesto: "Investment Execution Analyst", educacion: "Estudiante Lic. Economía", incluido: true },
];

const CATEGORIAS = ["Renta Fija", "Multi Activo", "Renta Variable", "Alternativos Líquidos"];

function Section({ title, subtitle, children }) {
  return (
    <div style={{ marginBottom: 28 }}>
      <h3 style={{ color: NAVY, fontSize: 13, fontWeight: 700, letterSpacing: 1, textTransform: "uppercase", marginBottom: subtitle ? 4 : 14 }}>
        {title}
      </h3>
      {subtitle && <p style={{ fontSize: 12.5, color: "#78776f", marginTop: 0, marginBottom: 16 }}>{subtitle}</p>}
      {children}
    </div>
  );
}

function Field({ label, children, hint }) {
  return (
    <div style={{ marginBottom: 14 }}>
      <label style={{ display: "block", fontSize: 12.5, color: "#5b5b55", marginBottom: 5 }}>{label}</label>
      {children}
      {hint && <div style={{ fontSize: 11, color: "#a5a399", marginTop: 3 }}>{hint}</div>}
    </div>
  );
}

// Campo chico con su etiqueta arriba, pensado para usarse en grillas de
// varias columnas (portafolio actual) sin depender de recordar qué va en
// cada casillero.
function MiniField({ label, children }) {
  return (
    <div>
      <div style={{ fontSize: 10.5, color: "#9b9993", marginBottom: 3, textTransform: "uppercase", letterSpacing: 0.4 }}>{label}</div>
      {children}
    </div>
  );
}

const inputStyle = {
  width: "100%", padding: "9px 11px", borderRadius: 6, border: "1px solid #d8d5cc",
  fontSize: 14, fontFamily: "inherit", boxSizing: "border-box", background: "#fff",
};

const miniInputStyle = { ...inputStyle, padding: "7px 9px", fontSize: 13 };

function fileToBase64(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result.split(",")[1]);
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

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
  const [cashValorRevision, setCashValorRevision] = useState(0);
  const [montoInvertir, setMontoInvertir] = useState(500000);
  const [fondoQuery, setFondoQuery] = useState("");
  const [fondoResultados, setFondoResultados] = useState([]);
  const [proposedAssets, setProposedAssets] = useState([]);
  const [comentarios, setComentarios] = useState("");

  const [evolucionFileName, setEvolucionFileName] = useState("");
  const [evolucionImageBase64, setEvolucionImageBase64] = useState(null);

  const [generando, setGenerando] = useState(false);
  const [resultado, setResultado] = useState(null);
  const [error, setError] = useState("");

  const currentSteps = tipo === "Revision" ? STEPS_REVISION : STEPS_PROPUESTA;
  const stepName = currentSteps[step];

  // si se cambia el tipo de documento a mitad de camino, el paso actual
  // puede dejar de existir en la lista nueva — volvemos al principio para
  // no quedar en un paso inválido
  useEffect(() => { setStep(0); }, [tipo]);

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

  async function handleEvolucionFile(file) {
    if (!file) { setEvolucionFileName(""); setEvolucionImageBase64(null); return; }
    setEvolucionFileName(file.name);
    const b64 = await fileToBase64(file);
    setEvolucionImageBase64(b64);
  }

  // rendimiento y % de portafolio para Revisión se calculan solos a partir
  // de costo/valor actual — no hace falta que nadie los tipee ni se
  // equivoque cargándolos a mano
  function activoConCalculos(a, totalValor) {
    const rendimiento = a.costo ? ((a.valor_actual - a.costo) / a.costo) * 100 : 0;
    const pct = totalValor ? (a.valor_actual / totalValor) * 100 : 0;
    return { ...a, rendimiento, pct };
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

    // --- Revisión: los % y rendimiento salen solos de costo/valor actual ---
    const totalRevision = currentAssets.reduce((s, a) => s + (Number(a.valor_actual) || 0), 0) + Number(cashValorRevision || 0);
    const revisionAssets = currentAssets.map((a) => activoConCalculos({
      ...a,
      valor_actual: Number(a.valor_actual) || 0,
      costo: Number(a.costo) || 0,
    }, totalRevision));

    const rf = revisionAssets.filter((a) => a.categoria === "Renta Fija").reduce((s, a) => s + a.valor_actual, 0);
    const rv = revisionAssets.filter((a) => a.categoria === "Renta Variable").reduce((s, a) => s + a.valor_actual, 0);
    const assetAllocationRevision = totalRevision
      ? { "Renta Fija": rf / totalRevision, "Renta Variable": rv / totalRevision, "Cash": Number(cashValorRevision || 0) / totalRevision }
      : { "Renta Fija": 0, "Renta Variable": 0, "Cash": 0 };

    return {
      cliente, nro_cuenta: nroCuenta,
      incluir_pagina2: incluirPagina2, incluir_valor:
