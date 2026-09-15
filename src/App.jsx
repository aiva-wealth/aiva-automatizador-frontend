import React, { useState, useEffect } from "react";
import { supabase } from "./supabaseClient";
import * as XLSX from "xlsx";

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
  const [cashManualPropuesta, setCashManualPropuesta] = useState(0);
  const [cashActualPropuesta, setCashActualPropuesta] = useState(0);
  const [nuevoActivoNombre, setNuevoActivoNombre] = useState("");
  const [nuevoActivoIsin, setNuevoActivoIsin] = useState("");
  const [comentarios, setComentarios] = useState("");

  const [evolucionFileName, setEvolucionFileName] = useState("");
  const [evolucionImageBase64, setEvolucionImageBase64] = useState(null);
  const [evolucionInputKey, setEvolucionInputKey] = useState(0);

  const [generando, setGenerando] = useState(false);
  const [resultado, setResultado] = useState(null);
  const [error, setError] = useState("");

  const [vista, setVista] = useState("nueva"); // "nueva" | "registro" | "biblioteca"
  const [registro, setRegistro] = useState([]);
  const [registroCargando, setRegistroCargando] = useState(false);

  const [bibliotecaQuery, setBibliotecaQuery] = useState("");
  const [bibliotecaResultados, setBibliotecaResultados] = useState([]);
  const [bibliotecaSel, setBibliotecaSel] = useState(null);
  const [bibliotecaLogoFile, setBibliotecaLogoFile] = useState(null);
  const [bibliotecaGuardando, setBibliotecaGuardando] = useState(false);
  const [bibliotecaMensaje, setBibliotecaMensaje] = useState("");

  // --- Importación masiva de biblioteca (desde el paquete extraído de un PDF/PPTX) ---
  const [importModo, setImportModo] = useState(false);
  const [importEntradas, setImportEntradas] = useState([]);
  const [importLogos, setImportLogos] = useState({});
  const [importPreparando, setImportPreparando] = useState(false);
  const [importAplicando, setImportAplicando] = useState(false);
  const [importResumen, setImportResumen] = useState("");

  async function cargarRegistro() {
    setRegistroCargando(true);
    const { data } = await supabase
      .from("propuestas")
      .select("id, tipo, creado_por, repcode, cliente, nro_cuenta, monto, status, archivo_pptx_url, archivo_pdf_url, created_at")
      .order("created_at", { ascending: false })
      .limit(100);
    setRegistro(data || []);
    setRegistroCargando(false);
  }

  async function cambiarStatus(id, nuevoStatus) {
    await supabase.from("propuestas").update({ status: nuevoStatus }).eq("id", id);
    setRegistro((prev) => prev.map((r) => r.id === id ? { ...r, status: nuevoStatus } : r));
  }

  useEffect(() => {
    if (vista === "registro") cargarRegistro();
  }, [vista]);

  // --- Biblioteca de fondos ---
  useEffect(() => {
    if (vista !== "biblioteca" || bibliotecaQuery.trim().length < 2) { setBibliotecaResultados([]); return; }
    const t = setTimeout(async () => {
      const { data } = await supabase
        .from("fondos")
        .select("isin, nombre, sector, categoria, uso_frecuente, descripcion, factsheet_url, logo_url")
        .or(`isin.ilike.%${bibliotecaQuery}%,nombre.ilike.%${bibliotecaQuery}%`)
        .limit(15);
      setBibliotecaResultados(data || []);
    }, 250);
    return () => clearTimeout(t);
  }, [bibliotecaQuery, vista]);

  function seleccionarFondoBiblioteca(f) {
    setBibliotecaSel({ ...f });
    setBibliotecaLogoFile(null);
    setBibliotecaMensaje("");
  }

  async function guardarFondoBiblioteca() {
    setBibliotecaGuardando(true);
    setBibliotecaMensaje("");
    try {
      let logo_url = bibliotecaSel.logo_url;
      if (bibliotecaLogoFile) {
        const ext = bibliotecaLogoFile.name.split(".").pop();
        const path = `${bibliotecaSel.isin}.${ext}`;
        const { error: uploadError } = await supabase.storage.from("logos-fondos").upload(path, bibliotecaLogoFile, { upsert: true });
        if (uploadError) throw uploadError;
        const { data } = supabase.storage.from("logos-fondos").getPublicUrl(path);
        logo_url = data.publicUrl;
      }
      await supabase.from("fondos").update({
        descripcion: bibliotecaSel.descripcion || null,
        factsheet_url: bibliotecaSel.factsheet_url || null,
        logo_url,
      }).eq("isin", bibliotecaSel.isin);
      setBibliotecaSel((prev) => ({ ...prev, logo_url }));
      setBibliotecaMensaje("✓ Guardado — queda así para todas las próximas propuestas.");
    } catch (e) {
      setBibliotecaMensaje("Error al guardar: " + (e.message || e));
    } finally {
      setBibliotecaGuardando(false);
    }
  }

  // --- Importación masiva ---
  async function handleImportJson(file) {
    if (!file) return;
    const texto = await file.text();
    const data = JSON.parse(texto);
    setImportPreparando(true);
    const conMatch = [];
    for (const entry of data) {
      let candidatos = [];
      if (entry.isin_detectado) {
        const { data: exacto } = await supabase.from("fondos").select("isin, nombre").eq("isin", entry.isin_detectado);
        candidatos = exacto || [];
      }
      if (candidatos.length === 0) {
        const primeras = entry.nombre.split(" ").slice(0, 2).join(" ");
        const { data: porNombre } = await supabase.from("fondos").select("isin, nombre").ilike("nombre", `%${primeras}%`).limit(6);
        candidatos = porNombre || [];
      }
      conMatch.push({
        ...entry,
        candidatos,
        isins_elegidos: candidatos.length === 1 ? [candidatos[0].isin] : [],
        omitir: candidatos.length === 0,
      });
    }
    setImportEntradas(conMatch);
    setImportPreparando(false);
  }

  function handleImportLogoFiles(fileList) {
    const mapa = {};
    for (const f of fileList) mapa[f.name] = f;
    setImportLogos(mapa);
  }

  function actualizarImportEntrada(idx, campo, valor) {
    setImportEntradas((prev) => prev.map((e, i) => i === idx ? { ...e, [campo]: valor } : e));
  }

  async function aplicarImportacion() {
    setImportAplicando(true);
    let aplicados = 0, saltados = 0, errores = 0;
    for (const entry of importEntradas) {
      if (entry.omitir || !entry.isins_elegidos || entry.isins_elegidos.length === 0) { saltados++; continue; }
      // el mismo logo/descripción/factsheet puede aplicar a varios fondos
      // de una misma familia (ej: una descripción genérica de "Money
      // Market" que cubre a todos los fondos de esa categoría)
      let logo_url = undefined;
      try {
        const nombreArchivo = entry.logo_file ? entry.logo_file.split("/").pop() : null;
        const archivo = nombreArchivo ? importLogos[nombreArchivo] : null;
        if (archivo) {
          const ext = archivo.name.split(".").pop();
          // se sube una sola vez y se reutiliza el mismo link para todos los ISIN de este grupo
          const path = `${entry.isins_elegidos[0]}.${ext}`;
          await supabase.storage.from("logos-fondos").upload(path, archivo, { upsert: true });
          const { data } = supabase.storage.from("logos-fondos").getPublicUrl(path);
          logo_url = data.publicUrl;
        }
      } catch (e) {
        errores++;
        continue;
      }
      for (const isin of entry.isins_elegidos) {
        try {
          const update = { descripcion: entry.descripcion || null, factsheet_url: entry.factsheet_url || null };
          if (logo_url) update.logo_url = logo_url;
          await supabase.from("fondos").update(update).eq("isin", isin);
          aplicados++;
        } catch (e) {
          errores++;
        }
      }
    }
    setImportResumen(`${aplicados} fondos actualizados, ${saltados} grupos omitidos, ${errores} con error.`);
    setImportAplicando(false);
  }

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
        .select("isin, nombre, sector, categoria, ter, uso_frecuente, descripcion, factsheet_url, logo_url")
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

  async function toggleFavorito(idx) {
    const asset = proposedAssets[idx];
    const nuevoValor = !asset.uso_frecuente;
    // se guarda en Supabase de una — la próxima vez que se busque este
    // fondo (en esta propuesta o en cualquier otra) ya va a aparecer
    // marcado como frecuente
    await supabase.from("fondos").update({ uso_frecuente: nuevoValor }).eq("isin", asset.isin);
    setProposedAssets((prev) => prev.map((a, i) => i === idx ? { ...a, uso_frecuente: nuevoValor } : a));
  }

  // Agrega un activo que NO está en la biblioteca (una acción, un bono,
  // cualquier cosa) — lo guarda en Supabase (upsert por ISIN) para que a
  // partir de ahora quede buscable como cualquier otro fondo, y lo suma a
  // esta propuesta.
  async function agregarActivoNuevo() {
    if (!nuevoActivoNombre.trim() || !nuevoActivoIsin.trim()) return;
    const nuevoFondo = { isin: nuevoActivoIsin.trim(), nombre: nuevoActivoNombre.trim(), uso_frecuente: false };
    await supabase.from("fondos").upsert(nuevoFondo, { onConflict: "isin" });
    addProposedAsset(nuevoFondo);
    setNuevoActivoNombre("");
    setNuevoActivoIsin("");
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

  function quitarEvolucionFile() {
    setEvolucionFileName("");
    setEvolucionImageBase64(null);
    setEvolucionInputKey((k) => k + 1);
  }

  // Importa el excel "Open Tax Lots" de StoneX para el portafolio actual de
  // una Propuesta nueva (a diferencia de Revisión, acá solo hace falta
  // nombre + importe — el % se calcula solo sobre el total).
  async function handleExcelImportPropuestaActual(file) {
    if (!file) return;
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: "array" });
    const sheetName = wb.SheetNames.includes("By Security") ? "By Security" : wb.SheetNames[0];
    const ws = wb.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(ws, { defval: null });
    const nuevos = rows
      .map((r) => {
        const cantidad = Number(r["Quantity"]) || 0;
        const usdPrice = Number(r["USD Price"]) || 0;
        const importe = Math.round(Number(r["Mkt Value"]) || (usdPrice * cantidad) || 0);
        return { nombre: r["Description"] || "", importe };
      })
      .filter((a) => a.nombre);
    setCurrentAssets((prev) => [...prev, ...nuevos]);
  }

  // Importa el excel "Open Tax Lots" de StoneX (hoja "By Security") para el
  // portafolio actual de Revisión — Symbol/ID, Description, Adjusted Cost y
  // Mkt Value son exactamente lo que necesitamos.
  async function handleExcelImport(file) {
    if (!file) return;
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: "array" });
    const sheetName = wb.SheetNames.includes("By Security") ? "By Security" : wb.SheetNames[0];
    const ws = wb.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(ws, { defval: null });
    const nuevos = rows
      .map((r) => {
        const cantidad = Number(r["Quantity"]) || 0;
        const unitCost = Number(r["Unit Cost"]) || 0;
        const usdPrice = Number(r["USD Price"]) || 0;
        const costo = Number(r["Adjusted Cost"]) || (unitCost * cantidad) || 0;
        const valor_actual = Number(r["Mkt Value"]) || (usdPrice * cantidad) || 0;
        return {
          isin: r["Symbol/ID"] || "",
          nombre: r["Description"] || "",
          categoria: "Renta Variable",
          costo,
          valor_actual,
          precio_unidad: "",
          cantidad: "",
        };
      })
      .filter((a) => a.isin || a.nombre);
    setCurrentAssets((prev) => [...prev, ...nuevos]);
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

    const cashMonto = Number(cashManualPropuesta) || 0;

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
      "Renta Fija & Multi Activo": proposedAssets.filter((a) => a.categoria === "Renta Fija" || a.categoria === "Multi Activo").map((a) => ({ nombre: a.nombre, descripcion: a.descripcion || "", factsheet_url: a.factsheet_url || "", logo_url: a.logo_url || "" })),
      "Renta Variable": proposedAssets.filter((a) => a.categoria === "Renta Variable").map((a) => ({ nombre: a.nombre, descripcion: a.descripcion || "", factsheet_url: a.factsheet_url || "", logo_url: a.logo_url || "" })),
      "Alternativos Líquidos": proposedAssets.filter((a) => a.categoria === "Alternativos Líquidos").map((a) => ({ nombre: a.nombre, descripcion: a.descripcion || "", factsheet_url: a.factsheet_url || "", logo_url: a.logo_url || "" })),
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
      incluir_pagina2: incluirPagina2, incluir_valor: incluirValueProp,
      // el template de Propuesta solo tiene 4 casilleros armados para el
      // equipo (no incluye a Belén); el de Revisión sí tiene 5. Se recorta
      // acá para no depender de que el usuario se acuerde de destildar a
      // alguien.
      equipo: team.filter((m) => m.incluido).slice(0, 5).map((m) => ({ nombre: m.nombre, puesto: m.puesto, educacion: m.educacion })),
      perfil_riesgo: perfil,
      portafolio_actual: tipo === "Revision"
        ? revisionAssets.map((a) => ({ isin: a.isin, nombre: a.nombre, pct: a.pct, costo: a.costo, valor_actual: a.valor_actual, rendimiento: a.rendimiento }))
        : (() => {
            const total = currentAssets.reduce((s, a) => s + (Number(a.importe) || 0), 0) + (Number(cashActualPropuesta) || 0);
            const filas = currentAssets.map((a) => ({
              nombre: a.nombre,
              importe: Math.round(Number(a.importe) || 0),
              pct: total ? Math.round((Number(a.importe) || 0) / total * 100) : 0,
            }));
            if (cashActualPropuesta) {
              filas.push({ nombre: "Cash", importe: Math.round(Number(cashActualPropuesta)), pct: total ? Math.round(Number(cashActualPropuesta) / total * 100) : 0 });
            }
            return filas;
          })(),
      cash_valor: Number(cashValorRevision || 0),
      columnas_visibles: ["pct", "isin", "nombre", "costo", "valor_actual", "rendimiento"],
      asset_allocation: assetAllocationRevision,
      evolucion_image_base64: evolucionImageBase64,
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

  const totalRevisionPreview = currentAssets.reduce((s, a) => s + (Number(a.valor_actual) || 0), 0) + Number(cashValorRevision || 0);

  return (
    <div style={{ fontFamily: "Montserrat, sans-serif", background: CREAM, minHeight: "100vh" }}>
      <div style={{ background: NAVY, color: "#fff", padding: "14px 24px", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
        <span style={{ fontSize: 14, fontWeight: 600 }}>Automatizador de propuestas</span>
        <div style={{ display: "flex", alignItems: "center", gap: 18 }}>
          <button onClick={() => setVista("nueva")} style={{ background: "none", border: "none", color: vista === "nueva" ? "#fff" : "#9fb0c9", fontWeight: vista === "nueva" ? 600 : 400, cursor: "pointer", fontSize: 13 }}>Nueva propuesta</button>
          <button onClick={() => setVista("registro")} style={{ background: "none", border: "none", color: vista === "registro" ? "#fff" : "#9fb0c9", fontWeight: vista === "registro" ? 600 : 400, cursor: "pointer", fontSize: 13 }}>Registro</button>
          <button onClick={() => setVista("biblioteca")} style={{ background: "none", border: "none", color: vista === "biblioteca" ? "#fff" : "#9fb0c9", fontWeight: vista === "biblioteca" ? 600 : 400, cursor: "pointer", fontSize: 13 }}>Biblioteca de fondos</button>
          <span style={{ fontSize: 12.5, opacity: 0.85 }}>{usuario} · {repcode} — {asesorSel.nombre}</span>
        </div>
      </div>

      {vista === "biblioteca" ? (
        importModo ? (
          <div style={{ padding: "28px 36px" }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
              <h3 style={{ color: NAVY, fontSize: 16, margin: 0 }}>Importar biblioteca masiva</h3>
              <button onClick={() => setImportModo(false)} style={{ border: "none", background: "none", color: "#78776f", fontSize: 12.5, cursor: "pointer" }}>← Volver a la biblioteca</button>
            </div>
            <p style={{ fontSize: 12.5, color: "#78776f", marginBottom: 16 }}>Subí el JSON extraído y las imágenes de los logos. Matcheamos automático por ISIN o por nombre — revisá los casos dudosos antes de aplicar.</p>

            <div style={{ display: "flex", gap: 16, marginBottom: 20 }}>
              <Field label="Archivo JSON de extracción">
                <input type="file" accept=".json" onChange={(e) => handleImportJson(e.target.files[0])} style={{ ...inputStyle, padding: "8px" }} />
              </Field>
              <Field label="Imágenes de logo (seleccioná todas juntas)">
                <input type="file" accept="image/*" multiple onChange={(e) => handleImportLogoFiles(e.target.files)} style={{ ...inputStyle, padding: "8px" }} />
              </Field>
            </div>

            {importPreparando && <div style={{ fontSize: 13, color: "#78776f" }}>Buscando coincidencias contra la biblioteca…</div>}

            {importEntradas.length > 0 && !importPreparando && (
              <>
                <div style={{ fontSize: 12.5, marginBottom: 12, color: "#78776f" }}>
                  {importEntradas.length} fondos leídos — {importEntradas.filter(e => e.candidatos.length === 1).length} con match automático, {importEntradas.filter(e => e.candidatos.length !== 1 && !e.omitir).length} para revisar.
                </div>
                <div style={{ maxHeight: 480, overflowY: "auto", background: "#fff", border: "1px solid #eae7dc", borderRadius: 8 }}>
                  {importEntradas.map((entry, i) => (
                    <div key={i} style={{ display: "grid", gridTemplateColumns: "1fr 1fr auto", gap: 10, padding: "10px 14px", borderBottom: "1px solid #eae7dc", alignItems: "center", opacity: entry.omitir ? 0.45 : 1 }}>
                      <div>
                        <div style={{ fontSize: 13 }}>{entry.nombre}</div>
                        <div style={{ fontSize: 11, color: "#a5a399" }}>{entry.categoria_pptx} {entry.isin_detectado ? `· ISIN: ${entry.isin_detectado}` : ""}{entry.logo_file ? " · con logo" : " · sin logo"}</div>
                      </div>
                      <div style={{ display: "flex", flexDirection: "column", gap: 4, maxHeight: 90, overflowY: "auto" }}>
                        {entry.candidatos.length === 0 && <div style={{ fontSize: 11.5, color: "#a5a399" }}>Sin coincidencias en la biblioteca</div>}
                        {entry.candidatos.map((c) => (
                          <label key={c.isin} style={{ fontSize: 11.5, display: "flex", alignItems: "center", gap: 5 }}>
                            <input
                              type="checkbox"
                              checked={entry.isins_elegidos?.includes(c.isin) || false}
                              onChange={(e) => {
                                const actuales = entry.isins_elegidos || [];
                                const nuevos = e.target.checked ? [...actuales, c.isin] : actuales.filter((x) => x !== c.isin);
                                actualizarImportEntrada(i, "isins_elegidos", nuevos);
                              }}
                            />
                            {c.isin} — {c.nombre}
                          </label>
                        ))}
                      </div>
                      <label style={{ fontSize: 11.5, display: "flex", alignItems: "center", gap: 4 }}>
                        <input type="checkbox" checked={entry.omitir} onChange={(e) => actualizarImportEntrada(i, "omitir", e.target.checked)} /> Omitir
                      </label>
                    </div>
                  ))}
                </div>
                <button onClick={aplicarImportacion} disabled={importAplicando} style={{ marginTop: 16, padding: "10px 20px", borderRadius: 6, border: "none", background: NAVY, color: "#fff", fontWeight: 600, cursor: "pointer" }}>
                  {importAplicando ? "Aplicando…" : "Aplicar importación"}
                </button>
                {importResumen && <div style={{ marginTop: 10, fontSize: 13, color: "#3a7d44" }}>✓ {importResumen}</div>}
              </>
            )}
          </div>
        ) : (
        <div style={{ padding: "28px 36px", display: "flex", gap: 24 }}>
          <div style={{ width: 360 }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
              <h3 style={{ color: NAVY, fontSize: 16, marginBottom: 8 }}>Biblioteca de fondos</h3>
              <button onClick={() => setImportModo(true)} style={{ border: "none", background: "none", color: TEAL, fontSize: 12, cursor: "pointer", textDecoration: "underline" }}>Importar masivo</button>
            </div>
            <p style={{ fontSize: 12.5, color: "#78776f", marginBottom: 14 }}>Buscá un fondo para cargarle logo, descripción y factsheet — queda guardado para todas las próximas propuestas, no hay que repetirlo.</p>
            <input style={inputStyle} value={bibliotecaQuery} onChange={(e) => setBibliotecaQuery(e.target.value)} placeholder="Buscar por ISIN o nombre" />
            <div style={{ marginTop: 10, maxHeight: 480, overflowY: "auto" }}>
              {bibliotecaResultados.map((f) => (
                <div key={f.isin} onClick={() => seleccionarFondoBiblioteca(f)} style={{ padding: "8px 10px", fontSize: 12.5, cursor: "pointer", borderBottom: "1px solid #eae7dc", background: bibliotecaSel?.isin === f.isin ? "#fff" : "transparent" }}>
                  <b>{f.isin}</b> — {f.nombre} {f.logo_url && <span style={{ color: TEAL }}>✓ logo</span>}
                </div>
              ))}
            </div>
          </div>

          {bibliotecaSel && (
            <div style={{ flex: 1, maxWidth: 480, background: "#fff", border: "1px solid #eae7dc", borderRadius: 8, padding: 20 }}>
              <div style={{ fontSize: 14, fontWeight: 600, marginBottom: 4 }}>{bibliotecaSel.nombre}</div>
              <div style={{ fontSize: 12, color: "#78776f", marginBottom: 16 }}>{bibliotecaSel.isin}</div>

              {bibliotecaSel.logo_url && (
                <img src={bibliotecaSel.logo_url} alt="logo" style={{ maxHeight: 60, marginBottom: 12, display: "block" }} />
              )}
              <Field label="Logo (imagen)">
                <input type="file" accept="image/*" onChange={(e) => setBibliotecaLogoFile(e.target.files[0])} style={{ ...inputStyle, padding: "8px" }} />
              </Field>
              <Field label="Descripción">
                <textarea value={bibliotecaSel.descripcion || ""} onChange={(e) => setBibliotecaSel((prev) => ({ ...prev, descripcion: e.target.value }))} rows={4} style={{ ...inputStyle, resize: "vertical" }} />
              </Field>
              <Field label="Link al factsheet">
                <input style={inputStyle} value={bibliotecaSel.factsheet_url || ""} onChange={(e) => setBibliotecaSel((prev) => ({ ...prev, factsheet_url: e.target.value }))} placeholder="https://..." />
              </Field>
              <button onClick={guardarFondoBiblioteca} disabled={bibliotecaGuardando} style={{ padding: "9px 18px", borderRadius: 6, border: "none", background: NAVY, color: "#fff", fontWeight: 600, cursor: "pointer" }}>
                {bibliotecaGuardando ? "Guardando…" : "Guardar"}
              </button>
              {bibliotecaMensaje && <div style={{ marginTop: 10, fontSize: 12.5, color: bibliotecaMensaje.startsWith("Error") ? "#b23b3b" : "#3a7d44" }}>{bibliotecaMensaje}</div>}
            </div>
          )}
        </div>
        )
      ) : vista === "registro" ? (
        <div style={{ padding: "28px 36px" }}>
          <h3 style={{ color: NAVY, fontSize: 16, marginBottom: 16 }}>Registro de propuestas</h3>
          {registroCargando && <div style={{ fontSize: 13, color: "#78776f" }}>Cargando…</div>}
          {!registroCargando && registro.length === 0 && <div style={{ fontSize: 13, color: "#78776f" }}>Todavía no hay propuestas generadas.</div>}
          {!registroCargando && registro.length > 0 && (
            <table style={{ width: "100%", borderCollapse: "collapse", background: "#fff", fontSize: 13 }}>
              <thead>
                <tr style={{ background: NAVY, color: "#fff" }}>
                  {["Fecha", "Tipo", "Cliente / Cuenta", "Hecha por", "Asesor", "Monto", "Status", "Archivos"].map((h) => (
                    <th key={h} style={{ padding: "8px 10px", textAlign: "left" }}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {registro.map((r) => (
                  <tr key={r.id} style={{ borderBottom: "1px solid #eae7dc" }}>
                    <td style={{ padding: "8px 10px" }}>{new Date(r.created_at).toLocaleDateString()}</td>
                    <td style={{ padding: "8px 10px" }}>{r.tipo === "Revision" ? "Revisión" : "Propuesta"}</td>
                    <td style={{ padding: "8px 10px" }}>{r.tipo === "Revision" ? (r.nro_cuenta || "—") : (r.cliente || "—")}</td>
                    <td style={{ padding: "8px 10px" }}>{r.creado_por}</td>
                    <td style={{ padding: "8px 10px" }}>{r.repcode}</td>
                    <td style={{ padding: "8px 10px" }}>{r.monto ? `$${Number(r.monto).toLocaleString()}` : "—"}</td>
                    <td style={{ padding: "8px 10px" }}>
                      <select value={r.status} onChange={(e) => cambiarStatus(r.id, e.target.value)} style={{ ...miniInputStyle, padding: "4px 6px" }}>
                        <option value="en_proceso">En proceso</option>
                        <option value="enviada">Enviada</option>
                        <option value="confirmada">Confirmada</option>
                        <option value="invertida">Invertida</option>
                      </select>
                    </td>
                    <td style={{ padding: "8px 10px" }}>
                      {r.archivo_pptx_url && <a href={r.archivo_pptx_url} target="_blank" rel="noreferrer" style={{ color: NAVY, marginRight: 10 }}>PPTX</a>}
                      {r.archivo_pdf_url && <a href={r.archivo_pdf_url} target="_blank" rel="noreferrer" style={{ color: NAVY }}>PDF</a>}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          )}
        </div>
      ) : (
      <div style={{ display: "flex" }}>
        <div style={{ width: 210, padding: "24px 0", borderRight: "1px solid #e4e1d6" }}>
          {currentSteps.map((s, i) => (
            <div key={s} onClick={() => setStep(i)} style={{
              padding: "10px 22px", fontSize: 13.5, cursor: "pointer",
              color: i === step ? NAVY : "#8c8b83", fontWeight: i === step ? 600 : 400,
              borderLeft: i === step ? `3px solid ${NAVY}` : "3px solid transparent",
              background: i === step ? "#fff" : "transparent",
            }}>{i + 1}. {s}</div>
          ))}
        </div>

        <div style={{ flex: 1, padding: "28px 36px", maxWidth: 720 }}>
          {stepName === "Portada" && (
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

          {stepName === "Equipo" && (
            <Section title="Equipo" subtitle="El template tiene 5 lugares armados para el equipo.">
              {team.map((m) => (
                <div key={m.id} style={{ display: "flex", alignItems: "center", gap: 10, padding: "8px 0" }}>
                  <input type="checkbox" checked={m.incluido} onChange={() => setTeam((prev) => prev.map((x) => x.id === m.id ? { ...x, incluido: !x.incluido } : x))} />
                  <div style={{ fontSize: 13.5 }}>{m.nombre} — <span style={{ color: "#78776f" }}>{m.puesto}</span></div>
                </div>
              ))}
            </Section>
          )}

          {stepName === "Propuesta de valor" && (
            <Section title="Propuesta de valor">
              <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                <input type="checkbox" checked={incluirValueProp} onChange={(e) => setIncluirValueProp(e.target.checked)} />
                <span style={{ fontSize: 13.5 }}>Incluir esta página</span>
              </div>
            </Section>
          )}

          {stepName === "Estrategia" && (
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

          {stepName === "Evolución de la cuenta" && (
            <Section title="Evolución de la cuenta" subtitle="Todavía no tenemos conectado el circuito de datos reales de PowerBI, así que esta página se arma a partir de una imagen que subís vos (una captura del resumen + gráfico de esta cuenta). Si no subís nada, esta página directamente no va a aparecer en el documento final.">
              <Field label="Imagen de evolución de la cuenta (PNG o JPG)">
                <input
                  key={evolucionInputKey}
                  type="file"
                  accept="image/*"
                  onChange={(e) => handleEvolucionFile(e.target.files[0])}
                  style={{ ...inputStyle, padding: "8px" }}
                />
              </Field>
              {evolucionFileName && (
                <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 8 }}>
                  <div style={{ fontSize: 12.5, color: TEAL }}>✓ {evolucionFileName} — se va a incluir la página</div>
                  <button onClick={quitarEvolucionFile} style={{ border: "none", background: "none", color: "#b23b3b", fontSize: 12, cursor: "pointer", padding: 0, textDecoration: "underline" }}>Quitar archivo</button>
                </div>
              )}
              {!evolucionFileName && (
                <div style={{ fontSize: 12.5, color: "#a5a399" }}>Sin archivo subido — esta página no va a aparecer en el documento.</div>
              )}
            </Section>
          )}

          {stepName === "Portafolio actual" && tipo === "Revision" && (
            <Section title="Portafolio actual" subtitle="Cargá cada activo de la cuenta. El % y el rendimiento se calculan solos a partir del costo y el valor actual — no hace falta tipearlos.">
              <Field label="Cash / equivalentes (USD)" hint="Lo que está en efectivo o cuasi-efectivo, no en un activo puntual.">
                <input type="number" style={{ ...inputStyle, maxWidth: 220 }} value={cashValorRevision} onChange={(e) => setCashValorRevision(+e.target.value)} />
              </Field>

              <Field label="Importar desde Excel (Open Tax Lots de StoneX)" hint="Toma Symbol/ID, Description, Adjusted Cost y Mkt Value de la hoja 'By Security' y agrega una fila por activo.">
                <input type="file" accept=".xlsx,.xls" onChange={(e) => handleExcelImport(e.target.files[0])} style={{ ...inputStyle, padding: "8px" }} />
              </Field>

              <button onClick={() => setCurrentAssets((prev) => [...prev, { isin: "", nombre: "", categoria: "Renta Variable", costo: 0, valor_actual: 0, precio_unidad: "", cantidad: "" }])} style={{ marginBottom: 14, padding: "6px 12px", borderRadius: 6, border: "1px dashed #b8b5a9", background: "none", cursor: "pointer", fontSize: 12.5 }}>+ Agregar activo a mano</button>

              {currentAssets.map((a, i) => {
                const rendimiento = a.costo ? (((a.valor_actual - a.costo) / a.costo) * 100).toFixed(1) : "0.0";
                const pct = totalRevisionPreview ? ((a.valor_actual / totalRevisionPreview) * 100).toFixed(1) : "0.0";
                return (
                  <div key={i} style={{ background: "#fff", border: "1px solid #eae7dc", borderRadius: 8, padding: 12, marginBottom: 10 }}>
                    <div style={{ display: "grid", gridTemplateColumns: "1fr 2fr 1fr", gap: 10, marginBottom: 8 }}>
                      <MiniField label="ISIN">
                        <input style={miniInputStyle} value={a.isin} onChange={(e) => setCurrentAssets((prev) => prev.map((x, j) => j === i ? { ...x, isin: e.target.value } : x))} />
                      </MiniField>
                      <MiniField label="Nombre del fondo">
                        <input style={miniInputStyle} value={a.nombre} onChange={(e) => setCurrentAssets((prev) => prev.map((x, j) => j === i ? { ...x, nombre: e.target.value } : x))} />
                      </MiniField>
                      <MiniField label="Categoría">
                        <select style={miniInputStyle} value={a.categoria} onChange={(e) => setCurrentAssets((prev) => prev.map((x, j) => j === i ? { ...x, categoria: e.target.value } : x))}>
                          <option value="Renta Fija">Renta Fija</option>
                          <option value="Renta Variable">Renta Variable</option>
                        </select>
                      </MiniField>
                    </div>
                    <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr 1fr", gap: 10, alignItems: "end", marginBottom: 8 }}>
                      <MiniField label="Precio unidad (opcional)">
                        <input type="number" style={miniInputStyle} value={a.precio_unidad} onChange={(e) => {
                          const precio_unidad = e.target.value;
                          setCurrentAssets((prev) => prev.map((x, j) => {
                            if (j !== i) return x;
                            const cantidad = +x.cantidad || 0;
                            const costo = precio_unidad && cantidad ? +precio_unidad * cantidad : x.costo;
                            return { ...x, precio_unidad, costo };
                          }));
                        }} />
                      </MiniField>
                      <MiniField label="Cantidad (opcional)">
                        <input type="number" style={miniInputStyle} value={a.cantidad} onChange={(e) => {
                          const cantidad = e.target.value;
                          setCurrentAssets((prev) => prev.map((x, j) => {
                            if (j !== i) return x;
                            const precio_unidad = +x.precio_unidad || 0;
                            const costo = precio_unidad && cantidad ? precio_unidad * +cantidad : x.costo;
                            return { ...x, cantidad, costo };
                          }));
                        }} />
                      </MiniField>
                      <div style={{ gridColumn: "span 2", fontSize: 11, color: "#a5a399", paddingBottom: 8 }}>
                        Completá estos dos si querés que el costo se calcule solo. Si no, cargá el costo total directo abajo.
                      </div>
                    </div>
                    <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr 1fr", gap: 10, alignItems: "end" }}>
                      <MiniField label="Costo total (USD)">
                        <input type="number" style={miniInputStyle} value={a.costo} onChange={(e) => setCurrentAssets((prev) => prev.map((x, j) => j === i ? { ...x, costo: +e.target.value } : x))} />
                      </MiniField>
                      <MiniField label="Valor actual (USD)">
                        <input type="number" style={miniInputStyle} value={a.valor_actual} onChange={(e) => setCurrentAssets((prev) => prev.map((x, j) => j === i ? { ...x, valor_actual: +e.target.value } : x))} />
                      </MiniField>
                      <MiniField label="Rendimiento (calculado)">
                        <div style={{ ...miniInputStyle, background: CREAM, fontWeight: 600, color: rendimiento >= 0 ? "#3a7d44" : "#b23b3b" }}>{rendimiento}%</div>
                      </MiniField>
                      <MiniField label="% del portafolio (calculado)">
                        <div style={{ ...miniInputStyle, background: CREAM, fontWeight: 600 }}>{pct}%</div>
                      </MiniField>
                    </div>
                    <button onClick={() => setCurrentAssets((prev) => prev.filter((_, j) => j !== i))} style={{ marginTop: 8, border: "none", background: "none", color: "#b23b3b", fontSize: 12, cursor: "pointer", padding: 0 }}>Quitar activo</button>
                  </div>
                );
              })}
            </Section>
          )}

          {stepName === "Portafolio actual" && tipo === "Propuesta" && (
            <Section title="Portafolio actual" subtitle="El % se calcula solo sobre el total (activos + cash) — no hace falta tipearlo.">
              <Field label="Cash (USD)" hint="StoneX no incluye el efectivo en el excel de posiciones — hay que cargarlo aparte.">
                <input type="number" style={{ ...inputStyle, maxWidth: 220 }} value={cashActualPropuesta} onChange={(e) => setCashActualPropuesta(+e.target.value)} />
              </Field>
              <Field label="Importar desde Excel (Open Tax Lots de StoneX)" hint="Toma Description y Mkt Value de la hoja 'By Security' y agrega una fila por activo.">
                <input type="file" accept=".xlsx,.xls" onChange={(e) => handleExcelImportPropuestaActual(e.target.files[0])} style={{ ...inputStyle, padding: "8px" }} />
              </Field>
              <button onClick={() => setCurrentAssets((prev) => [...prev, { nombre: "", importe: 0 }])} style={{ marginBottom: 12, padding: "6px 12px", borderRadius: 6, border: "1px dashed #b8b5a9", background: "none", cursor: "pointer", fontSize: 12.5 }}>+ Agregar activo a mano</button>
              {(() => {
                const total = currentAssets.reduce((s, a) => s + (Number(a.importe) || 0), 0) + (Number(cashActualPropuesta) || 0);
                return currentAssets.map((a, i) => {
                  const pct = total ? Math.round((Number(a.importe) || 0) / total * 100) : 0;
                  return (
                    <div key={i} style={{ display: "grid", gridTemplateColumns: "2fr 1fr 1fr auto", gap: 10, marginBottom: 8, alignItems: "end" }}>
                      <MiniField label="Nombre">
                        <input style={miniInputStyle} value={a.nombre} onChange={(e) => setCurrentAssets((prev) => prev.map((x, j) => j === i ? { ...x, nombre: e.target.value } : x))} />
                      </MiniField>
                      <MiniField label="% (calculado)">
                        <div style={{ ...miniInputStyle, background: CREAM, fontWeight: 600 }}>{pct}%</div>
                      </MiniField>
                      <MiniField label="Importe USD">
                        <input style={miniInputStyle} type="number" value={a.importe} onChange={(e) => setCurrentAssets((prev) => prev.map((x, j) => j === i ? { ...x, importe: +e.target.value } : x))} />
                      </MiniField>
                      <button onClick={() => setCurrentAssets((prev) => prev.filter((_, j) => j !== i))} style={{ border: "none", background: "none", color: "#b23b3b", fontSize: 12, cursor: "pointer" }}>Quitar</button>
                    </div>
                  );
                });
              })()}
            </Section>
          )}

          {stepName === "Portafolio propuesto" && (
            <Section title="Portafolio propuesto" subtitle="Buscá fondos en la biblioteca, asigná % o monto (se calculan solos entre sí), y completá los rendimientos históricos si los tenés a mano.">
              <Field label="Monto total a invertir (USD)">
                <input type="number" style={{ ...inputStyle, maxWidth: 220 }} value={montoInvertir} onChange={(e) => setMontoInvertir(+e.target.value)} />
              </Field>
              <Field label="Cash (USD)" hint="La parte del monto que se deja en efectivo, sin invertir en ningún fondo.">
                <input type="number" style={{ ...inputStyle, maxWidth: 220 }} value={cashManualPropuesta} onChange={(e) => setCashManualPropuesta(+e.target.value)} />
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

              <div style={{ background: "#fff", border: "1px dashed #d8d5cc", borderRadius: 8, padding: 12, marginBottom: 18 }}>
                <div style={{ fontSize: 12, color: "#78776f", marginBottom: 8 }}>¿No está en la biblioteca? Puede ser cualquier cosa — una acción, un bono, una alternativa. Se agrega acá y queda guardado para la próxima vez.</div>
                <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr auto", gap: 8 }}>
                  <input style={miniInputStyle} placeholder="Nombre (ej: Apple Inc / AAPL)" value={nuevoActivoNombre} onChange={(e) => setNuevoActivoNombre(e.target.value)} />
                  <input style={miniInputStyle} placeholder="ISIN / Ticker" value={nuevoActivoIsin} onChange={(e) => setNuevoActivoIsin(e.target.value)} />
                  <button onClick={agregarActivoNuevo} style={{ padding: "0 16px", borderRadius: 6, border: "none", background: NAVY, color: "#fff", fontSize: 12.5, cursor: "pointer" }}>Agregar</button>
                </div>
              </div>

              {(() => {
                const sumaFondos = proposedAssets.reduce((s, a) => s + (Number(a.monto) || 0), 0);
                const falta = montoInvertir - sumaFondos - (Number(cashManualPropuesta) || 0);
                return (
                  <div style={{ fontSize: 12.5, marginBottom: 14, color: Math.abs(falta) < 1 ? "#3a7d44" : "#b23b3b" }}>
                    {Math.abs(falta) < 1 ? "✓ Asignado el 100% del monto." : falta > 0 ? `Falta asignar ${falta.toLocaleString()} USD para llegar al monto total.` : `Te pasaste por ${Math.abs(falta).toLocaleString()} USD del monto total.`}
                  </div>
                );
              })()}

              {proposedAssets.map((a, i) => (
                <div key={i} style={{ background: "#fff", border: "1px solid #eae7dc", borderRadius: 8, padding: 12, marginBottom: 10 }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 8 }}>
                    <div style={{ flex: 1, fontSize: 13, fontWeight: 600 }}>{a.nombre}</div>
                    <button onClick={() => toggleFavorito(i)} title="Marcar/desmarcar como fondo frecuente" style={{ border: "none", background: "none", cursor: "pointer", fontSize: 16, color: a.uso_frecuente ? TEAL : "#d8d5cc", padding: 0 }}>★</button>
                    <button onClick={() => setProposedAssets((prev) => prev.filter((_, j) => j !== i))} style={{ border: "none", background: "none", color: "#b23b3b", fontSize: 12, cursor: "pointer" }}>Quitar</button>
                  </div>
                  <div style={{ display: "grid", gridTemplateColumns: "1.3fr 0.8fr 1fr", gap: 10, marginBottom: 8 }}>
                    <MiniField label="Categoría">
                      <select style={miniInputStyle} value={a.categoria} onChange={(e) => updateProposedField(i, "categoria", e.target.value)}>
                        {CATEGORIAS.map((c) => <option key={c} value={c}>{c}</option>)}
                      </select>
                    </MiniField>
                    <MiniField label="% del monto">
                      <input type="number" style={miniInputStyle} value={a.pct} onChange={(e) => updateProposedField(i, "pct", +e.target.value)} />
                    </MiniField>
                    <MiniField label="Monto (USD)">
                      <input type="number" style={miniInputStyle} value={a.monto} onChange={(e) => updateProposedField(i, "monto", +e.target.value)} />
                    </MiniField>
                  </div>
                  <div style={{ fontSize: 10.5, color: "#9b9993", marginBottom: 4, textTransform: "uppercase", letterSpacing: 0.4 }}>Rendimientos históricos (%, opcional)</div>
                  <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr 1fr", gap: 10 }}>
                    <MiniField label="YTD">
                      <input type="number" style={miniInputStyle} value={a.ytd} onChange={(e) => updateProposedField(i, "ytd", +e.target.value)} />
                    </MiniField>
                    <MiniField label="1 año">
                      <input type="number" style={miniInputStyle} value={a.y1} onChange={(e) => updateProposedField(i, "y1", +e.target.value)} />
                    </MiniField>
                    <MiniField label="3 años">
                      <input type="number" style={miniInputStyle} value={a.y3} onChange={(e) => updateProposedField(i, "y3", +e.target.value)} />
                    </MiniField>
                    <MiniField label="5 años">
                      <input type="number" style={miniInputStyle} value={a.y5} onChange={(e) => updateProposedField(i, "y5", +e.target.value)} />
                    </MiniField>
                  </div>
                </div>
              ))}
            </Section>
          )}

          {stepName === "Comentarios" && (
            <Section title="Comentarios">
              <textarea value={comentarios} onChange={(e) => setComentarios(e.target.value)} rows={8} style={{ ...inputStyle, resize: "vertical" }} />
            </Section>
          )}

          <div style={{ display: "flex", justifyContent: "space-between", marginTop: 24, paddingTop: 18, borderTop: "1px solid #eae7dc" }}>
            <button disabled={step === 0} onClick={() => setStep((s) => s - 1)} style={{ padding: "9px 18px", borderRadius: 6, border: "1px solid #d8d5cc", background: "#fff", opacity: step === 0 ? 0.4 : 1 }}>← Atrás</button>
            {step < currentSteps.length - 1 ? (
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
      )}
    </div>
  );
}
