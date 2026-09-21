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
const STEPS_PROPUESTA = ["Portada", "Equipo", "Propuesta de valor", "Estrategia", "Portafolio actual", "Portafolio propuesto", "Descripción de activos", "Comentarios"];
const STEPS_REVISION = ["Portada", "Equipo", "Propuesta de valor", "Estrategia", "Evolución de la cuenta", "Portafolio actual", "Comentarios"];

const DEFAULT_TEAM = [
  { id: 1, nombre: "Carmela Hernández", puesto: "Investment Specialist", educacion: "Lic. Economía, Candidate CFA III", incluido: true },
  { id: 2, nombre: "Gregorio Anza", puesto: "Dealing Desk Specialist", educacion: "Lic. Administración de Empresas", incluido: true },
  { id: 3, nombre: "Julieta Broggi", puesto: "Investment Analyst", educacion: "Contadora Pública, Bsc. in Finance, University of London", incluido: true },
  { id: 4, nombre: "Federico Cañette", puesto: "Investment Support", educacion: "Lic. Economía", incluido: true },
  { id: 5, nombre: "Belén Rodriguez", puesto: "Investment Execution Analyst", educacion: "Estudiante Lic. Economía", incluido: true },
];

const CATEGORIAS = ["Renta Fija", "Multi Activo", "Renta Variable", "Alternativos Líquidos"];

// Columnas por tipo de instrumento — mismo set y mismo orden que arma
// engine_propuesto_extra.py en el PPTX, para que la tabla en pantalla sea
// un espejo real de lo que va a salir en el documento. "fijo: true" son
// columnas que siempre se muestran (no se pueden ocultar ni reordenar).
// "calculado: true" es de solo lectura, se deriva de otros campos.
const COLUMNAS_POR_TIPO = {
  fondo: [
    { key: "isin", label: "Código", tipo: "text", fijo: true },
    { key: "nombre", label: "Nombre", tipo: "text", fijo: true },
    { key: "sector", label: "Sector", tipo: "text" },
    { key: "ytd", label: "Rend. YTD", tipo: "number" },
    { key: "y1", label: "Rend. 1 año", tipo: "number" },
    { key: "y3", label: "Rend. 3 años", tipo: "number" },
    { key: "y5", label: "Rend. 5 años", tipo: "number" },
    { key: "ter", label: "TER", tipo: "number" },
  ],
  fondo_distributivo: [
    { key: "isin", label: "Código", tipo: "text", fijo: true },
    { key: "nombre", label: "Nombre", tipo: "text", fijo: true },
    { key: "sector", label: "Sector", tipo: "text" },
    { key: "dividendo_pct", label: "Dividendo (%)", tipo: "number" },
    { key: "frecuencia_dividendo", label: "Frec. Dividendo", tipo: "text" },
    { key: "ytd", label: "Rend. YTD", tipo: "number" },
    { key: "y1", label: "Rend. 1 año", tipo: "number" },
    { key: "y3", label: "Rend. 3 años", tipo: "number" },
    { key: "y5", label: "Rend. 5 años", tipo: "number" },
    { key: "dividendo_anual", label: "Dividendo Anual", tipo: "number", calculado: true },
    { key: "ter", label: "TER", tipo: "number" },
  ],
  accion: [
    { key: "isin", label: "Ticker", tipo: "text", fijo: true },
    { key: "nombre", label: "Nombre", tipo: "text", fijo: true },
    { key: "sector", label: "Sector", tipo: "text" },
    { key: "ytd", label: "Rend. YTD", tipo: "number" },
    { key: "y1", label: "Rend. 1 año", tipo: "number" },
    { key: "y3", label: "Rend. 3 años", tipo: "number" },
    { key: "y5", label: "Rend. 5 años", tipo: "number" },
  ],
  bono: [
    { key: "isin", label: "Código", tipo: "text", fijo: true },
    { key: "nombre", label: "Nombre", tipo: "text", fijo: true },
    { key: "sector", label: "Sector", tipo: "text" },
    { key: "cupon_pct", label: "Cupón (%)", tipo: "number" },
    { key: "rating", label: "Rating S&P", tipo: "text" },
    { key: "price", label: "Price", tipo: "number" },
    { key: "yield_pct", label: "Yield", tipo: "number" },
    { key: "maturity", label: "Maturity", tipo: "date" },
    { key: "cupon_anual", label: "Cupón Anual", tipo: "number", calculado: true },
  ],
};

const TIPO_LABELS = { fondo: "Fondos", fondo_distributivo: "Fondos distributivos", accion: "Acciones", bono: "Bonos" };
const TIPO_ORDEN = ["fondo", "fondo_distributivo", "accion", "bono"];

// Columnas para la tabla de revisión de la importación del Excel base —
// mismas que COLUMNAS_POR_TIPO pero sin los campos calculados (dividendo
// anual, cupón anual), que dependen de un monto de propuesta que acá no
// existe (esto es la biblioteca, no una propuesta puntual).
const LIBRERIA_COLUMNAS_POR_TIPO = {};
Object.keys(COLUMNAS_POR_TIPO).forEach((t) => {
  LIBRERIA_COLUMNAS_POR_TIPO[t] = COLUMNAS_POR_TIPO[t].filter((c) => !c.calculado);
});

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

// Nombre de archivo simple y estable a partir del nombre de marca (para el
// storage) — no necesita ser bonito, solo único y sin caracteres raros.
function slugify(texto) {
  return texto
    .toString()
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/(^-|-$)/g, "") || "marca";
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
  const [equipoExpandidoId, setEquipoExpandidoId] = useState(null);
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

  // --- Columnas visibles/orden por tipo de instrumento en "Portafolio
  // propuesto" (no se guarda en Supabase, es solo cómo se ve mientras se
  // arma esta propuesta puntual) — arranca con todas las columnas no-fijas
  // visibles, en el orden por defecto.
  const [columnasConfig, setColumnasConfig] = useState(() => {
    const init = {};
    Object.keys(COLUMNAS_POR_TIPO).forEach((t) => {
      init[t] = COLUMNAS_POR_TIPO[t].filter((c) => !c.fijo).map((c) => c.key);
    });
    return init;
  });
  const [columnasAbiertoPara, setColumnasAbiertoPara] = useState(null); // tipo cuyo panel de columnas está abierto

  // --- Descripción de activos: selección manual por categoría, ---
  // independiente de lo que se haya cargado en Portafolio propuesto
  const DESC_CATEGORIAS = ["Renta Fija & Multi Activo", "Renta Variable", "Alternativos Líquidos"];
  const [descSeleccion, setDescSeleccion] = useState({ "Renta Fija & Multi Activo": [], "Renta Variable": [], "Alternativos Líquidos": [] });
  const [descQuery, setDescQuery] = useState("");
  const [descResultados, setDescResultados] = useState([]);
  const [descCategoriaDestino, setDescCategoriaDestino] = useState("Renta Fija & Multi Activo");

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

  // --- Auditoría de logos: agrupa TODOS los fondos que tienen logo_url por
  // el logo real (comparando por URL), para detectar rápido casos como el
  // de Calamos/Morgan Stanley — un fondo con el logo de otro pegado por
  // error en la importación masiva (se agrupaban varios ISIN bajo un mismo
  // logo a mano, y ahí es fácil equivocarse de fila).
  const [logosVista, setLogosVista] = useState(false);
  const [logosCargando, setLogosCargando] = useState(false);
  const [logosGrupos, setLogosGrupos] = useState([]);
  const [logosQuery, setLogosQuery] = useState("");

  // --- Marcas de logo: biblioteca de logos reutilizables, independiente de
  // los fondos. Se buscan por nombre (ej: "MFS", "BlackRock") y se asignan
  // a cualquier fondo sin volver a subir el archivo — así un mismo logo
  // real nunca queda duplicado en el storage, y cargar un logo que falta
  // no depende de tener un fondo puntual a mano primero.
  const [marcaNombreNuevo, setMarcaNombreNuevo] = useState("");
  const [marcaArchivoNuevo, setMarcaArchivoNuevo] = useState(null);
  const [marcaGuardandoNueva, setMarcaGuardandoNueva] = useState(false);
  const [marcaMensajeNueva, setMarcaMensajeNueva] = useState("");
  const [marcasTodas, setMarcasTodas] = useState([]);
  const [marcaNombrePorGrupo, setMarcaNombrePorGrupo] = useState({}); // logo_url -> texto que se está tipeando
  const [marcaGuardandoGrupo, setMarcaGuardandoGrupo] = useState(""); // logo_url en proceso de guardarse

  // buscador de marca dentro del editor de un fondo puntual
  const [marcaFondoQuery, setMarcaFondoQuery] = useState("");
  const [marcaFondoResultados, setMarcaFondoResultados] = useState([]);
  const [marcaFondoAsignada, setMarcaFondoAsignada] = useState(""); // nombre de la marca recién elegida, solo para mostrar feedback

  // --- Importación de marcas en bloque: subir varios archivos de logo de
  // una — el nombre de marca sale del propio nombre de archivo (limpiando
  // el número/guiones que le puso el ZIP que armamos), editable antes de
  // aplicar por si hay que corregir alguno.
  const [marcaImportEntradas, setMarcaImportEntradas] = useState([]); // [{file, nombre}]
  const [marcaImportAplicando, setMarcaImportAplicando] = useState(false);
  const [marcaImportResumen, setMarcaImportResumen] = useState("");
  // --- Fila unificada expandible (reemplaza las dos secciones separadas de
  // antes: "marcas registradas" + "grupos por logo"). Una sola lista,
  // colapsada por defecto (solo logo + nombre); al abrir una fila se ve la
  // lista de fondos, se pueden agregar/quitar, y cambiar el logo.
  const [filaExpandidaKey, setFilaExpandidaKey] = useState(null);
  const [marcaLogoNuevoArchivo, setMarcaLogoNuevoArchivo] = useState(null);
  const [marcaLogoActualizando, setMarcaLogoActualizando] = useState(false);
  const [marcaLogoMensaje, setMarcaLogoMensaje] = useState("");
  const [marcaAsocQuery, setMarcaAsocQuery] = useState("");
  const [marcaAsocResultados, setMarcaAsocResultados] = useState([]);
  const [marcaAsocSeleccionados, setMarcaAsocSeleccionados] = useState({}); // isin -> nombre
  const [marcaAsocGuardando, setMarcaAsocGuardando] = useState(false);
  const [marcaAsocMensaje, setMarcaAsocMensaje] = useState("");
  const [marcaFusionSeleccion, setMarcaFusionSeleccion] = useState({}); // logo_url -> id de marca elegida
  const [marcaFusionando, setMarcaFusionando] = useState("");

  // --- Importación del Excel base de instrumentos (3 pestañas: Fondos,
  // Acciones, Bonos — Fondos distributivos NO es una pestaña aparte, es la
  // misma pestaña Fondos con 3 columnas extra opcionales, ver
  // descargarPlantillaBase) — cada pestaña define columnas distintas según
  // lo que necesita ese tipo de instrumento. Hace upsert directo por
  // ISIN/Ticker, no pide revisión fila por fila.
  const [baseImportCargando, setBaseImportCargando] = useState(false);
  const [baseImportResumen, setBaseImportResumen] = useState("");
  const [descargaConDividendos, setDescargaConDividendos] = useState(false);
  const [baseImportPreview, setBaseImportPreview] = useState([]); // [{tipo_instrumento, isin, nombre, sector, categoria, ...campos}] — revisable/editable antes de guardar
  const [propuestaImportResumen, setPropuestaImportResumen] = useState(""); // resultado de importar el Excel directo a ESTA propuesta (no toca la biblioteca)

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

  async function eliminarPropuesta(id) {
    if (!window.confirm("¿Eliminar este registro? Esta acción no se puede deshacer (no borra el PPTX/PDF ya generado en Storage, solo la fila del registro).")) return;
    await supabase.from("propuestas").delete().eq("id", id);
    setRegistro((prev) => prev.filter((r) => r.id !== id));
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
    setMarcaFondoQuery("");
    setMarcaFondoResultados([]);
    setMarcaFondoAsignada("");
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

  // --- Auditoría de logos: trae TODOS los fondos con logo_url cargado y
  // los agrupa por URL de logo — cada grupo debería ser un solo fondo (o
  // una familia real que comparte marca); si aparece un grupo con nombres
  // de fondos que no tienen nada que ver entre sí, ahí hay un error de
  // carga para corregir.
  async function cargarAuditoriaLogos() {
    setLogosCargando(true);
    const { data } = await supabase
      .from("fondos")
      .select("isin, nombre, logo_url")
      .not("logo_url", "is", null)
      .order("logo_url");
    const porLogo = new Map();
    (data || []).forEach((f) => {
      if (!porLogo.has(f.logo_url)) porLogo.set(f.logo_url, []);
      porLogo.get(f.logo_url).push(f);
    });
    const grupos = Array.from(porLogo.entries()).map(([logo_url, fondos]) => ({ logo_url, fondos }));
    // los grupos con más de un fondo van primero — son los que más vale la
    // pena revisar (o son una familia real, o es un error de carga)
    grupos.sort((a, b) => b.fondos.length - a.fondos.length);
    setLogosGrupos(grupos);
    setLogosCargando(false);
  }

  useEffect(() => {
    if (vista === "biblioteca" && logosVista) { cargarAuditoriaLogos(); cargarTodasLasMarcas(); }
  }, [vista, logosVista]);

  // --- Lista unificada: cada marca registrada + los grupos de logo que
  // todavía no tienen nombre asignado (para poder bautizarlos ahí mismo) —
  // una sola lista en vez de dos secciones separadas y redundantes.
  const gruposPorLogoUrl = new Map(logosGrupos.map((g) => [g.logo_url, g]));
  const marcaLogoUrls = new Set(marcasTodas.map((m) => m.logo_url));

  const filasUnificadasSinFiltrar = [
    ...marcasTodas.map((m) => ({
      key: `m:${m.id}`,
      tipo: "marca",
      marca: m,
      logo_url: m.logo_url,
      fondos: gruposPorLogoUrl.get(m.logo_url)?.fondos || [],
    })),
    ...logosGrupos.filter((g) => !marcaLogoUrls.has(g.logo_url)).map((g) => ({
      key: `g:${g.logo_url}`,
      tipo: "sinNombre",
      marca: null,
      logo_url: g.logo_url,
      fondos: g.fondos,
    })),
  ].sort((a, b) => {
    if (a.tipo !== b.tipo) return a.tipo === "sinNombre" ? -1 : 1; // sin nombre primero, necesitan atención
    const na = a.marca?.nombre || "";
    const nb = b.marca?.nombre || "";
    return na.localeCompare(nb);
  });

  const filasUnificadas = logosQuery.trim().length < 2
    ? filasUnificadasSinFiltrar
    : filasUnificadasSinFiltrar.filter((fila) =>
        (fila.marca?.nombre || "").toLowerCase().includes(logosQuery.toLowerCase()) ||
        fila.fondos.some((f) =>
          f.nombre.toLowerCase().includes(logosQuery.toLowerCase()) || f.isin.toLowerCase().includes(logosQuery.toLowerCase())
        )
      );

  // trae todas las marcas ya registradas, para saber qué grupos de la
  // auditoría ya tienen nombre asignado (y no mostrarles el formulario de
  // "ponerle nombre" de nuevo)
  async function cargarTodasLasMarcas() {
    const { data } = await supabase.from("marcas_logo").select("id, nombre, logo_url").order("nombre");
    setMarcasTodas(data || []);
  }

  const marcaPorLogoUrl = (logo_url) => marcasTodas.find((m) => m.logo_url === logo_url);

  // Sube un logo COMPLETAMENTE NUEVO (de una marca que ningún fondo tiene
  // todavía) y lo registra en marcas_logo con un nombre buscable — así se
  // puede dejar cargado de antemano un logo que "falta" sin necesitar un
  // fondo puntual a mano.
  async function crearMarcaNueva() {
    if (!marcaNombreNuevo.trim() || !marcaArchivoNuevo) {
      setMarcaMensajeNueva("Completá el nombre y elegí un archivo de imagen.");
      return;
    }
    setMarcaGuardandoNueva(true);
    setMarcaMensajeNueva("");
    try {
      const ext = marcaArchivoNuevo.name.split(".").pop();
      const path = `marcas/${slugify(marcaNombreNuevo)}-${Date.now()}.${ext}`;
      const { error: uploadError } = await supabase.storage.from("logos-fondos").upload(path, marcaArchivoNuevo, { upsert: true });
      if (uploadError) throw uploadError;
      const { data } = supabase.storage.from("logos-fondos").getPublicUrl(path);
      // OJO: supabase-js NO tira excepción si el insert falla (por RLS, por
      // ejemplo) — hay que revisar el campo "error" del resultado a mano,
      // si no el código sigue de largo creyendo que se guardó bien.
      const { error: insertError } = await supabase.from("marcas_logo").insert({ nombre: marcaNombreNuevo.trim(), logo_url: data.publicUrl });
      if (insertError) throw insertError;
      setMarcaMensajeNueva(`✓ Marca "${marcaNombreNuevo.trim()}" guardada — ya se puede buscar al editar cualquier fondo.`);
      setMarcaNombreNuevo("");
      setMarcaArchivoNuevo(null);
      cargarTodasLasMarcas();
    } catch (e) {
      setMarcaMensajeNueva("Error al guardar: " + (e.message || JSON.stringify(e)));
    } finally {
      setMarcaGuardandoNueva(false);
    }
  }

  // Bautiza un logo que YA existe (uno de los grupos de la auditoría) como
  // una marca buscable — reutiliza la misma URL, no sube nada de nuevo, así
  // que nunca duplica el archivo en el storage.
  async function guardarGrupoComoMarca(logo_url) {
    const nombre = (marcaNombrePorGrupo[logo_url] || "").trim();
    if (!nombre) return;
    setMarcaGuardandoGrupo(logo_url);
    try {
      const { error: insertError } = await supabase.from("marcas_logo").insert({ nombre, logo_url });
      if (insertError) throw insertError;
      await cargarTodasLasMarcas();
      setMarcaNombrePorGrupo((prev) => ({ ...prev, [logo_url]: "" }));
    } catch (e) {
      alert("Error al guardar la marca: " + (e.message || JSON.stringify(e)));
    } finally {
      setMarcaGuardandoGrupo("");
    }
  }

  // --- Buscador de marca dentro del editor de un fondo (Biblioteca) ---
  useEffect(() => {
    if (marcaFondoQuery.trim().length < 2) { setMarcaFondoResultados([]); return; }
    const t = setTimeout(async () => {
      const { data } = await supabase.from("marcas_logo").select("id, nombre, logo_url").ilike("nombre", `%${marcaFondoQuery}%`).limit(10);
      setMarcaFondoResultados(data || []);
    }, 250);
    return () => clearTimeout(t);
  }, [marcaFondoQuery]);

  // Asigna el logo de una marca elegida al fondo que se está editando —
  // solo en memoria; hay que apretar "Guardar" para que quede en Supabase,
  // igual que con cualquier otro cambio del formulario.
  function asignarMarcaAFondo(marca) {
    setBibliotecaSel((prev) => ({ ...prev, logo_url: marca.logo_url }));
    setBibliotecaLogoFile(null);
    setMarcaFondoAsignada(marca.nombre);
    setMarcaFondoQuery("");
    setMarcaFondoResultados([]);
  }

  // --- Asociación en bloque de una marca a varios fondos ---
  // --- Fila unificada expandible ---
  function toggleFila(key) {
    setFilaExpandidaKey((prev) => (prev === key ? null : key));
    setMarcaAsocQuery("");
    setMarcaAsocResultados([]);
    setMarcaAsocSeleccionados({});
    setMarcaAsocMensaje("");
    setMarcaLogoNuevoArchivo(null);
    setMarcaLogoMensaje("");
  }

  useEffect(() => {
    if (marcaAsocQuery.trim().length < 2) { setMarcaAsocResultados([]); return; }
    const t = setTimeout(async () => {
      const { data } = await supabase
        .from("fondos")
        .select("isin, nombre, logo_url")
        .or(`isin.ilike.%${marcaAsocQuery}%,nombre.ilike.%${marcaAsocQuery}%`)
        .limit(25);
      setMarcaAsocResultados(data || []);
    }, 250);
    return () => clearTimeout(t);
  }, [marcaAsocQuery]);

  function toggleFondoAsoc(fondo) {
    setMarcaAsocSeleccionados((prev) => {
      const next = { ...prev };
      if (next[fondo.isin]) delete next[fondo.isin];
      else next[fondo.isin] = fondo.nombre;
      return next;
    });
  }

  // Aplica el logo de la marca a TODOS los ISIN tildados en un solo update.
  // A partir de ahora, cuando cualquiera de esos fondos se agregue a un
  // Portafolio propuesto o a Descripción de activos, ya va a traer el logo
  // solo (el refresco que ya existe en handleGenerar y en las búsquedas de
  // fondos toma logo_url directo de la tabla `fondos`).
  async function aplicarAsociacionMarca(marca) {
    const isins = Object.keys(marcaAsocSeleccionados);
    if (isins.length === 0) return;
    setMarcaAsocGuardando(true);
    setMarcaAsocMensaje("");
    try {
      await supabase.from("fondos").update({ logo_url: marca.logo_url }).in("isin", isins);
      setMarcaAsocMensaje(`✓ Logo asignado a ${isins.length} fondo${isins.length > 1 ? "s" : ""}.`);
      setMarcaAsocSeleccionados({});
      cargarAuditoriaLogos();
    } catch (e) {
      setMarcaAsocMensaje("Error al guardar: " + (e.message || e));
    } finally {
      setMarcaAsocGuardando(false);
    }
  }

  // cuántos fondos ya tienen el logo de esta marca puesto (para mostrar en
  // la tarjeta de la marca, se apoya en el agrupado de la auditoría)
  function cantidadFondosConLogo(logo_url) {
    const grupo = logosGrupos.find((g) => g.logo_url === logo_url);
    return grupo ? grupo.fondos.length : 0;
  }

  // Cambia la imagen de una marca ya registrada. Clave: además de actualizar
  // marcas_logo, también actualiza TODOS los fondos que ya tenían la URL
  // vieja para que apunten a la nueva — así ningún fondo asociado pierde su
  // logo por el simple hecho de haber actualizado la imagen.
  async function actualizarLogoMarca(marca) {
    if (!marcaLogoNuevoArchivo) return;
    setMarcaLogoActualizando(true);
    setMarcaLogoMensaje("");
    try {
      const ext = marcaLogoNuevoArchivo.name.split(".").pop();
      const path = `marcas/${slugify(marca.nombre)}-${Date.now()}.${ext}`;
      const { error: uploadError } = await supabase.storage.from("logos-fondos").upload(path, marcaLogoNuevoArchivo, { upsert: true });
      if (uploadError) throw uploadError;
      const { data } = supabase.storage.from("logos-fondos").getPublicUrl(path);
      const nuevaUrl = data.publicUrl;
      const viejaUrl = marca.logo_url;

      const { error: errMarca } = await supabase.from("marcas_logo").update({ logo_url: nuevaUrl }).eq("id", marca.id);
      if (errMarca) throw errMarca;

      // reasignar a la nueva URL todos los fondos que tenían la vieja
      const { error: errFondos } = await supabase.from("fondos").update({ logo_url: nuevaUrl }).eq("logo_url", viejaUrl);
      if (errFondos) throw errFondos;

      setMarcaLogoMensaje("✓ Logo actualizado — los fondos ya asociados lo siguen teniendo.");
      setMarcaLogoNuevoArchivo(null);
      await cargarTodasLasMarcas();
      await cargarAuditoriaLogos();
    } catch (e) {
      setMarcaLogoMensaje("Error al actualizar: " + (e.message || JSON.stringify(e)));
    } finally {
      setMarcaLogoActualizando(false);
    }
  }

  // Sacarle el logo a un fondo puntual (sin borrar la marca ni afectar a
  // los demás fondos que la tengan).
  async function quitarFondoDeGrupo(isin) {
    await supabase.from("fondos").update({ logo_url: null }).eq("isin", isin);
    await cargarAuditoriaLogos();
  }

  // Fusiona un grupo "sin nombre" (fondos que llegaron con su propio logo
  // individual, de antes del sistema de marcas) con una marca YA existente
  // — reasigna esos fondos a la URL de la marca elegida. No borra nada: el
  // grupo "sin nombre" simplemente desaparece de la lista porque ningún
  // fondo vuelve a apuntar a esa URL vieja.
  async function fusionarConMarcaExistente(fila) {
    const marcaId = marcaFusionSeleccion[fila.logo_url];
    if (!marcaId) return;
    const marca = marcasTodas.find((m) => String(m.id) === String(marcaId));
    if (!marca) return;
    setMarcaFusionando(fila.logo_url);
    try {
      const isins = fila.fondos.map((f) => f.isin);
      await supabase.from("fondos").update({ logo_url: marca.logo_url }).in("isin", isins);
      await cargarAuditoriaLogos();
      setMarcaFusionSeleccion((prev) => ({ ...prev, [fila.logo_url]: "" }));
    } finally {
      setMarcaFusionando("");
    }
  }

  // --- Importación de marcas en bloque ---
  function handleMarcaImportFiles(fileList) {
    const entradas = Array.from(fileList).map((file) => {
      let nombre = file.name.replace(/\.[^.]+$/, ""); // sacar extensión
      nombre = nombre.replace(/^\d+[_\-\s]*/, ""); // sacar el número inicial (ej: "07_")
      nombre = nombre.replace(/[_\-]+/g, " ").trim(); // guiones/underscores -> espacios
      return { file, nombre };
    });
    setMarcaImportEntradas(entradas);
    setMarcaImportResumen("");
  }

  function actualizarMarcaImportNombre(idx, valor) {
    setMarcaImportEntradas((prev) => prev.map((e, i) => i === idx ? { ...e, nombre: valor } : e));
  }

  function quitarMarcaImportEntrada(idx) {
    setMarcaImportEntradas((prev) => prev.filter((_, i) => i !== idx));
  }

  async function aplicarImportMarcas() {
    setMarcaImportAplicando(true);
    let aplicados = 0, errores = 0, primerError = "";
    for (let i = 0; i < marcaImportEntradas.length; i++) {
      const { file, nombre } = marcaImportEntradas[i];
      if (!nombre.trim()) continue;
      try {
        const ext = file.name.split(".").pop();
        const path = `marcas/${slugify(nombre)}-${Date.now()}-${i}.${ext}`;
        const { error: uploadError } = await supabase.storage.from("logos-fondos").upload(path, file, { upsert: true });
        if (uploadError) throw uploadError;
        const { data } = supabase.storage.from("logos-fondos").getPublicUrl(path);
        // OJO: supabase-js NO tira excepción si el insert falla — hay que
        // revisar "error" del resultado a mano, si no el conteo de
        // "aplicados" queda mintiendo aunque no se haya guardado nada.
        const { error: insertError } = await supabase.from("marcas_logo").insert({ nombre: nombre.trim(), logo_url: data.publicUrl });
        if (insertError) throw insertError;
        aplicados++;
      } catch (e) {
        errores++;
        if (!primerError) primerError = e.message || JSON.stringify(e);
      }
    }
    setMarcaImportResumen(
      `✓ ${aplicados} marca(s) cargada(s)${errores ? `, ${errores} con error` : ""}.` +
      (primerError ? ` Primer error: ${primerError}` : "")
    );
    setMarcaImportEntradas([]);
    await cargarTodasLasMarcas();
    setMarcaImportAplicando(false);
  }

  // Excel de celdas con fecha (Maturity de bonos) puede venir como objeto
  // Date de JS (si la columna está formateada como fecha en Excel) o como
  // texto plano — normalizamos siempre a "AAAA-MM-DD" para Supabase.
  function excelValueToDateStr(v) {
    if (v instanceof Date) return v.toISOString().slice(0, 10);
    if (!v) return null;
    return String(v).trim() || null;
  }

  function excelValueToNumber(v) {
    if (v === null || v === undefined || v === "") return null;
    const n = Number(v);
    return Number.isNaN(n) ? null : n;
  }

  // El "%" puede venir como número entero (10 = 10%, formato que usa el
  // resto de la app) o como fracción de Excel (0.10 = 10%, si la celda
  // viene formateada como porcentaje desde otra planilla). Se asume
  // fracción solo si es un valor entre 0 y 1 — una posición real de
  // portafolio casi nunca pesa menos de 1%.
  function normalizarPct(v) {
    if (v === null || v === undefined) return null;
    return v > 0 && v < 1 ? v * 100 : v;
  }

  // Parsea el Excel base de instrumentos (3 pestañas: Fondos, Acciones,
  // Bonos — Fondos distributivos NO es una pestaña aparte, ver
  // descargarPlantillaBase) y devuelve la lista de instrumentos leídos, sin
  // tocar nada de Supabase ni de la propuesta — lo usan tanto el importador
  // de la Biblioteca (que sí guarda en `fondos`) como el de Portafolio
  // propuesto (que arma esta propuesta puntual y no toca la biblioteca).
  async function parseExcelBaseInstrumentos(file) {
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: "array", cellDates: true });

    const nombresPestana = ["Fondos", "Acciones", "Bonos"];
    const filas = [];

    for (const sheet of nombresPestana) {
      const ws = wb.Sheets[sheet];
      if (!ws) continue; // pestaña no presente en este Excel puntual, se saltea

      const filasHoja = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
      const headers = (filasHoja[1] || []).map((h) => (h == null ? "" : String(h).trim()));
      const filasDatos = filasHoja.slice(2).filter((r) => r.some((v) => v !== null && v !== ""));
      // la pestaña Fondos puede o no traer las columnas de dividendo,
      // según se haya tildado "Incluir dividendos" al descargarla — acá no
      // importa cuál se usó, se detecta solo mirando los headers.
      const tieneColumnasDividendo = headers.includes("Dividendo (%)");

      filasDatos.forEach((r) => {
        const obj = {};
        headers.forEach((h, i) => { obj[h] = r[i]; });

        const codigo = sheet === "Acciones" ? obj["Ticker"] : obj["Código"];
        const registro = {
          isin: codigo ? String(codigo).trim() : "",
          nombre: obj["Nombre"] ? String(obj["Nombre"]).trim() : "",
          sector: obj["Sector"] || "",
          categoria: obj["Categoría"] || CATEGORIAS[0],
          // % e Inversión (USD): se leen si están — quien los use decide
          // qué hacer si faltan (la Biblioteca los ignora directamente, ya
          // que "fondos" no tiene esas columnas; Portafolio propuesto los
          // usa o los calcula, ver handleImportExcelPropuesta).
          pct: normalizarPct(excelValueToNumber(obj["%"])),
          monto: excelValueToNumber(obj["Inversión (USD)"]),
        };

        if (sheet === "Fondos") {
          // distributivo o no se decide FILA POR FILA: si esta fila
          // puntual trae un Dividendo (%) cargado, es distributivo — así
          // una misma planilla puede tener de los dos tipos mezclados.
          const dividendoPct = tieneColumnasDividendo ? excelValueToNumber(obj["Dividendo (%)"]) : null;
          const esDistributivo = dividendoPct !== null && dividendoPct > 0;
          registro.tipo_instrumento = esDistributivo ? "fondo_distributivo" : "fondo";
          registro.ytd = excelValueToNumber(obj["Rend. YTD"]) || 0;
          registro.y1 = excelValueToNumber(obj["Rend. 1 año"]) || 0;
          registro.y3 = excelValueToNumber(obj["Rend. 3 años"]) || 0;
          registro.y5 = excelValueToNumber(obj["Rend. 5 años"]) || 0;
          registro.ter = excelValueToNumber(obj["TER"]) || 0;
          registro.dividendo_pct = esDistributivo ? dividendoPct : 0;
          registro.frecuencia_dividendo = esDistributivo ? (obj["Frec. Dividendo"] || "") : "";
        } else if (sheet === "Acciones") {
          registro.tipo_instrumento = "accion";
          registro.ytd = excelValueToNumber(obj["Rend. YTD"]) || 0;
          registro.y1 = excelValueToNumber(obj["Rend. 1 año"]) || 0;
          registro.y3 = excelValueToNumber(obj["Rend. 3 años"]) || 0;
          registro.y5 = excelValueToNumber(obj["Rend. 5 años"]) || 0;
        } else if (sheet === "Bonos") {
          registro.tipo_instrumento = "bono";
          registro.cupon_pct = excelValueToNumber(obj["Cupón (%)"]) || 0;
          registro.rating = obj["Rating S&P"] || "";
          registro.price = excelValueToNumber(obj["Price**"] ?? obj["Price"]) || 0;
          registro.yield_pct = excelValueToNumber(obj["Yield"]) || 0;
          registro.maturity = excelValueToDateStr(obj["Maturity"]) || "";
        }
        if (registro.isin && registro.nombre) filas.push(registro);
      });
    }
    return filas;
  }

  // --- Importador de la Biblioteca de fondos (screen aparte): guarda en
  // Supabase, para que quede buscable en cualquier propuesta futura. Pide
  // revisión antes de aplicar. ---
  async function handleImportBibliotecaBase(file) {
    if (!file) return;
    setBaseImportResumen("");
    try {
      const filasPreview = await parseExcelBaseInstrumentos(file);
      setBaseImportPreview(filasPreview);
      if (filasPreview.length === 0) setBaseImportResumen("No se encontró ninguna fila para importar — revisá que el archivo tenga las pestañas Fondos/Acciones/Bonos con datos desde la fila 3.");
    } catch (e) {
      setBaseImportResumen("Error al leer el archivo: " + (e.message || e));
    }
  }

  // --- Importador de Portafolio propuesto: NO toca la biblioteca ni
  // Supabase — agrega los instrumentos directo a ESTA propuesta. Si el
  // Excel trae % y/o Inversión (USD), se usan tal cual; si falta uno de
  // los dos, se calcula solo a partir del otro y del "Monto total a
  // invertir" ya cargado en este paso. Si el cash todavía no está
  // definido, se completa solo con lo que sobre del monto total. ---
  async function handleImportExcelPropuesta(file) {
    if (!file) return;
    setPropuestaImportResumen("");
    try {
      const filas = await parseExcelBaseInstrumentos(file);
      if (filas.length === 0) {
        setPropuestaImportResumen("No se encontró ninguna fila para importar — revisá que el archivo tenga las pestañas Fondos/Acciones/Bonos con datos desde la fila 3.");
        return;
      }
      const nuevos = filas.map((f) => {
        let pct = f.pct;
        let monto = f.monto;
        // falta uno de los dos: se calcula a partir del otro + el monto
        // total a invertir de esta propuesta
        if ((monto === null || monto === undefined) && pct !== null && montoInvertir) {
          monto = Math.round((pct / 100) * montoInvertir);
        }
        if ((pct === null || pct === undefined) && monto !== null && montoInvertir) {
          pct = +(100 * monto / montoInvertir).toFixed(1);
        }
        return {
          ...f,
          pct: pct || 0, monto: monto || 0,
          rating: f.rating || "", price: f.price || 0, yield_pct: f.yield_pct || 0, maturity: f.maturity || "",
        };
      });

      setProposedAssets((prev) => {
        const todos = [...prev, ...nuevos];
        // cash: si todavía no se cargó nada a mano, se completa solo con
        // lo que falte para llegar al monto total, ahora que ya sabemos
        // cuánto quedó asignado en instrumentos.
        if (!cashManualPropuesta) {
          const sumaMontos = todos.reduce((s, a) => s + (Number(a.monto) || 0), 0);
          const cashCalculado = Math.max(Math.round(montoInvertir - sumaMontos), 0);
          if (cashCalculado > 0) setCashManualPropuesta(cashCalculado);
        }
        return todos;
      });

      const huboCalculo = filas.some((f) => f.pct === null || f.monto === null);
      setPropuestaImportResumen(
        `✓ ${nuevos.length} instrumento(s) agregado(s) a esta propuesta.` +
        (huboCalculo ? " Completé % o monto donde faltaba uno de los dos, a partir del monto total a invertir." : "")
      );
    } catch (e) {
      setPropuestaImportResumen("Error al leer el archivo: " + (e.message || e));
    }
  }

  function actualizarBaseImportPreview(idx, campo, valor) {
    setBaseImportPreview((prev) => prev.map((r, i) => i === idx ? { ...r, [campo]: valor } : r));
  }

  function quitarBaseImportPreview(idx) {
    setBaseImportPreview((prev) => prev.filter((_, i) => i !== idx));
  }

  // Recién acá se guarda de verdad en Supabase — todo lo de arriba fue
  // solo lectura y edición en memoria. OJO: baseImportPreview puede traer
  // "pct"/"monto" (se leen del Excel para el caso de uso de Portafolio
  // propuesto) — la tabla `fondos` no tiene esas columnas, así que acá se
  // arma explícitamente solo con los campos que sí existen ahí.
  async function aplicarBaseImportPreview() {
    if (baseImportPreview.length === 0) return;
    setBaseImportCargando(true);
    setBaseImportResumen("");
    try {
      const registrosParaGuardar = baseImportPreview.map(({ pct, monto, ...resto }) => resto);
      const { error } = await supabase.from("fondos").upsert(registrosParaGuardar, { onConflict: "isin" });
      if (error) throw error;
      setBaseImportResumen(`✓ ${baseImportPreview.length} instrumento(s) cargado(s) a la biblioteca.`);
      setBaseImportPreview([]);
    } catch (e) {
      setBaseImportResumen("Error al guardar: " + (e.message || e));
    } finally {
      setBaseImportCargando(false);
    }
  }

  // Genera y descarga la plantilla del Excel base (3 pestañas) desde acá
  // mismo — no depende de un archivo guardado en ningún lado, así siempre
  // está actualizada con las columnas que el importador realmente espera.
  function descargarPlantillaBase(incluirDividendos) {
    const categoriaNota = `Categoría: ${CATEGORIAS.join(" / ")}.`;
    const notaComun = " % e Inversión (USD): si subís este archivo en Portafolio propuesto, se usan (o se calculan solos entre sí y con el cash); si lo subís en Biblioteca de fondos, se ignoran — ahí no hacen falta.";

    const fondosCols = incluirDividendos
      ? ["%", "Código", "Nombre", "Sector", "Dividendo (%)", "Frec. Dividendo", "Rend. YTD", "Rend. 1 año", "Rend. 3 años", "Rend. 5 años", "Inversión (USD)", "Dividendo anual", "TER", "Categoría"]
      : ["%", "Código", "Nombre", "Sector", "Rend. YTD", "Rend. 1 año", "Rend. 3 años", "Rend. 5 años", "Inversión (USD)", "TER", "Categoría"];
    const fondosEjemplo = incluirDividendos
      ? [null, "IE00B8K7V925", "PIMCO GIS Income Fund", "Global", 6.31, "Mensual", 0.39, 5.15, 5.87, 2.56, null, null, 1.45, "Renta Fija"]
      : [null, "LU2750480548", "Wellington Total Credit", "Global Flexible", 0.22, 3.8, 6.1, 3.2, null, 1.25, "Renta Fija"];
    const fondosNota = (incluirDividendos
      ? "Fila por fila: si un fondo tiene algo cargado en Dividendo (%), se guarda como fondo distributivo — si lo dejás vacío, es un fondo normal. Se pueden mezclar los dos tipos en la misma pestaña."
      : "Sin columnas de dividendo — si algún fondo reparte, descargá de nuevo tildando 'Incluir dividendos' antes de bajar el archivo.") + " " + categoriaNota + notaComun;

    const specs = [
      { sheet: "Fondos", cols: fondosCols, nota: fondosNota, ejemplo: fondosEjemplo },
      {
        sheet: "Acciones",
        cols: ["%", "Ticker", "Nombre", "Sector", "Rend. YTD", "Rend. 1 año", "Rend. 3 años", "Rend. 5 años", "Inversión (USD)", "Categoría"],
        nota: "Acciones individuales — sin TER (no tienen costo de gestión)." + " " + categoriaNota + notaComun,
        ejemplo: [null, "ORCL", "Oracle Corp", "Tecnología", -21.45, -34.83, 10.67, 12.71, null, "Renta Variable"],
      },
      {
        sheet: "Bonos",
        cols: ["%", "Código", "Nombre", "Sector", "Cupón (%)", "Rating S&P", "Price**", "Yield", "Maturity", "Cupón Anual", "Inversión (USD)", "Categoría"],
        nota: "Bonos individuales. Maturity en formato AAAA-MM-DD." + " " + categoriaNota + notaComun,
        ejemplo: [null, "US03938LBE39", "Arcelormittal SA", "Basic Materials", 6.55, "BBB", 102.27, 4.79, "2027-11-29", null, null, "Renta Fija"],
      },
    ];

    const wb = XLSX.utils.book_new();
    specs.forEach(({ sheet, cols, nota, ejemplo }) => {
      const ws = XLSX.utils.aoa_to_sheet([[nota], cols, ejemplo]);
      ws["!cols"] = cols.map((c) => ({ wch: c === "Nombre" ? 32 : Math.max(12, c.length + 3) }));
      ws["!merges"] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: cols.length - 1 } }];
      XLSX.utils.book_append_sheet(wb, ws, sheet);
    });
    XLSX.writeFile(wb, `Biblioteca_instrumentos_AIVA${incluirDividendos ? "_con_dividendos" : ""}.xlsx`);
  }

  // --- Importación masiva ---
  async function handleImportJson(file) {    if (!file) return;
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
        .select("isin, nombre, sector, categoria, tipo_instrumento, ter, ytd, y1, y3, y5, dividendo_pct, frecuencia_dividendo, cupon_pct, rating, price, yield_pct, maturity, uso_frecuente, descripcion, factsheet_url, logo_url")
        .or(`isin.ilike.%${fondoQuery}%,nombre.ilike.%${fondoQuery}%`)
        .order("uso_frecuente", { ascending: false })
        .limit(8);
      setFondoResultados(data || []);
    }, 250);
    return () => clearTimeout(t);
  }, [fondoQuery]);

  function addProposedAsset(fondo) {
    const tipo = fondo.tipo_instrumento || "fondo";
    setProposedAssets((prev) => [...prev, {
      ...fondo,
      tipo_instrumento: tipo,
      categoria: fondo.categoria || "Renta Variable",
      pct: 0, monto: 0,
      ytd: fondo.ytd || 0, y1: fondo.y1 || 0, y3: fondo.y3 || 0, y5: fondo.y5 || 0,
      ter: fondo.ter ?? null,
      dividendo_pct: fondo.dividendo_pct || 0, frecuencia_dividendo: fondo.frecuencia_dividendo || "",
      cupon_pct: fondo.cupon_pct || 0, rating: fondo.rating || "", price: fondo.price || 0,
      yield_pct: fondo.yield_pct || 0, maturity: fondo.maturity || "",
    }]);
    setFondoQuery("");
    setFondoResultados([]);
  }

  // --- Columnas visibles/orden de la tabla de Portafolio propuesto ---
  function toggleColumnaVisible(tipo, key) {
    setColumnasConfig((prev) => {
      const actual = prev[tipo];
      const nueva = actual.includes(key) ? actual.filter((k) => k !== key) : [...actual, key];
      return { ...prev, [tipo]: nueva };
    });
  }

  function moverColumna(tipo, key, direccion) {
    setColumnasConfig((prev) => {
      const actual = [...prev[tipo]];
      const idx = actual.indexOf(key);
      const nuevoIdx = idx + direccion;
      if (idx === -1 || nuevoIdx < 0 || nuevoIdx >= actual.length) return prev;
      [actual[idx], actual[nuevoIdx]] = [actual[nuevoIdx], actual[idx]];
      return { ...prev, [tipo]: actual };
    });
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
  function moverTeam(idx, direccion) {
    setTeam((prev) => {
      const nuevo = [...prev];
      const destino = idx + direccion;
      if (destino < 0 || destino >= nuevo.length) return prev;
      [nuevo[idx], nuevo[destino]] = [nuevo[destino], nuevo[idx]];
      return nuevo;
    });
  }

  async function handleTeamFoto(id, file) {
    if (!file) return;
    const b64 = await fileToBase64(file);
    setTeam((prev) => prev.map((x) => x.id === id ? { ...x, foto_base64: b64 } : x));
  }

  async function agregarActivoNuevo() {
    if (!nuevoActivoNombre.trim() || !nuevoActivoIsin.trim()) return;
    const nuevoFondo = { isin: nuevoActivoIsin.trim(), nombre: nuevoActivoNombre.trim(), uso_frecuente: false, tipo_instrumento: "accion" };
    await supabase.from("fondos").upsert(nuevoFondo, { onConflict: "isin" });
    addProposedAsset(nuevoFondo);
    setNuevoActivoNombre("");
    setNuevoActivoIsin("");
  }

  // --- Descripción de activos: búsqueda + selección manual por categoría ---
  useEffect(() => {
    if (descQuery.trim().length < 2) { setDescResultados([]); return; }
    const t = setTimeout(async () => {
      const { data } = await supabase
        .from("fondos")
        .select("isin, nombre, descripcion, factsheet_url, logo_url")
        .or(`isin.ilike.%${descQuery}%,nombre.ilike.%${descQuery}%`)
        .limit(8);
      setDescResultados(data || []);
    }, 250);
    return () => clearTimeout(t);
  }, [descQuery]);

  function addDescManual(fondo) {
    setDescSeleccion((prev) => {
      const lista = prev[descCategoriaDestino];
      if (lista.some((f) => f.isin === fondo.isin)) return prev;
      return { ...prev, [descCategoriaDestino]: [...lista, fondo] };
    });
    setDescQuery("");
    setDescResultados([]);
  }

  function quitarDescManual(categoria, isin) {
    setDescSeleccion((prev) => ({ ...prev, [categoria]: prev[categoria].filter((f) => f.isin !== isin) }));
  }

  // Toma lo que ya está cargado en Portafolio propuesto y lo vuelca en las
  // 3 categorías de descripción — punto de partida rápido, después se
  // puede seguir ajustando a mano.
  function prellenarDescDesdePortafolio() {
    const grupos = { "Renta Fija & Multi Activo": [], "Renta Variable": [], "Alternativos Líquidos": [] };
    proposedAssets.forEach((a) => {
      const destino = (a.categoria === "Renta Fija" || a.categoria === "Multi Activo") ? "Renta Fija & Multi Activo"
        : a.categoria === "Renta Variable" ? "Renta Variable"
        : a.categoria === "Alternativos Líquidos" ? "Alternativos Líquidos" : null;
      if (destino && a.isin && !grupos[destino].some((f) => f.isin === a.isin)) {
        grupos[destino].push({ isin: a.isin, nombre: a.nombre, descripcion: a.descripcion, factsheet_url: a.factsheet_url, logo_url: a.logo_url });
      }
    });
    setDescSeleccion(grupos);
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

  function buildConfig(proposedAssetsOverride, descSeleccionOverride) {
    const assetsAUsar = proposedAssetsOverride || proposedAssets;
    const descAUsar = descSeleccionOverride || descSeleccion;
    // La tabla "estándar" del PPT (build_portafolio_propuesto) es la que
    // usan Fondos y Acciones — Acciones comparte el mismo layout de
    // columnas, con TER en blanco. Bonos y Fondos distributivos arman sus
    // propias slides aparte (ver más abajo), así que quedan afuera de acá.
    const categorias = CATEGORIAS.map((label) => ({
      label: label === "Renta Fija" ? "Fondos Renta Fija" : label === "Multi Activo" ? "Fondo Multi Activo" : label === "Renta Variable" ? "Fondo Renta Variable" : "Fondos Alternativos Líquidos",
      fondos: assetsAUsar
        .filter((a) => a.categoria === label && (a.tipo_instrumento === "fondo" || a.tipo_instrumento === "accion" || !a.tipo_instrumento))
        .map((a) => ({
          isin: a.isin, nombre: a.nombre, sector: a.sector || "", ytd: a.ytd || 0, y1: a.y1 || 0, y3: a.y3 || 0, y5: a.y5 || 0, pct: a.pct, monto: a.monto, ter: a.ter || 0,
        })),
    })).filter((c) => c.fondos.length > 0);

    // Bonos y Fondos distributivos: cada uno arma su propia página en el
    // PPT (agregar_slide_bonos / agregar_slide_distributivos), solo si hay
    // al menos uno cargado. Cupón anual y Dividendo anual se calculan acá
    // mismo a partir del monto asignado, no vienen de la biblioteca.
    const bonosPropuesto = assetsAUsar.filter((a) => a.tipo_instrumento === "bono").map((a) => ({
      isin: a.isin, nombre: a.nombre, sector: a.sector || "",
      cupon_pct: a.cupon_pct || 0, rating: a.rating || "", price: a.price || 0,
      yield_pct: a.yield_pct || 0, maturity: a.maturity || "",
      cupon_anual: Math.round((a.monto || 0) * (a.cupon_pct || 0) / 100),
      pct: a.pct, monto: a.monto,
    }));

    const fondosDistributivosPropuesto = assetsAUsar.filter((a) => a.tipo_instrumento === "fondo_distributivo").map((a) => ({
      isin: a.isin, nombre: a.nombre, sector: a.sector || "",
      ytd: a.ytd || 0, y1: a.y1 || 0, y3: a.y3 || 0, y5: a.y5 || 0, ter: a.ter || 0,
      dividendo_pct: a.dividendo_pct || 0, frecuencia_dividendo: a.frecuencia_dividendo || "",
      dividendo_anual: Math.round((a.monto || 0) * (a.dividendo_pct || 0) / 100),
      pct: a.pct, monto: a.monto,
    }));

    const cashMonto = Number(cashManualPropuesta) || 0;

    const byCat = {};
    assetsAUsar.forEach((a) => { byCat[a.categoria] = (byCat[a.categoria] || 0) + (a.pct || 0) / 100; });
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
      "Renta Fija & Multi Activo": descAUsar["Renta Fija & Multi Activo"].map((a) => ({ nombre: a.nombre, descripcion: a.descripcion || "", factsheet_url: a.factsheet_url || "", logo_url: a.logo_url || "" })),
      "Renta Variable": descAUsar["Renta Variable"].map((a) => ({ nombre: a.nombre, descripcion: a.descripcion || "", factsheet_url: a.factsheet_url || "", logo_url: a.logo_url || "" })),
      "Alternativos Líquidos": descAUsar["Alternativos Líquidos"].map((a) => ({ nombre: a.nombre, descripcion: a.descripcion || "", factsheet_url: a.factsheet_url || "", logo_url: a.logo_url || "" })),
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
      // hasta 6 integrantes, en el orden que haya quedado en la lista —
      // el backend ya arma la grilla de 3x2 con esa cantidad. foto_base64
      // solo viaja si se subió una nueva (si no, el backend hereda la foto
      // original por nombre, o dibuja un círculo con las iniciales).
      equipo: team.filter((m) => m.incluido).slice(0, 6).map((m) => ({
        nombre: m.nombre, puesto: m.puesto, educacion: m.educacion,
        ...(m.foto_base64 ? { foto_base64: m.foto_base64 } : {}),
      })),
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
      bonos_propuesto: bonosPropuesto,
      fondos_distributivos_propuesto: fondosDistributivosPropuesto,
      comentarios,
    };
  }

  async function handleGenerar() {
    setGenerando(true); setError(""); setResultado(null);
    try {
      // antes de armar el config, se refrescan descripción/logo/factsheet
      // de cada fondo directo desde la biblioteca — así no importa si el
      // fondo se agregó a la propuesta antes o después de actualizarlo en
      // la biblioteca, siempre viaja la versión más reciente
      const isinsUsados = proposedAssets.map((a) => a.isin).filter(Boolean);
      const isinsDesc = Object.values(descSeleccion).flat().map((a) => a.isin).filter(Boolean);
      const todosLosIsin = [...new Set([...isinsUsados, ...isinsDesc])];
      let frescos = {};
      if (todosLosIsin.length > 0) {
        const { data } = await supabase.from("fondos").select("isin, descripcion, factsheet_url, logo_url").in("isin", todosLosIsin);
        (data || []).forEach((f) => { frescos[f.isin] = f; });
      }
      const proposedAssetsFrescos = proposedAssets.map((a) => ({
        ...a,
        descripcion: frescos[a.isin]?.descripcion ?? a.descripcion,
        factsheet_url: frescos[a.isin]?.factsheet_url ?? a.factsheet_url,
        logo_url: frescos[a.isin]?.logo_url ?? a.logo_url,
      }));
      const descSeleccionFresca = Object.fromEntries(
        Object.entries(descSeleccion).map(([cat, lista]) => [cat, lista.map((a) => ({
          ...a,
          descripcion: frescos[a.isin]?.descripcion ?? a.descripcion,
          factsheet_url: frescos[a.isin]?.factsheet_url ?? a.factsheet_url,
          logo_url: frescos[a.isin]?.logo_url ?? a.logo_url,
        }))])
      );

      const config = buildConfig(proposedAssetsFrescos, descSeleccionFresca);
      const res = await fetch(`${BACKEND_URL}/generar`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ tipo, creado_por: usuario, repcode: repcode || "sin-rep", cliente, config }),
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

  if (!usuario) {
    return (
      <div style={{ fontFamily: "Montserrat, sans-serif", background: CREAM, minHeight: "100vh", padding: 40, display: "flex", alignItems: "center", justifyContent: "center" }}>
        <div style={{ background: "#fff", borderRadius: 12, padding: "40px 44px", width: 440, boxShadow: "0 1px 3px rgba(0,0,0,0.08)" }}>
          <h2 style={{ color: NAVY, margin: "0 0 6px", fontSize: 20 }}>Automatizador de propuestas</h2>
          <p style={{ color: "#78776f", fontSize: 13.5, margin: "0 0 26px" }}>Elegí quién sos para empezar. El asesor / RepCode se elige después, en la Portada — es opcional.</p>

          <Field label="¿Quién sos?">
            <select style={inputStyle} value={usuario} onChange={(e) => setUsuario(e.target.value)}>
              <option value="">Seleccioná tu nombre</option>
              {USUARIOS.map((u) => <option key={u} value={u}>{u}</option>)}
            </select>
          </Field>
        </div>
      </div>
    );
  }

  const totalRevisionPreview = currentAssets.reduce((s, a) => s + (Number(a.valor_actual) || 0), 0) + Number(cashValorRevision || 0);

  // Tabla de revisión de la importación del Excel base — se usa en los dos
  // lugares donde está el cargador (Biblioteca de fondos y Portafolio
  // propuesto), agrupada por tipo igual que la tabla de Portafolio
  // propuesto, para que sea consistente visualmente.
  function renderBaseImportPreview() {
    if (baseImportPreview.length === 0) return null;
    return (
      <div style={{ marginTop: 12 }}>
        {TIPO_ORDEN.filter((tipo) => baseImportPreview.some((r) => r.tipo_instrumento === tipo)).map((tipo) => {
          const filas = baseImportPreview.map((r, i) => ({ r, i })).filter(({ r }) => r.tipo_instrumento === tipo);
          const cols = LIBRERIA_COLUMNAS_POR_TIPO[tipo];
          return (
            <div key={tipo} style={{ marginBottom: 16 }}>
              <div style={{ fontSize: 11.5, fontWeight: 700, color: NAVY, marginBottom: 6 }}>{TIPO_LABELS[tipo]} ({filas.length})</div>
              <div style={{ overflowX: "auto", background: "#fff", border: "1px solid #eae7dc", borderRadius: 8 }}>
                <table style={{ borderCollapse: "collapse", width: "100%", fontSize: 11.5 }}>
                  <thead>
                    <tr style={{ background: CREAM }}>
                      <th style={{ padding: "5px 7px", textAlign: "left", whiteSpace: "nowrap" }}>Categoría</th>
                      {cols.map((c) => <th key={c.key} style={{ padding: "5px 7px", textAlign: "left", whiteSpace: "nowrap" }}>{c.label}</th>)}
                      <th></th>
                    </tr>
                  </thead>
                  <tbody>
                    {filas.map(({ r, i }) => (
                      <tr key={i} style={{ borderTop: "1px solid #f2f0e9" }}>
                        <td style={{ padding: "3px 7px" }}>
                          <select style={{ ...miniInputStyle, padding: "3px 5px", fontSize: 11 }} value={r.categoria} onChange={(e) => actualizarBaseImportPreview(i, "categoria", e.target.value)}>
                            {CATEGORIAS.map((c) => <option key={c} value={c}>{c}</option>)}
                          </select>
                        </td>
                        {cols.map((c) => (
                          <td key={c.key} style={{ padding: "3px 7px" }}>
                            {c.fijo ? (
                              <span>{r[c.key]}</span>
                            ) : c.tipo === "date" ? (
                              <input type="date" style={{ ...miniInputStyle, padding: "3px 5px", fontSize: 11 }} value={r[c.key] || ""} onChange={(e) => actualizarBaseImportPreview(i, c.key, e.target.value)} />
                            ) : (
                              <input
                                type={c.tipo === "number" ? "number" : "text"}
                                style={{ ...miniInputStyle, padding: "3px 5px", width: c.tipo === "number" ? 60 : 90, fontSize: 11 }}
                                value={r[c.key] ?? ""}
                                onChange={(e) => actualizarBaseImportPreview(i, c.key, c.tipo === "number" ? +e.target.value : e.target.value)}
                              />
                            )}
                          </td>
                        ))}
                        <td style={{ padding: "3px 7px" }}>
                          <button onClick={() => quitarBaseImportPreview(i)} style={{ border: "none", background: "none", color: "#b23b3b", fontSize: 11, cursor: "pointer" }}>Quitar</button>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          );
        })}
        <button onClick={aplicarBaseImportPreview} disabled={baseImportCargando} style={{ padding: "8px 16px", borderRadius: 6, border: "none", background: NAVY, color: "#fff", fontSize: 12.5, cursor: "pointer" }}>
          {baseImportCargando ? "Guardando…" : `Aplicar (${baseImportPreview.length} instrumento${baseImportPreview.length === 1 ? "" : "s"})`}
        </button>
      </div>
    );
  }

  return (
    <div style={{ fontFamily: "Montserrat, sans-serif", background: CREAM, minHeight: "100vh" }}>
      <div style={{ background: NAVY, color: "#fff", padding: "14px 24px", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
        <span style={{ fontSize: 14, fontWeight: 600 }}>Automatizador de propuestas</span>
        <div style={{ display: "flex", alignItems: "center", gap: 18 }}>
          <button onClick={() => setVista("nueva")} style={{ background: "none", border: "none", color: vista === "nueva" ? "#fff" : "#9fb0c9", fontWeight: vista === "nueva" ? 600 : 400, cursor: "pointer", fontSize: 13 }}>Nueva propuesta</button>
          <button onClick={() => setVista("registro")} style={{ background: "none", border: "none", color: vista === "registro" ? "#fff" : "#9fb0c9", fontWeight: vista === "registro" ? 600 : 400, cursor: "pointer", fontSize: 13 }}>Registro</button>
          <button onClick={() => setVista("biblioteca")} style={{ background: "none", border: "none", color: vista === "biblioteca" ? "#fff" : "#9fb0c9", fontWeight: vista === "biblioteca" ? 600 : 400, cursor: "pointer", fontSize: 13 }}>Biblioteca de fondos</button>
          <span style={{ fontSize: 12.5, opacity: 0.85 }}>{usuario}{asesorSel ? ` · ${repcode} — ${asesorSel.nombre}` : ""}</span>
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
        ) : logosVista ? (
          <div style={{ padding: "28px 36px" }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
              <h3 style={{ color: NAVY, fontSize: 16, margin: 0 }}>Auditoría de logos</h3>
              <button onClick={() => setLogosVista(false)} style={{ border: "none", background: "none", color: "#78776f", fontSize: 12.5, cursor: "pointer" }}>← Volver a la biblioteca</button>
            </div>
            <p style={{ fontSize: 12.5, color: "#78776f", marginBottom: 16 }}>
              Agrupa todos los fondos por su logo real. Un grupo con un solo fondo es lo normal. Un grupo con varios fondos que no son de la misma familia (como pasó con Calamos y Morgan Stanley) suele ser un error de carga — el logo de uno quedó pegado en el otro.
            </p>

            <div style={{ background: "#fff", border: "1px dashed #d8d5cc", borderRadius: 8, padding: 14, marginBottom: 22, maxWidth: 520 }}>
              <div style={{ fontSize: 12.5, fontWeight: 600, color: NAVY, marginBottom: 8 }}>Agregar un logo de marca que todavía no está cargado</div>
              <div style={{ fontSize: 11.5, color: "#78776f", marginBottom: 10 }}>Para logos que ningún fondo tiene todavía (ej: uno que se perdió en la importación). Después se busca por este nombre al editar cualquier fondo.</div>
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr auto", gap: 8, alignItems: "end" }}>
                <MiniField label="Nombre de la marca">
                  <input style={miniInputStyle} value={marcaNombreNuevo} onChange={(e) => setMarcaNombreNuevo(e.target.value)} placeholder="Ej: MFS" />
                </MiniField>
                <MiniField label="Imagen del logo">
                  <input type="file" accept="image/*" onChange={(e) => setMarcaArchivoNuevo(e.target.files[0])} style={{ ...miniInputStyle, padding: "6px" }} />
                </MiniField>
                <button onClick={crearMarcaNueva} disabled={marcaGuardandoNueva} style={{ padding: "8px 16px", borderRadius: 6, border: "none", background: NAVY, color: "#fff", fontSize: 12.5, cursor: "pointer" }}>
                  {marcaGuardandoNueva ? "Guardando…" : "Guardar"}
                </button>
              </div>
              {marcaMensajeNueva && <div style={{ marginTop: 8, fontSize: 12, color: marcaMensajeNueva.startsWith("Error") ? "#b23b3b" : "#3a7d44" }}>{marcaMensajeNueva}</div>}
            </div>

            <div style={{ background: "#fff", border: "1px dashed #d8d5cc", borderRadius: 8, padding: 14, marginBottom: 22, maxWidth: 640 }}>
              <div style={{ fontSize: 12.5, fontWeight: 600, color: NAVY, marginBottom: 8 }}>Importar varias marcas a la vez</div>
              <div style={{ fontSize: 11.5, color: "#78776f", marginBottom: 10 }}>Descomprimí el ZIP en una carpeta y seleccioná todas las imágenes juntas — el nombre de cada marca sale del nombre del archivo (editable antes de aplicar).</div>
              <input type="file" accept="image/*" multiple onChange={(e) => handleMarcaImportFiles(e.target.files)} style={{ ...inputStyle, padding: "8px", marginBottom: 12 }} />

              {marcaImportEntradas.length > 0 && (
                <>
                  <div style={{ maxHeight: 320, overflowY: "auto", border: "1px solid #eae7dc", borderRadius: 6, marginBottom: 12 }}>
                    {marcaImportEntradas.map((e, i) => (
                      <div key={i} style={{ display: "grid", gridTemplateColumns: "40px 1fr auto", gap: 10, alignItems: "center", padding: "8px 10px", borderBottom: "1px solid #f2f0e9" }}>
                        <img src={URL.createObjectURL(e.file)} alt="" style={{ height: 28, maxWidth: 40, objectFit: "contain" }} />
                        <input style={miniInputStyle} value={e.nombre} onChange={(ev) => actualizarMarcaImportNombre(i, ev.target.value)} placeholder="Nombre de la marca" />
                        <button onClick={() => quitarMarcaImportEntrada(i)} style={{ border: "none", background: "none", color: "#b23b3b", fontSize: 12, cursor: "pointer" }}>Quitar</button>
                      </div>
                    ))}
                  </div>
                  <button onClick={aplicarImportMarcas} disabled={marcaImportAplicando} style={{ padding: "8px 16px", borderRadius: 6, border: "none", background: NAVY, color: "#fff", fontSize: 12.5, cursor: "pointer" }}>
                    {marcaImportAplicando ? "Aplicando…" : `Aplicar (${marcaImportEntradas.length} marca${marcaImportEntradas.length === 1 ? "" : "s"})`}
                  </button>
                </>
              )}
              {marcaImportResumen && <div style={{ marginTop: 8, fontSize: 12, color: "#3a7d44" }}>{marcaImportResumen}</div>}
            </div>

            {logosCargando && <div style={{ fontSize: 13, color: "#78776f" }}>Cargando…</div>}

            <input style={{ ...inputStyle, maxWidth: 360, marginBottom: 18 }} value={logosQuery} onChange={(e) => setLogosQuery(e.target.value)} placeholder="Filtrar por ISIN, nombre de fondo o de marca" />

            {filasUnificadas.length === 0 ? (
              <div style={{ fontSize: 13, color: "#78776f" }}>No hay logos cargados{logosQuery ? " que coincidan con ese filtro" : ""}.</div>
            ) : (
              <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
                {filasUnificadas.map((fila) => {
                  const abierta = filaExpandidaKey === fila.key;
                  return (
                    <div key={fila.key} style={{ background: "#fff", border: `1px solid ${fila.tipo === "sinNombre" ? "#e0b96a" : "#eae7dc"}`, borderRadius: 8, overflow: "hidden" }}>
                      <div onClick={() => toggleFila(fila.key)} style={{ display: "flex", alignItems: "center", gap: 12, padding: 12, cursor: "pointer" }}>
                        <img src={fila.logo_url} alt="" style={{ height: 26, maxWidth: 90, objectFit: "contain" }} />
                        <div style={{ flex: 1, fontSize: 13, fontWeight: 600, color: fila.tipo === "sinNombre" ? "#8a6d1f" : "#333" }}>
                          {fila.marca?.nombre || "(sin nombre — click para asignarle uno)"}
                        </div>
                        <span style={{ fontSize: 11, color: "#a5a399" }}>{fila.fondos.length} fondo{fila.fondos.length === 1 ? "" : "s"}</span>
                        <span style={{ fontSize: 12, color: "#78776f" }}>{abierta ? "▲" : "▼"}</span>
                      </div>

                      {abierta && (
                        <div style={{ padding: "0 14px 14px 14px", borderTop: "1px solid #f2f0e9" }}>

                          {fila.tipo === "marca" && (
                            <div style={{ marginTop: 12, marginBottom: 14 }}>
                              <div style={{ fontSize: 11.5, fontWeight: 600, color: NAVY, marginBottom: 6 }}>Cambiar el logo de esta marca</div>
                              <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
                                <input type="file" accept="image/*" onChange={(e) => setMarcaLogoNuevoArchivo(e.target.files[0])} style={{ ...miniInputStyle, padding: "6px", flex: 1 }} />
                                <button onClick={() => actualizarLogoMarca(fila.marca)} disabled={marcaLogoActualizando || !marcaLogoNuevoArchivo} style={{ padding: "7px 14px", borderRadius: 6, border: "none", background: NAVY, color: "#fff", fontSize: 12, cursor: "pointer", whiteSpace: "nowrap" }}>
                                  {marcaLogoActualizando ? "…" : "Actualizar"}
                                </button>
                              </div>
                              <div style={{ fontSize: 10.5, color: "#a5a399", marginTop: 4 }}>Los fondos que ya tiene asociados van a seguir teniéndolo — no se pierden.</div>
                              {marcaLogoMensaje && <div style={{ marginTop: 6, fontSize: 11.5, color: marcaLogoMensaje.startsWith("Error") ? "#b23b3b" : "#3a7d44" }}>{marcaLogoMensaje}</div>}
                            </div>
                          )}

                          <div style={{ fontSize: 11.5, fontWeight: 600, color: NAVY, marginTop: 12, marginBottom: 6 }}>Fondos asociados</div>
                          {fila.fondos.length === 0 && <div style={{ fontSize: 12, color: "#a5a399", marginBottom: 8 }}>Ninguno todavía.</div>}
                          {fila.fondos.map((f) => (
                            <div key={f.isin} style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 12, padding: "5px 0", borderTop: "1px solid #f2f0e9" }}>
                              <span style={{ flex: 1 }}><b>{f.isin}</b> — {f.nombre}</span>
                              <button onClick={() => quitarFondoDeGrupo(f.isin)} style={{ border: "none", background: "none", color: "#b23b3b", fontSize: 11.5, cursor: "pointer" }}>Quitar</button>
                            </div>
                          ))}

                          {fila.tipo === "sinNombre" ? (
                            <>
                              {marcasTodas.length > 0 && (
                                <div style={{ display: "flex", gap: 6, marginTop: 12, borderTop: "1px solid #f2f0e9", paddingTop: 12, alignItems: "center" }}>
                                  <select
                                    style={{ ...miniInputStyle, padding: "6px 9px", flex: 1 }}
                                    value={marcaFusionSeleccion[fila.logo_url] || ""}
                                    onChange={(e) => setMarcaFusionSeleccion((prev) => ({ ...prev, [fila.logo_url]: e.target.value }))}
                                  >
                                    <option value="">Fusionar con una marca ya existente…</option>
                                    {marcasTodas.map((m) => <option key={m.id} value={m.id}>{m.nombre}</option>)}
                                  </select>
                                  <button
                                    onClick={() => fusionarConMarcaExistente(fila)}
                                    disabled={marcaFusionando === fila.logo_url || !marcaFusionSeleccion[fila.logo_url]}
                                    style={{ padding: "6px 14px", borderRadius: 6, border: "none", background: NAVY, color: "#fff", fontSize: 12, cursor: "pointer", whiteSpace: "nowrap" }}
                                  >
                                    {marcaFusionando === fila.logo_url ? "…" : "Fusionar"}
                                  </button>
                                </div>
                              )}
                              <div style={{ fontSize: 10.5, color: "#a5a399", margin: "6px 0" }}>o, si es una marca nueva de verdad:</div>
                              <div style={{ display: "flex", gap: 6 }}>
                                <input
                                  style={{ ...miniInputStyle, padding: "6px 9px" }}
                                  placeholder="Nombre de marca (ej: MFS)"
                                  value={marcaNombrePorGrupo[fila.logo_url] || ""}
                                  onChange={(e) => setMarcaNombrePorGrupo((prev) => ({ ...prev, [fila.logo_url]: e.target.value }))}
                                />
                                <button
                                  onClick={() => guardarGrupoComoMarca(fila.logo_url)}
                                  disabled={marcaGuardandoGrupo === fila.logo_url || !(marcaNombrePorGrupo[fila.logo_url] || "").trim()}
                                  style={{ padding: "6px 14px", borderRadius: 6, border: "none", background: NAVY, color: "#fff", fontSize: 12, cursor: "pointer", whiteSpace: "nowrap" }}
                                >
                                  {marcaGuardandoGrupo === fila.logo_url ? "…" : "Guardar nombre"}
                                </button>
                              </div>
                            </>
                          ) : (
                            <div style={{ marginTop: 12, borderTop: "1px solid #f2f0e9", paddingTop: 12 }}>
                              <div style={{ fontSize: 11.5, fontWeight: 600, color: NAVY, marginBottom: 6 }}>Agregar más fondos</div>
                              <input
                                style={inputStyle}
                                value={marcaAsocQuery}
                                onChange={(e) => setMarcaAsocQuery(e.target.value)}
                                placeholder={`Buscar por ISIN o nombre (ej: "${fila.marca.nombre.split(" ")[0]}")`}
                              />
                              {marcaAsocResultados.length > 0 && (
                                <div style={{ marginTop: 8, maxHeight: 220, overflowY: "auto", border: "1px solid #eae7dc", borderRadius: 6 }}>
                                  {marcaAsocResultados.map((f) => (
                                    <label key={f.isin} style={{ display: "flex", alignItems: "center", gap: 8, padding: "6px 10px", fontSize: 12.5, borderBottom: "1px solid #f2f0e9", cursor: "pointer" }}>
                                      <input type="checkbox" checked={!!marcaAsocSeleccionados[f.isin]} onChange={() => toggleFondoAsoc(f)} />
                                      {f.logo_url && <img src={f.logo_url} alt="" style={{ height: 16, opacity: f.logo_url === fila.logo_url ? 1 : 0.5 }} />}
                                      <b>{f.isin}</b> — {f.nombre}
                                      {f.logo_url && f.logo_url !== fila.logo_url && <span style={{ color: "#b23b3b", fontSize: 11 }}>(ya tiene otro logo)</span>}
                                    </label>
                                  ))}
                                </div>
                              )}
                              {Object.keys(marcaAsocSeleccionados).length > 0 && (
                                <div style={{ marginTop: 10, fontSize: 12, color: "#78776f" }}>
                                  {Object.keys(marcaAsocSeleccionados).length} fondo(s) tildado(s): {Object.values(marcaAsocSeleccionados).join(", ")}
                                </div>
                              )}
                              <button
                                onClick={() => aplicarAsociacionMarca(fila.marca)}
                                disabled={marcaAsocGuardando || Object.keys(marcaAsocSeleccionados).length === 0}
                                style={{ marginTop: 10, padding: "8px 16px", borderRadius: 6, border: "none", background: NAVY, color: "#fff", fontSize: 12.5, cursor: "pointer" }}
                              >
                                {marcaAsocGuardando ? "Guardando…" : `Asociar a ${Object.keys(marcaAsocSeleccionados).length || ""} fondo(s)`}
                              </button>
                              {marcaAsocMensaje && <div style={{ marginTop: 8, fontSize: 12, color: marcaAsocMensaje.startsWith("Error") ? "#b23b3b" : "#3a7d44" }}>{marcaAsocMensaje}</div>}
                            </div>
                          )}
                        </div>
                      )}
                    </div>
                  );
                })}
              </div>
            )}
          </div>
        ) : (
        <div style={{ padding: "28px 36px", display: "flex", gap: 24 }}>
          <div style={{ width: 360 }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
              <h3 style={{ color: NAVY, fontSize: 16, marginBottom: 8 }}>Biblioteca de fondos</h3>
              <div style={{ display: "flex", gap: 12 }}>
                <button onClick={() => setLogosVista(true)} style={{ border: "none", background: "none", color: TEAL, fontSize: 12, cursor: "pointer", textDecoration: "underline" }}>Ver todos los logos</button>
                <button onClick={() => setImportModo(true)} style={{ border: "none", background: "none", color: TEAL, fontSize: 12, cursor: "pointer", textDecoration: "underline" }}>Importar masivo</button>
              </div>
            </div>

            <div style={{ background: "#fff", border: "1px dashed #d8d5cc", borderRadius: 8, padding: 12, margin: "12px 0" }}>
              <div style={{ fontSize: 12, fontWeight: 600, color: NAVY, marginBottom: 4 }}>Importar biblioteca base (Fondos / Fondos distributivos / Acciones / Bonos)</div>
              <div style={{ fontSize: 11, color: "#78776f", marginBottom: 8 }}>El Excel con las pestañas que mantiene el equipo — carga directo por ISIN/Ticker, sin pedir revisión.</div>
              <div style={{ display: "flex", alignItems: "center", gap: 6, marginBottom: 6 }}>
                <input type="checkbox" id="divid-biblioteca" checked={descargaConDividendos} onChange={(e) => setDescargaConDividendos(e.target.checked)} />
                <label htmlFor="divid-biblioteca" style={{ fontSize: 11, color: "#78776f" }}>Incluir columnas de dividendos en la pestaña Fondos</label>
                <button onClick={() => descargarPlantillaBase(descargaConDividendos)} style={{ border: "none", background: "none", color: TEAL, fontSize: 11, cursor: "pointer", textDecoration: "underline", padding: 0, marginLeft: 4 }}>Descargar plantilla en blanco</button>
              </div>
              <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
                <input type="file" accept=".xlsx,.xls" onChange={(e) => handleImportBibliotecaBase(e.target.files[0])} style={{ ...miniInputStyle, padding: "6px", flex: 1 }} />
                {baseImportCargando && <span style={{ fontSize: 11.5, color: "#78776f" }}>Cargando…</span>}
              </div>
              {baseImportResumen && <div style={{ marginTop: 8, fontSize: 11.5, color: baseImportResumen.startsWith("Error") ? "#b23b3b" : "#3a7d44" }}>{baseImportResumen}</div>}
              {renderBaseImportPreview()}
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

              <Field label="Buscar logo por marca (ej: MFS, BlackRock, Vontobel)" hint="Reutiliza un logo ya cargado — no sube ningún archivo nuevo.">
                <input style={inputStyle} value={marcaFondoQuery} onChange={(e) => { setMarcaFondoQuery(e.target.value); setMarcaFondoAsignada(""); }} placeholder="Escribí el nombre de la marca" />
              </Field>
              {marcaFondoResultados.length > 0 && (
                <div style={{ border: "1px solid #eae7dc", borderRadius: 6, marginTop: -8, marginBottom: 14, maxHeight: 160, overflowY: "auto" }}>
                  {marcaFondoResultados.map((m) => (
                    <div key={m.id} onClick={() => asignarMarcaAFondo(m)} style={{ display: "flex", alignItems: "center", gap: 8, padding: "6px 10px", fontSize: 12.5, cursor: "pointer", borderBottom: "1px solid #f2f0e9" }}>
                      <img src={m.logo_url} alt="" style={{ height: 20, maxWidth: 70, objectFit: "contain" }} />
                      {m.nombre}
                    </div>
                  ))}
                </div>
              )}
              {marcaFondoAsignada && (
                <div style={{ fontSize: 12, color: TEAL, marginBottom: 14 }}>✓ Logo de "{marcaFondoAsignada}" asignado — no olvides apretar Guardar.</div>
              )}

              <Field label="O subir un logo nuevo para este fondo puntual (imagen)">
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
                  {["Fecha", "Tipo", "Cliente / Cuenta", "Hecha por", "Asesor", "Monto", "Status", "Archivos", ""].map((h) => (
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
                    <td style={{ padding: "8px 10px" }}>
                      <button onClick={() => eliminarPropuesta(r.id)} style={{ border: "none", background: "none", color: "#b23b3b", fontSize: 12, cursor: "pointer" }}>Eliminar</button>
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
              <Field label="Asesor / RepCode (opcional)" hint="Se puede armar la propuesta sin asociarla a un asesor puntual.">
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
                <div style={{ display: "flex", alignItems: "center", gap: 10, fontSize: 12.5, color: "#78776f", background: CREAM, borderRadius: 6, padding: "8px 10px", marginBottom: 18 }}>
                  <span style={{ flex: 1 }}>{asesorSel.email}</span>
                  <button onClick={() => { setAsesorSel(null); setRepcode(""); setAsesorQuery(""); }} style={{ border: "none", background: "none", color: "#b23b3b", fontSize: 12, cursor: "pointer", textDecoration: "underline" }}>Quitar</button>
                </div>
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
            <Section title="Equipo" subtitle="Hasta 6 personas, en 2 filas de 3 (1-2-3 arriba, 4-5-6 abajo) según el orden de la lista. Click en un nombre para editarlo. Agregá, sacá y reordená con las flechas.">
              {team.map((m, idx) => {
                const abierto = equipoExpandidoId === m.id;
                return (
                  <div key={m.id} style={{ background: "#fff", border: "1px solid #eae7dc", borderRadius: 8, marginBottom: 10, overflow: "hidden" }}>
                    <div style={{ display: "flex", alignItems: "center", gap: 8, padding: 12 }}>
                      <input type="checkbox" checked={m.incluido} onChange={() => setTeam((prev) => prev.map((x) => x.id === m.id ? { ...x, incluido: !x.incluido } : x))} />
                      <div onClick={() => setEquipoExpandidoId((prev) => (prev === m.id ? null : m.id))} style={{ flex: 1, cursor: "pointer" }}>
                        <div style={{ fontSize: 13.5, fontWeight: 600 }}>{m.nombre || "(sin nombre — click para completar)"}</div>
                        <div style={{ fontSize: 11.5, color: "#78776f" }}>{m.puesto}{idx < 6 ? ` · Posición ${idx + 1} de 6` : " · No entra en el documento (solo los primeros 6)"}</div>
                      </div>
                      <button onClick={() => moverTeam(idx, -1)} disabled={idx === 0} style={{ border: "none", background: "none", cursor: idx === 0 ? "default" : "pointer", color: idx === 0 ? "#ccc" : "#78776f", fontSize: 13, padding: "0 4px" }}>▲</button>
                      <button onClick={() => moverTeam(idx, 1)} disabled={idx === team.length - 1} style={{ border: "none", background: "none", cursor: idx === team.length - 1 ? "default" : "pointer", color: idx === team.length - 1 ? "#ccc" : "#78776f", fontSize: 13, padding: "0 4px" }}>▼</button>
                      <button onClick={() => setTeam((prev) => prev.filter((x) => x.id !== m.id))} style={{ border: "none", background: "none", color: "#b23b3b", fontSize: 12, cursor: "pointer" }}>Quitar</button>
                    </div>

                    {abierto && (
                      <div style={{ padding: "0 12px 12px 12px", borderTop: "1px solid #f2f0e9" }}>
                        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8, marginTop: 10, marginBottom: 8 }}>
                          <MiniField label="Nombre">
                            <input style={miniInputStyle} value={m.nombre} onChange={(e) => setTeam((prev) => prev.map((x) => x.id === m.id ? { ...x, nombre: e.target.value } : x))} />
                          </MiniField>
                          <MiniField label="Puesto">
                            <input style={miniInputStyle} value={m.puesto} onChange={(e) => setTeam((prev) => prev.map((x) => x.id === m.id ? { ...x, puesto: e.target.value } : x))} />
                          </MiniField>
                        </div>
                        <MiniField label="Educación">
                          <input style={miniInputStyle} value={m.educacion} onChange={(e) => setTeam((prev) => prev.map((x) => x.id === m.id ? { ...x, educacion: e.target.value } : x))} />
                        </MiniField>
                        <div style={{ marginTop: 8 }}>
                          <MiniField label="Foto (opcional — si no se sube, se usa la del template si el nombre coincide, o iniciales)">
                            <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                              <input type="file" accept="image/*" onChange={(e) => handleTeamFoto(m.id, e.target.files[0])} style={{ ...miniInputStyle, padding: "6px", flex: 1 }} />
                              {m.foto_base64 && <span style={{ fontSize: 11, color: TEAL }}>✓ foto cargada</span>}
                            </div>
                          </MiniField>
                        </div>
                      </div>
                    )}
                  </div>
                );
              })}
              <button onClick={() => { const nuevoId = Date.now(); setTeam((prev) => [...prev, { id: nuevoId, nombre: "", puesto: "", educacion: "", incluido: true }]); setEquipoExpandidoId(nuevoId); }} style={{ marginTop: 4, padding: "6px 12px", borderRadius: 6, border: "1px dashed #b8b5a9", background: "none", cursor: "pointer", fontSize: 12.5 }}>+ Agregar integrante</button>
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
                <div style={{ fontSize: 12, fontWeight: 600, color: NAVY, marginBottom: 4 }}>Cargar instrumentos desde Excel para esta propuesta</div>
                <div style={{ fontSize: 11.5, color: "#78776f", marginBottom: 8 }}>
                  Subí el Excel base (Fondos / Acciones / Bonos) — se agregan directo a la tabla de abajo. Si el archivo trae % o Inversión (USD), se usan; si falta alguno de los dos, se calcula solo. No toca la Biblioteca de fondos.
                </div>
                <div style={{ display: "flex", alignItems: "center", gap: 6, marginBottom: 8 }}>
                  <input type="checkbox" id="divid-portafolio" checked={descargaConDividendos} onChange={(e) => setDescargaConDividendos(e.target.checked)} />
                  <label htmlFor="divid-portafolio" style={{ fontSize: 11.5, color: "#78776f" }}>Incluir dividendos</label>
                  <button onClick={() => descargarPlantillaBase(descargaConDividendos)} style={{ border: "none", background: "none", color: TEAL, fontSize: 11.5, cursor: "pointer", textDecoration: "underline", padding: 0, marginLeft: 4 }}>Descargar plantilla en blanco</button>
                </div>
                <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
                  <input type="file" accept=".xlsx,.xls" onChange={(e) => handleImportExcelPropuesta(e.target.files[0])} style={{ ...miniInputStyle, padding: "6px", flex: 1 }} />
                </div>
                {propuestaImportResumen && <div style={{ marginTop: 8, fontSize: 11.5, color: propuestaImportResumen.startsWith("Error") || propuestaImportResumen.startsWith("No se") ? "#b23b3b" : "#3a7d44" }}>{propuestaImportResumen}</div>}
              </div>

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

              {TIPO_ORDEN.filter((tipo) => proposedAssets.some((a) => (a.tipo_instrumento || "fondo") === tipo)).map((tipo) => {
                const filas = proposedAssets.map((a, i) => ({ a, i })).filter(({ a }) => (a.tipo_instrumento || "fondo") === tipo);
                const colsOrdenadas = [
                  ...COLUMNAS_POR_TIPO[tipo].filter((c) => c.fijo),
                  ...columnasConfig[tipo].map((key) => COLUMNAS_POR_TIPO[tipo].find((c) => c.key === key)).filter(Boolean),
                ];
                return (
                  <div key={tipo} style={{ marginBottom: 26 }}>
                    <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 8 }}>
                      <div style={{ fontSize: 12.5, fontWeight: 700, color: NAVY }}>{TIPO_LABELS[tipo]} ({filas.length})</div>
                      <button onClick={() => setColumnasAbiertoPara((prev) => (prev === tipo ? null : tipo))} style={{ border: "none", background: "none", color: TEAL, fontSize: 11.5, cursor: "pointer", textDecoration: "underline" }}>
                        {columnasAbiertoPara === tipo ? "Cerrar columnas" : "Elegir columnas"}
                      </button>
                    </div>

                    {columnasAbiertoPara === tipo && (
                      <div style={{ background: "#fff", border: "1px solid #eae7dc", borderRadius: 6, padding: 10, marginBottom: 10, maxWidth: 320 }}>
                        <div style={{ fontSize: 11, fontWeight: 600, color: NAVY, marginBottom: 6 }}>Tildar para mostrar — flechas para ordenar</div>
                        {COLUMNAS_POR_TIPO[tipo].filter((c) => !c.fijo).map((c) => {
                          const idx = columnasConfig[tipo].indexOf(c.key);
                          const visible = idx !== -1;
                          return (
                            <div key={c.key} style={{ display: "flex", alignItems: "center", gap: 6, padding: "3px 0" }}>
                              <input type="checkbox" checked={visible} onChange={() => toggleColumnaVisible(tipo, c.key)} />
                              <span style={{ fontSize: 12, flex: 1 }}>{c.label}</span>
                              {visible && (
                                <>
                                  <button onClick={() => moverColumna(tipo, c.key, -1)} disabled={idx === 0} style={{ border: "none", background: "none", cursor: idx === 0 ? "default" : "pointer", color: idx === 0 ? "#ccc" : "#78776f", fontSize: 12, padding: "0 3px" }}>▲</button>
                                  <button onClick={() => moverColumna(tipo, c.key, 1)} disabled={idx === columnasConfig[tipo].length - 1} style={{ border: "none", background: "none", cursor: idx === columnasConfig[tipo].length - 1 ? "default" : "pointer", color: idx === columnasConfig[tipo].length - 1 ? "#ccc" : "#78776f", fontSize: 12, padding: "0 3px" }}>▼</button>
                                </>
                              )}
                            </div>
                          );
                        })}
                      </div>
                    )}

                    <div style={{ overflowX: "auto", background: "#fff", border: "1px solid #eae7dc", borderRadius: 8 }}>
                      <table style={{ borderCollapse: "collapse", width: "100%", fontSize: 12 }}>
                        <thead>
                          <tr style={{ background: CREAM }}>
                            <th style={{ padding: "6px 8px", textAlign: "left", whiteSpace: "nowrap" }}>Categoría</th>
                            <th style={{ padding: "6px 8px", textAlign: "left", whiteSpace: "nowrap" }}>%</th>
                            <th style={{ padding: "6px 8px", textAlign: "left", whiteSpace: "nowrap" }}>Monto (USD)</th>
                            {colsOrdenadas.map((c) => (
                              <th key={c.key} style={{ padding: "6px 8px", textAlign: "left", whiteSpace: "nowrap" }}>{c.label}</th>
                            ))}
                            <th></th>
                          </tr>
                        </thead>
                        <tbody>
                          {filas.map(({ a, i }) => (
                            <tr key={i} style={{ borderTop: "1px solid #f2f0e9" }}>
                              <td style={{ padding: "4px 8px" }}>
                                <select style={{ ...miniInputStyle, padding: "4px 6px", fontSize: 11.5 }} value={a.categoria} onChange={(e) => updateProposedField(i, "categoria", e.target.value)}>
                                  {CATEGORIAS.map((c) => <option key={c} value={c}>{c}</option>)}
                                </select>
                              </td>
                              <td style={{ padding: "4px 8px" }}>
                                <input type="number" style={{ ...miniInputStyle, padding: "4px 6px", width: 60, fontSize: 11.5 }} value={a.pct} onChange={(e) => updateProposedField(i, "pct", +e.target.value)} />
                              </td>
                              <td style={{ padding: "4px 8px" }}>
                                <input type="number" style={{ ...miniInputStyle, padding: "4px 6px", width: 90, fontSize: 11.5 }} value={a.monto} onChange={(e) => updateProposedField(i, "monto", +e.target.value)} />
                              </td>
                              {colsOrdenadas.map((c) => (
                                <td key={c.key} style={{ padding: "4px 8px" }}>
                                  {c.fijo ? (
                                    <span>{a[c.key]}</span>
                                  ) : c.calculado ? (
                                    <span style={{ color: "#78776f" }}>
                                      {c.key === "dividendo_anual"
                                        ? Math.round((a.monto || 0) * (a.dividendo_pct || 0) / 100).toLocaleString()
                                        : c.key === "cupon_anual"
                                        ? Math.round((a.monto || 0) * (a.cupon_pct || 0) / 100).toLocaleString()
                                        : ""}
                                    </span>
                                  ) : c.tipo === "date" ? (
                                    <input type="date" style={{ ...miniInputStyle, padding: "4px 6px", fontSize: 11.5 }} value={a[c.key] || ""} onChange={(e) => updateProposedField(i, c.key, e.target.value)} />
                                  ) : (
                                    <input
                                      type={c.tipo === "number" ? "number" : "text"}
                                      style={{ ...miniInputStyle, padding: "4px 6px", width: c.tipo === "number" ? 70 : 100, fontSize: 11.5 }}
                                      value={a[c.key] ?? ""}
                                      onChange={(e) => updateProposedField(i, c.key, c.tipo === "number" ? +e.target.value : e.target.value)}
                                    />
                                  )}
                                </td>
                              ))}
                              <td style={{ padding: "4px 8px", whiteSpace: "nowrap" }}>
                                <button onClick={() => toggleFavorito(i)} title="Marcar/desmarcar como fondo frecuente" style={{ border: "none", background: "none", cursor: "pointer", fontSize: 14, color: a.uso_frecuente ? TEAL : "#d8d5cc", padding: 0, marginRight: 8 }}>★</button>
                                <button onClick={() => setProposedAssets((prev) => prev.filter((_, j) => j !== i))} style={{ border: "none", background: "none", color: "#b23b3b", fontSize: 11, cursor: "pointer" }}>Quitar</button>
                              </td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  </div>
                );
              })}
            </Section>
          )}

          {stepName === "Descripción de activos" && (
            <Section title="Descripción de activos" subtitle="Elegí a mano, por categoría, qué fondos van en cada página de descripción — con su logo, descripción y factsheet ya asociados desde la biblioteca. No depende de lo que hayas cargado en Portafolio propuesto.">
              <button onClick={prellenarDescDesdePortafolio} style={{ marginBottom: 16, padding: "6px 12px", borderRadius: 6, border: "1px dashed #b8b5a9", background: "none", cursor: "pointer", fontSize: 12.5 }}>Rellenar automático desde Portafolio propuesto</button>

              <Field label="Buscar fondo en la biblioteca">
                <div style={{ display: "flex", gap: 8 }}>
                  <select style={{ ...miniInputStyle, maxWidth: 220 }} value={descCategoriaDestino} onChange={(e) => setDescCategoriaDestino(e.target.value)}>
                    {DESC_CATEGORIAS.map((c) => <option key={c} value={c}>{c}</option>)}
                  </select>
                  <input style={inputStyle} value={descQuery} onChange={(e) => setDescQuery(e.target.value)} placeholder="Buscar por ISIN o nombre, se agrega a la categoría de arriba" />
                </div>
              </Field>
              {descResultados.length > 0 && (
                <div style={{ border: "1px solid #eae7dc", borderRadius: 6, marginBottom: 18, maxHeight: 180, overflowY: "auto" }}>
                  {descResultados.map((f) => (
                    <div key={f.isin} onClick={() => addDescManual(f)} style={{ padding: "8px 10px", fontSize: 12.5, cursor: "pointer", borderBottom: "1px solid #f2f0e9", display: "flex", alignItems: "center", gap: 8 }}>
                      {f.logo_url && <img src={f.logo_url} alt="" style={{ height: 18 }} />}
                      <b>{f.isin}</b> — {f.nombre} {!f.descripcion && <span style={{ color: "#b23b3b" }}>(sin descripción todavía)</span>}
                    </div>
                  ))}
                </div>
              )}

              {DESC_CATEGORIAS.map((cat) => (
                <div key={cat} style={{ marginBottom: 22 }}>
                  <div style={{ fontSize: 12.5, fontWeight: 700, color: NAVY, marginBottom: 8 }}>{cat} ({descSeleccion[cat].length})</div>
                  {descSeleccion[cat].length === 0 && <div style={{ fontSize: 12, color: "#a5a399", marginBottom: 8 }}>Sin fondos elegidos — esta página no va a aparecer en el documento.</div>}
                  {descSeleccion[cat].map((f) => (
                    <div key={f.isin} style={{ display: "flex", alignItems: "center", gap: 10, background: "#fff", border: "1px solid #eae7dc", borderRadius: 6, padding: "8px 10px", marginBottom: 6 }}>
                      {f.logo_url ? <img src={f.logo_url} alt="" style={{ height: 24 }} /> : <div style={{ width: 24 }} />}
                      <div style={{ flex: 1 }}>
                        <div style={{ fontSize: 13 }}>{f.nombre}</div>
                        <div style={{ fontSize: 11, color: f.descripcion ? "#78776f" : "#b23b3b" }}>{f.descripcion ? f.descripcion.slice(0, 90) + (f.descripcion.length > 90 ? "…" : "") : "Sin descripción cargada en la biblioteca todavía"}</div>
                      </div>
                      <button onClick={() => quitarDescManual(cat, f.isin)} style={{ border: "none", background: "none", color: "#b23b3b", fontSize: 12, cursor: "pointer" }}>Quitar</button>
                    </div>
                  ))}
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
