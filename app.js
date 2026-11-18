/* ==========================================================================
   Esencial Studio — Creador de Presupuestos y Cotizaciones (HTML → DOCX / PDF)
   ========================================================================== */

const STORAGE_CURRENT_KEY = "esencial_presupuesto_current_v1";
const STORAGE_HISTORY_KEY = "esencial_presupuesto_history_v1";

// Activos gráficos oficiales embebidos en base64 (100% compatibles offline y en file://)
const ISOTIPO_B64 =
  "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";

const ESENCIAL_B64 =
  "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";

const STUDIO_B64 =
  "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";

const ISOTIPO_DATA_URI = "data:image/png;base64," + ISOTIPO_B64;
const ESENCIAL_DATA_URI = "data:image/png;base64," + ESENCIAL_B64;
const STUDIO_DATA_URI = "data:image/png;base64," + STUDIO_B64;

const DEFAULT_INTRO = {
  introGreeting: "De antemano, agradecemos su amable solicitud de presupuesto para este proyecto.",
  introBody:
    "El costo de la creación del concepto, las decisiones de diseño y la producción de los planos del proyecto se sintetiza en determinados conceptos de trabajo. Los costos se basan en las dimensiones del proyecto, la calidad de nuestro trabajo y servicio, así como en los tiempos de entrega requeridos por el cliente, entre otros factores.",
  introLead: "A continuación, el desglose de los conceptos:"
};

const COMMON_DELIVERABLES = [
  "Plantas arquitectónicas.",
  "Cortes arquitectónicos.",
  "Fachadas arquitectónicas.",
  "Acabados y despiece de pisos",
  "Albañilería, Trazo y Detalles especiales.",
  "Carpintería de mobiliario fijo.",
  "Cancelería y cristalería fija.",
  "Herrería fija.",
  "Detalles de baños",
  "Criterio de iluminación.",
  "Criterio de salidas y contactos.",
  "Perspectivas calidad HD",
  "Plantas editadas para venta"
];

const SCOPE_PRESETS = [
  {
    id: "integral",
    name: "Conceptual + Arq. + Comercial",
    badge: "Paquete integral",
    data: {
      title: "Propuesta Conceptual, Proyecto Arquitectónico y Material Comercial.",
      subtitle: "",
      leadText: "Diseño integral del proyecto arquitectónico, definición de concepto y material visual para presentación o venta.",
      groups: [
        {
          heading: "Propuesta conceptual",
          items: [
            "Investigación sobre el giro del proyecto y cuadros de áreas.",
            "Definición de concepto.",
            "Definición de esquema.",
            "Definición de programa arquitectónico."
          ]
        },
        {
          heading: "Proyecto arquitectónico",
          items: [
            "Plantas arquitectónicas.",
            "Cortes arquitectónicos.",
            "Fachadas arquitectónicas."
          ]
        },
        {
          heading: "Material comercial",
          items: [
            "8 perspectivas calidad HD",
            "Plantas editadas para venta"
          ]
        }
      ],
      notes: "",
      amount: 120000,
      currencyLabel: "MXN",
      included: true
    }
  },
  {
    id: "conceptual",
    name: "Propuesta Conceptual",
    badge: "Etapa 01",
    data: {
      title: "Propuesta Conceptual",
      subtitle: "",
      leadText: "Conceptualización arquitectónica inicial, análisis del terreno y programa de necesidades.",
      groups: [
        {
          heading: "Propuesta conceptual",
          items: [
            "Investigación sobre el giro del proyecto y cuadros de áreas.",
            "Definición de concepto y partido arquitectónico.",
            "Definición de zonificación y esquema.",
            "Definición de programa arquitectónico."
          ]
        }
      ],
      notes: "",
      amount: 45000,
      currencyLabel: "MXN",
      included: true
    }
  },
  {
    id: "arquitectonico",
    name: "Proyecto Arquitectónico",
    badge: "Etapa 02",
    data: {
      title: "Proyecto Arquitectónico",
      subtitle: "",
      leadText: "Desarrollo arquitectónico completo a nivel anteproyecto definitivo.",
      groups: [
        {
          heading: "Proyecto arquitectónico",
          items: [
            "Plantas arquitectónicas acotadas y ambientadas.",
            "Cortes arquitectónicos generales.",
            "Fachadas arquitectónicas.",
            "Planta de conjunto y azoteas."
          ]
        }
      ],
      notes: "",
      amount: 65000,
      currencyLabel: "MXN",
      included: true
    }
  },
  {
    id: "comercial",
    name: "Material Comercial / Renders",
    badge: "Visualización",
    data: {
      title: "Material Comercial y Visualización",
      subtitle: "",
      leadText: "Producción de visualizaciones fotorrealistas y láminas comerciales de presentación.",
      groups: [
        {
          heading: "Material comercial",
          items: [
            "10 perspectivas calidad HD (interiores y exteriores)",
            "Plantas editadas y ambientadas para venta"
          ]
        }
      ],
      notes: "",
      amount: 35000,
      currencyLabel: "MXN",
      included: true
    }
  },
  {
    id: "ejecutivo",
    name: "Proyecto Ejecutivo",
    badge: "Sub-especialidades",
    data: {
      title: "Proyecto Ejecutivo",
      subtitle: "",
      leadText: "",
      groups: [
        {
          heading: "Planos de sub-especialidades (con ubicación y especificaciones)",
          items: [
            "Acabados y despiece de pisos",
            "Albañilería, Trazo y Detalles especiales.",
            "Carpintería de mobiliario fijo.",
            "Cancelería y cristalería fija.",
            "Herrería fija.",
            "Detalles de baños",
            "Demolición (en caso de ser necesarios)"
          ]
        },
        {
          heading: "Planos de criterios técnicos (no incluyen cálculo o diseño)",
          items: [
            "Criterio de iluminación.",
            "Criterio de salidas y contactos."
          ]
        }
      ],
      notes:
        "Adecuaciones de memorias y planos requeridos por desarrollo urbano e INAH. (No incluye gestiones ante las dependencias)\n\nSe incluye coordinación y revisión de ingenierías.",
      amount: 50000,
      currencyLabel: "MXN",
      included: true
    }
  },
  {
    id: "interiores",
    name: "Diseño de Interiores",
    badge: "Interiorismo",
    data: {
      title: "Diseño de Interiores y Mobiliario",
      subtitle: "",
      leadText: "Selección de paleta de materiales, iluminación ambiental, carpinterías y catálogo de mobiliario.",
      groups: [
        {
          heading: "Alcance de interiorismo",
          items: [
            "Moodboards y paleta oficial de acabados y texturas.",
            "Alzados interiores por espacio.",
            "Diseño de carpintería y mobiliario fijo a medida.",
            "Catálogo de mobiliario suelto, luminarias y accesorios."
          ]
        }
      ],
      notes: "",
      amount: 55000,
      currencyLabel: "MXN",
      included: true
    }
  },
  {
    id: "supervision",
    name: "Supervisión Arquitectónica",
    badge: "En obra",
    data: {
      title: "Supervisión Arquitectónica de Obra",
      subtitle: "",
      leadText: "Acompañamiento técnico y estético mediante visitas programadas durante la ejecución de la obra.",
      groups: [
        {
          heading: "Actividades incluidas",
          items: [
            "Visitas semanales de revisión de calidad y apego a proyecto.",
            "Resolución de dudas en sitio con contratistas y residentes.",
            "Validación de muestras de acabados, cancelería y carpintería.",
            "Emisión de minutas de seguimiento de obra."
          ]
        }
      ],
      notes: "No incluye administración de obra ni control presupuestal de contratistas.",
      amount: 30000,
      currencyLabel: "MXN",
      included: true
    }
  },
  {
    id: "levantamiento",
    name: "Levantamiento Arquitectónico",
    badge: "Estado actual",
    data: {
      title: "Levantamiento Arquitectónico y Topográfico",
      subtitle: "",
      leadText: "Medición física y digitalización del estado actual del inmueble o predio existente.",
      groups: [
        {
          heading: "Entregables incluidos",
          items: [
            "Medición física en sitio y registro fotográfico.",
            "Planos arquitectónicos de estado actual.",
            "Identificación de estructura e instalaciones visibles."
          ]
        }
      ],
      notes: "",
      amount: 18000,
      currencyLabel: "MXN",
      included: true
    }
  }
];

const PAYMENT_PRESETS = {
  "40-30-20-10": [
    { percent: 40, text: "El primer pago constituye al 40% del valor total por el proyecto y corresponde al anticipo." },
    { percent: 30, text: "El segundo pago constituye al 30% del valor total por el proyecto y corresponde a la entrega del anteproyecto." },
    { percent: 20, text: "El tercer pago constituye al 20% del valor total por el proyecto y corresponde a la entrega del proyecto arquitectónico." },
    { percent: 10, text: "Se saldará con el 10% restante del valor del proyecto, y corresponde a la entrega del proyecto ejecutivo completo." }
  ],
  "50-50": [
    { percent: 50, text: "El primer pago constituye al 50% del valor total por el proyecto y corresponde al anticipo para el inicio de los trabajos." },
    { percent: 50, text: "El segundo pago constituye al 50% restante del valor del proyecto y corresponde a la entrega final del paquete completo." }
  ],
  "40-30-30": [
    { percent: 40, text: "El primer pago constituye al 40% del valor total por el proyecto y corresponde al anticipo." },
    { percent: 30, text: "El segundo pago constituye al 30% del valor total por el proyecto y corresponde a la entrega de la propuesta conceptual y anteproyecto." },
    { percent: 30, text: "Se saldará con el 30% restante del valor del proyecto contra la entrega del proyecto definitivo." }
  ],
  "30-30-30-10": [
    { percent: 30, text: "El primer pago constituye al 30% del valor total por el proyecto y corresponde al anticipo." },
    { percent: 30, text: "El segundo pago constituye al 30% del valor total por el proyecto y corresponde a la entrega del anteproyecto." },
    { percent: 30, text: "El tercer pago constituye al 30% del valor total por el proyecto y corresponde a la entrega del proyecto arquitectónico." },
    { percent: 10, text: "Se saldará con el 10% restante del valor del proyecto, y corresponde a la entrega del proyecto ejecutivo completo." }
  ]
};

const DEFAULT_OBSERVATIONS = [
  {
    id: "rounds",
    enabled: true,
    text: "El proyecto contempla hasta 2 rondas de ajustes conceptuales posteriores a la primera propuesta. Rondas adicionales serán cotizadas por separado."
  },
  {
    id: "pause",
    enabled: true,
    text: "Los trabajos podrán pausarse en caso de no realizar el pago correspondiente a los entregables. Los días de gracia se negociarán con Esencial Studio."
  },
  {
    id: "city",
    enabled: true,
    text: "Las revisiones están consideradas en la ciudad de Mérida o por videollamada."
  },
  {
    id: "obra",
    enabled: true,
    text: "Los ajustes de proyecto solicitados durante el proceso de obra no están comprendidos en este presupuesto."
  },
  {
    id: "validity",
    enabled: true,
    text: "Los precios incluidos en este presupuesto tienen una vigencia de 90 días naturales para su aceptación y firma. Una vez iniciado el proyecto y pagado el anticipo, los costos se mantendrán fijos durante todo el periodo de desarrollo del proyecto."
  },
  {
    id: "permits",
    enabled: true,
    text: "Este presupuesto incluye adecuaciones de los planos de permiso, pero no incluye ni las firmas ni gestiones requeridas por parte del ayuntamiento de la ciudad o alguna otra dependencia, para la obtención de las licencias de construcción y permisos."
  },
  {
    id: "engineering",
    enabled: true,
    text: "Los cálculos estructurales y cálculo de las demás ingenierías del edificio no están comprendidos en este presupuesto."
  },
  {
    id: "copyright",
    enabled: true,
    text: "Los derechos de autor de este proyecto son propiedad de los autores, y serán propiedad del cliente cuando el proyecto haya sido saldado en su totalidad el concepto correspondiente."
  },
  {
    id: "delays",
    enabled: true,
    text: "Las revisiones retrasadas por parte del cliente que excedan los 5 días hábiles, permitirán extender los tiempos de entrega."
  }
];

const DEFAULT_PRESUPUESTO = {
  quoteEyebrow: "PRESUPUESTO",
  projectName: "MARBELLO AJUSTES",
  clientName: "Marbello",
  quoteCode: "ES_MARBELLO AJUSTES_26 OCT 2026",
  quoteDateText: "13 de octubre del 2026",
  paperTheme: "white",
  includeCoverPage: true,
  includeIntroText: true,
  showPageNumbers: false,
  ...DEFAULT_INTRO,
  scopes: [
    {
      title: "Propuesta Conceptual, Proyecto Arquitectónico y Material Comercial.",
      subtitle: "(Entrega anticipada para el 10 de diciembre)",
      leadText:
        "Diseño de amenidades, siendo ajustes de proyecto, tanto en planos arquitectónicos, de ejecutivo y perspectivas comerciales.",
      groups: [
        {
          heading: "Propuesta conceptual",
          items: [
            "Investigación sobre el giro del proyecto y cuadros de áreas.",
            "Definición de concepto.",
            "Definición de esquema.",
            "Definición de programa arquitectónico."
          ]
        },
        {
          heading: "Proyecto arquitectónico",
          items: [
            "Plantas arquitectónicas.",
            "Cortes arquitectónicos.",
            "Fachadas arquitectónicas."
          ]
        },
        {
          heading: "Material comercial",
          items: [
            "12 perspectivas calidad HD",
            "Plantas editadas para venta"
          ]
        }
      ],
      notes: "",
      amount: 147000,
      currencyLabel: "MXN",
      included: true
    },
    {
      title: "Proyecto Ejecutivo",
      subtitle: "",
      leadText: "",
      groups: [
        {
          heading: "Planos de sub-especialidades (con ubicación y especificaciones)",
          items: [
            "Acabados y despiece de pisos",
            "Albañilería, Trazo y Detalles especiales.",
            "Carpintería de mobiliario fijo.",
            "Cancelería y cristalería fija.",
            "Herrería fija.",
            "Detalles de baños",
            "Demolición (en caso de ser necesarios)"
          ]
        },
        {
          heading: "Planos de criterios técnicos (no incluyen cálculo o diseño)",
          items: [
            "Criterio de iluminación.",
            "Criterio de salidas y contactos."
          ]
        }
      ],
      notes:
        "Adecuaciones de memorias y planos requeridos por desarrollo urbano e INAH. (No incluye gestiones ante las dependencias)\n\nSe incluye coordinación y revisión de ingenierías.",
      amount: 50000,
      currencyLabel: "MXN",
      included: true
    }
  ],
  discountMode: "custom_final",
  discountPercent: 12,
  discountFixed: 24000,
  discountCustomLabel: "DESC. 12%",
  discountFinalAmount: 173000,
  showIvaBreakdown: false,
  ivaRatePercent: 16,
  weeksMin: 14,
  weeksMax: 16,
  deliveryMainText:
    "La entrega del trabajo completo tiene un tiempo estimado de 14 a 16 semanas (dependiendo del número de revisiones que sean necesarias).",
  deliveryEarlyNote:
    "***La entrega del material comercial se irá entregando anticipadamente por paquetes, contando con la primera entrega el 16 de noviembre con 5 renders de las modificaciones solicitadas y 2 plantas ambientadas con las modificaciones representadas.",
  deliveryFootnote: "*El tiempo de entrega empieza a correr a partir del pago del anticipo.",
  showPaymentAmounts: false,
  payments: JSON.parse(JSON.stringify(PAYMENT_PRESETS["40-30-20-10"])),
  ivaStatus: "Precios NO INCLUYEN IVA.",
  obsAdjustRounds: 2,
  obsValidityDays: 90,
  obsDelayDays: 5,
  obsCity: "Mérida",
  observations: JSON.parse(JSON.stringify(DEFAULT_OBSERVATIONS)),
  closingLine:
    "Dejando a su consideración lo anteriormente expuesto, quedamos a su disposición para cualquier observación.",
  signoffGreeting: "Atentamente:",
  signers: [
    { name: "M. ARQ. SEBASTIÁN MONTALVO MILLET", role: "" },
    { name: "M. ARQ. ESTEBAN HERNÁNDEZ MARTÍN", role: "" }
  ],
  footerPhone: "999 275 8369",
  footerEmail: "proyectos@esencial.studio"
};

const SAMPLE_QUICK_NOTES = `Proyecto: CASA MONTEJO
Cliente: Familia García
Descuento: 10%
Entrega: 10 a 12 semanas
Vigencia: 60 días naturales
Alcance: Propuesta Conceptual y Proyecto Arquitectónico | $95,000 | Plantas, cortes, fachadas, definición de concepto y 8 renders HD
Alcance: Proyecto Ejecutivo | $65,000 | Acabados, albañilería, carpintería, cancelería, detalles de baños y criterios de iluminación`;

let state = loadInitialState();
let currentZoom = 0.78;
let speechRecognition = null;
let isDictating = false;
let toastTimer = null;

/* ==========================================================================
   Inicialización y Estado
   ========================================================================== */
function deepClone(obj) {
  return JSON.parse(JSON.stringify(obj));
}

function loadInitialState() {
  try {
    const saved = localStorage.getItem(STORAGE_CURRENT_KEY);
    if (saved) {
      const parsed = JSON.parse(saved);
      if (parsed && typeof parsed === "object") {
        return normalizeState(parsed);
      }
    }
  } catch (err) {
    console.warn("No se pudo leer el presupuesto guardado:", err);
  }
  return deepClone(DEFAULT_PRESUPUESTO);
}

function normalizeState(raw) {
  const d = deepClone(DEFAULT_PRESUPUESTO);
  return {
    quoteEyebrow: String(raw.quoteEyebrow ?? d.quoteEyebrow),
    projectName: String(raw.projectName ?? d.projectName),
    clientName: String(raw.clientName ?? ""),
    quoteCode: String(raw.quoteCode ?? d.quoteCode),
    quoteDateText: String(raw.quoteDateText ?? d.quoteDateText),
    paperTheme: raw.paperTheme === "chukum" ? "chukum" : "white",
    includeCoverPage: raw.includeCoverPage !== false,
    includeIntroText: raw.includeIntroText !== false,
    showPageNumbers: Boolean(raw.showPageNumbers),
    introGreeting: String(raw.introGreeting ?? d.introGreeting),
    introBody: String(raw.introBody ?? d.introBody),
    introLead: String(raw.introLead ?? d.introLead),
    scopes: Array.isArray(raw.scopes)
      ? raw.scopes.map((sc) => ({
          title: String(sc.title ?? ""),
          subtitle: String(sc.subtitle ?? ""),
          leadText: String(sc.leadText ?? ""),
          groups: Array.isArray(sc.groups)
            ? sc.groups.map((g) => ({
                heading: String(g.heading ?? ""),
                items: Array.isArray(g.items) ? g.items.map((it) => String(it ?? "")) : []
              }))
            : [],
          notes: String(sc.notes ?? ""),
          amount: Number.isFinite(Number(sc.amount)) ? Number(sc.amount) : 0,
          currencyLabel: String(sc.currencyLabel ?? "MXN"),
          included: sc.included !== false
        }))
      : d.scopes,
    discountMode: ["none", "percent", "fixed", "custom_final"].includes(raw.discountMode)
      ? raw.discountMode
      : "none",
    discountPercent: Number.isFinite(Number(raw.discountPercent)) ? Number(raw.discountPercent) : 0,
    discountFixed: Number.isFinite(Number(raw.discountFixed)) ? Number(raw.discountFixed) : 0,
    discountCustomLabel: String(raw.discountCustomLabel ?? "DESC. 10%"),
    discountFinalAmount: Number.isFinite(Number(raw.discountFinalAmount))
      ? Number(raw.discountFinalAmount)
      : 0,
    showIvaBreakdown: Boolean(raw.showIvaBreakdown),
    ivaRatePercent: Number.isFinite(Number(raw.ivaRatePercent)) ? Number(raw.ivaRatePercent) : 16,
    weeksMin: Number.isFinite(Number(raw.weeksMin)) ? Number(raw.weeksMin) : d.weeksMin,
    weeksMax: Number.isFinite(Number(raw.weeksMax)) ? Number(raw.weeksMax) : d.weeksMax,
    deliveryMainText: String(raw.deliveryMainText ?? d.deliveryMainText),
    deliveryEarlyNote: String(raw.deliveryEarlyNote ?? ""),
    deliveryFootnote: String(raw.deliveryFootnote ?? d.deliveryFootnote),
    showPaymentAmounts: Boolean(raw.showPaymentAmounts),
    payments: Array.isArray(raw.payments)
      ? raw.payments.map((p) => ({
          percent: Number.isFinite(Number(p.percent)) ? Number(p.percent) : 0,
          text: String(p.text ?? "")
        }))
      : d.payments,
    ivaStatus: String(raw.ivaStatus ?? d.ivaStatus),
    obsAdjustRounds: Number.isFinite(Number(raw.obsAdjustRounds)) ? Number(raw.obsAdjustRounds) : 2,
    obsValidityDays: Number.isFinite(Number(raw.obsValidityDays)) ? Number(raw.obsValidityDays) : 90,
    obsDelayDays: Number.isFinite(Number(raw.obsDelayDays)) ? Number(raw.obsDelayDays) : 5,
    obsCity: String(raw.obsCity ?? "Mérida"),
    observations: Array.isArray(raw.observations)
      ? raw.observations.map((o, idx) => ({
          id: String(o.id ?? `obs_${idx}`),
          enabled: o.enabled !== false,
          text: String(o.text ?? "")
        }))
      : d.observations,
    closingLine: String(raw.closingLine ?? d.closingLine),
    signoffGreeting: String(raw.signoffGreeting ?? "Atentamente:"),
    signers: Array.isArray(raw.signers)
      ? raw.signers.map((sg) => ({
          name: String(sg.name ?? ""),
          role: String(sg.role ?? "")
        }))
      : d.signers,
    footerPhone: String(raw.footerPhone ?? d.footerPhone),
    footerEmail: String(raw.footerEmail ?? d.footerEmail)
  };
}

function persistCurrentState() {
  try {
    localStorage.setItem(STORAGE_CURRENT_KEY, JSON.stringify(state));
    const statusEl = document.getElementById("saveStatusText");
    if (statusEl) {
      const now = new Date().toLocaleTimeString("es-MX", { hour: "2-digit", minute: "2-digit" });
      statusEl.textContent = `Autoguardado · ${now}`;
    }
  } catch (err) {
    console.warn("No se pudo guardar en localStorage:", err);
  }
}

document.addEventListener("DOMContentLoaded", () => {
  if (window.lucide && typeof window.lucide.createIcons === "function") {
    window.lucide.createIcons();
  }
  renderPresetCatalog();
  bindStaticInputs();
  bindTopbarActions();
  bindTabNavigation();
  bindCalculatorAndNotes();
  bindZoomControls();
  syncEditorFromState();
  applyPaperTheme();
  renderHistoryList();
  renderDocumentPages();

  if (document.fonts && document.fonts.ready) {
    document.fonts.ready.then(() => {
      renderDocumentPages();
    });
  }
});

/* ==========================================================================
   Cálculos Económicos y Formato de Moneda
   ========================================================================== */
function formatMoney(num, withSpaceAfterDollar = false) {
  const n = Number(num) || 0;
  const formatted = n.toLocaleString("es-MX", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  });
  return withSpaceAfterDollar ? `$ ${formatted}` : `$${formatted}`;
}

function computeTotals(s = state) {
  const subtotal = s.scopes.reduce((acc, sc) => {
    if (sc.included === false) return acc;
    return acc + (Number(sc.amount) || 0);
  }, 0);

  let hasDiscount = false;
  let discountAmount = 0;
  let discountLabel = "";
  let finalTotal = subtotal;

  if (s.discountMode === "percent" && Number(s.discountPercent) > 0) {
    const pct = Math.min(100, Math.max(0, Number(s.discountPercent) || 0));
    discountAmount = subtotal * (pct / 100);
    finalTotal = Math.max(0, subtotal - discountAmount);
    discountLabel = `DESC. ${pct}%`;
    hasDiscount = true;
  } else if (s.discountMode === "fixed" && Number(s.discountFixed) > 0) {
    discountAmount = Math.min(subtotal, Math.max(0, Number(s.discountFixed) || 0));
    finalTotal = Math.max(0, subtotal - discountAmount);
    discountLabel = "DESCUENTO";
    hasDiscount = true;
  } else if (s.discountMode === "custom_final") {
    finalTotal = Math.max(0, Number(s.discountFinalAmount) || 0);
    discountAmount = Math.max(0, subtotal - finalTotal);
    discountLabel = (s.discountCustomLabel || "DESC.").trim() || "DESC.";
    hasDiscount = true;
  }

  const ivaRate = Math.max(0, Number(s.ivaRatePercent) || 16) / 100;
  const ivaAmount = s.showIvaBreakdown ? finalTotal * ivaRate : 0;
  const grandTotal = finalTotal + ivaAmount;

  return {
    subtotal,
    hasDiscount,
    discountAmount,
    discountLabel,
    finalTotal,
    ivaAmount,
    grandTotal
  };
}

function updateFinancialSummaryUI() {
  const totals = computeTotals(state);
  document.getElementById("summarySubtotal").textContent = `${formatMoney(totals.subtotal, true)} MXN`;
  document.getElementById("summaryDiscount").textContent = totals.hasDiscount
    ? `-${formatMoney(totals.discountAmount, true)}`
    : "$ 0.00";
  const displayFinal = state.showIvaBreakdown ? totals.grandTotal : totals.finalTotal;
  document.getElementById("summaryTotal").textContent = `${formatMoney(displayFinal, true)} MXN`;
  document.getElementById("previewTotalBadge").textContent = `${formatMoney(displayFinal, true)} MXN`;

  // Mostrar/ocultar controles de descuento según modo
  document.getElementById("discountPercentField").classList.toggle("hidden", state.discountMode !== "percent");
  document.getElementById("discountFixedField").classList.toggle("hidden", state.discountMode !== "fixed");
  document.getElementById("discountCustomFields").classList.toggle("hidden", state.discountMode !== "custom_final");
  document.getElementById("ivaRateWrap").classList.toggle("hidden", !state.showIvaBreakdown);
}

/* ==========================================================================
   Catálogo de Paquetes Rápidos y Enlace del Formulario
   ========================================================================== */
function renderPresetCatalog() {
  const grid = document.getElementById("quickPresetGrid");
  grid.innerHTML = "";
  SCOPE_PRESETS.forEach((preset) => {
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = "preset-chip-btn";
    btn.innerHTML = `
      <strong>+ ${escapeHtmlText(preset.name)}</strong>
      <span>${escapeHtmlText(preset.badge)} · ${formatMoney(preset.data.amount)}</span>
    `;
    btn.addEventListener("click", () => {
      state.scopes.push(deepClone(preset.data));
      persistCurrentState();
      renderScopesEditor();
      updateFinancialSummaryUI();
      renderDocumentPages();
      showToast(`Agregado: ${preset.name}`);
    });
    grid.appendChild(btn);
  });
}

function bindStaticInputs() {
  const textBindings = [
    ["projectName", "projectName"],
    ["clientName", "clientName"],
    ["quoteCode", "quoteCode"],
    ["quoteDateText", "quoteDateText"],
    ["quoteEyebrow", "quoteEyebrow"],
    ["introGreeting", "introGreeting"],
    ["introBody", "introBody"],
    ["introLead", "introLead"],
    ["discountCustomLabel", "discountCustomLabel"],
    ["deliveryMainText", "deliveryMainText"],
    ["deliveryEarlyNote", "deliveryEarlyNote"],
    ["deliveryFootnote", "deliveryFootnote"],
    ["closingLine", "closingLine"],
    ["signoffGreeting", "signoffGreeting"],
    ["footerPhone", "footerPhone"],
    ["footerEmail", "footerEmail"],
    ["obsCity", "obsCity"]
  ];

  textBindings.forEach(([id, key]) => {
    const el = document.getElementById(id);
    if (!el) return;
    el.addEventListener("input", () => {
      state[key] = el.value;
      persistCurrentState();
      if (key === "discountCustomLabel") updateFinancialSummaryUI();
      renderDocumentPages();
    });
  });

  const numberBindings = [
    ["discountPercent", "discountPercent"],
    ["discountFixed", "discountFixed"],
    ["discountFinalAmount", "discountFinalAmount"],
    ["ivaRatePercent", "ivaRatePercent"],
    ["weeksMin", "weeksMin"],
    ["weeksMax", "weeksMax"],
    ["obsAdjustRounds", "obsAdjustRounds"],
    ["obsValidityDays", "obsValidityDays"],
    ["obsDelayDays", "obsDelayDays"]
  ];

  numberBindings.forEach(([id, key]) => {
    const el = document.getElementById(id);
    if (!el) return;
    el.addEventListener("input", () => {
      state[key] = Number(el.value) || 0;
      persistCurrentState();
      updateFinancialSummaryUI();
      if (state.showPaymentAmounts) renderPaymentsEditor();
      renderDocumentPages();
    });
  });

  const checkboxBindings = [
    ["includeCoverPage", "includeCoverPage"],
    ["includeIntroText", "includeIntroText"],
    ["showPageNumbers", "showPageNumbers"],
    ["showIvaBreakdown", "showIvaBreakdown"],
    ["showPaymentAmounts", "showPaymentAmounts"]
  ];

  checkboxBindings.forEach(([id, key]) => {
    const el = document.getElementById(id);
    if (!el) return;
    el.addEventListener("change", () => {
      state[key] = el.checked;
      persistCurrentState();
      updateFinancialSummaryUI();
      renderDocumentPages();
    });
  });

  document.getElementById("discountMode").addEventListener("change", (e) => {
    state.discountMode = e.target.value;
    persistCurrentState();
    updateFinancialSummaryUI();
    if (state.showPaymentAmounts) renderPaymentsEditor();
    renderDocumentPages();
  });

  document.getElementById("ivaStatus").addEventListener("change", (e) => {
    state.ivaStatus = e.target.value;
    persistCurrentState();
    renderDocumentPages();
  });

  document.getElementById("btnTodayDate").addEventListener("click", () => {
    state.quoteDateText = formatTodayLongSpanish();
    document.getElementById("quoteDateText").value = state.quoteDateText;
    persistCurrentState();
    renderDocumentPages();
    showToast("Fecha actualizada al día de hoy");
  });

  document.getElementById("btnAutoCode").addEventListener("click", () => {
    state.quoteCode = buildAutoQuoteCode(state.projectName);
    document.getElementById("quoteCode").value = state.quoteCode;
    persistCurrentState();
    showToast(`Folio generado: ${state.quoteCode}`);
  });

  document.getElementById("btnRestoreIntroText").addEventListener("click", () => {
    state.introGreeting = DEFAULT_INTRO.introGreeting;
    state.introBody = DEFAULT_INTRO.introBody;
    state.introLead = DEFAULT_INTRO.introLead;
    document.getElementById("introGreeting").value = state.introGreeting;
    document.getElementById("introBody").value = state.introBody;
    document.getElementById("introLead").value = state.introLead;
    persistCurrentState();
    renderDocumentPages();
    showToast("Redacción introductoria oficial restaurada");
  });

  document.getElementById("btnAddBlankScope").addEventListener("click", () => {
    state.scopes.push({
      title: "Nuevo Concepto / Alcance",
      subtitle: "",
      leadText: "",
      groups: [{ heading: "Entregables incluidos", items: ["Plantas arquitectónicas."] }],
      notes: "",
      amount: 0,
      currencyLabel: "MXN",
      included: true
    });
    persistCurrentState();
    renderScopesEditor();
    updateFinancialSummaryUI();
    renderDocumentPages();
  });

  document.getElementById("btnApplyWeeks").addEventListener("click", () => {
    const min = Math.max(1, Number(state.weeksMin) || 1);
    const max = Math.max(min, Number(state.weeksMax) || min);
    const rangeStr = min === max ? `${min} semanas` : `${min} a ${max} semanas`;
    state.deliveryMainText = `La entrega del trabajo completo tiene un tiempo estimado de ${rangeStr} (dependiendo del número de revisiones que sean necesarias).`;
    document.getElementById("deliveryMainText").value = state.deliveryMainText;
    persistCurrentState();
    renderDocumentPages();
    showToast("Semanas aplicadas al calendario");
  });

  document.querySelectorAll("#paymentPresetRow .chip-pill").forEach((btn) => {
    btn.addEventListener("click", () => {
      const preset = PAYMENT_PRESETS[btn.dataset.preset];
      if (!preset) return;
      state.payments = deepClone(preset);
      persistCurrentState();
      renderPaymentsEditor();
      renderDocumentPages();
      showToast(`Esquema ${btn.textContent} aplicado`);
    });
  });

  document.getElementById("btnAddPayment").addEventListener("click", () => {
    state.payments.push({
      percent: 10,
      text: "Pago parcial correspondiente al 10% del valor total del proyecto contra entregable."
    });
    persistCurrentState();
    renderPaymentsEditor();
    renderDocumentPages();
  });

  document.getElementById("btnApplyObsParams").addEventListener("click", () => {
    const rounds = Math.max(0, Number(state.obsAdjustRounds) || 2);
    const validity = Math.max(5, Number(state.obsValidityDays) || 90);
    const delay = Math.max(1, Number(state.obsDelayDays) || 5);
    const city = (state.obsCity || "Mérida").trim();

    state.observations.forEach((obs) => {
      if (obs.id === "rounds") {
        obs.text = `El proyecto contempla hasta ${rounds} ${rounds === 1 ? "ronda" : "rondas"} de ajustes conceptuales posteriores a la primera propuesta. Rondas adicionales serán cotizadas por separado.`;
      } else if (obs.id === "city") {
        obs.text = `Las revisiones están consideradas en la ciudad de ${city} o por videollamada.`;
      } else if (obs.id === "validity") {
        obs.text = `Los precios incluidos en este presupuesto tienen una vigencia de ${validity} días naturales para su aceptación y firma. Una vez iniciado el proyecto y pagado el anticipo, los costos se mantendrán fijos durante todo el periodo de desarrollo del proyecto.`;
      } else if (obs.id === "delays") {
        obs.text = `Las revisiones retrasadas por parte del cliente que excedan los ${delay} días hábiles, permitirán extender los tiempos de entrega.`;
      }
    });

    persistCurrentState();
    renderObservationsEditor();
    renderDocumentPages();
    showToast("Parámetros actualizados en cláusulas");
  });

  document.getElementById("btnAddObservation").addEventListener("click", () => {
    state.observations.push({
      id: `custom_${Date.now()}`,
      enabled: true,
      text: "Observación adicional para este proyecto."
    });
    persistCurrentState();
    renderObservationsEditor();
    renderDocumentPages();
  });

  document.getElementById("btnAddSigner").addEventListener("click", () => {
    state.signers.push({ name: "M. ARQ. NUEVO FIRMANTE", role: "" });
    persistCurrentState();
    renderSignersEditor();
    renderDocumentPages();
  });

  document.getElementById("btnAddClientSigner").addEventListener("click", () => {
    const clientLabel = (state.clientName || "CLIENTE").trim().toUpperCase();
    state.signers.push({ name: clientLabel, role: "ACEPTACIÓN DE PRESUPUESTO" });
    persistCurrentState();
    renderSignersEditor();
    renderDocumentPages();
    showToast("Firma de aceptación del cliente agregada");
  });
}

function syncEditorFromState() {
  const valMap = {
    projectName: state.projectName,
    clientName: state.clientName,
    quoteCode: state.quoteCode,
    quoteDateText: state.quoteDateText,
    quoteEyebrow: state.quoteEyebrow,
    introGreeting: state.introGreeting,
    introBody: state.introBody,
    introLead: state.introLead,
    discountMode: state.discountMode,
    discountPercent: state.discountPercent,
    discountFixed: state.discountFixed,
    discountCustomLabel: state.discountCustomLabel,
    discountFinalAmount: state.discountFinalAmount,
    ivaRatePercent: state.ivaRatePercent,
    weeksMin: state.weeksMin,
    weeksMax: state.weeksMax,
    deliveryMainText: state.deliveryMainText,
    deliveryEarlyNote: state.deliveryEarlyNote,
    deliveryFootnote: state.deliveryFootnote,
    ivaStatus: state.ivaStatus,
    obsAdjustRounds: state.obsAdjustRounds,
    obsValidityDays: state.obsValidityDays,
    obsDelayDays: state.obsDelayDays,
    obsCity: state.obsCity,
    closingLine: state.closingLine,
    signoffGreeting: state.signoffGreeting,
    footerPhone: state.footerPhone,
    footerEmail: state.footerEmail
  };

  Object.entries(valMap).forEach(([id, val]) => {
    const el = document.getElementById(id);
    if (el) el.value = val ?? "";
  });

  document.getElementById("includeCoverPage").checked = state.includeCoverPage;
  document.getElementById("includeIntroText").checked = state.includeIntroText;
  document.getElementById("showPageNumbers").checked = state.showPageNumbers;
  document.getElementById("showIvaBreakdown").checked = state.showIvaBreakdown;
  document.getElementById("showPaymentAmounts").checked = state.showPaymentAmounts;

  renderScopesEditor();
  renderPaymentsEditor();
  renderObservationsEditor();
  renderSignersEditor();
  updateFinancialSummaryUI();
}

/* ==========================================================================
   Editores Dinámicos: Alcances, Pagos, Observaciones, Firmas
   ========================================================================== */
function renderScopesEditor() {
  const container = document.getElementById("scopesList");
  container.innerHTML = "";
  document.getElementById("scopeCountBadge").textContent =
    `${state.scopes.length} ${state.scopes.length === 1 ? "concepto" : "conceptos"}`;

  if (state.scopes.length === 0) {
    container.innerHTML = `<p class="helper-text">No hay conceptos todavía. Usa los paquetes rápidos de arriba o haz clic en "+ Partida vacía".</p>`;
    return;
  }

  state.scopes.forEach((sc, idx) => {
    const box = document.createElement("div");
    box.className = "item-box";
    box.innerHTML = `
      <div class="item-box-head">
        <span class="item-num-tag">Concepto ${String(idx + 1).padStart(2, "0")}</span>
        <div class="item-actions">
          <label class="checkbox-inline" title="Incluir en total">
            <input type="checkbox" class="sc-inc" ${sc.included !== false ? "checked" : ""} />
            Sumar
          </label>
          <button type="button" class="mini-btn sc-up" title="Subir" ${idx === 0 ? "disabled" : ""}>↑</button>
          <button type="button" class="mini-btn sc-down" title="Bajar" ${idx === state.scopes.length - 1 ? "disabled" : ""}>↓</button>
          <button type="button" class="mini-btn sc-dup" title="Duplicar concepto">⧉</button>
          <button type="button" class="mini-btn danger sc-del" title="Eliminar">✕</button>
        </div>
      </div>

      <div class="form-grid-2">
        <div class="field">
          <label>Alcance (Columna izquierda)</label>
          <input type="text" class="sc-title" value="${escapeHtmlAttr(sc.title)}" placeholder="Ej. Proyecto Ejecutivo" />
        </div>
        <div class="field">
          <label>Nota cursiva opcional (debajo del título)</label>
          <input type="text" class="sc-sub" value="${escapeHtmlAttr(sc.subtitle)}" placeholder="Ej. (Entrega anticipada para el 10 de dic.)" />
        </div>
      </div>

      <div class="field">
        <label>Párrafo introductorio del concepto (opcional)</label>
        <textarea rows="2" class="sc-lead" placeholder="Breve descripción general de la etapa...">${escapeHtmlText(sc.leadText)}</textarea>
      </div>

      <div class="field">
        <div class="item-box-head">
          <label>Bloques de entregables y viñetas</label>
          <button type="button" class="mini-btn sage sc-add-group">+ Bloque</button>
        </div>
        <div class="subgroups-wrap sc-groups-container"></div>
      </div>

      <div class="field">
        <label>Notas al pie del concepto (opcional)</label>
        <textarea rows="2" class="sc-notes" placeholder="Ej. Se incluye coordinación y revisión de ingenierías.">${escapeHtmlText(sc.notes)}</textarea>
      </div>

      <div class="form-grid-2">
        <div class="field">
          <label>Importe del concepto ($)</label>
          <input type="number" class="sc-amount" min="0" step="500" value="${Number(sc.amount) || 0}" />
        </div>
        <div class="field">
          <label>Moneda</label>
          <input type="text" class="sc-curr" value="${escapeHtmlAttr(sc.currencyLabel || "MXN")}" placeholder="MXN" />
        </div>
      </div>
    `;

    box.querySelector(".sc-inc").addEventListener("change", (e) => {
      state.scopes[idx].included = e.target.checked;
      persistCurrentState();
      updateFinancialSummaryUI();
      if (state.showPaymentAmounts) renderPaymentsEditor();
      renderDocumentPages();
    });

    box.querySelector(".sc-title").addEventListener("input", (e) => {
      state.scopes[idx].title = e.target.value;
      persistCurrentState();
      renderDocumentPages();
    });

    box.querySelector(".sc-sub").addEventListener("input", (e) => {
      state.scopes[idx].subtitle = e.target.value;
      persistCurrentState();
      renderDocumentPages();
    });

    box.querySelector(".sc-lead").addEventListener("input", (e) => {
      state.scopes[idx].leadText = e.target.value;
      persistCurrentState();
      renderDocumentPages();
    });

    box.querySelector(".sc-notes").addEventListener("input", (e) => {
      state.scopes[idx].notes = e.target.value;
      persistCurrentState();
      renderDocumentPages();
    });

    box.querySelector(".sc-amount").addEventListener("input", (e) => {
      state.scopes[idx].amount = Number(e.target.value) || 0;
      persistCurrentState();
      updateFinancialSummaryUI();
      if (state.showPaymentAmounts) renderPaymentsEditor();
      renderDocumentPages();
    });

    box.querySelector(".sc-curr").addEventListener("input", (e) => {
      state.scopes[idx].currencyLabel = e.target.value || "MXN";
      persistCurrentState();
      renderDocumentPages();
    });

    box.querySelector(".sc-up").addEventListener("click", () => {
      if (idx === 0) return;
      [state.scopes[idx - 1], state.scopes[idx]] = [state.scopes[idx], state.scopes[idx - 1]];
      persistCurrentState();
      renderScopesEditor();
      renderDocumentPages();
    });

    box.querySelector(".sc-down").addEventListener("click", () => {
      if (idx === state.scopes.length - 1) return;
      [state.scopes[idx], state.scopes[idx + 1]] = [state.scopes[idx + 1], state.scopes[idx]];
      persistCurrentState();
      renderScopesEditor();
      renderDocumentPages();
    });

    box.querySelector(".sc-dup").addEventListener("click", () => {
      const copy = deepClone(state.scopes[idx]);
      copy.title = `${copy.title} (Copia)`;
      state.scopes.splice(idx + 1, 0, copy);
      persistCurrentState();
      renderScopesEditor();
      updateFinancialSummaryUI();
      renderDocumentPages();
      showToast("Concepto duplicado");
    });

    box.querySelector(".sc-del").addEventListener("click", () => {
      state.scopes.splice(idx, 1);
      persistCurrentState();
      renderScopesEditor();
      updateFinancialSummaryUI();
      renderDocumentPages();
    });

    box.querySelector(".sc-add-group").addEventListener("click", () => {
      state.scopes[idx].groups.push({ heading: "Nuevo bloque", items: [""] });
      persistCurrentState();
      renderScopeGroupsEditor(box.querySelector(".sc-groups-container"), idx);
      renderDocumentPages();
    });

    renderScopeGroupsEditor(box.querySelector(".sc-groups-container"), idx);
    container.appendChild(box);
  });
}

function renderScopeGroupsEditor(container, scopeIdx) {
  container.innerHTML = "";
  const groups = state.scopes[scopeIdx].groups;

  groups.forEach((grp, gIdx) => {
    const gBox = document.createElement("div");
    gBox.className = "subgroup-box";
    const itemsText = (grp.items || []).join("\n");

    gBox.innerHTML = `
      <div class="subgroup-head">
        <input type="text" class="grp-heading" value="${escapeHtmlAttr(grp.heading)}" placeholder="Título del bloque (ej. Propuesta conceptual)" />
        <button type="button" class="mini-btn grp-chips-toggle" title="Insertar entregables rápidos">+ Rápido</button>
        <button type="button" class="mini-btn danger grp-del" title="Quitar bloque">✕</button>
      </div>
      <div class="chips-row grp-chips hidden"></div>
      <textarea rows="3" class="grp-items" placeholder="Un entregable por línea (el guion '-' se agrega automático)">${escapeHtmlText(itemsText)}</textarea>
    `;

    const chipsWrap = gBox.querySelector(".grp-chips");
    COMMON_DELIVERABLES.forEach((itemText) => {
      const chip = document.createElement("button");
      chip.type = "button";
      chip.className = "chip-pill";
      chip.textContent = `+ ${itemText.replace(/\.$/, "")}`;
      chip.addEventListener("click", () => {
        const cur = state.scopes[scopeIdx].groups[gIdx].items.filter((x) => x.trim() !== "");
        if (!cur.includes(itemText)) {
          cur.push(itemText);
          state.scopes[scopeIdx].groups[gIdx].items = cur;
          gBox.querySelector(".grp-items").value = cur.join("\n");
          persistCurrentState();
          renderDocumentPages();
        }
      });
      chipsWrap.appendChild(chip);
    });

    gBox.querySelector(".grp-chips-toggle").addEventListener("click", () => {
      chipsWrap.classList.toggle("hidden");
    });

    gBox.querySelector(".grp-heading").addEventListener("input", (e) => {
      state.scopes[scopeIdx].groups[gIdx].heading = e.target.value;
      persistCurrentState();
      renderDocumentPages();
    });

    gBox.querySelector(".grp-items").addEventListener("input", (e) => {
      state.scopes[scopeIdx].groups[gIdx].items = e.target.value
        .split(/\r?\n/)
        .map((line) => line.replace(/^[-•*]\s*/, "").trimEnd())
        .filter((line) => line.trim() !== "");
      persistCurrentState();
      renderDocumentPages();
    });

    gBox.querySelector(".grp-del").addEventListener("click", () => {
      state.scopes[scopeIdx].groups.splice(gIdx, 1);
      persistCurrentState();
      renderScopeGroupsEditor(container, scopeIdx);
      renderDocumentPages();
    });

    container.appendChild(gBox);
  });
}

function renderPaymentsEditor() {
  const container = document.getElementById("paymentsList");
  container.innerHTML = "";

  const totals = computeTotals(state);
  const baseTotal = state.showIvaBreakdown ? totals.grandTotal : totals.finalTotal;
  const sumPct = state.payments.reduce((acc, p) => acc + (Number(p.percent) || 0), 0);

  const badge = document.getElementById("paymentSumBadge");
  badge.textContent = `Suma: ${sumPct}%`;
  badge.classList.toggle("badge-ok", Math.abs(sumPct - 100) < 0.01);
  badge.classList.toggle("badge-warn", Math.abs(sumPct - 100) >= 0.01);

  state.payments.forEach((pay, idx) => {
    const moneyVal = baseTotal * ((Number(pay.percent) || 0) / 100);
    const box = document.createElement("div");
    box.className = "payment-row-box";
    box.innerHTML = `
      <div class="payment-row-top">
        <input type="number" class="pay-pct" min="0" max="100" step="1" value="${Number(pay.percent) || 0}" title="Porcentaje (%)" />
        <span class="payment-amount-tag">Pago ${idx + 1}: ${formatMoney(moneyVal, true)} MXN</span>
        <div class="item-actions">
          <button type="button" class="mini-btn pay-sync" title="Sincronizar % en el texto">Sync %</button>
          <button type="button" class="mini-btn danger pay-del" title="Eliminar pago">✕</button>
        </div>
      </div>
      <textarea rows="2" class="pay-text">${escapeHtmlText(pay.text)}</textarea>
    `;

    const pctInput = box.querySelector(".pay-pct");
    const tagEl = box.querySelector(".payment-amount-tag");
    const textInput = box.querySelector(".pay-text");

    pctInput.addEventListener("input", (e) => {
      const pct = Number(e.target.value) || 0;
      state.payments[idx].percent = pct;
      const newMoney = (state.showIvaBreakdown ? computeTotals().grandTotal : computeTotals().finalTotal) * (pct / 100);
      tagEl.textContent = `Pago ${idx + 1}: ${formatMoney(newMoney, true)} MXN`;
      const newSum = state.payments.reduce((acc, p) => acc + (Number(p.percent) || 0), 0);
      badge.textContent = `Suma: ${newSum}%`;
      badge.classList.toggle("badge-ok", Math.abs(newSum - 100) < 0.01);
      badge.classList.toggle("badge-warn", Math.abs(newSum - 100) >= 0.01);
      persistCurrentState();
      renderDocumentPages();
    });

    box.querySelector(".pay-sync").addEventListener("click", () => {
      const pct = Number(state.payments[idx].percent) || 0;
      state.payments[idx].text = state.payments[idx].text.replace(/\d+(\.\d+)?%/, `${pct}%`);
      textInput.value = state.payments[idx].text;
      persistCurrentState();
      renderDocumentPages();
      showToast("Porcentaje sincronizado en el texto");
    });

    textInput.addEventListener("input", (e) => {
      state.payments[idx].text = e.target.value;
      persistCurrentState();
      renderDocumentPages();
    });

    box.querySelector(".pay-del").addEventListener("click", () => {
      state.payments.splice(idx, 1);
      persistCurrentState();
      renderPaymentsEditor();
      renderDocumentPages();
    });

    container.appendChild(box);
  });
}

function renderObservationsEditor() {
  const container = document.getElementById("observationsList");
  container.innerHTML = "";

  state.observations.forEach((obs, idx) => {
    const box = document.createElement("div");
    box.className = "obs-item" + (obs.enabled ? "" : " is-disabled");
    box.innerHTML = `
      <div class="item-box-head">
        <label class="checkbox-inline">
          <input type="checkbox" class="obs-enabled" ${obs.enabled ? "checked" : ""} />
          <span>Cláusula ${String(idx + 1).padStart(2, "0")}</span>
        </label>
        <div class="item-actions">
          <button type="button" class="mini-btn obs-up" ${idx === 0 ? "disabled" : ""}>↑</button>
          <button type="button" class="mini-btn obs-down" ${idx === state.observations.length - 1 ? "disabled" : ""}>↓</button>
          <button type="button" class="mini-btn danger obs-del">✕</button>
        </div>
      </div>
      <textarea rows="2" class="obs-text">${escapeHtmlText(obs.text)}</textarea>
    `;

    box.querySelector(".obs-enabled").addEventListener("change", (e) => {
      state.observations[idx].enabled = e.target.checked;
      box.classList.toggle("is-disabled", !e.target.checked);
      persistCurrentState();
      renderDocumentPages();
    });

    box.querySelector(".obs-text").addEventListener("input", (e) => {
      state.observations[idx].text = e.target.value;
      persistCurrentState();
      renderDocumentPages();
    });

    box.querySelector(".obs-up").addEventListener("click", () => {
      if (idx === 0) return;
      [state.observations[idx - 1], state.observations[idx]] = [state.observations[idx], state.observations[idx - 1]];
      persistCurrentState();
      renderObservationsEditor();
      renderDocumentPages();
    });

    box.querySelector(".obs-down").addEventListener("click", () => {
      if (idx === state.observations.length - 1) return;
      [state.observations[idx], state.observations[idx + 1]] = [state.observations[idx + 1], state.observations[idx]];
      persistCurrentState();
      renderObservationsEditor();
      renderDocumentPages();
    });

    box.querySelector(".obs-del").addEventListener("click", () => {
      state.observations.splice(idx, 1);
      persistCurrentState();
      renderObservationsEditor();
      renderDocumentPages();
    });

    container.appendChild(box);
  });
}

function renderSignersEditor() {
  const container = document.getElementById("signersList");
  container.innerHTML = "";

  state.signers.forEach((sg, idx) => {
    const box = document.createElement("div");
    box.className = "item-box";
    box.innerHTML = `
      <div class="item-box-head">
        <span class="item-num-tag">Firmante ${String(idx + 1).padStart(2, "0")}</span>
        <button type="button" class="mini-btn danger sg-del">✕</button>
      </div>
      <div class="form-grid-2">
        <div class="field">
          <label>Nombre completo</label>
          <input type="text" class="sg-name" value="${escapeHtmlAttr(sg.name)}" />
        </div>
        <div class="field">
          <label>Cargo / Nota (opcional)</label>
          <input type="text" class="sg-role" value="${escapeHtmlAttr(sg.role)}" placeholder="Opcional" />
        </div>
      </div>
    `;

    box.querySelector(".sg-name").addEventListener("input", (e) => {
      state.signers[idx].name = e.target.value;
      persistCurrentState();
      renderDocumentPages();
    });

    box.querySelector(".sg-role").addEventListener("input", (e) => {
      state.signers[idx].role = e.target.value;
      persistCurrentState();
      renderDocumentPages();
    });

    box.querySelector(".sg-del").addEventListener("click", () => {
      state.signers.splice(idx, 1);
      persistCurrentState();
      renderSignersEditor();
      renderDocumentPages();
    });

    container.appendChild(box);
  });
}

/* ==========================================================================
   Motor de Renderizado y Paginación Dinámica Multipágina
   ========================================================================== */
function createCoverPageElement() {
  const page = document.createElement("article");
  page.className = "presupuesto-page page-cover";
  page.innerHTML = `
    <header class="cover-wordmark">
      <img class="cover-wordmark-esencial" src="${ESENCIAL_DATA_URI}" alt="ESENCIAL" />
      <img class="cover-wordmark-studio" src="${STUDIO_DATA_URI}" alt="STUDIO" />
    </header>
    <div class="cover-meta-block">
      <div class="cover-eyebrow">${escapeHtmlText(state.quoteEyebrow || "PRESUPUESTO")}</div>
      <h1 class="cover-project">${escapeHtmlText(state.projectName || "PROYECTO")}</h1>
      <p class="cover-date">${escapeHtmlText(state.quoteDateText || "")}</p>
      <img class="cover-iso" src="${ISOTIPO_DATA_URI}" alt="Esencial Studio" />
    </div>
  `;
  return page;
}

function createInnerPageShell() {
  const page = document.createElement("article");
  page.className = "presupuesto-page page-inner";

  const header = document.createElement("header");
  header.className = "doc-inner-header";
  header.innerHTML = `
    <img class="inner-wordmark-esencial" src="${ESENCIAL_DATA_URI}" alt="ESENCIAL" />
    <img class="inner-wordmark-studio" src="${STUDIO_DATA_URI}" alt="STUDIO" />
  `;

  const body = document.createElement("div");
  body.className = "presupuesto-page-body";

  const phoneParts = (state.footerPhone || "999 275 8369").trim().split(/\s+/);
  const emailParts = (state.footerEmail || "proyectos@esencial.studio").trim().split("@");
  const emailUser = emailParts[0] ? `${emailParts[0]}@` : "proyectos@";
  const emailDom = emailParts[1] || "esencial.studio";

  const contactSpans = [
    "<span>T.</span>",
    ...phoneParts.map((p) => `<span>${escapeHtmlText(p)}</span>`),
    "<span>·</span>",
    `<span>${escapeHtmlText(emailUser)}</span>`,
    `<span>${escapeHtmlText(emailDom)}</span>`
  ].join("");

  const footer = document.createElement("footer");
  footer.className = "doc-footer";
  footer.innerHTML = `
    <div class="doc-footer-contact">${contactSpans}</div>
    <div class="doc-footer-right">
      ${state.showPageNumbers ? `<span class="doc-page-num">PÁGINA 1</span>` : ""}
      <img class="doc-monogram-img" src="${ISOTIPO_DATA_URI}" alt="Esencial Studio" />
    </div>
  `;

  page.appendChild(header);
  page.appendChild(body);
  page.appendChild(footer);

  return { page, body, footer };
}

function createIntroElement() {
  const wrap = document.createElement("section");
  wrap.className = "doc-intro";
  const paragraphs = [state.introGreeting, state.introBody, state.introLead]
    .map((t) => (t || "").trim())
    .filter(Boolean);
  wrap.innerHTML = paragraphs.map((p) => `<p>${escapeHtmlText(p)}</p>`).join("");
  return wrap;
}

function createScopeTableShell(isContinuation) {
  const wrapper = document.createElement("section");
  wrapper.className = "scope-table-wrap" + (isContinuation ? " continuation" : "");
  wrapper.innerHTML = `
    <div class="scope-thead">
      <span>ALCANCE</span>
      <span>DESCRIPCIÓN</span>
    </div>
    <div class="scope-tbody"></div>
  `;
  return { wrapper, tbody: wrapper.querySelector(".scope-tbody") };
}

function createScopeRowElement(sc) {
  const row = document.createElement("div");
  row.className = "scope-row";

  const groupsHtml = (sc.groups || [])
    .filter((g) => (g.heading || "").trim() || (g.items || []).length > 0)
    .map((g) => {
      const headingHtml = (g.heading || "").trim()
        ? `<div class="scope-group-title">${escapeHtmlText(g.heading)}</div>`
        : "";
      const items = (g.items || []).filter((it) => it.trim() !== "");
      const listHtml = items.length
        ? `<ul class="scope-bullets">${items.map((it) => `<li>${escapeHtmlText(it)}</li>`).join("")}</ul>`
        : "";
      return `<div class="scope-group">${headingHtml}${listHtml}</div>`;
    })
    .join("");

  const notesParagraphs = (sc.notes || "")
    .split(/\n\s*\n|\r?\n/)
    .map((n) => n.trim())
    .filter(Boolean)
    .map((n) => `<p class="scope-note">${escapeHtmlText(n)}</p>`)
    .join("");

  const curr = (sc.currencyLabel || "MXN").trim();
  const priceText = `${formatMoney(sc.amount, false)} ${curr}`;

  row.innerHTML = `
    <div class="scope-left">
      <div class="scope-left-title">${escapeHtmlText(sc.title || "Concepto")}</div>
      ${(sc.subtitle || "").trim() ? `<div class="scope-left-sub">${escapeHtmlText(sc.subtitle)}</div>` : ""}
    </div>
    <div class="scope-right">
      ${(sc.leadText || "").trim() ? `<p class="scope-lead">${escapeHtmlText(sc.leadText)}</p>` : ""}
      ${groupsHtml}
      ${notesParagraphs}
      <div class="scope-price">${escapeHtmlText(priceText)}</div>
    </div>
  `;
  return row;
}

function createTotalsElement() {
  const totals = computeTotals(state);
  const wrap = document.createElement("div");
  wrap.className = "totals-wrap";

  if (totals.hasDiscount) {
    wrap.innerHTML += `
      <div class="total-line">
        <span>TOTAL</span>
        <span class="total-strike">${formatMoney(totals.subtotal, true)} MXN</span>
      </div>
      <div class="total-line is-bold">
        <span>${escapeHtmlText(totals.discountLabel)}</span>
        <span>${formatMoney(totals.finalTotal, true)} MXN</span>
      </div>
    `;
  } else {
    wrap.innerHTML += `
      <div class="total-line is-bold">
        <span>TOTAL</span>
        <span>${formatMoney(totals.subtotal, true)} MXN</span>
      </div>
    `;
  }

  if (state.showIvaBreakdown) {
    wrap.innerHTML += `
      <div class="total-line">
        <span>IVA (${Number(state.ivaRatePercent) || 16}%)</span>
        <span>${formatMoney(totals.ivaAmount, true)} MXN</span>
      </div>
      <div class="total-line is-bold">
        <span>TOTAL CON IVA</span>
        <span>${formatMoney(totals.grandTotal, true)} MXN</span>
      </div>
    `;
  }

  return wrap;
}

function createDeliverySectionElement() {
  const sec = document.createElement("section");
  sec.className = "doc-section";
  sec.innerHTML = `<h2 class="doc-section-title">CALENDARIO DE ENTREGA</h2>`;

  if ((state.deliveryMainText || "").trim()) {
    sec.innerHTML += `<p class="doc-p">${escapeHtmlText(state.deliveryMainText)}</p>`;
  }
  if ((state.deliveryEarlyNote || "").trim()) {
    sec.innerHTML += `<p class="doc-p">${escapeHtmlText(state.deliveryEarlyNote)}</p>`;
  }
  if ((state.deliveryFootnote || "").trim()) {
    sec.innerHTML += `<p class="doc-footnote">${escapeHtmlText(state.deliveryFootnote)}</p>`;
  }
  return sec;
}

function createPaymentsSectionElement() {
  const sec = document.createElement("section");
  sec.className = "doc-section";
  sec.innerHTML = `<h2 class="doc-section-title">CALENDARIO DE PAGOS</h2>`;

  const totals = computeTotals(state);
  const baseTotal = state.showIvaBreakdown ? totals.grandTotal : totals.finalTotal;

  const ol = document.createElement("ol");
  ol.className = "payment-list";

  state.payments.forEach((p) => {
    const li = document.createElement("li");
    let txt = (p.text || "").trim();
    if (state.showPaymentAmounts) {
      const amount = baseTotal * ((Number(p.percent) || 0) / 100);
      txt += ` (${formatMoney(amount, true)} MXN)`;
    }
    li.textContent = txt;
    ol.appendChild(li);
  });

  sec.appendChild(ol);
  return sec;
}

function formatObservationHtml(text) {
  const escaped = escapeHtmlText(text);
  // Resaltar automáticamente la vigencia en días naturales como en el PDF oficial
  return escaped.replace(/(\d+\s+d[ií]as\s+naturales)/gi, "<strong>$1</strong>");
}

function createObservationsShell(includeTitle) {
  const sec = document.createElement("section");
  sec.className = "doc-section obs-list";
  if (includeTitle) {
    sec.innerHTML = `<h2 class="doc-section-title">OBSERVACIONES</h2>`;
  }
  return sec;
}

function createSignoffElement() {
  const wrap = document.createElement("section");
  wrap.className = "doc-signoff-wrap";

  const activeSigners = state.signers.filter((s) => (s.name || "").trim());
  const single = activeSigners.length <= 1;

  wrap.innerHTML = `
    <p class="signoff-greeting">${escapeHtmlText(state.signoffGreeting || "Atentamente:")}</p>
    <div class="signoff-grid ${single ? "single-col" : ""}">
      ${activeSigners
        .map(
          (sg) => `
        <div class="signoff-item">
          <div class="signoff-name">${escapeHtmlText(sg.name)}</div>
          ${(sg.role || "").trim() ? `<div class="signoff-role">${escapeHtmlText(sg.role)}</div>` : ""}
        </div>
      `
        )
        .join("")}
    </div>
  `;
  return wrap;
}

function pageOverflows(pageEl) {
  return pageEl.scrollHeight > pageEl.clientHeight + 1;
}

function renderDocumentPages() {
  const sandbox = document.getElementById("measureSandbox");
  const container = document.getElementById("pagesContainer");
  sandbox.innerHTML = "";

  const renderedPages = [];

  if (state.includeCoverPage) {
    renderedPages.push(createCoverPageElement());
  }

  const innerShells = [];
  const startInnerPage = () => {
    const shell = createInnerPageShell();
    sandbox.appendChild(shell.page);
    innerShells.push(shell);
    renderedPages.push(shell.page);
    return shell;
  };

  let current = startInnerPage();

  // 1) Carta introductoria
  if (state.includeIntroText) {
    const introEl = createIntroElement();
    if (introEl.children.length > 0) {
      current.body.appendChild(introEl);
    }
  }

  // 2) Tabla de Alcances y Totales
  if (state.scopes.length > 0) {
    let tableShell = createScopeTableShell(false);
    current.body.appendChild(tableShell.wrapper);

    if (pageOverflows(current.page)) {
      current.body.removeChild(tableShell.wrapper);
      current = startInnerPage();
      tableShell = createScopeTableShell(true);
      current.body.appendChild(tableShell.wrapper);
    }

    state.scopes.forEach((sc, idx) => {
      const row = createScopeRowElement(sc);
      tableShell.tbody.appendChild(row);

      if (pageOverflows(current.page)) {
        tableShell.tbody.removeChild(row);
        if (tableShell.tbody.children.length === 0) {
          current.body.removeChild(tableShell.wrapper);
        }
        current = startInnerPage();
        tableShell = createScopeTableShell(true);
        current.body.appendChild(tableShell.wrapper);
        tableShell.tbody.appendChild(row);
      }

      if (idx === state.scopes.length - 1) {
        row.classList.add("no-bottom-border");
      }
    });

    const totalsEl = createTotalsElement();
    tableShell.wrapper.appendChild(totalsEl);
    if (pageOverflows(current.page)) {
      tableShell.wrapper.removeChild(totalsEl);
      current = startInnerPage();
      const totWrap = document.createElement("section");
      totWrap.className = "scope-table-wrap continuation";
      totWrap.appendChild(totalsEl);
      current.body.appendChild(totWrap);
    }
  }

  // 3) Calendario de Entrega
  const hasDelivery =
    (state.deliveryMainText || "").trim() ||
    (state.deliveryEarlyNote || "").trim() ||
    (state.deliveryFootnote || "").trim();

  if (hasDelivery) {
    const delEl = createDeliverySectionElement();
    current.body.appendChild(delEl);
    if (pageOverflows(current.page)) {
      current.body.removeChild(delEl);
      current = startInnerPage();
      delEl.classList.add("at-top");
      current.body.appendChild(delEl);
    }
  }

  // 4) Calendario de Pagos
  if (state.payments.length > 0) {
    const payEl = createPaymentsSectionElement();
    current.body.appendChild(payEl);
    if (pageOverflows(current.page)) {
      current.body.removeChild(payEl);
      current = startInnerPage();
      payEl.classList.add("at-top");
      current.body.appendChild(payEl);
    }
  }

  // 5) Observaciones (división inteligente si hay muchas cláusulas)
  const obsBlocks = [];
  if ((state.ivaStatus || "").trim()) {
    obsBlocks.push({ isIva: true, text: state.ivaStatus.trim() });
  }
  state.observations
    .filter((o) => o.enabled && (o.text || "").trim())
    .forEach((o) => obsBlocks.push({ isIva: false, text: o.text.trim() }));
  if ((state.closingLine || "").trim()) {
    obsBlocks.push({ isIva: false, text: state.closingLine.trim() });
  }

  if (obsBlocks.length > 0) {
    let obsShell = createObservationsShell(true);
    current.body.appendChild(obsShell);

    obsBlocks.forEach((blk, idx) => {
      const p = document.createElement("p");
      p.className = "doc-p" + (blk.isIva ? " obs-iva" : "");
      p.innerHTML = blk.isIva ? escapeHtmlText(blk.text) : formatObservationHtml(blk.text);
      obsShell.appendChild(p);

      if (pageOverflows(current.page)) {
        obsShell.removeChild(p);
        // Si ni el primer párrafo cupo bajo el título OBSERVACIONES, pasamos todo a nueva página
        if (idx === 0 || (obsShell.children.length === 1 && obsShell.querySelector(".doc-section-title"))) {
          current.body.removeChild(obsShell);
          current = startInnerPage();
          obsShell = createObservationsShell(true);
          obsShell.classList.add("at-top");
          current.body.appendChild(obsShell);
          obsShell.appendChild(p);
        } else {
          current = startInnerPage();
          obsShell = createObservationsShell(false);
          obsShell.classList.add("at-top");
          current.body.appendChild(obsShell);
          obsShell.appendChild(p);
        }
      }
    });
  }

  // 6) Firmas ("Atentamente:")
  if (state.signers.some((s) => (s.name || "").trim())) {
    const sigEl = createSignoffElement();
    current.body.appendChild(sigEl);
    if (pageOverflows(current.page)) {
      current.body.removeChild(sigEl);
      current = startInnerPage();
      sigEl.style.marginTop = "20px";
      current.body.appendChild(sigEl);
    }
  }

  // Numerar si aplica y trasladar al visor
  container.innerHTML = "";
  const totalPages = renderedPages.length;
  renderedPages.forEach((pageEl, idx) => {
    const numSpan = pageEl.querySelector(".doc-page-num");
    if (numSpan) {
      numSpan.textContent = `PÁGINA ${idx + 1} / ${totalPages}`;
    }
    container.appendChild(pageEl);
  });

  document.getElementById("pageCountBadge").textContent =
    `${totalPages} ${totalPages === 1 ? "página" : "páginas"} · Carta (8.5 × 11 in)`;
}

/* ==========================================================================
   Acciones Topbar, Papel, PDF y Exportador Nativo DOCX (.docx)
   ========================================================================== */
function applyPaperTheme() {
  document.body.setAttribute("data-paper", state.paperTheme);
  const btn = document.getElementById("btnTogglePaper");
  const span = btn.querySelector("span");
  if (span) {
    span.textContent = state.paperTheme === "chukum" ? "Papel: Chukum" : "Papel: Blanco";
  }
}

function bindTopbarActions() {
  document.getElementById("btnTogglePaper").addEventListener("click", () => {
    state.paperTheme = state.paperTheme === "white" ? "chukum" : "white";
    applyPaperTheme();
    persistCurrentState();
    showToast(state.paperTheme === "chukum" ? "Papel Chukum activado" : "Papel Blanco editorial activado");
  });

  document.getElementById("btnNewQuote").addEventListener("click", () => {
    const fresh = deepClone(DEFAULT_PRESUPUESTO);
    fresh.projectName = "NUEVO PROYECTO";
    fresh.clientName = "";
    fresh.quoteDateText = formatTodayLongSpanish();
    fresh.quoteCode = buildAutoQuoteCode("NUEVO PROYECTO");
    fresh.scopes = [deepClone(SCOPE_PRESETS[0].data)];
    fresh.discountMode = "none";
    fresh.deliveryEarlyNote = "";
    state = fresh;

    persistCurrentState();
    syncEditorFromState();
    renderDocumentPages();
    showToast("Nuevo presupuesto listo para editar");
  });

  document.getElementById("btnRestoreExample").addEventListener("click", () => {
    state = deepClone(DEFAULT_PRESUPUESTO);
    persistCurrentState();
    syncEditorFromState();
    applyPaperTheme();
    renderDocumentPages();
    showToast("Ejemplo oficial MARBELLO AJUSTES restaurado");
  });

  document.getElementById("btnPrintPdf").addEventListener("click", () => {
    const prevTitle = document.title;
    const safeName = (state.quoteCode || `ES_${state.projectName || "PRESUPUESTO"}`).replace(
      /[\\/:*?"<>|]/g,
      "_"
    );
    document.title = safeName;
    window.print();
    setTimeout(() => {
      document.title = prevTitle;
    }, 600);
  });

  document.getElementById("btnExportDocx").addEventListener("click", () => {
    try {
      exportQuoteToDocx();
      showToast("Archivo .docx generado y descargado");
    } catch (err) {
      console.error("Error al generar DOCX:", err);
      showToast("Error al generar el archivo .docx");
    }
  });
}

/* ==========================================================================
   Generador Nativo OpenXML (.docx) 100% Autónomo (Offline, Word + Google Docs)
   ========================================================================== */
function exportQuoteToDocx() {
  const files = buildDocxPackageFiles(state);
  const zipBytes = buildZipStore(files);
  const blob = new Blob([zipBytes], {
    type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
  });

  const safeName = (state.quoteCode || buildAutoQuoteCode(state.projectName))
    .replace(/[\\/:*?"<>|]/g, "_")
    .trim();
  triggerBlobDownload(blob, `${safeName}.docx`);
}

function xmlEscape(str) {
  return String(str ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function base64ToUint8Array(b64) {
  const bin = atob(b64);
  const arr = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) {
    arr[i] = bin.charCodeAt(i);
  }
  return arr;
}

function makeDocxDrawingXml(rId, docPrId, name, cxEmu, cyEmu) {
  return (
    `<w:drawing>` +
      `<wp:inline distT="0" distB="0" distL="0" distR="0">` +
        `<wp:extent cx="${cxEmu}" cy="${cyEmu}"/>` +
        `<wp:effectExtent l="0" t="0" r="0" b="0"/>` +
        `<wp:docPr id="${docPrId}" name="${xmlEscape(name)}"/>` +
        `<wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/></wp:cNvGraphicFramePr>` +
        `<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
          `<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
            `<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
              `<pic:nvPicPr>` +
                `<pic:cNvPr id="${docPrId}" name="${xmlEscape(name)}"/>` +
                `<pic:cNvPicPr/>` +
              `</pic:nvPicPr>` +
              `<pic:blipFill>` +
                `<a:blip r:embed="${rId}"/>` +
                `<a:stretch><a:fillRect/></a:stretch>` +
              `</pic:blipFill>` +
              `<pic:spPr>` +
                `<a:xfrm><a:off x="0" y="0"/><a:ext cx="${cxEmu}" cy="${cyEmu}"/></a:xfrm>` +
                `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>` +
              `</pic:spPr>` +
            `</pic:pic>` +
          `</a:graphicData>` +
        `</a:graphic>` +
      `</wp:inline>` +
    `</w:drawing>`
  );
}

function makeRunXml(text, opts = {}) {
  const font = opts.font || "Calibri";
  const sz = opts.sz || 21; // half-points (21 = 10.5pt)
  const color = opts.color || "1A1917";
  const rPr = [
    `<w:rFonts w:ascii="${font}" w:hAnsi="${font}" w:cs="${font}"/>`,
    opts.bold ? `<w:b/>` : "",
    opts.italic ? `<w:i/>` : "",
    opts.strike ? `<w:strike/>` : "",
    `<w:color w:val="${color}"/>`,
    opts.spacing ? `<w:spacing w:val="${opts.spacing}"/>` : "",
    `<w:sz w:val="${sz}"/><w:szCs w:val="${sz}"/>`,
    opts.underline ? `<w:u w:val="single"/>` : ""
  ]
    .filter(Boolean)
    .join("");
  return `<w:r><w:rPr>${rPr}</w:rPr><w:t xml:space="preserve">${xmlEscape(text)}</w:t></w:r>`;
}

function makeParaXml(runsXml, opts = {}) {
  const pPrParts = [];
  if (opts.keepNext) pPrParts.push(`<w:keepNext/>`);
  if (opts.keepLines) pPrParts.push(`<w:keepLines/>`);
  if (opts.pBdrBottom) {
    pPrParts.push(`<w:pBdr><w:bottom w:val="single" w:sz="6" w:space="4" w:color="141413"/></w:pBdr>`);
  }
  if (opts.spaceBefore !== undefined || opts.spaceAfter !== undefined || opts.line !== undefined) {
    pPrParts.push(
      `<w:spacing${opts.spaceBefore !== undefined ? ` w:before="${opts.spaceBefore}"` : ""}${opts.spaceAfter !== undefined ? ` w:after="${opts.spaceAfter}"` : ""}${opts.line !== undefined ? ` w:line="${opts.line}" w:lineRule="auto"` : ""}/>`
    );
  }
  if (opts.indLeft || opts.hanging) {
    pPrParts.push(
      `<w:ind${opts.indLeft ? ` w:left="${opts.indLeft}"` : ""}${opts.hanging ? ` w:hanging="${opts.hanging}"` : ""}/>`
    );
  }
  if (opts.align) pPrParts.push(`<w:jc w:val="${opts.align}"/>`);
  const pPr = pPrParts.length ? `<w:pPr>${pPrParts.join("")}</w:pPr>` : "";
  return `<w:p>${pPr}${runsXml}</w:p>`;
}

function buildObservationRunsXml(text) {
  const regex = /(\d+\s+d[ií]as\s+naturales)/gi;
  let lastIndex = 0;
  let out = "";
  let m;
  while ((m = regex.exec(text)) !== null) {
    if (m.index > lastIndex) {
      out += makeRunXml(text.slice(lastIndex, m.index), { sz: 21 });
    }
    out += makeRunXml(m[0], { sz: 21, bold: true });
    lastIndex = regex.lastIndex;
  }
  if (lastIndex < text.length) {
    out += makeRunXml(text.slice(lastIndex), { sz: 21 });
  }
  return out || makeRunXml(text, { sz: 21 });
}

function buildDocxPackageFiles(s) {
  // Dimensiones en EMU (1 inch = 914400 EMU)
  // ESENCIAL (312x49), STUDIO (253x49), ISOTIPO (125x96)
  const coverH = 490000; // ~0.536 in
  const coverEsencialW = Math.round(coverH * (312 / 49)); // ~3.41 in
  const coverStudioW = Math.round(coverH * (253 / 49));   // ~2.76 in

  const innerH = 165000; // ~0.18 in
  const innerEsencialW = Math.round(innerH * (312 / 49));
  const innerStudioW = Math.round(innerH * (253 / 49));

  const isoCoverW = 200000;
  const isoCoverH = Math.round(isoCoverW * (96 / 125));
  const isoFootW = 225000;
  const isoFootH = Math.round(isoFootW * (96 / 125));

  const bodyParts = [];

  // 1) Portada (Página 1)
  if (s.includeCoverPage) {
    const coverHeaderTbl =
      `<w:tbl>` +
        `<w:tblPr><w:tblW w:w="10080" w:type="dxa"/><w:tblBorders><w:top w:val="none"/><w:left w:val="none"/><w:bottom w:val="none"/><w:right w:val="none"/><w:insideH w:val="none"/><w:insideV w:val="none"/></w:tblBorders></w:tblPr>` +
        `<w:tblGrid><w:gridCol w:w="5400"/><w:gridCol w:w="4680"/></w:tblGrid>` +
        `<w:tr>` +
          `<w:tc><w:tcPr><w:tcW w:w="5400" w:type="dxa"/></w:tcPr>${makeParaXml(`<w:r>${makeDocxDrawingXml("rIdImgEsencial", 1, "ESENCIAL", coverEsencialW, coverH)}</w:r>`, { align: "left", spaceAfter: 0 })}</w:tc>` +
          `<w:tc><w:tcPr><w:tcW w:w="4680" w:type="dxa"/></w:tcPr>${makeParaXml(`<w:r>${makeDocxDrawingXml("rIdImgStudio", 2, "STUDIO", coverStudioW, coverH)}</w:r>`, { align: "right", spaceAfter: 0 })}</w:tc>` +
        `</w:tr>` +
      `</w:tbl>`;

    bodyParts.push(coverHeaderTbl);
    bodyParts.push(
      makeParaXml(
        makeRunXml((s.quoteEyebrow || "PRESUPUESTO").toUpperCase(), { bold: true, sz: 19, spacing: 20 }),
        { align: "center", spaceBefore: 640, spaceAfter: 160 }
      )
    );
    bodyParts.push(
      makeParaXml(
        makeRunXml((s.projectName || "PROYECTO").toUpperCase(), { font: "Georgia", sz: 34, spacing: 24 }),
        { align: "center", spaceBefore: 0, spaceAfter: 140 }
      )
    );
    bodyParts.push(
      makeParaXml(
        makeRunXml(s.quoteDateText || "", { font: "Georgia", italic: true, sz: 22 }),
        { align: "center", spaceBefore: 0, spaceAfter: 120 }
      )
    );
    bodyParts.push(
      makeParaXml(
        `<w:r>${makeDocxDrawingXml("rIdImgIsotipo", 3, "ISOTIPO", isoCoverW, isoCoverH)}</w:r>`,
        { align: "center", spaceBefore: 0, spaceAfter: 0 }
      )
    );
    bodyParts.push(`<w:p><w:r><w:br w:type="page"/></w:r></w:p>`);
  }

  // 2) Introducción
  if (s.includeIntroText) {
    const introParas = [s.introGreeting, s.introBody, s.introLead]
      .map((t) => (t || "").trim())
      .filter(Boolean);
    introParas.forEach((p, i) => {
      bodyParts.push(
        makeParaXml(makeRunXml(p, { font: "Georgia", sz: 21, color: "24221F" }), {
          spaceBefore: 0,
          spaceAfter: i === introParas.length - 1 ? 220 : 150,
          line: 264
        })
      );
    });
  }

  // 3) Tabla ALCANCE | DESCRIPCIÓN + Totales
  if (s.scopes.length > 0) {
    const tcMar = (top, left, bottom, right) =>
      `<w:tcMar><w:top w:w="${top}" w:type="dxa"/><w:left w:w="${left}" w:type="dxa"/><w:bottom w:w="${bottom}" w:type="dxa"/><w:right w:w="${right}" w:type="dxa"/></w:tcMar>`;

    const tblHeader =
      `<w:tr>` +
        `<w:trPr><w:tblHeader/><w:cantSplit/></w:trPr>` +
        `<w:tc><w:tcPr><w:tcW w:w="2760" w:type="dxa"/><w:tcBorders><w:bottom w:val="single" w:sz="8" w:space="0" w:color="141413"/></w:tcBorders>${tcMar(0, 0, 80, 140)}</w:tcPr>${makeParaXml(makeRunXml("ALCANCE", { sz: 19 }), { spaceAfter: 0 })}</w:tc>` +
        `<w:tc><w:tcPr><w:tcW w:w="7320" w:type="dxa"/><w:tcBorders><w:bottom w:val="single" w:sz="8" w:space="0" w:color="141413"/></w:tcBorders>${tcMar(0, 100, 80, 0)}</w:tcPr>${makeParaXml(makeRunXml("DESCRIPCIÓN", { sz: 19 }), { spaceAfter: 0 })}</w:tc>` +
      `</w:tr>`;

    const scopeRows = s.scopes.map((sc, idx) => {
      const isLast = idx === s.scopes.length - 1;
      const borderBottom = isLast
        ? `<w:bottom w:val="single" w:sz="10" w:space="0" w:color="141413"/>`
        : `<w:bottom w:val="dashed" w:sz="6" w:space="0" w:color="4A453F"/>`;

      const leftParas = [
        makeParaXml(makeRunXml(sc.title || "Concepto", { bold: true, sz: 21 }), {
          spaceBefore: 0,
          spaceAfter: (sc.subtitle || "").trim() ? 140 : 0,
          line: 256
        })
      ];
      if ((sc.subtitle || "").trim()) {
        leftParas.push(
          makeParaXml(makeRunXml(sc.subtitle.trim(), { bold: true, italic: true, sz: 20 }), {
            spaceBefore: 0,
            spaceAfter: 0,
            line: 250
          })
        );
      }

      const rightParas = [];
      if ((sc.leadText || "").trim()) {
        rightParas.push(
          makeParaXml(makeRunXml(sc.leadText.trim(), { sz: 21 }), {
            spaceBefore: 0,
            spaceAfter: 130,
            line: 256
          })
        );
      }

      (sc.groups || []).forEach((g) => {
        if ((g.heading || "").trim()) {
          rightParas.push(
            makeParaXml(makeRunXml(g.heading.trim(), { sz: 21 }), {
              spaceBefore: 60,
              spaceAfter: 30,
              line: 250
            })
          );
        }
        (g.items || [])
          .filter((it) => it.trim() !== "")
          .forEach((it) => {
            rightParas.push(
              makeParaXml(makeRunXml(`-   ${it.trim()}`, { sz: 21 }), {
                indLeft: 420,
                hanging: 220,
                spaceBefore: 0,
                spaceAfter: 25,
                line: 250
              })
            );
          });
      });

      (sc.notes || "")
        .split(/\n\s*\n|\r?\n/)
        .map((n) => n.trim())
        .filter(Boolean)
        .forEach((n) => {
          rightParas.push(
            makeParaXml(makeRunXml(n, { sz: 21 }), {
              spaceBefore: 110,
              spaceAfter: 30,
              line: 256
            })
          );
        });

      const curr = (sc.currencyLabel || "MXN").trim();
      rightParas.push(
        makeParaXml(makeRunXml(`${formatMoney(sc.amount, false)} ${curr}`, { bold: true, sz: 21 }), {
          spaceBefore: 150,
          spaceAfter: 0
        })
      );

      return (
        `<w:tr>` +
          `<w:tc><w:tcPr><w:tcW w:w="2760" w:type="dxa"/><w:tcBorders>${borderBottom}</w:tcBorders>${tcMar(150, 0, 160, 140)}</w:tcPr>${leftParas.join("")}</w:tc>` +
          `<w:tc><w:tcPr><w:tcW w:w="7320" w:type="dxa"/><w:tcBorders>${borderBottom}</w:tcBorders>${tcMar(150, 100, 160, 0)}</w:tcPr>${rightParas.join("")}</w:tc>` +
        `</w:tr>`
      );
    });

    // Filas de Totales
    const totals = computeTotals(s);
    const totalRows = [];

    if (totals.hasDiscount) {
      totalRows.push(
        `<w:tr><w:trPr><w:cantSplit/></w:trPr>` +
          `<w:tc><w:tcPr><w:tcW w:w="2760" w:type="dxa"/>${tcMar(140, 0, 20, 140)}</w:tcPr>${makeParaXml(makeRunXml("TOTAL", { sz: 21 }), { spaceAfter: 0 })}</w:tc>` +
          `<w:tc><w:tcPr><w:tcW w:w="7320" w:type="dxa"/>${tcMar(140, 100, 20, 0)}</w:tcPr>${makeParaXml(makeRunXml(`${formatMoney(totals.subtotal, true)} MXN`, { sz: 21, strike: true }), { spaceAfter: 0 })}</w:tc>` +
        `</w:tr>`
      );
      totalRows.push(
        `<w:tr><w:trPr><w:cantSplit/></w:trPr>` +
          `<w:tc><w:tcPr><w:tcW w:w="2760" w:type="dxa"/>${tcMar(20, 0, 40, 140)}</w:tcPr>${makeParaXml(makeRunXml(totals.discountLabel, { bold: true, sz: 21 }), { spaceAfter: 0 })}</w:tc>` +
          `<w:tc><w:tcPr><w:tcW w:w="7320" w:type="dxa"/>${tcMar(20, 100, 40, 0)}</w:tcPr>${makeParaXml(makeRunXml(`${formatMoney(totals.finalTotal, true)} MXN`, { bold: true, sz: 21 }), { spaceAfter: 0 })}</w:tc>` +
        `</w:tr>`
      );
    } else {
      totalRows.push(
        `<w:tr><w:trPr><w:cantSplit/></w:trPr>` +
          `<w:tc><w:tcPr><w:tcW w:w="2760" w:type="dxa"/>${tcMar(140, 0, 40, 140)}</w:tcPr>${makeParaXml(makeRunXml("TOTAL", { bold: true, sz: 21 }), { spaceAfter: 0 })}</w:tc>` +
          `<w:tc><w:tcPr><w:tcW w:w="7320" w:type="dxa"/>${tcMar(140, 100, 40, 0)}</w:tcPr>${makeParaXml(makeRunXml(`${formatMoney(totals.subtotal, true)} MXN`, { bold: true, sz: 21 }), { spaceAfter: 0 })}</w:tc>` +
        `</w:tr>`
      );
    }

    if (s.showIvaBreakdown) {
      totalRows.push(
        `<w:tr><w:trPr><w:cantSplit/></w:trPr>` +
          `<w:tc><w:tcPr><w:tcW w:w="2760" w:type="dxa"/>${tcMar(20, 0, 20, 140)}</w:tcPr>${makeParaXml(makeRunXml(`IVA (${Number(s.ivaRatePercent) || 16}%)`, { sz: 21 }), { spaceAfter: 0 })}</w:tc>` +
          `<w:tc><w:tcPr><w:tcW w:w="7320" w:type="dxa"/>${tcMar(20, 100, 20, 0)}</w:tcPr>${makeParaXml(makeRunXml(`${formatMoney(totals.ivaAmount, true)} MXN`, { sz: 21 }), { spaceAfter: 0 })}</w:tc>` +
        `</w:tr>`
      );
      totalRows.push(
        `<w:tr><w:trPr><w:cantSplit/></w:trPr>` +
          `<w:tc><w:tcPr><w:tcW w:w="2760" w:type="dxa"/>${tcMar(20, 0, 40, 140)}</w:tcPr>${makeParaXml(makeRunXml("TOTAL CON IVA", { bold: true, sz: 21 }), { spaceAfter: 0 })}</w:tc>` +
          `<w:tc><w:tcPr><w:tcW w:w="7320" w:type="dxa"/>${tcMar(20, 100, 40, 0)}</w:tcPr>${makeParaXml(makeRunXml(`${formatMoney(totals.grandTotal, true)} MXN`, { bold: true, sz: 21 }), { spaceAfter: 0 })}</w:tc>` +
        `</w:tr>`
      );
    }

    const scopeTable =
      `<w:tbl>` +
        `<w:tblPr><w:tblW w:w="10080" w:type="dxa"/><w:tblBorders><w:top w:val="none"/><w:left w:val="none"/><w:bottom w:val="none"/><w:right w:val="none"/><w:insideH w:val="none"/><w:insideV w:val="none"/></w:tblBorders></w:tblPr>` +
        `<w:tblGrid><w:gridCol w:w="2760"/><w:gridCol w:w="7320"/></w:tblGrid>` +
        tblHeader +
        scopeRows.join("") +
        totalRows.join("") +
      `</w:tbl>`;

    bodyParts.push(scopeTable);
  }

  // 4) Calendario de Entrega
  const hasDelivery =
    (s.deliveryMainText || "").trim() ||
    (s.deliveryEarlyNote || "").trim() ||
    (s.deliveryFootnote || "").trim();

  if (hasDelivery) {
    bodyParts.push(
      makeParaXml(makeRunXml("CALENDARIO DE ENTREGA", { font: "Georgia", sz: 24, underline: true, spacing: 14 }), {
        keepNext: true,
        spaceBefore: 260,
        spaceAfter: 120
      })
    );
    if ((s.deliveryMainText || "").trim()) {
      bodyParts.push(
        makeParaXml(makeRunXml(s.deliveryMainText.trim(), { sz: 21 }), {
          spaceBefore: 0,
          spaceAfter: 110,
          line: 256
        })
      );
    }
    if ((s.deliveryEarlyNote || "").trim()) {
      bodyParts.push(
        makeParaXml(makeRunXml(s.deliveryEarlyNote.trim(), { sz: 21 }), {
          spaceBefore: 0,
          spaceAfter: 90,
          line: 256
        })
      );
    }
    if ((s.deliveryFootnote || "").trim()) {
      bodyParts.push(
        makeParaXml(makeRunXml(s.deliveryFootnote.trim(), { sz: 19, italic: true }), {
          spaceBefore: 0,
          spaceAfter: 120,
          line: 240
        })
      );
    }
  }

  // 5) Calendario de Pagos
  if (s.payments.length > 0) {
    bodyParts.push(
      makeParaXml(makeRunXml("CALENDARIO DE PAGOS", { font: "Georgia", sz: 24, underline: true, spacing: 14 }), {
        keepNext: true,
        spaceBefore: 220,
        spaceAfter: 120
      })
    );

    const totals = computeTotals(s);
    const baseTotal = s.showIvaBreakdown ? totals.grandTotal : totals.finalTotal;

    s.payments.forEach((p, idx) => {
      let txt = `${idx + 1}.   ${(p.text || "").trim()}`;
      if (s.showPaymentAmounts) {
        const amount = baseTotal * ((Number(p.percent) || 0) / 100);
        txt += ` (${formatMoney(amount, true)} MXN)`;
      }
      bodyParts.push(
        makeParaXml(makeRunXml(txt, { sz: 21 }), {
          indLeft: 620,
          hanging: 340,
          spaceBefore: 0,
          spaceAfter: idx === s.payments.length - 1 ? 140 : 40,
          line: 256
        })
      );
    });
  }

  // 6) Observaciones
  const activeObs = s.observations.filter((o) => o.enabled && (o.text || "").trim());
  if ((s.ivaStatus || "").trim() || activeObs.length > 0 || (s.closingLine || "").trim()) {
    bodyParts.push(
      makeParaXml(makeRunXml("OBSERVACIONES", { font: "Georgia", sz: 24, underline: true, spacing: 14 }), {
        keepNext: true,
        spaceBefore: 220,
        spaceAfter: 130
      })
    );

    if ((s.ivaStatus || "").trim()) {
      bodyParts.push(
        makeParaXml(makeRunXml(s.ivaStatus.trim(), { sz: 21, bold: true }), {
          spaceBefore: 0,
          spaceAfter: 130
        })
      );
    }

    activeObs.forEach((o) => {
      bodyParts.push(
        makeParaXml(buildObservationRunsXml(o.text.trim()), {
          spaceBefore: 0,
          spaceAfter: 120,
          line: 256
        })
      );
    });

    if ((s.closingLine || "").trim()) {
      bodyParts.push(
        makeParaXml(makeRunXml(s.closingLine.trim(), { sz: 21 }), {
          spaceBefore: 40,
          spaceAfter: 160,
          line: 256
        })
      );
    }
  }

  // 7) Firmas ("Atentamente:")
  const activeSigners = s.signers.filter((sg) => (sg.name || "").trim());
  if (activeSigners.length > 0) {
    bodyParts.push(
      makeParaXml(makeRunXml(s.signoffGreeting || "Atentamente:", { font: "Georgia", italic: true, sz: 21 }), {
        keepNext: true,
        spaceBefore: 240,
        spaceAfter: 540
      })
    );

    const rows = [];
    for (let i = 0; i < activeSigners.length; i += 2) {
      const left = activeSigners[i];
      const right = activeSigners[i + 1];

      const buildCell = (sg) => {
        if (!sg) {
          return `<w:tc><w:tcPr><w:tcW w:w="5040" w:type="dxa"/></w:tcPr>${makeParaXml("", { spaceAfter: 0 })}</w:tc>`;
        }
        const paras = [
          makeParaXml(makeRunXml(sg.name.toUpperCase(), { sz: 20 }), { align: "center", spaceAfter: sg.role ? 40 : 0 })
        ];
        if ((sg.role || "").trim()) {
          paras.push(
            makeParaXml(makeRunXml(sg.role.toUpperCase(), { sz: 16, color: "6E6A63" }), { align: "center", spaceAfter: 0 })
          );
        }
        return `<w:tc><w:tcPr><w:tcW w:w="5040" w:type="dxa"/></w:tcPr>${paras.join("")}</w:tc>`;
      };

      if (activeSigners.length === 1) {
        rows.push(
          `<w:tr><w:trPr><w:cantSplit/></w:trPr>` +
            `<w:tc><w:tcPr><w:tcW w:w="10080" w:type="dxa"/><w:gridSpan w:val="2"/></w:tcPr>` +
              makeParaXml(makeRunXml(left.name.toUpperCase(), { sz: 20 }), { align: "center", spaceAfter: left.role ? 40 : 0 }) +
              ((left.role || "").trim()
                ? makeParaXml(makeRunXml(left.role.toUpperCase(), { sz: 16, color: "6E6A63" }), { align: "center", spaceAfter: 0 })
                : "") +
            `</w:tc>` +
          `</w:tr>`
        );
      } else {
        rows.push(`<w:tr><w:trPr><w:cantSplit/></w:trPr>${buildCell(left)}${buildCell(right)}</w:tr>`);
      }
    }

    bodyParts.push(
      `<w:tbl>` +
        `<w:tblPr><w:tblW w:w="10080" w:type="dxa"/><w:tblBorders><w:top w:val="none"/><w:left w:val="none"/><w:bottom w:val="none"/><w:right w:val="none"/><w:insideH w:val="none"/><w:insideV w:val="none"/></w:tblBorders></w:tblPr>` +
        `<w:tblGrid><w:gridCol w:w="5040"/><w:gridCol w:w="5040"/></w:tblGrid>` +
        rows.join("") +
      `</w:tbl>`
    );
  }

  // Header & Footer para páginas interiores (titlePg oculta header/footer en Portada)
  const titlePgXml = s.includeCoverPage ? `<w:titlePg/>` : "";
  const sectPrXml =
    `<w:sectPr>` +
      `<w:headerReference w:type="default" r:id="rIdHeader1"/>` +
      `<w:footerReference w:type="default" r:id="rIdFooter1"/>` +
      `<w:pgSz w:w="12240" w:h="15840"/>` +
      `<w:pgMar w:top="1160" w:right="1080" w:bottom="1020" w:left="1080" w:header="560" w:footer="520" w:gutter="0"/>` +
      titlePgXml +
    `</w:sectPr>`;

  const documentXml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
      `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ` +
      `xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ` +
      `xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ` +
      `xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
      `<w:body>${bodyParts.join("")}${sectPrXml}</w:body>` +
    `</w:document>`;

  const header1Xml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
      `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ` +
      `xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ` +
      `xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ` +
      `xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
      `<w:tbl>` +
        `<w:tblPr><w:tblW w:w="10080" w:type="dxa"/><w:tblBorders><w:top w:val="none"/><w:left w:val="none"/><w:bottom w:val="none"/><w:right w:val="none"/><w:insideH w:val="none"/><w:insideV w:val="none"/></w:tblBorders></w:tblPr>` +
        `<w:tblGrid><w:gridCol w:w="5040"/><w:gridCol w:w="5040"/></w:tblGrid>` +
        `<w:tr>` +
          `<w:tc><w:tcPr><w:tcW w:w="5040" w:type="dxa"/></w:tcPr>${makeParaXml(`<w:r>${makeDocxDrawingXml("rIdHdrEsencial", 11, "ESENCIAL_HDR", innerEsencialW, innerH)}</w:r>`, { align: "left", spaceAfter: 0 })}</w:tc>` +
          `<w:tc><w:tcPr><w:tcW w:w="5040" w:type="dxa"/></w:tcPr>${makeParaXml(`<w:r>${makeDocxDrawingXml("rIdHdrStudio", 12, "STUDIO_HDR", innerStudioW, innerH)}</w:r>`, { align: "right", spaceAfter: 0 })}</w:tc>` +
        `</w:tr>` +
      `</w:tbl>` +
    `</w:hdr>`;

  const phoneSpaced = (s.footerPhone || "999 275 8369").trim().split(/\s+/).join("         ");
  const emailParts = (s.footerEmail || "proyectos@esencial.studio").trim().split("@");
  const emailSpaced = emailParts.length === 2 ? `${emailParts[0]}@         ${emailParts[1]}` : s.footerEmail;
  const footerContactStr = `T.         ${phoneSpaced}         ·         ${emailSpaced}`;

  const footer1Xml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
      `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ` +
      `xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ` +
      `xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ` +
      `xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
      `<w:tbl>` +
        `<w:tblPr><w:tblW w:w="10080" w:type="dxa"/><w:tblBorders><w:top w:val="none"/><w:left w:val="none"/><w:bottom w:val="none"/><w:right w:val="none"/><w:insideH w:val="none"/><w:insideV w:val="none"/></w:tblBorders></w:tblPr>` +
        `<w:tblGrid><w:gridCol w:w="9000"/><w:gridCol w:w="1080"/></w:tblGrid>` +
        `<w:tr>` +
          `<w:tc><w:tcPr><w:tcW w:w="9000" w:type="dxa"/><w:vAlign w:val="bottom"/></w:tcPr>${makeParaXml(makeRunXml(footerContactStr, { sz: 18, color: "211F1C" }), { align: "left", spaceAfter: 0 })}</w:tc>` +
          `<w:tc><w:tcPr><w:tcW w:w="1080" w:type="dxa"/><w:vAlign w:val="bottom"/></w:tcPr>${makeParaXml(`<w:r>${makeDocxDrawingXml("rIdFtrIso", 13, "ISOTIPO_FTR", isoFootW, isoFootH)}</w:r>`, { align: "right", spaceAfter: 0 })}</w:tc>` +
        `</w:tr>` +
      `</w:tbl>` +
    `</w:ftr>`;

  const stylesXml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
      `<w:docDefaults><w:rPrDefault><w:rPr>` +
        `<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/>` +
        `<w:sz w:val="21"/><w:szCs w:val="21"/>` +
      `</w:rPr></w:rPrDefault></w:docDefaults>` +
      `<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>` +
    `</w:styles>`;

  const contentTypesXml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
      `<Default Extension="xml" ContentType="application/xml"/>` +
      `<Default Extension="png" ContentType="image/png"/>` +
      `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
      `<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>` +
      `<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>` +
      `<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>` +
    `</Types>`;

  const rootRelsXml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>` +
    `</Relationships>`;

  const docRelsXml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>` +
      `<Relationship Id="rIdHeader1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>` +
      `<Relationship Id="rIdFooter1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>` +
      `<Relationship Id="rIdImgEsencial" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>` +
      `<Relationship Id="rIdImgStudio" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image2.png"/>` +
      `<Relationship Id="rIdImgIsotipo" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image3.png"/>` +
    `</Relationships>`;

  const headerRelsXml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rIdHdrEsencial" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>` +
      `<Relationship Id="rIdHdrStudio" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image2.png"/>` +
    `</Relationships>`;

  const footerRelsXml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rIdFtrIso" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image3.png"/>` +
    `</Relationships>`;

  const enc = new TextEncoder();
  return [
    { name: "[Content_Types].xml", data: enc.encode(contentTypesXml) },
    { name: "_rels/.rels", data: enc.encode(rootRelsXml) },
    { name: "word/document.xml", data: enc.encode(documentXml) },
    { name: "word/styles.xml", data: enc.encode(stylesXml) },
    { name: "word/header1.xml", data: enc.encode(header1Xml) },
    { name: "word/footer1.xml", data: enc.encode(footer1Xml) },
    { name: "word/_rels/document.xml.rels", data: enc.encode(docRelsXml) },
    { name: "word/_rels/header1.xml.rels", data: enc.encode(headerRelsXml) },
    { name: "word/_rels/footer1.xml.rels", data: enc.encode(footerRelsXml) },
    { name: "word/media/image1.png", data: base64ToUint8Array(ESENCIAL_B64) },
    { name: "word/media/image2.png", data: base64ToUint8Array(STUDIO_B64) },
    { name: "word/media/image3.png", data: base64ToUint8Array(ISOTIPO_B64) }
  ];
}

// Tabla CRC32 y constructor PKZIP (STORE) sin dependencias externas
const CRC32_TABLE = (() => {
  const table = new Uint32Array(256);
  for (let i = 0; i < 256; i++) {
    let c = i;
    for (let k = 0; k < 8; k++) {
      c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    }
    table[i] = c >>> 0;
  }
  return table;
})();

function crc32(bytes) {
  let crc = 0xffffffff;
  for (let i = 0; i < bytes.length; i++) {
    crc = CRC32_TABLE[(crc ^ bytes[i]) & 0xff] ^ (crc >>> 8);
  }
  return (crc ^ 0xffffffff) >>> 0;
}

function buildZipStore(files) {
  const enc = new TextEncoder();
  const localChunks = [];
  const centralChunks = [];
  let offset = 0;

  files.forEach((f) => {
    const nameBytes = enc.encode(f.name);
    const data = f.data;
    const sum = crc32(data);

    const localHeader = new Uint8Array(30 + nameBytes.length);
    const lv = new DataView(localHeader.buffer);
    lv.setUint32(0, 0x04034b50, true);
    lv.setUint16(4, 20, true);
    lv.setUint16(6, 0x0800, true); // UTF-8 filename flag
    lv.setUint16(8, 0, true);      // STORE (sin compresión)
    lv.setUint16(10, 0, true);
    lv.setUint16(12, 0, true);
    lv.setUint32(14, sum, true);
    lv.setUint32(18, data.length, true);
    lv.setUint32(22, data.length, true);
    lv.setUint16(26, nameBytes.length, true);
    lv.setUint16(28, 0, true);
    localHeader.set(nameBytes, 30);

    localChunks.push(localHeader, data);

    const centralHeader = new Uint8Array(46 + nameBytes.length);
    const cv = new DataView(centralHeader.buffer);
    cv.setUint32(0, 0x02014b50, true);
    cv.setUint16(4, 20, true);
    cv.setUint16(6, 20, true);
    cv.setUint16(8, 0x0800, true);
    cv.setUint16(10, 0, true);
    cv.setUint16(12, 0, true);
    cv.setUint16(14, 0, true);
    cv.setUint32(16, sum, true);
    cv.setUint32(20, data.length, true);
    cv.setUint32(24, data.length, true);
    cv.setUint16(28, nameBytes.length, true);
    cv.setUint16(30, 0, true);
    cv.setUint16(32, 0, true);
    cv.setUint16(34, 0, true);
    cv.setUint16(36, 0, true);
    cv.setUint32(38, 0, true);
    cv.setUint32(42, offset, true);
    centralHeader.set(nameBytes, 46);

    centralChunks.push(centralHeader);
    offset += localHeader.length + data.length;
  });

  const centralSize = centralChunks.reduce((acc, c) => acc + c.length, 0);
  const eocd = new Uint8Array(22);
  const ev = new DataView(eocd.buffer);
  ev.setUint32(0, 0x06054b50, true);
  ev.setUint16(4, 0, true);
  ev.setUint16(6, 0, true);
  ev.setUint16(8, files.length, true);
  ev.setUint16(10, files.length, true);
  ev.setUint32(12, centralSize, true);
  ev.setUint32(16, offset, true);
  ev.setUint16(20, 0, true);

  const totalSize = offset + centralSize + eocd.length;
  const out = new Uint8Array(totalSize);
  let pos = 0;
  localChunks.forEach((c) => {
    out.set(c, pos);
    pos += c.length;
  });
  centralChunks.forEach((c) => {
    out.set(c, pos);
    pos += c.length;
  });
  out.set(eocd, pos);
  return out;
}

/* ==========================================================================
   Calculadora Orientativa, Notas Rápidas / Dictado e Historial
   ========================================================================== */
function bindCalculatorAndNotes() {
  const typologyDefaults = {
    residencial: 650,
    comercial: 750,
    remodelacion: 450,
    interiorismo: 580
  };

  const areaEl = document.getElementById("calcAreaM2");
  const typEl = document.getElementById("calcTypology");
  const rateEl = document.getElementById("calcRateM2");
  const rendersEl = document.getElementById("calcRendersCount");
  const execEl = document.getElementById("calcIncludeExec");

  const computeCalc = () => {
    const area = Math.max(0, Number(areaEl.value) || 0);
    const rate = Math.max(0, Number(rateEl.value) || 0);
    const renders = Math.max(0, Number(rendersEl.value) || 0);
    const includeExec = execEl.checked;

    const baseDesign = area * rate;
    const conceptual = Math.round((baseDesign * 0.55) / 500) * 500;
    const comercial = Math.round((renders * 3200 + 4500) / 500) * 500;
    const ejecutivo = includeExec ? Math.round((baseDesign * 0.45) / 500) * 500 : 0;
    const total = conceptual + comercial + ejecutivo;

    document.getElementById("calcConceptualOut").textContent = `${formatMoney(conceptual)} MXN`;
    document.getElementById("calcComercialOut").textContent = `${formatMoney(comercial)} MXN`;
    document.getElementById("calcEjecutivoOut").textContent = `${formatMoney(ejecutivo)} MXN`;
    document.getElementById("calcTotalOut").textContent = `${formatMoney(total)} MXN`;
    return { conceptual, comercial, ejecutivo, renders, includeExec };
  };

  typEl.addEventListener("change", () => {
    rateEl.value = typologyDefaults[typEl.value] || 650;
    computeCalc();
  });

  [areaEl, rateEl, rendersEl, execEl].forEach((el) => {
    el.addEventListener("input", computeCalc);
    el.addEventListener("change", computeCalc);
  });

  computeCalc();

  document.getElementById("btnApplyCalcToQuote").addEventListener("click", () => {
    const res = computeCalc();
    const firstScope = deepClone(SCOPE_PRESETS[0].data);
    firstScope.amount = res.conceptual + res.comercial;
    firstScope.groups[2].items = [
      `${res.renders} perspectivas calidad HD`,
      "Plantas editadas para venta"
    ];

    const newScopes = [firstScope];
    if (res.includeExec) {
      const execScope = deepClone(SCOPE_PRESETS[4].data);
      execScope.amount = res.ejecutivo;
      newScopes.push(execScope);
    }

    state.scopes = newScopes;
    persistCurrentState();
    syncEditorFromState();
    renderDocumentPages();
    showToast("Alcances generados desde la calculadora");
  });

  // Dictado Web Speech API (es-MX)
  const btnDictate = document.getElementById("btnToggleDictation");
  const statusEl = document.getElementById("recordingStatus");
  btnDictate.addEventListener("click", () => {
    const SpeechRec = window.SpeechRecognition || window.webkitSpeechRecognition;
    if (!SpeechRec) {
      showToast("Dictado en vivo disponible en Chrome o Edge. También puedes pegar notas.");
      return;
    }

    if (isDictating && speechRecognition) {
      speechRecognition.stop();
      return;
    }

    speechRecognition = new SpeechRec();
    speechRecognition.lang = "es-MX";
    speechRecognition.continuous = true;
    speechRecognition.interimResults = true;

    const textarea = document.getElementById("quickNotesInput");
    let committed = textarea.value ? textarea.value + "\n" : "";

    speechRecognition.onstart = () => {
      isDictating = true;
      btnDictate.classList.add("recording");
      btnDictate.querySelector("span").textContent = "Detener Dictado";
      statusEl.textContent = "Escuchando en español (es-MX)...";
      statusEl.classList.remove("hidden");
    };

    speechRecognition.onresult = (event) => {
      let interim = "";
      let newFinal = "";
      for (let i = event.resultIndex; i < event.results.length; i++) {
        const t = event.results[i][0].transcript;
        if (event.results[i].isFinal) newFinal += t + "\n";
        else interim += t;
      }
      if (newFinal) committed += newFinal;
      textarea.value = committed + interim;
    };

    const resetDictationUI = () => {
      isDictating = false;
      btnDictate.classList.remove("recording");
      btnDictate.querySelector("span").textContent = "Dictar a Texto";
      statusEl.classList.add("hidden");
    };

    speechRecognition.onerror = resetDictationUI;
    speechRecognition.onend = resetDictationUI;
    speechRecognition.start();
  });

  document.getElementById("btnLoadSampleNotes").addEventListener("click", () => {
    document.getElementById("quickNotesInput").value = SAMPLE_QUICK_NOTES;
    showToast("Ejemplo de notas rápidas cargado");
  });

  document.getElementById("btnParseQuickNotes").addEventListener("click", () => {
    const raw = document.getElementById("quickNotesInput").value.trim();
    if (!raw) {
      showToast("Escribe o dicta notas rápidas primero");
      return;
    }
    applyQuickNotes(raw);
  });
}

function applyQuickNotes(rawText) {
  const mode = document.getElementById("notesMergeMode").value;
  const lines = rawText.split(/\r?\n/).map((l) => l.trim()).filter(Boolean);
  const parsedScopes = [];

  lines.forEach((line) => {
    if (/^proyecto\s*:/i.test(line)) {
      state.projectName = line.split(":").slice(1).join(":").trim().toUpperCase();
      state.quoteCode = buildAutoQuoteCode(state.projectName);
    } else if (/^cliente\s*:/i.test(line)) {
      state.clientName = line.split(":").slice(1).join(":").trim();
    } else if (/^fecha\s*:/i.test(line)) {
      state.quoteDateText = line.split(":").slice(1).join(":").trim();
    } else if (/^descuento\s*:/i.test(line)) {
      const valStr = line.split(":").slice(1).join(":").trim();
      const pctMatch = valStr.match(/(\d+(?:\.\d+)?)\s*%/);
      if (pctMatch) {
        state.discountMode = "percent";
        state.discountPercent = Number(pctMatch[1]);
      } else {
        const num = Number(valStr.replace(/[^0-9.]/g, "")) || 0;
        if (num > 0) {
          state.discountMode = "fixed";
          state.discountFixed = num;
        }
      }
    } else if (/^entrega\s*:/i.test(line)) {
      const rest = line.split(":").slice(1).join(":").trim();
      state.deliveryMainText = `La entrega del trabajo completo tiene un tiempo estimado de ${rest} (dependiendo del número de revisiones que sean necesarias).`;
    } else if (/^vigencia\s*:/i.test(line)) {
      const m = line.match(/(\d+)/);
      if (m) {
        state.obsValidityDays = Number(m[1]);
        const valObs = state.observations.find((o) => o.id === "validity");
        if (valObs) {
          valObs.text = `Los precios incluidos en este presupuesto tienen una vigencia de ${state.obsValidityDays} días naturales para su aceptación y firma. Una vez iniciado el proyecto y pagado el anticipo, los costos se mantendrán fijos durante todo el periodo de desarrollo del proyecto.`;
        }
      }
    } else if (/^(alcance|concepto|partida)\s*:/i.test(line)) {
      const body = line.split(":").slice(1).join(":").trim();
      const parts = body.split("|").map((p) => p.trim());
      const title = parts[0] || "Nuevo Concepto";
      const amount = Number((parts[1] || "0").replace(/[^0-9.]/g, "")) || 0;
      const items = (parts[2] || "")
        .split(/,|;/)
        .map((x) => x.trim())
        .filter(Boolean);

      parsedScopes.push({
        title,
        subtitle: "",
        leadText: "",
        groups: items.length ? [{ heading: "Entregables", items }] : [],
        notes: "",
        amount,
        currencyLabel: "MXN",
        included: true
      });
    }
  });

  if (parsedScopes.length > 0) {
    state.scopes = mode === "replace" ? parsedScopes : [...state.scopes, ...parsedScopes];
  }

  persistCurrentState();
  syncEditorFromState();
  renderDocumentPages();
  showToast("Notas aplicadas al presupuesto");
}

/* ==========================================================================
   Pestañas e Historial Local + Backup JSON
   ========================================================================== */
function bindTabNavigation() {
  const buttons = document.querySelectorAll(".tab-btn");
  const tabs = document.querySelectorAll(".tab-content");

  buttons.forEach((btn) => {
    btn.addEventListener("click", () => {
      const target = btn.dataset.tab;
      buttons.forEach((b) => {
        b.classList.toggle("active", b === btn);
        b.setAttribute("aria-selected", String(b === btn));
      });
      tabs.forEach((tab) => {
        tab.classList.toggle("active", tab.id === `tab-${target}`);
      });
    });
  });

  document.getElementById("btnSaveCurrentToHistory").addEventListener("click", () => {
    const history = getHistory();
    history.unshift({
      id: Date.now(),
      savedAt: new Date().toLocaleString("es-MX"),
      data: deepClone(state)
    });
    localStorage.setItem(STORAGE_HISTORY_KEY, JSON.stringify(history.slice(0, 40)));
    renderHistoryList();
    showToast("Presupuesto guardado en historial");
  });

  document.getElementById("btnExportJson").addEventListener("click", () => {
    const payload = {
      exportedAt: new Date().toISOString(),
      current: state,
      history: getHistory()
    };
    const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" });
    triggerBlobDownload(blob, `ESENCIAL_PRESUPUESTOS_RESPALDO.json`);
    showToast("Respaldo JSON descargado");
  });

  const importInput = document.getElementById("importJsonInput");
  document.getElementById("btnImportJsonBtn").addEventListener("click", () => importInput.click());
  importInput.addEventListener("change", (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => {
      try {
        const parsed = JSON.parse(String(reader.result || "{}"));
        if (parsed.current) {
          state = normalizeState(parsed.current);
          persistCurrentState();
          syncEditorFromState();
          applyPaperTheme();
          renderDocumentPages();
        }
        if (Array.isArray(parsed.history)) {
          localStorage.setItem(STORAGE_HISTORY_KEY, JSON.stringify(parsed.history.slice(0, 40)));
          renderHistoryList();
        }
        showToast("Respaldo JSON importado correctamente");
      } catch {
        showToast("El archivo JSON no tiene un formato válido");
      }
    };
    reader.readAsText(file);
    e.target.value = "";
  });
}

function getHistory() {
  try {
    const raw = localStorage.getItem(STORAGE_HISTORY_KEY);
    return raw ? JSON.parse(raw) : [];
  } catch {
    return [];
  }
}

function renderHistoryList() {
  const listEl = document.getElementById("historyList");
  const history = getHistory();
  listEl.innerHTML = "";

  if (history.length === 0) {
    listEl.innerHTML = `<p class="helper-text">Aún no has archivado cotizaciones. Haz clic en "+ Guardar copia" para almacenar versiones de tus clientes.</p>`;
    return;
  }

  history.forEach((entry, idx) => {
    const norm = normalizeState(entry.data || {});
    const totals = computeTotals(norm);
    const item = document.createElement("div");
    item.className = "history-item";
    item.innerHTML = `
      <div class="history-item-info">
        <strong>${escapeHtmlText(norm.projectName || "Sin proyecto")} · ${formatMoney(totals.finalTotal)} MXN</strong>
        <span>${escapeHtmlText(norm.quoteDateText || "")} · Guardado ${escapeHtmlText(entry.savedAt)}</span>
      </div>
      <div class="history-item-btns">
        <button type="button" class="mini-btn hist-load">Abrir</button>
        <button type="button" class="mini-btn danger hist-del">✕</button>
      </div>
    `;

    item.querySelector(".hist-load").addEventListener("click", () => {
      state = normalizeState(entry.data);
      persistCurrentState();
      syncEditorFromState();
      applyPaperTheme();
      renderDocumentPages();
      showToast(`Cargado: ${state.projectName}`);
    });

    item.querySelector(".hist-del").addEventListener("click", () => {
      const updated = getHistory();
      updated.splice(idx, 1);
      localStorage.setItem(STORAGE_HISTORY_KEY, JSON.stringify(updated));
      renderHistoryList();
    });

    listEl.appendChild(item);
  });
}

/* ==========================================================================
   Zoom de Vista Previa
   ========================================================================== */
function bindZoomControls() {
  const applyZoom = (val) => {
    currentZoom = Math.min(1.25, Math.max(0.32, Number(val.toFixed(2))));
    document.documentElement.style.setProperty("--preview-zoom", String(currentZoom));
    document.getElementById("zoomLabel").textContent = `${Math.round(currentZoom * 100)}%`;
  };

  document.getElementById("btnZoomIn").addEventListener("click", () => applyZoom(currentZoom + 0.08));
  document.getElementById("btnZoomOut").addEventListener("click", () => applyZoom(currentZoom - 0.08));
  document.getElementById("btnZoomFit").addEventListener("click", () => {
    const pad = window.innerWidth <= 768 ? 28 : 64;
    const viewportW = document.getElementById("previewViewport").clientWidth - pad;
    const fit = Math.min(0.95, Math.max(0.36, viewportW / 816));
    applyZoom(fit);
  });

  if (window.innerWidth <= 768) {
    const mobileFit = Math.min(0.95, Math.max(0.36, (window.innerWidth - 28) / 816));
    applyZoom(mobileFit);
  } else {
    applyZoom(currentZoom);
  }
}

/* ==========================================================================
   Utilidades
   ========================================================================== */
function formatTodayLongSpanish() {
  const today = new Date();
  const months = [
    "enero", "febrero", "marzo", "abril", "mayo", "junio",
    "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"
  ];
  return `${today.getDate()} de ${months[today.getMonth()]} del ${today.getFullYear()}`;
}

function buildAutoQuoteCode(project) {
  const today = new Date();
  const monthsShort = ["ENE", "FEB", "MAR", "ABR", "MAY", "JUN", "JUL", "AGO", "SEP", "OCT", "NOV", "DIC"];
  const dd = String(today.getDate()).padStart(2, "0");
  const mmm = monthsShort[today.getMonth()];
  const yyyy = today.getFullYear();
  const cleanProj = (project || "PROYECTO")
    .trim()
    .toUpperCase()
    .replace(/[\\/:*?"<>|]/g, "")
    .slice(0, 32);
  return `ES_${cleanProj}_${dd} ${mmm} ${yyyy}`;
}

function triggerBlobDownload(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  setTimeout(() => {
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }, 500);
}

function escapeHtmlText(str) {
  return String(str ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function escapeHtmlAttr(str) {
  return escapeHtmlText(str).replace(/'/g, "&#39;");
}

function showToast(message) {
  const toast = document.getElementById("toast");
  toast.textContent = message;
  toast.classList.add("show");
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => {
    toast.classList.remove("show");
  }, 3000);
}
