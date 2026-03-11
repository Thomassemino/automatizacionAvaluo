const els = {
  fileInput: document.getElementById("fileInput"),
  jsonTextInput: document.getElementById("jsonTextInput"),
  loadTextBtn: document.getElementById("loadTextBtn"),
  selectedFile: document.getElementById("selectedFile"),
  messageBox: document.getElementById("messageBox"),
  editorSection: document.getElementById("editorSection"),
  previewSection: document.getElementById("previewSection"),
  actionsSection: document.getElementById("actionsSection"),
  metadataContainer: document.getElementById("metadataContainer"),
  periodsContainer: document.getElementById("periodsContainer"),
  alertsSummaryContainer: document.getElementById("alertsSummaryContainer"),
  addPeriodBtn: document.getElementById("addPeriodBtn"),
  generateBtn: document.getElementById("generateBtn"),
  downloadLink: document.getElementById("downloadLink"),
  templateInfo: document.getElementById("templateInfo"),
  refreshPreviewBtn: document.getElementById("refreshPreviewBtn"),
  previewStatus: document.getElementById("previewStatus"),
  previewTableContainer: document.getElementById("previewTableContainer"),
  previewWrap: document.getElementById("previewWrap"),
  previewNav: document.getElementById("previewNav"),
};

const metadataFields = [
  {
    path: "empresa_detectada",
    label: "Empresa detectada",
    type: "text",
    help: "Extraido del encabezado principal del documento (OCR o celda A1).",
  },
  {
    path: "moneda",
    label: "Moneda",
    type: "text",
    help: "Inferido por el formato de moneda encontrado en encabezados ($, MXN, Pesos).",
  },
  {
    path: "periodos_encontrados",
    label: "Periodos encontrados",
    type: "years",
    help: "Lista de aÃ±os fiscales unicos detectados en el lote procesado. Escribe aÃ±os separados por coma.",
  },
];

const periodFieldGroups = [
  {
    title: "A. Identificadores temporales",
    fields: [
      {
        path: "anio",
        label: "Anio fiscal",
        type: "integer",
        help: "Ano fiscal detectado en el documento.",
      },
      {
        path: "tipo_periodo",
        label: "Tipo de periodo",
        type: "select",
        options: ["ANUAL_CERRADO", "PARCIAL"],
        help: "ANUAL_CERRADO para enero-diciembre; PARCIAL para cortes intermedios.",
      },
    ],
  },
  {
    title: "B. Estado de resultados",
    fields: [
      {
        path: "estado_resultados.ingresos_operativos_netos",
        label: "Ingresos operativos netos",
        type: "number",
        help: "Ventas o ingresos netos operativos del periodo.",
      },
      {
        path: "estado_resultados.devoluciones_y_descuentos_sobre_ventas",
        label: "Devoluciones y descuentos",
        type: "number",
        help: "Devoluciones, rebajas o descuentos sobre ventas.",
      },
      {
        path: "estado_resultados.costo_de_ventas",
        label: "Costo de ventas",
        type: "number",
        help: "Costo de lo vendido o costo directo.",
      },
      {
        path: "estado_resultados.utilidad_bruta",
        label: "Utilidad bruta",
        type: "number",
        help: "Ingresos operativos netos menos costo de ventas.",
      },
      {
        path: "estado_resultados.gastos_operativos",
        label: "Gastos de operacion",
        type: "number",
        help: "Gasto operativo general (si aplica).",
      },
      {
        path: "estado_resultados.gastos_generales",
        label: "Gastos generales",
        type: "number",
        help: "Gastos generales del periodo.",
      },
      {
        path: "estado_resultados.gastos_de_administracion",
        label: "Gastos de administracion",
        type: "number",
        help: "Gastos administrativos del periodo.",
      },
      {
        path: "estado_resultados.gastos_de_venta",
        label: "Gastos de venta",
        type: "number",
        help: "Gastos comerciales o de distribucion.",
      },
      {
        path: "estado_resultados.gastos_de_personal",
        label: "Gastos de personal",
        type: "number",
        help: "Nomina y prestaciones.",
      },
      {
        path: "estado_resultados.gastos_por_arrendamientos",
        label: "Gastos por arrendamientos",
        type: "number",
        help: "Rentas o arrendamientos del periodo.",
      },
      {
        path: "estado_resultados.servicios_externos_y_honorarios",
        label: "Servicios externos",
        type: "number",
        help: "Honorarios y servicios de terceros.",
      },
      {
        path: "estado_resultados.otros_ingresos_operativos",
        label: "Otros ingresos operativos",
        type: "number",
        help: "Ingresos operativos no recurrentes o accesorios.",
      },
      {
        path: "estado_resultados.otros_gastos_operativos",
        label: "Otros gastos operativos",
        type: "number",
        help: "Gastos operativos no recurrentes.",
      },
      {
        path: "estado_resultados.ebitda",
        label: "EBITDA",
        type: "number",
        help: "Resultado antes de intereses, impuestos, depreciacion y amortizacion.",
      },
      {
        path: "estado_resultados.depreciacion_del_periodo",
        label: "Depreciacion del periodo",
        type: "number",
        help: "Depreciacion reconocida en resultados.",
      },
      {
        path: "estado_resultados.amortizacion_del_periodo",
        label: "Amortizacion del periodo",
        type: "number",
        help: "Amortizacion reconocida en resultados.",
      },
      {
        path: "estado_resultados.depreciacion_y_amortizacion",
        label: "Depreciacion y amortizacion",
        type: "number",
        help: "Total de depreciacion y amortizacion.",
      },
      {
        path: "estado_resultados.resultado_financiero_neto",
        label: "Resultado financiero neto",
        type: "number",
        help: "RIF o neto financiero.",
      },
      {
        path: "estado_resultados.ingresos_financieros",
        label: "Ingresos financieros",
        type: "number",
        help: "Intereses ganados y productos financieros.",
      },
      {
        path: "estado_resultados.gastos_financieros",
        label: "Gastos financieros",
        type: "number",
        help: "Intereses pagados y costos financieros.",
      },
      {
        path: "estado_resultados.otros_ingresos_no_operativos",
        label: "Otros ingresos no operativos",
        type: "number",
        help: "Ingresos no operativos o extraordinarios.",
      },
      {
        path: "estado_resultados.otros_gastos_no_operativos",
        label: "Otros gastos no operativos",
        type: "number",
        help: "Gastos no operativos o extraordinarios.",
      },
      {
        path: "estado_resultados.utilidad_operativa_ebit",
        label: "Utilidad operativa EBIT",
        type: "number",
        help: "Resultado operativo (EBIT).",
      },
      {
        path: "estado_resultados.utilidad_antes_de_impuestos",
        label: "Utilidad antes de impuestos",
        type: "number",
        help: "Resultado antes de ISR/PTU.",
      },
      {
        path: "estado_resultados.isr_diferido",
        label: "ISR diferido",
        type: "number",
        help: "Impuesto diferido del periodo.",
      },
      {
        path: "estado_resultados.isr_corriente",
        label: "ISR corriente",
        type: "number",
        help: "ISR del ejercicio.",
      },
      {
        path: "estado_resultados.provision_ptu",
        label: "Provision PTU",
        type: "number",
        help: "Provision de participacion a trabajadores.",
      },
      {
        path: "estado_resultados.total_impuestos_generico",
        label: "Total impuestos generico",
        type: "number",
        help: "Usar solo cuando el origen no separa impuestos.",
      },
      {
        path: "estado_resultados.impuesto_a_la_utilidad",
        label: "Impuesto a la utilidad",
        type: "number",
        help: "Impuesto total del periodo cuando venga identificado.",
      },
      {
        path: "estado_resultados.utilidad_neta",
        label: "Utilidad neta",
        type: "number",
        help: "Resultado final del ejercicio.",
      },
    ],
  },
  {
    title: "C. Balance general",
    fields: [
      {
        path: "balance_general.activos.circulante.efectivo_y_equivalentes",
        label: "Efectivo y equivalentes",
        type: "number",
        help: "Caja, bancos e inversiones temporales.",
      },
      {
        path: "balance_general.activos.circulante.cuentas_por_cobrar_clientes",
        label: "Cuentas por cobrar",
        type: "number",
        help: "Clientes comerciales.",
      },
      {
        path: "balance_general.activos.circulante.impuestos_a_favor_cp",
        label: "Impuestos a favor CP",
        type: "number",
        help: "Saldos a favor de impuestos de corto plazo.",
      },
      {
        path: "balance_general.activos.circulante.deudores_diversos_cp",
        label: "Deudores diversos CP",
        type: "number",
        help: "Cuentas por cobrar no comerciales de corto plazo.",
      },
      {
        path: "balance_general.activos.circulante.inventarios",
        label: "Inventarios",
        type: "number",
        help: "Existencias al cierre.",
      },
      {
        path: "balance_general.activos.circulante.pagos_anticipados",
        label: "Pagos anticipados",
        type: "number",
        help: "Rentas, seguros y otros prepagados.",
      },
      {
        path: "balance_general.activos.circulante.otros_activos_circulantes",
        label: "Otros activos circulantes",
        type: "number",
        help: "Activos circulantes no clasificados.",
      },
      {
        path: "balance_general.activos.circulante.total_activo_circulante",
        label: "Total activo circulante",
        type: "number",
        help: "Total de activo circulante.",
      },
      {
        path: "balance_general.activos.no_circulante.equipo_de_transporte",
        label: "Equipo de transporte",
        type: "number",
        help: "Vehiculos y flota.",
      },
      {
        path: "balance_general.activos.no_circulante.equipo_de_computo",
        label: "Equipo de computo",
        type: "number",
        help: "Hardware y TI.",
      },
      {
        path: "balance_general.activos.no_circulante.mobiliario_y_equipo_de_oficina",
        label: "Mobiliario y equipo",
        type: "number",
        help: "Muebles y equipo de oficina.",
      },
      {
        path: "balance_general.activos.no_circulante.propiedad_planta_y_equipo_neto",
        label: "PP&E neto",
        type: "number",
        help: "Propiedad, planta y equipo neto.",
      },
      {
        path: "balance_general.activos.no_circulante.depreciacion_acumulada_historica",
        label: "Depreciacion acumulada",
        type: "number",
        help: "Depreciacion historica acumulada.",
      },
      {
        path: "balance_general.activos.no_circulante.activos_intangibles_neto",
        label: "Activos intangibles neto",
        type: "number",
        help: "Intangibles netos al cierre.",
      },
      {
        path: "balance_general.activos.no_circulante.activos_diferidos",
        label: "Activos diferidos",
        type: "number",
        help: "Activos diferidos de largo plazo.",
      },
      {
        path: "balance_general.activos.no_circulante.total_activo_no_circulante",
        label: "Total activo no circulante",
        type: "number",
        help: "Total de activo no circulante.",
      },
      {
        path: "balance_general.activos.total_activos",
        label: "Total activos",
        type: "number",
        help: "Total del activo.",
      },
      {
        path: "balance_general.pasivos.corto_plazo.proveedores",
        label: "Proveedores",
        type: "number",
        help: "Pasivo con proveedores.",
      },
      {
        path: "balance_general.pasivos.corto_plazo.deuda_financiera_cp",
        label: "Deuda financiera CP",
        type: "number",
        help: "Pasivos financieros de corto plazo.",
      },
      {
        path: "balance_general.pasivos.corto_plazo.impuestos_y_cuotas_por_pagar",
        label: "Impuestos y cuotas por pagar",
        type: "number",
        help: "Obligaciones fiscales por pagar.",
      },
      {
        path: "balance_general.pasivos.corto_plazo.anticipo_de_clientes",
        label: "Anticipo de clientes",
        type: "number",
        help: "Cobros anticipados de clientes.",
      },
      {
        path: "balance_general.pasivos.corto_plazo.acreedores_diversos",
        label: "Acreedores diversos",
        type: "number",
        help: "Otros acreedores operativos.",
      },
      {
        path: "balance_general.pasivos.corto_plazo.provisiones",
        label: "Provisiones",
        type: "number",
        help: "Provisiones del periodo.",
      },
      {
        path: "balance_general.pasivos.corto_plazo.otros_pasivos_corto_plazo",
        label: "Otros pasivos CP",
        type: "number",
        help: "Pasivos de corto plazo no clasificados.",
      },
      {
        path: "balance_general.pasivos.corto_plazo.total_pasivo_corto_plazo",
        label: "Total pasivo CP",
        type: "number",
        help: "Total pasivo corto plazo.",
      },
      {
        path: "balance_general.pasivos.largo_plazo.dividendos_decretados",
        label: "Dividendos decretados",
        type: "number",
        help: "Dividendos por pagar.",
      },
      {
        path: "balance_general.pasivos.largo_plazo.pasivo_por_arrendamiento",
        label: "Pasivo por arrendamiento",
        type: "number",
        help: "Arrendamientos de largo plazo.",
      },
      {
        path: "balance_general.pasivos.largo_plazo.deuda_financiera_lp",
        label: "Deuda financiera LP",
        type: "number",
        help: "Pasivos financieros de largo plazo.",
      },
      {
        path: "balance_general.pasivos.largo_plazo.total_pasivo_largo_plazo",
        label: "Total pasivo LP",
        type: "number",
        help: "Total pasivo largo plazo.",
      },
      {
        path: "balance_general.pasivos.total_pasivos",
        label: "Total pasivos",
        type: "number",
        help: "Suma de pasivos CP y LP.",
      },
      {
        path: "balance_general.capital_contable.capital_social",
        label: "Capital social",
        type: "number",
        help: "Aportaciones de socios.",
      },
      {
        path: "balance_general.capital_contable.utilidades_ejercicios_anteriores",
        label: "Utilidades ejercicios anteriores",
        type: "number",
        help: "Resultados acumulados.",
      },
      {
        path: "balance_general.capital_contable.resultado_del_ejercicio_balance",
        label: "Resultado del ejercicio",
        type: "number",
        help: "Resultado del ejercicio en balance.",
      },
      {
        path: "balance_general.capital_contable.total_capital_contable",
        label: "Total capital contable",
        type: "number",
        help: "Total de capital contable.",
      },
    ],
  },
  {
    title: "D. Alertas",
    fields: [
      {
        path: "alertaDeAI",
        label: "Alerta de IA",
        type: "textarea",
        help: "Observaciones de auditoria y validaciones contables.",
      },
    ],
  },
];

const numericPeriodPaths = periodFieldGroups
  .flatMap((group) => group.fields)
  .filter((field) => field.type === "number")
  .map((field) => field.path);

const rerenderPeriodPaths = new Set([
  "anio",
  "tipo_periodo",
  "balance_general.activos.total_activos",
  "balance_general.pasivos.total_pasivos",
  "balance_general.capital_contable.total_capital_contable",
]);

let appState = null;

function createDefaultPeriod(year = new Date().getFullYear()) {
  return {
    anio: year,
    tipo_periodo: "ANUAL_CERRADO",
    estado_resultados: {
      ingresos_operativos_netos: 0,
      devoluciones_y_descuentos_sobre_ventas: 0,
      costo_de_ventas: 0,
      utilidad_bruta: 0,
      gastos_de_venta: 0,
      gastos_de_administracion: 0,
      gastos_generales: 0,
      gastos_operativos: 0,
      gastos_de_personal: 0,
      gastos_por_arrendamientos: 0,
      servicios_externos_y_honorarios: 0,
      otros_ingresos_operativos: 0,
      otros_gastos_operativos: 0,
      ebitda: 0,
      depreciacion_del_periodo: 0,
      amortizacion_del_periodo: 0,
      depreciacion_y_amortizacion: 0,
      utilidad_operativa_ebit: 0,
      ingresos_financieros: 0,
      gastos_financieros: 0,
      resultado_financiero_neto: 0,
      otros_ingresos_no_operativos: 0,
      otros_gastos_no_operativos: 0,
      utilidad_antes_de_impuestos: 0,
      isr_diferido: 0,
      isr_corriente: 0,
      provision_ptu: 0,
      total_impuestos_generico: 0,
      impuesto_a_la_utilidad: 0,
      utilidad_neta: 0,
    },
    balance_general: {
      activos: {
        circulante: {
          efectivo_y_equivalentes: 0,
          cuentas_por_cobrar_clientes: 0,
          impuestos_a_favor_cp: 0,
          deudores_diversos_cp: 0,
          inventarios: 0,
          pagos_anticipados: 0,
          otros_activos_circulantes: 0,
          total_activo_circulante: 0,
        },
        no_circulante: {
          equipo_de_transporte: 0,
          equipo_de_computo: 0,
          mobiliario_y_equipo_de_oficina: 0,
          propiedad_planta_y_equipo_neto: 0,
          depreciacion_acumulada_historica: 0,
          activos_intangibles_neto: 0,
          activos_diferidos: 0,
          total_activo_no_circulante: 0,
        },
        total_activos: 0,
      },
      pasivos: {
        corto_plazo: {
          proveedores: 0,
          deuda_financiera_cp: 0,
          impuestos_y_cuotas_por_pagar: 0,
          anticipo_de_clientes: 0,
          acreedores_diversos: 0,
          provisiones: 0,
          otros_pasivos_corto_plazo: 0,
          total_pasivo_corto_plazo: 0,
        },
        largo_plazo: {
          deuda_financiera_lp: 0,
          pasivo_por_arrendamiento: 0,
          dividendos_decretados: 0,
          total_pasivo_largo_plazo: 0,
        },
        total_pasivos: 0,
      },
      capital_contable: {
        capital_social: 0,
        utilidades_ejercicios_anteriores: 0,
        resultado_del_ejercicio_balance: 0,
        total_capital_contable: 0,
      },
    },
    alertaDeAI: "",
  };
}

function isObject(value) {
  return value && typeof value === "object" && !Array.isArray(value);
}

function deepMerge(target, source) {
  if (!isObject(source)) {
    return target;
  }

  Object.entries(source).forEach(([key, value]) => {
    if (isObject(value) && isObject(target[key])) {
      deepMerge(target[key], value);
      return;
    }
    target[key] = value;
  });

  return target;
}

function toNumber(value, fallback = 0) {
  if (value === null || value === undefined || value === "") {
    return fallback;
  }
  const n = Number(value);
  return Number.isFinite(n) ? n : fallback;
}

function toInteger(value, fallback = new Date().getFullYear()) {
  const parsed = parseInt(value, 10);
  return Number.isFinite(parsed) ? parsed : fallback;
}

function parseYears(value) {
  if (Array.isArray(value)) {
    return [...new Set(value.map((y) => toInteger(y, NaN)).filter(Number.isFinite))];
  }

  if (typeof value === "string") {
    return [
      ...new Set(
        value
          .split(",")
          .map((chunk) => toInteger(chunk.trim(), NaN))
          .filter(Number.isFinite)
      ),
    ];
  }

  return [];
}

function readPath(obj, path) {
  return path.split(".").reduce((acc, key) => (acc ? acc[key] : undefined), obj);
}

function writePath(obj, path, value) {
  const keys = path.split(".");
  let current = obj;
  for (let i = 0; i < keys.length - 1; i += 1) {
    const key = keys[i];
    if (!isObject(current[key])) {
      current[key] = {};
    }
    current = current[key];
  }
  current[keys[keys.length - 1]] = value;
}

function normalizePeriod(period) {
  const guessedYear = toInteger(period?.anio, new Date().getFullYear());
  const normalized = createDefaultPeriod(guessedYear);
  deepMerge(normalized, period || {});

  normalized.anio = toInteger(normalized.anio, guessedYear);
  normalized.tipo_periodo =
    normalized.tipo_periodo === "PARCIAL" ? "PARCIAL" : "ANUAL_CERRADO";
  normalized.alertaDeAI = String(normalized.alertaDeAI ?? "");

  numericPeriodPaths.forEach((path) => {
    writePath(normalized, path, toNumber(readPath(normalized, path), 0));
  });

  return normalized;
}

function normalizeData(rawData) {
  if (!isObject(rawData)) {
    throw new Error("El archivo JSON debe contener un objeto raiz.");
  }

  const normalized = {
    metadata: {
      empresa_detectada: "",
      moneda: "",
      periodos_encontrados: [],
    },
    datos_financieros: [],
  };

  if (isObject(rawData.metadata)) {
    normalized.metadata.empresa_detectada = String(
      rawData.metadata.empresa_detectada ?? ""
    );
    normalized.metadata.moneda = String(rawData.metadata.moneda ?? "");
    normalized.metadata.periodos_encontrados = parseYears(
      rawData.metadata.periodos_encontrados
    );
  }

  if (!Array.isArray(rawData.datos_financieros) || rawData.datos_financieros.length === 0) {
    throw new Error("El JSON debe incluir datos_financieros con al menos un periodo.");
  }

  normalized.datos_financieros = rawData.datos_financieros
    .map((period) => normalizePeriod(period))
    .sort((a, b) => a.anio - b.anio);

  if (normalized.metadata.periodos_encontrados.length === 0) {
    normalized.metadata.periodos_encontrados = normalized.datos_financieros.map(
      (period) => period.anio
    );
  }

  return normalized;
}

function syncMetadataYears() {
  if (!appState) {
    return;
  }
  const years = appState.datos_financieros
    .map((period) => toInteger(period.anio, NaN))
    .filter(Number.isFinite)
    .sort((a, b) => a - b);
  appState.metadata.periodos_encontrados = [...new Set(years)];
}

function formatNumber(value) {
  const n = Number(value);
  if (!Number.isFinite(n)) {
    return "0";
  }
  return n.toLocaleString("es-MX", {
    minimumFractionDigits: 0,
    maximumFractionDigits: 2,
  });
}

function setMessage(text, type = "info") {
  els.messageBox.textContent = text;
  els.messageBox.classList.remove("hidden", "error");
  if (type === "error") {
    els.messageBox.classList.add("error");
  }
}

function clearMessage() {
  els.messageBox.textContent = "";
  els.messageBox.classList.add("hidden");
  els.messageBox.classList.remove("error");
}

function createField(field, value, onChange) {
  const wrapper = document.createElement("div");
  wrapper.className = field.type === "textarea" ? "field full" : "field";

  const label = document.createElement("label");
  label.textContent = field.label;
  wrapper.appendChild(label);

  let input;
  if (field.type === "select") {
    input = document.createElement("select");
    field.options.forEach((option) => {
      const opt = document.createElement("option");
      opt.value = option;
      opt.textContent = option;
      input.appendChild(opt);
    });
    input.value = String(value ?? field.options[0]);
    input.addEventListener("change", () => onChange(input.value));
  } else if (field.type === "textarea") {
    input = document.createElement("textarea");
    input.rows = 3;
    input.value = String(value ?? "");
    input.addEventListener("input", () => onChange(input.value));
  } else if (field.type === "text") {
    input = document.createElement("input");
    input.type = "text";
    input.value = String(value ?? "");
    input.addEventListener("input", () => onChange(input.value));
  } else if (field.type === "years") {
    input = document.createElement("input");
    input.type = "text";
    input.value = Array.isArray(value) ? value.join(", ") : "";
    input.addEventListener("change", () => onChange(parseYears(input.value)));
  } else {
    input = document.createElement("input");
    input.type = "number";
    input.step = field.type === "integer" ? "1" : "any";
    input.value = value ?? 0;
    input.addEventListener("change", () => {
      if (field.type === "integer") {
        onChange(toInteger(input.value, new Date().getFullYear()));
      } else {
        onChange(toNumber(input.value, 0));
      }
    });
  }

  wrapper.appendChild(input);

  const help = document.createElement("p");
  help.className = "help";
  help.textContent = field.help;
  wrapper.appendChild(help);

  return wrapper;
}

function renderMetadata() {
  const card = document.createElement("div");
  card.className = "subcard";

  const title = document.createElement("h3");
  title.textContent = "Metadata";
  card.appendChild(title);

  const grid = document.createElement("div");
  grid.className = "grid";
  card.appendChild(grid);

  metadataFields.forEach((field) => {
    const currentValue = appState.metadata[field.path];
    const fieldEl = createField(field, currentValue, (nextValue) => {
      appState.metadata[field.path] = nextValue;
    });
    grid.appendChild(fieldEl);
  });

  els.metadataContainer.replaceChildren(card);
}

function renderPeriods() {
  const fragment = document.createDocumentFragment();

  appState.datos_financieros.forEach((period, index) => {
    const details = document.createElement("details");
    details.className = "period-card";
    details.open = index === 0;

    const summary = document.createElement("summary");
    const totalActivos = toNumber(
      readPath(period, "balance_general.activos.total_activos"),
      0
    );
    const totalPasivos = toNumber(
      readPath(period, "balance_general.pasivos.total_pasivos"),
      0
    );
    const totalCapital = toNumber(
      readPath(period, "balance_general.capital_contable.total_capital_contable"),
      0
    );
    const diferencia = totalActivos - (totalPasivos + totalCapital);
    const equilibrio = Math.abs(diferencia) < 0.01 ? "Balanceado" : `Dif: ${formatNumber(diferencia)}`;

    summary.innerHTML = `
      <span>${period.anio} - ${period.tipo_periodo}</span>
      <span>${equilibrio}</span>
    `;
    details.appendChild(summary);

    const body = document.createElement("div");
    body.className = "period-body";

    const actions = document.createElement("div");
    actions.className = "period-actions";

    const removeBtn = document.createElement("button");
    removeBtn.type = "button";
    removeBtn.className = "btn small danger";
    removeBtn.textContent = "Eliminar aÃ±o";
    removeBtn.addEventListener("click", () => {
      appState.datos_financieros.splice(index, 1);
      if (appState.datos_financieros.length === 0) {
        appState.datos_financieros.push(createDefaultPeriod());
      }
      syncMetadataYears();
      renderEditor();
    });
    actions.appendChild(removeBtn);
    body.appendChild(actions);

    periodFieldGroups.forEach((group) => {
      const groupCard = document.createElement("div");
      groupCard.className = "subcard";

      const groupTitle = document.createElement("h3");
      groupTitle.textContent = group.title;
      groupCard.appendChild(groupTitle);

      const groupGrid = document.createElement("div");
      groupGrid.className = "grid";

      group.fields.forEach((field) => {
        const currentValue = readPath(period, field.path);
        const fieldEl = createField(field, currentValue, (nextValue) => {
          writePath(period, field.path, nextValue);
          if (field.path === "anio") {
            syncMetadataYears();
          }
          if (field.path === "alertaDeAI") {
            renderAlertsSummary();
          }
          if (rerenderPeriodPaths.has(field.path)) {
            renderEditor();
          }
        });
        groupGrid.appendChild(fieldEl);
      });

      groupCard.appendChild(groupGrid);
      body.appendChild(groupCard);
    });

    details.appendChild(body);
    fragment.appendChild(details);
  });

  els.periodsContainer.replaceChildren(fragment);
}

function renderAlertsSummary() {
  if (!appState) {
    els.alertsSummaryContainer.replaceChildren();
    return;
  }

  const card = document.createElement("div");
  card.className = "subcard";

  const title = document.createElement("h3");
  title.textContent = "Resumen final de alertas por aÃ±o";
  card.appendChild(title);

  const subtitle = document.createElement("p");
  subtitle.className = "muted";
  subtitle.textContent = "Este bloque consolida las alertas de todos los periodos antes del paso 3.";
  card.appendChild(subtitle);

  const list = document.createElement("div");
  list.className = "alert-summary-list";

  appState.datos_financieros
    .slice()
    .sort((a, b) => a.anio - b.anio)
    .forEach((period) => {
      const item = document.createElement("div");
      item.className = "alert-summary-item";

      const heading = document.createElement("p");
      heading.className = "alert-summary-year";
      heading.textContent = `AÃ±o ${period.anio}`;
      item.appendChild(heading);

      const text = document.createElement("p");
      const alertText = String(period.alertaDeAI ?? "").trim();
      text.className = alertText ? "alert-summary-text" : "alert-summary-text muted";
      text.textContent = alertText || "Sin alerta registrada";
      item.appendChild(text);

      list.appendChild(item);
    });

  card.appendChild(list);
  els.alertsSummaryContainer.replaceChildren(card);
}

function renderEditor() {
  if (!appState) {
    els.editorSection.classList.add("hidden");
    els.previewSection.classList.add("hidden");
    els.actionsSection.classList.add("hidden");
    els.alertsSummaryContainer.replaceChildren();
    return;
  }

  appState.datos_financieros.sort((a, b) => a.anio - b.anio);
  renderMetadata();
  renderPeriods();
  renderAlertsSummary();
  els.editorSection.classList.remove("hidden");
  els.previewSection.classList.remove("hidden");
  els.actionsSection.classList.remove("hidden");
}

// ---------------------------------------------------------------------------
// Preview (renderizado directo desde appState, sin llamada al servidor)
// ---------------------------------------------------------------------------

function loadPreview() {
  if (!appState) return;
  els.previewStatus.classList.add("hidden");
  els.previewTableContainer.innerHTML = "";
  renderPreviewStatic();
}

function renderPreviewStatic() {
  const periods = [...appState.datos_financieros].sort((a, b) => a.anio - b.anio);

  function yearLabel(p) {
    return (p.tipo_periodo || "").toUpperCase() === "PARCIAL"
      ? `${p.anio} (Parcial)`
      : String(p.anio);
  }

  function getVal(p, row) {
    if (row.special === "pasivo_capital") {
      return (
        toNumber(readPath(p, "balance_general.pasivos.total_pasivos"), 0) +
        toNumber(readPath(p, "balance_general.capital_contable.total_capital_contable"), 0)
      );
    }
    if (row.special === "activo_circulante_calc") {
      return (
        toNumber(readPath(p, "balance_general.activos.circulante.efectivo_y_equivalentes"), 0) +
        toNumber(readPath(p, "balance_general.activos.circulante.cuentas_por_cobrar_clientes"), 0) +
        toNumber(readPath(p, "balance_general.activos.circulante.impuestos_a_favor_cp"), 0) +
        toNumber(readPath(p, "balance_general.activos.circulante.deudores_diversos_cp"), 0) +
        toNumber(readPath(p, "balance_general.activos.circulante.inventarios"), 0) +
        toNumber(readPath(p, "balance_general.activos.circulante.pagos_anticipados"), 0) +
        toNumber(readPath(p, "balance_general.activos.circulante.otros_activos_circulantes"), 0)
      );
    }
    if (row.special === "activo_no_circulante_calc") {
      return (
        toNumber(readPath(p, "balance_general.activos.no_circulante.equipo_de_transporte"), 0) +
        toNumber(readPath(p, "balance_general.activos.no_circulante.equipo_de_computo"), 0) +
        toNumber(readPath(p, "balance_general.activos.no_circulante.mobiliario_y_equipo_de_oficina"), 0) +
        toNumber(readPath(p, "balance_general.activos.no_circulante.propiedad_planta_y_equipo_neto"), 0) +
        toNumber(readPath(p, "balance_general.activos.no_circulante.depreciacion_acumulada_historica"), 0) +
        toNumber(readPath(p, "balance_general.activos.no_circulante.activos_intangibles_neto"), 0) +
        toNumber(readPath(p, "balance_general.activos.no_circulante.activos_diferidos"), 0)
      );
    }
    if (row.special === "pasivo_cp_calc") {
      return (
        toNumber(readPath(p, "balance_general.pasivos.corto_plazo.proveedores"), 0) +
        toNumber(readPath(p, "balance_general.pasivos.corto_plazo.deuda_financiera_cp"), 0) +
        toNumber(readPath(p, "balance_general.pasivos.corto_plazo.impuestos_y_cuotas_por_pagar"), 0) +
        toNumber(readPath(p, "balance_general.pasivos.corto_plazo.anticipo_de_clientes"), 0) +
        toNumber(readPath(p, "balance_general.pasivos.corto_plazo.acreedores_diversos"), 0) +
        toNumber(readPath(p, "balance_general.pasivos.corto_plazo.provisiones"), 0) +
        toNumber(readPath(p, "balance_general.pasivos.corto_plazo.otros_pasivos_corto_plazo"), 0)
      );
    }
    if (row.special === "pasivo_lp_calc") {
      return (
        toNumber(readPath(p, "balance_general.pasivos.largo_plazo.deuda_financiera_lp"), 0) +
        toNumber(readPath(p, "balance_general.pasivos.largo_plazo.pasivo_por_arrendamiento"), 0) +
        toNumber(readPath(p, "balance_general.pasivos.largo_plazo.dividendos_decretados"), 0)
      );
    }
    return toNumber(readPath(p, row.path), 0);
  }

  function buildTable(id, title, rowDefs) {
    const block = document.createElement("div");
    block.className = "preview-static-block";
    block.id = id;

    const h = document.createElement("h4");
    h.className = "preview-static-title";
    h.textContent = title;
    block.appendChild(h);

    const scroll = document.createElement("div");
    scroll.className = "preview-static-scroll";

    const table = document.createElement("table");
    table.className = "preview-static-table";

    const thead = table.createTHead();
    const htr = thead.insertRow();
    const thC = document.createElement("th");
    thC.className = "pst-label";
    thC.textContent = "Concepto";
    htr.appendChild(thC);
    periods.forEach((p) => {
      const th = document.createElement("th");
      th.textContent = yearLabel(p);
      htr.appendChild(th);
    });

    const tbody = table.createTBody();
    rowDefs.forEach((row) => {
      const tr = tbody.insertRow();
      if (row.separator) {
        const td = tr.insertCell();
        td.colSpan = periods.length + 1;
        td.className = "pst-separator";
        td.textContent = row.separator;
        return;
      }
      if (row.total) tr.classList.add("pst-total");
      const tdL = tr.insertCell();
      tdL.className = "pst-label";
      tdL.textContent = row.label;
      periods.forEach((p) => {
        const td = tr.insertCell();
        const v = getVal(p, row);
        td.className = v === 0 ? "pst-num pst-zero" : "pst-num";
        td.textContent = v === 0 ? "-" : formatNumber(v);
      });
    });

    scroll.appendChild(table);
    block.appendChild(scroll);
    return block;
  }

  const erRows = [
    { label: "Ingresos operativos netos", path: "estado_resultados.ingresos_operativos_netos", total: true },
    { label: "Costo de ventas", path: "estado_resultados.costo_de_ventas" },
    { label: "Utilidad bruta", path: "estado_resultados.utilidad_bruta", total: true },
    { label: "Gastos de operacion", path: "estado_resultados.gastos_operativos" },
    { label: "Gastos generales", path: "estado_resultados.gastos_generales" },
    { label: "Gastos de administracion", path: "estado_resultados.gastos_de_administracion" },
    { label: "Gastos de venta", path: "estado_resultados.gastos_de_venta" },
    { label: "Gastos de personal", path: "estado_resultados.gastos_de_personal" },
    { label: "Utilidad operativa (EBIT)", path: "estado_resultados.utilidad_operativa_ebit", total: true },
    { label: "Resultado financiero neto", path: "estado_resultados.resultado_financiero_neto" },
    { label: "Utilidad antes de impuestos", path: "estado_resultados.utilidad_antes_de_impuestos" },
    { label: "ISR diferido", path: "estado_resultados.isr_diferido" },
    { label: "ISR corriente", path: "estado_resultados.isr_corriente" },
    { label: "Provision PTU", path: "estado_resultados.provision_ptu" },
    { label: "Total impuestos generico", path: "estado_resultados.total_impuestos_generico" },
    { label: "Utilidad neta", path: "estado_resultados.utilidad_neta", total: true },
  ];

  const bgRows = [
    { separator: "ACTIVO" },
    { label: "Activo Circulante", special: "activo_circulante_calc", total: true },
    { label: "  Efectivo", path: "balance_general.activos.circulante.efectivo_y_equivalentes" },
    { label: "  Cuentas por cobrar", path: "balance_general.activos.circulante.cuentas_por_cobrar_clientes" },
    { label: "  Impuestos a favor CP", path: "balance_general.activos.circulante.impuestos_a_favor_cp" },
    { label: "  Deudores diversos CP", path: "balance_general.activos.circulante.deudores_diversos_cp" },
    { label: "  Inventarios", path: "balance_general.activos.circulante.inventarios" },
    { label: "  Pagos anticipados", path: "balance_general.activos.circulante.pagos_anticipados" },
    { label: "  Otros circulantes", path: "balance_general.activos.circulante.otros_activos_circulantes" },
    { label: "Activo No Circulante", special: "activo_no_circulante_calc", total: true },
    { label: "  Equipo de transporte", path: "balance_general.activos.no_circulante.equipo_de_transporte" },
    { label: "  Equipo de computo", path: "balance_general.activos.no_circulante.equipo_de_computo" },
    { label: "  Mobiliario y equipo", path: "balance_general.activos.no_circulante.mobiliario_y_equipo_de_oficina" },
    { label: "  Depreciacion acumulada", path: "balance_general.activos.no_circulante.depreciacion_acumulada_historica" },
    { label: "  Activos diferidos", path: "balance_general.activos.no_circulante.activos_diferidos" },
    { label: "TOTAL ACTIVOS", path: "balance_general.activos.total_activos", total: true },
    { separator: "PASIVO" },
    { label: "Pasivo Corto Plazo", special: "pasivo_cp_calc", total: true },
    { label: "  Proveedores", path: "balance_general.pasivos.corto_plazo.proveedores" },
    { label: "  Deuda financiera CP", path: "balance_general.pasivos.corto_plazo.deuda_financiera_cp" },
    { label: "  Impuestos y cuotas por pagar", path: "balance_general.pasivos.corto_plazo.impuestos_y_cuotas_por_pagar" },
    { label: "  Anticipo de clientes", path: "balance_general.pasivos.corto_plazo.anticipo_de_clientes" },
    { label: "  Acreedores diversos", path: "balance_general.pasivos.corto_plazo.acreedores_diversos" },
    { label: "  Provisiones", path: "balance_general.pasivos.corto_plazo.provisiones" },
    { label: "  Otros pasivos CP", path: "balance_general.pasivos.corto_plazo.otros_pasivos_corto_plazo" },
    { label: "Pasivo Largo Plazo", special: "pasivo_lp_calc", total: true },
    { label: "  Dividendos decretados", path: "balance_general.pasivos.largo_plazo.dividendos_decretados" },
    { label: "  Pasivo por arrendamiento", path: "balance_general.pasivos.largo_plazo.pasivo_por_arrendamiento" },
    { label: "  Deuda financiera LP", path: "balance_general.pasivos.largo_plazo.deuda_financiera_lp" },
    { label: "TOTAL PASIVOS", path: "balance_general.pasivos.total_pasivos", total: true },
    { separator: "CAPITAL" },
    { label: "Capital social", path: "balance_general.capital_contable.capital_social" },
    { label: "Utilidades ejercicios anteriores", path: "balance_general.capital_contable.utilidades_ejercicios_anteriores" },
    { label: "Resultado del ejercicio", path: "balance_general.capital_contable.resultado_del_ejercicio_balance" },
    { label: "TOTAL CAPITAL", path: "balance_general.capital_contable.total_capital_contable", total: true },
    { label: "TOTAL PASIVO + CAPITAL", special: "pasivo_capital", total: true },
  ];

  els.previewTableContainer.appendChild(buildTable("preview-er", "Estado de Resultados", erRows));
  els.previewTableContainer.appendChild(buildTable("preview-bg", "Balance General", bgRows));
}

function parseFile(file) {
  return file
    .text()
    .then((text) => JSON.parse((text || "").replace(/^\uFEFF/, "")))
    .then((rawData) => normalizeData(rawData));
}

function parseRawJsonText(rawText) {
  return normalizeData(JSON.parse((rawText || "").replace(/^\uFEFF/, "")));
}

function parseFilenameFromHeader(contentDisposition) {
  if (!contentDisposition) {
    return null;
  }

  const utf = contentDisposition.match(/filename\*=UTF-8''([^;]+)/i);
  if (utf?.[1]) {
    return decodeURIComponent(utf[1]);
  }

  const basic = contentDisposition.match(/filename="?([^"]+)"?/i);
  if (basic?.[1]) {
    return basic[1];
  }

  return null;
}

async function loadTemplateInfo() {
  try {
    const response = await fetch("/api/template-info");
    const data = await response.json();
    els.templateInfo.textContent = `Plantilla detectada: ${data.template_name}`;
  } catch (error) {
    els.templateInfo.textContent = "Plantilla detectada: no disponible";
  }
}

async function loadSampleData() {
  clearMessage();
  try {
    const response = await fetch("/api/sample-data");
    const payload = await response.json();
    if (!response.ok) {
      throw new Error(payload.error || "No se pudo cargar el ejemplo local.");
    }
    appState = normalizeData(payload);
    syncMetadataYears();
    renderEditor();
    loadPreview();
    els.jsonTextInput.value = JSON.stringify(appState, null, 2);
    els.selectedFile.textContent = "Archivo cargado: ejemplo local";
    setMessage("Ejemplo cargado correctamente. Ya puedes revisar campos y generar Excel.");
  } catch (error) {
    setMessage(error.message, "error");
  }
}

async function handleFileSelection(event) {
  clearMessage();
  const file = event.target.files?.[0];
  if (!file) {
    return;
  }

  try {
    appState = await parseFile(file);
    syncMetadataYears();
    renderEditor();
    loadPreview();
    els.jsonTextInput.value = JSON.stringify(appState, null, 2);
    els.selectedFile.textContent = `Archivo cargado: ${file.name}`;
    setMessage("JSON cargado correctamente. Revisa y corrige antes de generar.");
  } catch (error) {
    setMessage(`No se pudo leer el archivo: ${error.message}`, "error");
  }
}

function handleTextLoad() {
  clearMessage();
  const rawText = (els.jsonTextInput.value || "").trim();
  if (!rawText) {
    setMessage("Pega un JSON en el cuadro de texto antes de cargar.", "error");
    return;
  }

  try {
    appState = parseRawJsonText(rawText);
    syncMetadataYears();
    renderEditor();
    loadPreview();
    els.jsonTextInput.value = JSON.stringify(appState, null, 2);
    els.selectedFile.textContent = "Archivo cargado: JSON pegado manualmente";
    setMessage("JSON pegado cargado correctamente. Revisa y corrige antes de generar.");
  } catch (error) {
    setMessage(`JSON invalido: ${error.message}`, "error");
  }
}

function addNewPeriod() {
  if (!appState) {
    return;
  }

  const currentYears = appState.datos_financieros.map((period) =>
    toInteger(period.anio, NaN)
  );
  const baseYear = currentYears.length
    ? Math.max(...currentYears.filter(Number.isFinite))
    : new Date().getFullYear();

  appState.datos_financieros.push(createDefaultPeriod(baseYear + 1));
  syncMetadataYears();
  renderEditor();
}

async function generateExcel() {
  if (!appState) {
    setMessage("Primero carga un JSON para poder generar el Excel.", "error");
    return;
  }

  clearMessage();
  els.generateBtn.disabled = true;
  els.generateBtn.textContent = "Generando...";
  els.downloadLink.classList.add("hidden");

  try {
    const response = await fetch("/api/generate-excel", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify(appState),
    });

    if (!response.ok) {
      const errorPayload = await response.json();
      throw new Error(
        errorPayload.details || errorPayload.error || "Error desconocido."
      );
    }

    const blob = await response.blob();
    const contentDisposition = response.headers.get("content-disposition");
    const filename =
      parseFilenameFromHeader(contentDisposition) || "Valuacion_Completada.xlsx";

    const url = URL.createObjectURL(blob);
    els.downloadLink.href = url;
    els.downloadLink.download = filename;
    els.downloadLink.classList.remove("hidden");
    els.downloadLink.click();

    setMessage("Excel generado con exito. Si no bajo automaticamente, usa el boton Descargar Excel.");
  } catch (error) {
    setMessage(`No se pudo generar el Excel: ${error.message}`, "error");
  } finally {
    els.generateBtn.disabled = false;
    els.generateBtn.textContent = "Confirmar datos y generar Excel";
  }
}

function bindEvents() {
  els.fileInput.addEventListener("change", handleFileSelection);
  els.loadTextBtn.addEventListener("click", handleTextLoad);
  els.jsonTextInput.addEventListener("keydown", (event) => {
    if ((event.ctrlKey || event.metaKey) && event.key === "Enter") {
      event.preventDefault();
      handleTextLoad();
    }
  });
  els.addPeriodBtn.addEventListener("click", addNewPeriod);
  els.generateBtn.addEventListener("click", generateExcel);
  els.refreshPreviewBtn.addEventListener("click", loadPreview);
  els.previewNav.addEventListener("click", (e) => {
    const btn = e.target.closest(".preview-jump");
    if (!btn) return;
    const target = document.getElementById(btn.dataset.target);
    if (target) target.scrollIntoView({ behavior: "smooth", block: "start" });
  });
}

bindEvents();
loadTemplateInfo();

