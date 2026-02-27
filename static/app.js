const els = {
  fileInput: document.getElementById("fileInput"),
  loadSampleBtn: document.getElementById("loadSampleBtn"),
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
    help: "Lista de anos fiscales unicos detectados en el lote procesado. Escribe anos separados por coma.",
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
        help: "Ano fiscal especifico extraido de la columna de fecha.",
      },
      {
        path: "tipo_periodo",
        label: "Tipo de periodo",
        type: "select",
        options: ["ANUAL_CERRADO", "PARCIAL"],
        help: "ANUAL_CERRADO si cubre enero-diciembre. PARCIAL si corta antes (por ejemplo, octubre).",
      },
    ],
  },
  {
    title: "B. Estado de resultados (flujo acumulado)",
    fields: [
      {
        path: "estado_resultados.ventas_netas",
        label: "Ventas netas",
        type: "number",
        help: "Suma acumulada anual de Ingresos, Ventas o Servicios.",
      },
      {
        path: "estado_resultados.costo_ventas",
        label: "Costo de ventas",
        type: "number",
        help: "Costos directos del servicio/producto. Puede ser 0 si se carga a gastos operativos.",
      },
      {
        path: "estado_resultados.utilidad_bruta",
        label: "Utilidad bruta",
        type: "number",
        help: "Calculo esperado: Ventas Netas - Costo de Ventas.",
      },
      {
        path: "estado_resultados.gastos_operativos_totales",
        label: "Gastos operativos totales",
        type: "number",
        help: "Suma de gastos de administracion, venta y generales; se guardan como magnitud positiva.",
      },
      {
        path: "estado_resultados.utilidad_operativa_ebit",
        label: "Utilidad operativa (EBIT)",
        type: "number",
        help: "Dato clave de valuacion. Calculo esperado: Utilidad Bruta - Gastos Operativos.",
      },
      {
        path: "estado_resultados.depreciacion_amortizacion_periodo",
        label: "Depreciacion y amortizacion del periodo",
        type: "number",
        help: "Partida de depreciacion del ejercicio. Si no aparece explicita, se deja en 0.0.",
      },
      {
        path: "estado_resultados.otros_ingresos_gastos_neto",
        label: "Otros ingresos/gastos neto",
        type: "number",
        help: "Partidas no operativas. Negativo = gasto extraordinario; positivo = ingreso ajeno al giro.",
      },
      {
        path: "estado_resultados.resultado_integral_financiamiento",
        label: "Resultado integral de financiamiento (RIF)",
        type: "number",
        help: "Intereses pagados/ganados y efecto cambiario.",
      },
      {
        path: "estado_resultados.impuestos",
        label: "Impuestos",
        type: "number",
        help: "ISR y PTU del ejercicio.",
      },
      {
        path: "estado_resultados.utilidad_neta",
        label: "Utilidad neta",
        type: "number",
        help: "Resultado final del ejercicio (bottom line).",
      },
    ],
  },
  {
    title: "C. Balance general (foto al cierre)",
    fields: [
      {
        path: "balance_general.activos.circulante.efectivo_y_equivalentes",
        label: "Efectivo y equivalentes",
        type: "number",
        help: "Suma de Caja + Bancos + Inversiones temporales al cierre.",
      },
      {
        path: "balance_general.activos.circulante.cuentas_por_cobrar_clientes",
        label: "Cuentas por cobrar clientes",
        type: "number",
        help: "Clientes comerciales al cierre.",
      },
      {
        path: "balance_general.activos.circulante.impuestos_a_favor",
        label: "Impuestos a favor",
        type: "number",
        help: "IVA acreditable, ISR a favor y similares.",
      },
      {
        path: "balance_general.activos.circulante.deudores_diversos",
        label: "Deudores diversos",
        type: "number",
        help: "Cuentas por cobrar no comerciales.",
      },
      {
        path: "balance_general.activos.circulante.pagos_anticipados",
        label: "Pagos anticipados",
        type: "number",
        help: "Seguros, rentas u otros pagos prepagados.",
      },
      {
        path: "balance_general.activos.circulante.otros_activos_circulantes",
        label: "Otros activos circulantes",
        type: "number",
        help: "Activos circulantes que no encajan en rubros principales.",
      },
      {
        path: "balance_general.activos.circulante.total_activo_circulante",
        label: "Total activo circulante",
        type: "number",
        help: "Suma de los activos circulantes.",
      },
      {
        path: "balance_general.activos.no_circulante.propiedad_planta_equipo_bruto",
        label: "Propiedad, planta y equipo bruto",
        type: "number",
        help: "Valor original de activos fijos (mobiliario, equipo, etc.).",
      },
      {
        path: "balance_general.activos.no_circulante.depreciacion_acumulada",
        label: "Depreciacion acumulada",
        type: "number",
        help: "Valor negativo del desgaste historico acumulado, tomado del balance.",
      },
      {
        path: "balance_general.activos.no_circulante.activos_diferidos",
        label: "Activos diferidos",
        type: "number",
        help: "Partidas diferidas de largo plazo.",
      },
      {
        path: "balance_general.activos.no_circulante.total_activo_no_circulante",
        label: "Total activo no circulante",
        type: "number",
        help: "Suma de activos no circulantes.",
      },
      {
        path: "balance_general.activos.total_activos",
        label: "Total activos",
        type: "number",
        help: "Total de activos. Debe cuadrar con Pasivo + Capital.",
      },
      {
        path: "balance_general.pasivos.corto_plazo.proveedores_cuentas_por_pagar",
        label: "Proveedores / cuentas por pagar",
        type: "number",
        help: "Deuda operativa con proveedores.",
      },
      {
        path: "balance_general.pasivos.corto_plazo.impuestos_por_pagar",
        label: "Impuestos por pagar",
        type: "number",
        help: "IVA trasladado, ISR por pagar y retenciones.",
      },
      {
        path: "balance_general.pasivos.corto_plazo.acreedores_diversos",
        label: "Acreedores diversos",
        type: "number",
        help: "Otros acreedores de corto plazo.",
      },
      {
        path: "balance_general.pasivos.corto_plazo.provisiones",
        label: "Provisiones",
        type: "number",
        help: "Reservas para obligaciones futuras.",
      },
      {
        path: "balance_general.pasivos.corto_plazo.otros_pasivos_corto_plazo",
        label: "Otros pasivos corto plazo",
        type: "number",
        help: "Pasivos de corto plazo que no encajan en rubros previos.",
      },
      {
        path: "balance_general.pasivos.corto_plazo.total_pasivo_corto_plazo",
        label: "Total pasivo corto plazo",
        type: "number",
        help: "Suma del pasivo exigible a corto plazo.",
      },
      {
        path: "balance_general.pasivos.largo_plazo.total_pasivo_largo_plazo",
        label: "Total pasivo largo plazo",
        type: "number",
        help: "Suma del pasivo de largo plazo.",
      },
      {
        path: "balance_general.pasivos.total_pasivos",
        label: "Total pasivos",
        type: "number",
        help: "Total de deudas (corto + largo plazo).",
      },
      {
        path: "balance_general.capital_contable.capital_social",
        label: "Capital social",
        type: "number",
        help: "Aportaciones de socios.",
      },
      {
        path: "balance_general.capital_contable.utilidades_acumuladas",
        label: "Utilidades acumuladas",
        type: "number",
        help: "Resultados acumulados de ejercicios anteriores.",
      },
      {
        path: "balance_general.capital_contable.resultado_ejercicio_balance",
        label: "Resultado del ejercicio (balance)",
        type: "number",
        help: "Debe coincidir con utilidad_neta del estado de resultados.",
      },
      {
        path: "balance_general.capital_contable.total_capital_contable",
        label: "Total capital contable",
        type: "number",
        help: "Calculo esperado: Total Activos - Total Pasivos.",
      },
      {
        path: "alertaDeAI",
        label: "Alerta de IA",
        type: "textarea",
        help: "Autodiagnostico de la IA sobre ecuacion contable, periodos incompletos o capital negativo.",
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
      ventas_netas: 0,
      costo_ventas: 0,
      utilidad_bruta: 0,
      gastos_operativos_totales: 0,
      utilidad_operativa_ebit: 0,
      depreciacion_amortizacion_periodo: 0,
      otros_ingresos_gastos_neto: 0,
      resultado_integral_financiamiento: 0,
      impuestos: 0,
      utilidad_neta: 0,
    },
    balance_general: {
      activos: {
        circulante: {
          efectivo_y_equivalentes: 0,
          cuentas_por_cobrar_clientes: 0,
          impuestos_a_favor: 0,
          deudores_diversos: 0,
          pagos_anticipados: 0,
          otros_activos_circulantes: 0,
          total_activo_circulante: 0,
        },
        no_circulante: {
          propiedad_planta_equipo_bruto: 0,
          depreciacion_acumulada: 0,
          activos_diferidos: 0,
          total_activo_no_circulante: 0,
        },
        total_activos: 0,
      },
      pasivos: {
        corto_plazo: {
          proveedores_cuentas_por_pagar: 0,
          impuestos_por_pagar: 0,
          acreedores_diversos: 0,
          provisiones: 0,
          otros_pasivos_corto_plazo: 0,
          total_pasivo_corto_plazo: 0,
        },
        largo_plazo: {
          total_pasivo_largo_plazo: 0,
        },
        total_pasivos: 0,
      },
      capital_contable: {
        capital_social: 0,
        utilidades_acumuladas: 0,
        resultado_ejercicio_balance: 0,
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
    removeBtn.textContent = "Eliminar ano";
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
  title.textContent = "Resumen final de alertas por año";
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
      heading.textContent = `Ano ${period.anio}`;
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
    if (row.special === "otros_gastos") {
      const v = toNumber(readPath(p, "estado_resultados.otros_ingresos_gastos_neto"), 0);
      return v < 0 ? Math.abs(v) : 0;
    }
    if (row.special === "otros_ingresos") {
      const v = toNumber(readPath(p, "estado_resultados.otros_ingresos_gastos_neto"), 0);
      return v >= 0 ? v : 0;
    }
    if (row.special === "pasivo_capital") {
      return (
        toNumber(readPath(p, "balance_general.pasivos.total_pasivos"), 0) +
        toNumber(readPath(p, "balance_general.capital_contable.total_capital_contable"), 0)
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
    { label: "Ventas Netas",              path: "estado_resultados.ventas_netas",                        total: true },
    { label: "Costo de Ventas",           path: "estado_resultados.costo_ventas" },
    { label: "UTILIDAD BRUTA",            path: "estado_resultados.utilidad_bruta",                      total: true },
    { label: "Costos y Gastos",           path: "estado_resultados.gastos_operativos_totales",           total: true },
    { label: "UTILIDAD OPERACION (EBIT)", path: "estado_resultados.utilidad_operativa_ebit",             total: true },
    { label: "Otros Gastos",              special: "otros_gastos" },
    { label: "Otros Ingresos",            special: "otros_ingresos" },
    { label: "RIF",                       path: "estado_resultados.resultado_integral_financiamiento" },
    { label: "Impuestos",                 path: "estado_resultados.impuestos" },
    { label: "UTILIDAD NETA",             path: "estado_resultados.utilidad_neta",                       total: true },
    { label: "Depreciacion periodo",      path: "estado_resultados.depreciacion_amortizacion_periodo" },
  ];

  const bgRows = [
    { separator: "ACTIVO" },
    { label: "Activo Circulante",        path: "balance_general.activos.circulante.total_activo_circulante",        total: true },
    { label: "  Efectivo",               path: "balance_general.activos.circulante.efectivo_y_equivalentes" },
    { label: "  Cuentas por cobrar",     path: "balance_general.activos.circulante.cuentas_por_cobrar_clientes" },
    { label: "  Impuestos a favor",      path: "balance_general.activos.circulante.impuestos_a_favor" },
    { label: "  Deudores diversos",      path: "balance_general.activos.circulante.deudores_diversos" },
    { label: "  Pagos anticipados",      path: "balance_general.activos.circulante.pagos_anticipados" },
    { label: "  Otros circulantes",      path: "balance_general.activos.circulante.otros_activos_circulantes" },
    { label: "Activo No Circulante",     path: "balance_general.activos.no_circulante.total_activo_no_circulante",  total: true },
    { label: "  PPE Bruto",              path: "balance_general.activos.no_circulante.propiedad_planta_equipo_bruto" },
    { label: "  Depreciacion acumulada", path: "balance_general.activos.no_circulante.depreciacion_acumulada" },
    { label: "  Activos diferidos",      path: "balance_general.activos.no_circulante.activos_diferidos" },
    { label: "TOTAL ACTIVOS",            path: "balance_general.activos.total_activos",                             total: true },
    { separator: "PASIVO" },
    { label: "Pasivo Corto Plazo",       path: "balance_general.pasivos.corto_plazo.total_pasivo_corto_plazo",      total: true },
    { label: "  Proveedores",            path: "balance_general.pasivos.corto_plazo.proveedores_cuentas_por_pagar" },
    { label: "  Impuestos por pagar",    path: "balance_general.pasivos.corto_plazo.impuestos_por_pagar" },
    { label: "  Acreedores diversos",    path: "balance_general.pasivos.corto_plazo.acreedores_diversos" },
    { label: "  Provisiones",            path: "balance_general.pasivos.corto_plazo.provisiones" },
    { label: "  Otros pasivos CP",       path: "balance_general.pasivos.corto_plazo.otros_pasivos_corto_plazo" },
    { label: "Pasivo Largo Plazo",       path: "balance_general.pasivos.largo_plazo.total_pasivo_largo_plazo",      total: true },
    { label: "TOTAL PASIVOS",            path: "balance_general.pasivos.total_pasivos",                             total: true },
    { separator: "CAPITAL" },
    { label: "Capital social",           path: "balance_general.capital_contable.capital_social" },
    { label: "Utilidades acumuladas",    path: "balance_general.capital_contable.utilidades_acumuladas" },
    { label: "Resultado ejercicio",      path: "balance_general.capital_contable.resultado_ejercicio_balance" },
    { label: "TOTAL CAPITAL",            path: "balance_general.capital_contable.total_capital_contable",           total: true },
    { label: "TOTAL PASIVO + CAPITAL",   special: "pasivo_capital",                                                 total: true },
  ];

  els.previewTableContainer.appendChild(buildTable("preview-er", "Estado de Resultados", erRows));
  els.previewTableContainer.appendChild(buildTable("preview-bg", "Balance General", bgRows));
}

function parseFile(file) {
  return file
    .text()
    .then((text) => JSON.parse(text))
    .then((rawData) => normalizeData(rawData));
}

function parseRawJsonText(rawText) {
  return normalizeData(JSON.parse(rawText));
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
  els.loadSampleBtn.addEventListener("click", loadSampleData);
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
