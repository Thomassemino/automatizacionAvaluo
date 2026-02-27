import json
import re
import openpyxl

JSON_FILE     = "datos_extraidos_ia.json"
TEMPLATE_FILE = "Grupo Ovando.xlsx"
SHEET_NAME    = "1. Datos"

# Filas donde se escriben los encabezados de año
HEADER_ROWS = [5, 42, 102]

# ---------------------------------------------------------------------------
# Mapeo fila Excel -> ruta JSON (dot-separated)
# Incluye items de detalle + subtotales
# ---------------------------------------------------------------------------
ROW_MAP = {
    # --- Estado de Resultados ---
    8:   "estado_resultados.ventas_netas",
    9:   "estado_resultados.costo_ventas",
    10:  "estado_resultados.utilidad_bruta",
    12:  "estado_resultados.gastos_operativos_totales",
    19:  "estado_resultados.utilidad_operativa_ebit",
    # filas 20 y 21 se manejan aparte (split de otros_ingresos_gastos_neto)
    24:  "estado_resultados.resultado_integral_financiamiento",
    29:  "estado_resultados.impuestos",
    30:  "estado_resultados.utilidad_neta",
    98:  "estado_resultados.depreciacion_amortizacion_periodo",

    # --- Balance: Activo Circulante (items) ---
    45:  "balance_general.activos.circulante.efectivo_y_equivalentes",
    46:  "balance_general.activos.circulante.cuentas_por_cobrar_clientes",
    47:  "balance_general.activos.circulante.impuestos_a_favor",
    48:  "balance_general.activos.circulante.otros_activos_circulantes",
    49:  "balance_general.activos.circulante.deudores_diversos",
    50:  "balance_general.activos.circulante.pagos_anticipados",
    # subtotal
    44:  "balance_general.activos.circulante.total_activo_circulante",

    # --- Balance: Activo No Circulante (items) ---
    67:  "balance_general.activos.no_circulante.propiedad_planta_equipo_bruto",
    68:  "balance_general.activos.no_circulante.depreciacion_acumulada",
    83:  "balance_general.activos.no_circulante.activos_diferidos",
    97:  "balance_general.activos.no_circulante.depreciacion_acumulada",  # repetido
    # subtotal
    64:  "balance_general.activos.no_circulante.total_activo_no_circulante",

    # --- Total Activos ---
    95:  "balance_general.activos.total_activos",

    # --- Balance: Pasivo Corto Plazo (items) ---
    105: "balance_general.pasivos.corto_plazo.proveedores_cuentas_por_pagar",
    106: "balance_general.pasivos.corto_plazo.impuestos_por_pagar",
    107: "balance_general.pasivos.corto_plazo.otros_pasivos_corto_plazo",
    108: "balance_general.pasivos.corto_plazo.acreedores_diversos",
    109: "balance_general.pasivos.corto_plazo.provisiones",
    # subtotal
    104: "balance_general.pasivos.corto_plazo.total_pasivo_corto_plazo",

    # --- Balance: Pasivo Largo Plazo ---
    119: "balance_general.pasivos.largo_plazo.total_pasivo_largo_plazo",

    # --- Total Pasivos ---
    123: "balance_general.pasivos.total_pasivos",

    # --- Capital Contable (items) ---
    126: "balance_general.capital_contable.capital_social",
    127: "balance_general.capital_contable.utilidades_acumuladas",
    128: "balance_general.capital_contable.resultado_ejercicio_balance",
    # subtotal
    135: "balance_general.capital_contable.total_capital_contable",
}

# Todas las filas de datos (para limpiar columnas sobrantes)
ALL_DATA_ROWS = list(ROW_MAP.keys()) + [20, 21, 137]


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def get_nested(d, path):
    """Devuelve el valor numerico siguiendo path dot-separated. Siempre retorna numero."""
    val = d
    for key in path.split("."):
        if isinstance(val, dict) and key in val:
            val = val[key]
        else:
            return 0
    if val is None or val == "":
        return 0
    try:
        return float(val)
    except (TypeError, ValueError):
        return 0


def find_data_columns(ws):
    """Lee fila 5 y devuelve (columna_base, n_slots).
    Cuenta desde la primera celda numerica/formula hasta encontrar
    una celda de texto no-formula (ej. '2025 proyecctado') o None final."""
    base_col = None
    n_cols   = 0

    for cell in ws[5]:
        v = cell.value
        # Texto no-formula = fin del bloque (ej. "2025 proyecctado")
        if isinstance(v, str) and not v.startswith("="):
            if base_col is not None:
                break
            continue
        # Numero o formula: forma parte del bloque de datos
        if isinstance(v, (int, float)) or (isinstance(v, str) and v.startswith("=")):
            if base_col is None:
                base_col = cell.column
            n_cols += 1

    return base_col, n_cols


def etiqueta_anio(periodo):
    """Retorna el label de cabecera: int para años cerrados, 'YYYY (Parcial)' si es parcial."""
    anio = periodo["anio"]
    if "PARCIAL" in periodo.get("tipo_periodo", "").upper():
        return f"{anio} (Parcial)"
    return anio


def nombre_a_archivo(nombre):
    nombre = re.sub(r'[\\/:*?"<>|]', '', nombre).strip()
    return nombre + ".xlsx"


# ---------------------------------------------------------------------------
# Funcion principal (usada por app.py y por CLI)
# ---------------------------------------------------------------------------

def inyectar_datos_financieros(json_path, template_path, output_path):
    # --- Cargar JSON ---
    with open(json_path, encoding="utf-8") as f:
        data = json.load(f)

    periodos = data["datos_financieros"]
    empresa  = data.get("metadata", {}).get("empresa_detectada", "Sin Nombre")

    print(f"Empresa:  {empresa}")
    print(f"Periodos: {[p['anio'] for p in periodos]}")

    # --- Cargar template ---
    wb = openpyxl.load_workbook(template_path)
    ws = wb[SHEET_NAME]

    # --- Detectar rango de columnas disponibles ---
    base_col, n_slots = find_data_columns(ws)
    if base_col is None:
        print("ERROR: No se encontraron columnas de datos en fila 5.")
        return

    print(f"Slots disponibles: columna {base_col} a {base_col + n_slots - 1} ({n_slots} slots)")

    n_periodos  = len(periodos)
    n_sobrantes = n_slots - n_periodos

    if n_sobrantes < 0:
        print(f"AVISO: JSON tiene {n_periodos} periodos pero plantilla solo tiene {n_slots} slots. "
              f"Se truncan los ultimos {-n_sobrantes} periodos.")
        periodos    = periodos[:n_slots]
        n_sobrantes = 0

    # --- Inyectar datos (right-aligned: ultimo año -> ultima col del template) ---
    total_celdas = 0
    for i, periodo in enumerate(periodos):
        col  = base_col + n_slots - n_periodos + i
        anio = periodo["anio"]
        label = etiqueta_anio(periodo)
        print(f"  {label} -> columna {col}")

        # Encabezados de año en las 3 filas de cabecera
        for hr in HEADER_ROWS:
            ws.cell(row=hr, column=col).value = label

        # Filas estandar + subtotales
        for row, path in ROW_MAP.items():
            ws.cell(row=row, column=col).value = get_nested(periodo, path)
            total_celdas += 1

        # Split otros_ingresos_gastos_neto -> fila 20 (gastos) / fila 21 (ingresos)
        otros = get_nested(periodo, "estado_resultados.otros_ingresos_gastos_neto")
        if otros < 0:
            ws.cell(row=20, column=col).value = abs(otros)
            ws.cell(row=21, column=col).value = 0
        else:
            ws.cell(row=20, column=col).value = 0
            ws.cell(row=21, column=col).value = otros
        total_celdas += 2

        # Fila 137: Total Pasivo + Capital (suma calculada, no existe como campo directo)
        total_pasivos = get_nested(periodo, "balance_general.pasivos.total_pasivos")
        total_capital = get_nested(periodo, "balance_general.capital_contable.total_capital_contable")
        ws.cell(row=137, column=col).value = total_pasivos + total_capital
        total_celdas += 1

    # --- Limpiar columnas sobrantes con 0 (al inicio, NO eliminar para no romper #REF!) ---
    if n_sobrantes > 0:
        print(f"\n  Limpiando {n_sobrantes} columna(s) sobrante(s) con 0...")
        for j in range(n_sobrantes):
            col = base_col + j
            for hr in HEADER_ROWS:
                ws.cell(row=hr, column=col).value = None
            for row in ALL_DATA_ROWS:
                ws.cell(row=row, column=col).value = 0

    # --- Guardar ---
    wb.save(output_path)
    print(f"\nListo. {total_celdas} celdas escritas en '{output_path}'.")


def main():
    with open(JSON_FILE, encoding="utf-8") as f:
        empresa = json.load(f).get("metadata", {}).get("empresa_detectada", "Sin Nombre")
    inyectar_datos_financieros(JSON_FILE, TEMPLATE_FILE, nombre_a_archivo(empresa))


if __name__ == "__main__":
    main()
