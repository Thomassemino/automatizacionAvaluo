import json
import re

import openpyxl
from openpyxl.utils import column_index_from_string, get_column_letter

JSON_FILE = "datos_extraidos_ia.json"
TEMPLATE_FILE = "Grupo Ovando.xlsx"
SHEET_NAME = "1. Datos"

# Filas de encabezado de anos (se actualizan solo si la celda no contiene formula).
HEADER_ROWS = [5, 42, 102]

# Regla de mapeo fijo de columnas.
YEAR_TO_COLUMN = {
    2019: "D",
    2020: "E",
    2021: "F",
    2022: "G",
    2023: "H",
    2024: "I",
}
COL_2025_ANUAL = "J"
# Solicitud de negocio: permitir sobreescritura de formulas en celdas de inyeccion.
FORCE_OVERWRITE_INJECTED_FORMULAS = True

# Filas que reciben dato duro (se usan tambien para limpiar columnas no mapeadas).
INJECTION_ROWS = [
    6, 8, 9, 13, 14, 15, 16, 17, 20, 21, 24, 25, 26, 27,
    45, 46, 47, 48, 49, 50,
    65, 66, 67, 68, 84,
    105, 106, 107, 108, 109,
    110, 111, 112, 113, 114, 115, 116, 117, 118,
    120, 121, 122,
    126, 127, 128,
]


def nombre_a_archivo(nombre):
    nombre = re.sub(r'[\\/:*?"<>|]', "", str(nombre or "")).strip()
    return (nombre or "Sin_Nombre") + ".xlsx"


def to_float(value):
    """Convierte a float de forma tolerante. Si falla, devuelve 0.0."""
    if value is None:
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        cleaned = value.strip().replace(",", "").replace("$", "")
        if not cleaned:
            return 0.0
        if cleaned.startswith("(") and cleaned.endswith(")"):
            cleaned = "-" + cleaned[1:-1]
        try:
            return float(cleaned)
        except ValueError:
            return 0.0
    return 0.0


def has_formula(value):
    return isinstance(value, str) and value.startswith("=")


def write_cell(ws, row, col, value, allow_formula_overwrite=False):
    """
    Escribe en celda respetando formulas (salvo excepciones explicitas).
    Retorna True si escribio, False si se omitio por proteccion de formula.
    """
    cell = ws.cell(row=row, column=col)
    if (
        has_formula(cell.value)
        and not allow_formula_overwrite
        and not FORCE_OVERWRITE_INJECTED_FORMULAS
    ):
        return False
    cell.value = value
    return True


def _set_formula_if_empty(ws, row, col, formula):
    """Escribe formula solo cuando la celda esta vacia y no tiene formula."""
    cell = ws.cell(row=row, column=col)
    if has_formula(cell.value):
        return
    if cell.value is None or cell.value == "":
        cell.value = formula


def ensure_required_formulas(ws):
    """
    Completa formulas faltantes de filas calculadas en D:K.
    Solo rellena celdas vacias para no romper formulas/valores existentes.
    """
    col_start = column_index_from_string("D")
    col_end = column_index_from_string("K")

    for col in range(col_start, col_end + 1):
        col_l = get_column_letter(col)
        prev_l = get_column_letter(col - 1) if col > col_start else None

        formulas = {
            10: f"={col_l}8-{col_l}9",
            12: f"=SUM({col_l}13:{col_l}17)",
            19: f"={col_l}10-{col_l}12",
            29: f"=SUM({col_l}25:{col_l}28)",
            30: f"={col_l}19+{col_l}24-{col_l}29",
            44: f"=SUM({col_l}45:{col_l}63)",
            64: f"=SUM({col_l}65:{col_l}88)",
            83: f"=SUM({col_l}84:{col_l}92)",
            95: f"=SUM({col_l}44,{col_l}64,{col_l}83)",
            97: f"={col_l}68",
            104: f"=SUM({col_l}105:{col_l}118)",
            119: f"=SUM({col_l}120:{col_l}122)",
            123: f"={col_l}104+{col_l}119",
            125: f"=SUM({col_l}126:{col_l}134)",
            135: f"={col_l}125",
            137: f"={col_l}123+{col_l}135",
        }

        if prev_l:
            formulas[98] = f"={col_l}97-{prev_l}97"

        for row, formula in formulas.items():
            _set_formula_if_empty(ws, row, col, formula)


def clear_unmapped_columns(ws, mapped_cols):
    """
    Limpia columnas historicas no incluidas en el JSON actual.
    Se limpia solo D:J; la K se preserva para formulas automaticas del modelo.
    """
    d_col = column_index_from_string("D")
    j_col = column_index_from_string("J")

    for col in range(d_col, j_col + 1):
        if col in mapped_cols:
            continue

        for row in HEADER_ROWS:
            write_cell(ws, row, col, None, allow_formula_overwrite=True)

        for row in INJECTION_ROWS:
            write_cell(ws, row, col, 0.0, allow_formula_overwrite=True)


def resolve_target_column(periodo):
    """Resuelve la columna destino segun anio y tipo_periodo."""
    anio = int(to_float(periodo.get("anio", 0)))

    if anio in YEAR_TO_COLUMN:
        return column_index_from_string(YEAR_TO_COLUMN[anio]), anio

    if anio == 2025:
        # Por instruccion de negocio, 2025 siempre se inyecta en J.
        return column_index_from_string(COL_2025_ANUAL), anio

    return None, None


def inject_headers(ws, col, label):
    for row in HEADER_ROWS:
        write_cell(ws, row, col, label, allow_formula_overwrite=False)


def inject_estado_resultados(ws, col, periodo):
    er = periodo.get("estado_resultados") or {}

    ingresos = to_float(er.get("ingresos_operativos_netos"))
    costo_ventas = abs(to_float(er.get("costo_de_ventas")))
    gastos_operativos = abs(to_float(er.get("gastos_operativos")))
    gastos_generales = abs(to_float(er.get("gastos_generales")))
    gastos_admin = abs(to_float(er.get("gastos_de_administracion")))
    gastos_venta = abs(to_float(er.get("gastos_de_venta")))
    gastos_personal = abs(to_float(er.get("gastos_de_personal")))
    rif = to_float(er.get("resultado_financiero_neto"))
    isr_diferido = abs(to_float(er.get("isr_diferido")))
    isr_corriente = abs(to_float(er.get("isr_corriente")))
    provision_ptu = abs(to_float(er.get("provision_ptu")))
    total_impuestos_generico = abs(to_float(er.get("total_impuestos_generico")))

    # Fila 6 y 8: ingresos_operativos_netos.
    write_cell(ws, 6, col, ingresos)
    write_cell(ws, 8, col, ingresos)

    # Fila 9: costo de ventas en valor absoluto.
    write_cell(ws, 9, col, costo_ventas)

    # Filas 13 a 15: gastos operativos en absoluto.
    write_cell(ws, 13, col, gastos_operativos)
    write_cell(ws, 14, col, gastos_generales)
    write_cell(ws, 15, col, gastos_admin)

    # Filas 16 y 17: opcionales (si vienen gastos_de_venta / gastos_de_personal).
    ws.cell(row=16, column=3).value = "Gastos de venta"
    ws.cell(row=17, column=3).value = "Gastos de personal"
    write_cell(ws, 16, col, gastos_venta)
    write_cell(ws, 17, col, gastos_personal)

    # Filas 20 y 21: en cero por regla de negocio.
    write_cell(ws, 20, col, 0.0)
    write_cell(ws, 21, col, 0.0)

    # Fila 24: RIF inyectado como dato duro (sobrescribe formula por instruccion).
    write_cell(
        ws,
        24,
        col,
        rif,
        allow_formula_overwrite=True,
    )

    # Filas 25, 26 y 27: impuestos en absoluto.
    write_cell(ws, 25, col, isr_diferido)
    write_cell(ws, 26, col, isr_corriente)
    write_cell(ws, 27, col, provision_ptu)

    # Fila 29: NO TOCAR salvo caso atipico de total agrupado.
    if (
        abs(isr_diferido) < 1e-9
        and abs(isr_corriente) < 1e-9
        and abs(provision_ptu) < 1e-9
        and abs(total_impuestos_generico) >= 1e-9
    ):
        write_cell(
            ws,
            29,
            col,
            total_impuestos_generico,
            allow_formula_overwrite=True,
        )


def inject_balance_general(ws, col, periodo):
    bg = periodo.get("balance_general") or {}
    activos = bg.get("activos") or {}
    pasivos = bg.get("pasivos") or {}
    capital = bg.get("capital_contable") or {}

    activos_circulante = activos.get("circulante") or {}
    activos_no_circulante = activos.get("no_circulante") or {}
    pasivo_cp = pasivos.get("corto_plazo") or {}
    pasivo_lp = pasivos.get("largo_plazo") or {}

    # Activo circulante (NO TOCAR fila 44).
    write_cell(
        ws,
        45,
        col,
        to_float(activos_circulante.get("efectivo_y_equivalentes")),
    )
    write_cell(
        ws,
        46,
        col,
        to_float(activos_circulante.get("cuentas_por_cobrar_clientes")),
    )
    write_cell(
        ws,
        47,
        col,
        to_float(activos_circulante.get("impuestos_a_favor_cp")),
    )
    write_cell(
        ws,
        48,
        col,
        to_float(activos_circulante.get("otros_activos_circulantes")),
    )
    write_cell(
        ws,
        49,
        col,
        to_float(activos_circulante.get("deudores_diversos_cp")),
    )
    write_cell(
        ws,
        50,
        col,
        to_float(activos_circulante.get("pagos_anticipados")),
    )

    # Activo no circulante (NO TOCAR fila 64).
    write_cell(
        ws,
        65,
        col,
        to_float(activos_no_circulante.get("equipo_de_transporte")),
    )
    write_cell(
        ws,
        66,
        col,
        to_float(activos_no_circulante.get("equipo_de_computo")),
    )
    write_cell(
        ws,
        67,
        col,
        to_float(activos_no_circulante.get("mobiliario_y_equipo_de_oficina")),
    )
    # Depreciacion acumulada historica SIEMPRE se inyecta en negativo.
    write_cell(
        ws,
        68,
        col,
        -abs(to_float(activos_no_circulante.get("depreciacion_acumulada_historica"))),
    )

    # Activos diferidos (NO TOCAR fila 83).
    ws.cell(row=84, column=3).value = "Activos Diferidos"
    write_cell(
        ws,
        84,
        col,
        to_float(activos_no_circulante.get("activos_diferidos")),
    )

    # Pasivo circulante (NO TOCAR fila 104).
    pasivo_cp_map = [
        (105, "proveedores", "Proveedores"),
        (106, "deuda_financiera_cp", "Deuda financiera CP"),
        (
            107,
            "impuestos_y_cuotas_por_pagar",
            "Impuestos y cuotas por pagar",
        ),
        (108, "acreedores_diversos", "Acreedores diversos"),
        (109, "provisiones", "Provisiones"),
    ]
    for row, key, label in pasivo_cp_map:
        ws.cell(row=row, column=3).value = label
        write_cell(ws, row, col, to_float(pasivo_cp.get(key)))

    # Limpia filas no utilizadas del bloque 110-118 para no arrastrar basura.
    for row in range(110, 119):
        write_cell(ws, row, col, 0.0)

    # Pasivo largo plazo (NO TOCAR filas 119 y 123).
    write_cell(ws, 120, col, to_float(pasivo_lp.get("dividendos_decretados")))
    write_cell(ws, 121, col, to_float(pasivo_lp.get("pasivo_por_arrendamiento")))
    write_cell(ws, 122, col, to_float(pasivo_lp.get("deuda_financiera_lp")))

    # Capital (NO TOCAR filas 125, 135 y 137).
    write_cell(ws, 126, col, to_float(capital.get("capital_social")))
    write_cell(
        ws,
        127,
        col,
        to_float(capital.get("utilidades_ejercicios_anteriores")),
    )
    write_cell(
        ws,
        128,
        col,
        to_float(capital.get("resultado_del_ejercicio_balance")),
    )


def inyectar_datos_financieros(json_path, template_path, output_path):
    with open(json_path, encoding="utf-8-sig") as f:
        data = json.load(f)

    periodos = data.get("datos_financieros") or []
    empresa = data.get("metadata", {}).get("empresa_detectada", "Sin Nombre")

    if not periodos:
        raise ValueError(
            "El JSON no contiene periodos en 'datos_financieros'."
        )

    print(f"Empresa:  {empresa}")
    print(f"Periodos recibidos: {[p.get('anio') for p in periodos]}")

    wb = openpyxl.load_workbook(template_path)
    ws = wb[SHEET_NAME]
    ensure_required_formulas(ws)

    mapped_cols = set()
    for periodo in periodos:
        col, _ = resolve_target_column(periodo)
        if col is not None:
            mapped_cols.add(col)
    clear_unmapped_columns(ws, mapped_cols)

    used_cols = {}
    skipped = 0

    for periodo in periodos:
        col, label = resolve_target_column(periodo)
        if col is None:
            print(
                f"AVISO: se omite periodo anio={periodo.get('anio')} "
                f"tipo_periodo={periodo.get('tipo_periodo')} (sin columna definida)."
            )
            skipped += 1
            continue

        prev = used_cols.get(col)
        if prev is not None:
            print(
                f"AVISO: columna {col} ya usada por anio={prev}. "
                f"Se sobrescribe con anio={periodo.get('anio')}."
            )
        used_cols[col] = periodo.get("anio")

        inject_headers(ws, col, label)
        inject_estado_resultados(ws, col, periodo)
        inject_balance_general(ws, col, periodo)

    wb.save(output_path)
    print(
        f"Listo. Archivo generado en '{output_path}'. "
        f"Columnas cargadas: {sorted(used_cols.keys())}. "
        f"Periodos omitidos: {skipped}."
    )


def main():
    with open(JSON_FILE, encoding="utf-8") as f:
        empresa = (
            json.load(f).get("metadata", {}).get("empresa_detectada", "Sin Nombre")
        )
    inyectar_datos_financieros(
        JSON_FILE,
        TEMPLATE_FILE,
        nombre_a_archivo(empresa),
    )


if __name__ == "__main__":
    main()
