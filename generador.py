from fpdf import FPDF

class StressTestPDF(FPDF):
    def header(self):
        self.set_font('Arial', 'B', 8)
        self.cell(0, 5, 'REPORTE DE AUDITORÍA FORENSE - SISTEMAS LOGÍSTICOS GLOBALES S.A.', 0, 1, 'R')
        self.ln(5)

    def chapter_title(self, year, context):
        self.set_font('Arial', 'B', 14)
        # Color rojo oscuro para indicar que es un test de estrés
        self.set_fill_color(100, 0, 0)
        self.set_text_color(255, 255, 255)
        self.cell(0, 10, f" EJERCICIO FISCAL {year} | ESCENARIO: {context}", 0, 1, 'L', 1)
        self.set_text_color(0, 0, 0)
        self.ln(4)

    def draw_table(self, data):
        self.set_font('Arial', '', 9)
        w_label = 120
        w_val = 60
        for label, val in data:
            if label.startswith("---"):
                self.set_font('Arial', 'B', 9)
                self.set_fill_color(230, 230, 230)
                self.cell(w_label + w_val, 8, label.replace("-", " "), 1, 1, 'L', 1)
                self.set_font('Arial', '', 9)
            else:
                self.cell(w_label, 7, f"  {label}", 1)
                # Alineamos a la derecha los valores
                self.cell(w_val, 7, str(val), 1, 1, 'R')
        self.ln(5)

def fmt(val):
    """Formatea números a string con comas y 2 decimales"""
    if isinstance(val, (int, float)):
        return f"{val:,.2f}"
    return val

pdf = StressTestPDF()

# ==========================================
# AÑO 2024: EL DESCUADRE BRUTAL
# Objetivo: Que la IA detecte que Activo (600M) != Pasivo+Capital (413M)
# ==========================================
data_2024 = [
    ("--- ESTADO DE RESULTADO INTEGRAL (ACUMULADO) ---", ""),
    ("Ventas Totales Netas", fmt(850432891.45)),
    ("Costo de Ventas y Servicios", fmt(420567123.10)),
    ("Utilidad Bruta", fmt(429865768.35)),
    ("Gastos de Administración", fmt(85432109.00)),
    ("Gastos de Venta", fmt(45231987.55)),
    ("Gastos Generales", fmt(12000000.00)),
    ("Depreciación del Periodo (Gasto)", fmt(12500000.00)),
    ("Utilidad de Operación (EBIT)", fmt(274701671.80)),
    ("Otros Ingresos No Operativos", fmt(5000000.00)),
    ("Otros Gastos No Operativos", fmt(2000000.00)),
    ("Resultado Integral de Financiamiento (RIF)", fmt(-15432987.00)),
    ("Utilidad Antes de Impuestos", fmt(262268684.80)),
    ("Impuestos a la Utilidad (ISR)", fmt(72380605.44)),
    ("UTILIDAD NETA DEL EJERCICIO", fmt(189888079.36)),
    
    ("--- ESTADO DE SITUACIÓN FINANCIERA (BALANCE) ---", ""),
    ("Efectivo y Equivalentes", fmt(45231987.11)),
    ("Inversiones Temporales", fmt(10000000.00)),
    ("Clientes (Cuentas por Cobrar)", fmt(125432987.44)),
    ("Inventarios", fmt(85432109.22)),
    ("IVA y Otros Impuestos a Favor", fmt(12345678.99)),
    ("Deudores Diversos", fmt(5000000.00)),
    ("Pagos Anticipados", fmt(2000000.00)),
    ("Otros Activos Circulantes", fmt(5432109.33)),
    ("Propiedad, Planta y Equipo (Bruto)", fmt(450678941.00)),
    ("Depreciación Acumulada", fmt(-141453913.09)), # Ajustado para cuadrar el Activo Fijo
    ("Activos Diferidos", fmt(8432109.11)),
    ("TOTAL ACTIVOS", fmt(600000000.00)), # <--- DATO FIJO: 600 MILLONES
    
    ("Proveedores y Cuentas por Pagar", fmt(150432987.00)),
    ("Impuestos por Pagar", fmt(25432987.44)),
    ("Acreedores Diversos CP", fmt(12543210.33)),
    ("Provisiones Operativas", fmt(5000000.00)),
    ("Otros Pasivos a Corto Plazo", fmt(8432109.11)),
    ("Deuda Bancaria Largo Plazo", fmt(150432987.22)),
    ("Capital Social", fmt(50000000.00)),
    ("Utilidades de Ejercicios Anteriores", fmt(15432987.55)),
    ("Resultado del Presente Ejercicio", fmt(189888079.36)),
    ("SUMA PASIVO Y CAPITAL", fmt(413002348.01)) # <--- ERROR: FALTAN CASI 187 MILLONES
]

# ==========================================
# AÑO 2023: QUIEBRA TÉCNICA Y ERROR ARITMÉTICO
# Objetivo: Ventas(120) - Costos(150) = -30. Pero EBIT dice -10.
# Objetivo 2: Pasivos > Activos = Capital Negativo
# ==========================================
data_2023 = [
    ("--- ESTADO DE RESULTADOS ---", ""),
    ("Ventas Totales Netas", fmt(120432987.00)),
    ("Costo de Ventas y Servicios", fmt(150678941.00)),
    ("Utilidad Bruta", fmt(-30245954.00)),
    ("Gastos de Administración", fmt(10000000.00)),
    ("Gastos de Venta", fmt(5000000.00)),
    ("Gastos Generales", fmt(2000000.00)),
    ("Depreciación del Periodo (Gasto)", fmt(3000000.00)),
    ("Utilidad de Operación (EBIT)", fmt(-10000000.00)), # <--- ERROR MATEMÁTICO EN EL PDF
    ("Otros Ingresos No Operativos", fmt(0.00)),
    ("Otros Gastos No Operativos", fmt(0.00)),
    ("Resultado Integral de Financiamiento (RIF)", fmt(-5000000.00)),
    ("Utilidad Antes de Impuestos", fmt(-15000000.00)),
    ("Impuestos a la Utilidad (ISR)", fmt(0.00)),
    ("UTILIDAD NETA DEL EJERCICIO", fmt(-15000000.00)),

    ("--- BALANCE GENERAL ---", ""),
    ("Efectivo y Equivalentes", fmt(5432109.00)),
    ("Inversiones Temporales", fmt(0.00)),
    ("Clientes (Cuentas por Cobrar)", fmt(1000000.00)),
    ("Inventarios", fmt(2000000.00)),
    ("IVA y Otros Impuestos a Favor", fmt(500000.00)),
    ("Deudores Diversos", fmt(0.00)),
    ("Pagos Anticipados", fmt(0.00)),
    ("Otros Activos Circulantes", fmt(0.00)),
    ("Propiedad, Planta y Equipo (Bruto)", fmt(50000000.00)),
    ("Depreciación Acumulada", fmt(-28066913.00)),
    ("Activos Diferidos", fmt(0.00)),
    ("TOTAL ACTIVOS", fmt(30865196.00)),

    ("Proveedores y Cuentas por Pagar", fmt(60000000.00)),
    ("Impuestos por Pagar", fmt(15432109.00)),
    ("Acreedores Diversos CP", fmt(5000000.00)),
    ("Provisiones Operativas", fmt(2000000.00)),
    ("Otros Pasivos a Corto Plazo", fmt(3000000.00)),
    ("Deuda Bancaria Largo Plazo", fmt(45231987.00)),
    ("Capital Social", fmt(10000000.00)),
    ("Utilidades de Ejercicios Anteriores", fmt(-94798900.00)),
    ("Resultado del Presente Ejercicio", fmt(-15000000.00)),
    ("SUMA PASIVO Y CAPITAL", fmt(30865196.00)) # Cuadra, pero con capital negativo masivo
]

# ==========================================
# AÑO 2022: DATOS OCULTOS (INFERENCIA)
# Objetivo: Costo de Ventas vacío. Depreciación en 0.00.
# ==========================================
data_2022 = [
    ("--- ESTADO DE RESULTADOS ---", ""),
    ("Ventas Totales Netas", fmt(95432109.00)),
    ("Costo de Ventas y Servicios", ""), # <--- VACÍO: IA debe inferir
    ("Utilidad Bruta", fmt(45000000.00)),
    ("Gastos de Administración", fmt(15000000.00)),
    ("Gastos de Venta", fmt(5000000.00)),
    ("Gastos Generales", fmt(5000000.00)),
    ("Depreciación del Periodo (Gasto)", "0.00"), # <--- TRAMPA: DATO EN NOTA AL PIE
    ("Utilidad de Operación (EBIT)", fmt(20000000.00)),
    ("Otros Ingresos No Operativos", fmt(0.00)),
    ("Otros Gastos No Operativos", fmt(0.00)),
    ("Resultado Integral de Financiamiento (RIF)", fmt(0.00)),
    ("Utilidad Antes de Impuestos", fmt(20000000.00)),
    ("Impuestos a la Utilidad (ISR)", fmt(0.00)),
    ("UTILIDAD NETA DEL EJERCICIO", fmt(20000000.00)),

    ("--- BALANCE GENERAL ---", ""),
    ("Efectivo y Equivalentes", fmt(40000000.00)),
    ("Inversiones Temporales", fmt(0.00)),
    ("Clientes (Cuentas por Cobrar)", fmt(20000000.00)),
    ("Inventarios", fmt(20000000.00)),
    ("IVA y Otros Impuestos a Favor", fmt(0.00)),
    ("Deudores Diversos", fmt(0.00)),
    ("Pagos Anticipados", fmt(0.00)),
    ("Otros Activos Circulantes", fmt(0.00)),
    ("Propiedad, Planta y Equipo (Bruto)", fmt(80000000.00)),
    ("Depreciación Acumulada", fmt(-40000000.00)),
    ("Activos Diferidos", fmt(0.00)),
    ("TOTAL ACTIVOS", fmt(120000000.00)),

    ("Proveedores y Cuentas por Pagar", fmt(30000000.00)),
    ("Impuestos por Pagar", fmt(10000000.00)),
    ("Acreedores Diversos CP", fmt(0.00)),
    ("Provisiones Operativas", fmt(0.00)),
    ("Otros Pasivos a Corto Plazo", fmt(0.00)),
    ("Deuda Bancaria Largo Plazo", fmt(30000000.00)),
    ("Capital Social", fmt(30000000.00)),
    ("Utilidades de Ejercicios Anteriores", fmt(0.00)),
    ("Resultado del Presente Ejercicio", fmt(20000000.00)),
    ("SUMA PASIVO Y CAPITAL", fmt(120000000.00))
]

# ==========================================
# AÑO 2021: SIGNOS MIXTOS
# Objetivo: Gastos con signo negativo.
# ==========================================
data_2021 = [
    ("--- ESTADO DE RESULTADOS ---", ""),
    ("Ventas Totales Netas", fmt(80432109.00)),
    ("Costo de Ventas y Servicios", fmt(-40000000.00)), # <--- SIGNO NEGATIVO
    ("Utilidad Bruta", fmt(40432109.00)),
    ("Gastos de Administración", fmt(-10000000.00)),
    ("Gastos de Venta", fmt(-5000000.00)),
    ("Gastos Generales", fmt(-5000000.00)),
    ("Depreciación del Periodo (Gasto)", fmt(-432109.00)),
    ("Utilidad de Operación (EBIT)", fmt(20000000.00)),
    ("Otros Ingresos No Operativos", fmt(0.00)),
    ("Otros Gastos No Operativos", fmt(0.00)),
    ("Resultado Integral de Financiamiento (RIF)", fmt(0.00)),
    ("Utilidad Antes de Impuestos", fmt(20000000.00)),
    ("Impuestos a la Utilidad (ISR)", fmt(0.00)),
    ("UTILIDAD NETA DEL EJERCICIO", fmt(20000000.00)),

    ("--- BALANCE GENERAL ---", ""),
    ("Efectivo y Equivalentes", fmt(30000000.00)),
    ("Inversiones Temporales", fmt(0.00)),
    ("Clientes (Cuentas por Cobrar)", fmt(20000000.00)),
    ("Inventarios", fmt(10000000.00)),
    ("IVA y Otros Impuestos a Favor", fmt(0.00)),
    ("Deudores Diversos", fmt(0.00)),
    ("Pagos Anticipados", fmt(0.00)),
    ("Otros Activos Circulantes", fmt(0.00)),
    ("Propiedad, Planta y Equipo (Bruto)", fmt(50000000.00)),
    ("Depreciación Acumulada", fmt(-10000000.00)),
    ("Activos Diferidos", fmt(0.00)),
    ("TOTAL ACTIVOS", fmt(100000000.00)),

    ("Proveedores y Cuentas por Pagar", fmt(20000000.00)),
    ("Impuestos por Pagar", fmt(5000000.00)),
    ("Acreedores Diversos CP", fmt(5000000.00)),
    ("Provisiones Operativas", fmt(0.00)),
    ("Otros Pasivos a Corto Plazo", fmt(10000000.00)),
    ("Deuda Bancaria Largo Plazo", fmt(20000000.00)),
    ("Capital Social", fmt(20000000.00)),
    ("Utilidades de Ejercicios Anteriores", fmt(0.00)),
    ("Resultado del Presente Ejercicio", fmt(20000000.00)),
    ("SUMA PASIVO Y CAPITAL", fmt(100000000.00))
]

# ==========================================
# AÑO 2020: PERIODO PARCIAL
# Objetivo: Enero a Mayo
# ==========================================
data_2020 = [
    ("--- ESTADO DE RESULTADOS (ENERO - MAYO) ---", ""),
    ("Ventas Totales Netas", fmt(25432987.00)),
    ("Costo de Ventas y Servicios", fmt(10000000.00)),
    ("Utilidad Bruta", fmt(15432987.00)),
    ("Gastos de Administración", fmt(2000000.00)),
    ("Gastos de Venta", fmt(1000000.00)),
    ("Gastos Generales", fmt(432987.00)),
    ("Depreciación del Periodo (Gasto)", fmt(2000000.00)),
    ("Utilidad de Operación (EBIT)", fmt(10000000.00)),
    ("Otros Ingresos No Operativos", fmt(0.00)),
    ("Otros Gastos No Operativos", fmt(0.00)),
    ("Resultado Integral de Financiamiento (RIF)", fmt(0.00)),
    ("Utilidad Antes de Impuestos", fmt(10000000.00)),
    ("Impuestos a la Utilidad (ISR)", fmt(0.00)),
    ("UTILIDAD NETA DEL EJERCICIO", fmt(10000000.00)),

    ("--- BALANCE GENERAL AL 31 DE MAYO ---", ""),
    ("Efectivo y Equivalentes", fmt(10000000.00)),
    ("Inversiones Temporales", fmt(0.00)),
    ("Clientes (Cuentas por Cobrar)", fmt(5000000.00)),
    ("Inventarios", fmt(5000000.00)),
    ("IVA y Otros Impuestos a Favor", fmt(0.00)),
    ("Deudores Diversos", fmt(0.00)),
    ("Pagos Anticipados", fmt(0.00)),
    ("Otros Activos Circulantes", fmt(0.00)),
    ("Propiedad, Planta y Equipo (Bruto)", fmt(40000000.00)),
    ("Depreciación Acumulada", fmt(-10000000.00)),
    ("Activos Diferidos", fmt(0.00)),
    ("TOTAL ACTIVOS", fmt(50000000.00)),

    ("Proveedores y Cuentas por Pagar", fmt(5000000.00)),
    ("Impuestos por Pagar", fmt(0.00)),
    ("Acreedores Diversos CP", fmt(0.00)),
    ("Provisiones Operativas", fmt(0.00)),
    ("Otros Pasivos a Corto Plazo", fmt(5000000.00)),
    ("Deuda Bancaria Largo Plazo", fmt(0.00)),
    ("Capital Social", fmt(30000000.00)),
    ("Utilidades de Ejercicios Anteriores", fmt(0.00)),
    ("Resultado del Presente Ejercicio", fmt(10000000.00)),
    ("SUMA PASIVO Y CAPITAL", fmt(50000000.00))
]

# Construcción de páginas
for year, data, context in [
    (2024, data_2024, "DESCUADRE MASIVO (187M DIFERENCIA)"),
    (2023, data_2023, "QUIEBRA TÉCNICA Y EBIT FALSO"),
    (2022, data_2022, "DATOS OCULTOS EN NOTAS"),
    (2021, data_2021, "FORMATO DE SIGNOS MIXTOS"),
    (2020, data_2020, "PERIODO PARCIAL (5 MESES)")
]:
    pdf.add_page()
    pdf.chapter_title(year, context)
    pdf.draw_table(data)
    
    if year == 2022:
        pdf.set_font('Arial', 'I', 8)
        pdf.multi_cell(0, 4, "Nota Importante: El gasto por depreciación de la maquinaria ascendió a $12,432,987.00 y se cargó directamente a resultados, aunque se omitió en el cuerpo de la tabla superior.")

pdf.output("MEGA_TEST_5_ANIOS_SUCIO.pdf")
print("Archivo generado: MEGA_TEST_5_ANIOS_SUCIO.pdf")