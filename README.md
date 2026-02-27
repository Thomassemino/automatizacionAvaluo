# Frontend de revision y generacion de Excel

## Requisitos
- Python 3.10+
- Microsoft Excel instalado (por `xlwings`)

## Instalacion
```bash
pip install -r requirements.txt
```

## Ejecutar
```bash
python app.py
```

Luego abre en navegador:

`http://localhost:8000`

## Flujo
1. Carga tu JSON extraido por IA.
2. Revisa y corrige campos.
3. Da clic en **Confirmar datos y generar Excel**.
4. Descarga el archivo generado.

## Plantilla
- Por defecto toma el archivo `.xlsx` cuyo nombre contenga `plantilla`.
- Si quieres fijar una plantilla exacta, define:

```bash
set AVALUOS_TEMPLATE_PATH=C:\ruta\tu_plantilla.xlsx
```
