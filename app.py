import io
import json
import os
import shutil
import tempfile
from datetime import datetime
from pathlib import Path

from flask import Flask, jsonify, render_template, request, send_file

from script import inyectar_datos_financieros

BASE_DIR = Path(__file__).resolve().parent
SAMPLE_JSON_PATH = BASE_DIR / "datos_extraidos_ia.json"
EXCEL_MIME = (
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

app = Flask(__name__)
app.config["JSON_AS_ASCII"] = False
app.config["MAX_CONTENT_LENGTH"] = 6 * 1024 * 1024


def _normalize_name(name: str) -> str:
    return "".join(name.split()).lower()


def _resolve_template_path() -> Path:
    configured = os.getenv("AVALUOS_TEMPLATE_PATH", "").strip()
    if configured:
        candidate = Path(configured)
        if candidate.exists():
            return candidate.resolve()
        raise FileNotFoundError(
            f"La plantilla configurada no existe: {configured}"
        )

    excel_files = list(BASE_DIR.glob("*.xlsx")) + list(BASE_DIR.glob("*.XLSX"))
    if not excel_files:
        raise FileNotFoundError(
            "No se encontro ninguna plantilla .xlsx en el proyecto."
        )

    for candidate in excel_files:
        if "grupoovando" in _normalize_name(candidate.stem):
            return candidate.resolve()

    return excel_files[0].resolve()


def _validate_payload(payload):
    if not isinstance(payload, dict):
        raise ValueError("El contenido recibido no es un JSON valido.")

    if "metadata" not in payload or not isinstance(payload["metadata"], dict):
        raise ValueError("Falta la seccion 'metadata' en el JSON.")

    periodos = payload.get("datos_financieros")
    if not isinstance(periodos, list) or not periodos:
        raise ValueError(
            "La seccion 'datos_financieros' debe ser una lista con al menos un periodo."
        )

    return payload


@app.get("/")
def index():
    return render_template("index.html", sample_exists=SAMPLE_JSON_PATH.exists())


@app.get("/api/template-info")
def template_info():
    try:
        template_path = _resolve_template_path()
        return jsonify({"template_name": template_path.name})
    except Exception as exc:
        return jsonify({"template_name": f"No disponible ({exc})"}), 200


@app.get("/api/sample-data")
def sample_data():
    if not SAMPLE_JSON_PATH.exists():
        return jsonify({"error": "No existe el archivo de ejemplo local."}), 404

    with SAMPLE_JSON_PATH.open("r", encoding="utf-8") as f:
        data = json.load(f)

    return jsonify(data)


@app.post("/api/generate-excel")
def generate_excel():
    payload = request.get_json(silent=True)
    if isinstance(payload, dict) and isinstance(payload.get("data"), dict):
        payload = payload["data"]

    try:
        data = _validate_payload(payload)
    except ValueError as exc:
        return jsonify({"error": str(exc)}), 400

    workdir = None
    try:
        template_path = _resolve_template_path()
        workdir = Path(tempfile.mkdtemp(prefix="avaluos_"))
        json_path = workdir / "datos_corregidos.json"
        filename = f"Valuacion_Completada_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
        output_path = workdir / filename

        with json_path.open("w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

        inyectar_datos_financieros(
            str(json_path),
            str(template_path),
            str(output_path),
        )

        excel_bytes = output_path.read_bytes()
        stream = io.BytesIO(excel_bytes)
        stream.seek(0)

        return send_file(
            stream,
            mimetype=EXCEL_MIME,
            as_attachment=True,
            download_name=filename,
        )
    except Exception as exc:
        return (
            jsonify(
                {
                    "error": "No se pudo generar el Excel.",
                    "details": str(exc),
                }
            ),
            500,
        )
    finally:
        if workdir and workdir.exists():
            shutil.rmtree(workdir, ignore_errors=True)


@app.post("/api/preview-sheet")
def preview_sheet():
    payload = request.get_json(silent=True)
    if isinstance(payload, dict) and isinstance(payload.get("data"), dict):
        payload = payload["data"]

    try:
        data = _validate_payload(payload)
    except ValueError as exc:
        return jsonify({"error": str(exc)}), 400

    workdir = None
    try:
        import openpyxl as _xl

        template_path = _resolve_template_path()
        workdir = Path(tempfile.mkdtemp(prefix="avaluos_prev_"))
        json_path = workdir / "preview.json"
        output_path = workdir / "preview.xlsx"

        with json_path.open("w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

        inyectar_datos_financieros(str(json_path), str(template_path), str(output_path))

        wb = _xl.load_workbook(str(output_path), read_only=True, data_only=True)
        ws = wb["1. Datos"]

        rows = []
        for r in range(1, 141):
            row = []
            for c in range(1, 12):
                v = ws.cell(row=r, column=c).value
                if v is None:
                    row.append(None)
                elif isinstance(v, (int, float)):
                    row.append(v)
                else:
                    row.append(str(v))
            rows.append(row)

        wb.close()
        return jsonify({"rows": rows})
    except Exception as exc:
        return jsonify({"error": "No se pudo generar la vista previa.", "details": str(exc)}), 500
    finally:
        if workdir and workdir.exists():
            shutil.rmtree(workdir, ignore_errors=True)


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8000, debug=True)
