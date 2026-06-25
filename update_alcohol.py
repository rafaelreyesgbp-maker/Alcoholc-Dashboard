#!/usr/bin/env python3
"""
update_alcohol.py
Descarga archivos XLS de Google Drive, calcula proyecciones de cierre y
actualiza el dashboard index.html del Impuesto a la Venta Final de Bebidas
con Contenido Alcohólico (AAFY 2026).

Uso:
  1. Exporta tu API Key de Drive como variable de entorno DRIVE_API_KEY, o
     edita la constante API_KEY directamente.
  2. Ejecuta: python update_alcohol.py
"""

import json
import os
import re
import sys
import unicodedata
from datetime import datetime

import requests
import xlrd

# ── Configuración ─────────────────────────────────────────────────────────────
FOLDER_ID = "1sqAqz7fBqOPhHkNHUXIQ-apDNj0FvOfz"
API_KEY   = os.environ.get("DRIVE_API_KEY", "YOUR_GOOGLE_DRIVE_API_KEY")
HTML_FILE = "index.html"
YEAR      = 2026

MONTHLY_GOALS = [
    5_971_005,  # Enero
    2_033_044,  # Febrero
    2_357_586,  # Marzo
    2_380_846,  # Abril
    2_204_240,  # Mayo
    2_495_661,  # Junio
    2_249_735,  # Julio
    2_324_560,  # Agosto
    2_360_836,  # Septiembre
    3_233_310,  # Octubre
    3_540_960,  # Noviembre
    4_050_214,  # Diciembre
]

MESES_ES = [
    "enero", "febrero", "marzo", "abril", "mayo", "junio",
    "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre",
]

# Umbrales de segmentación (sobre promedio mensual de referencia)
ALTA_MIN  = 50_000   # >= Alta
MEDIA_MIN = 10_000   # >= Media, < Alta  |  < Media -> Seguimiento

# ── Columnas del archivo (base 0 = columna A) ─────────────────────────────────
COL_RFC    = 0   # A
COL_NOMBRE = 1   # B
COL_PERIOD = 4   # E  (formato YYYYMM, guardado como float)
COL_SUMA   = 13  # N  (recaudación = N - I)
COL_RESTA  = 8   # I


# ── Utilidades ────────────────────────────────────────────────────────────────

def _norm(text: str) -> str:
    """Minúsculas sin acentos."""
    return "".join(
        c for c in unicodedata.normalize("NFD", text.lower())
        if unicodedata.category(c) != "Mn"
    )


def _cell_str(cell) -> str:
    v = cell.value
    if cell.ctype == 2 and v == int(v):  # número entero
        return str(int(v))
    return str(v).strip()


def _cell_float(cell) -> float:
    try:
        return float(cell.value) if cell.ctype in (2, 3) else 0.0
    except Exception:
        return 0.0


# ── Parser del archivo de alcohol ─────────────────────────────────────────────

def _find_data_start(ws) -> int:
    """Detecta la primera fila con un RFC válido (hasta 20 filas de encabezado)."""
    for r in range(min(20, ws.nrows)):
        try:
            val = _norm(_cell_str(ws.cell(r, COL_RFC)))
            if re.match(r"^[a-z]{3,4}\d{6}", val):
                return r
        except Exception:
            pass
    return 0


def _parse_alcohol(ws) -> list[dict]:
    """
    Parsea una hoja de alcohol.
    Recaudación = col N (13) − col I (8)
    Solo incluye filas con RFC válido y recaudación != 0.
    """
    start = _find_data_start(ws)
    rows = []
    for r in range(start, ws.nrows):
        try:
            rfc    = _cell_str(ws.cell(r, COL_RFC)).upper().strip()
            nombre = _cell_str(ws.cell(r, COL_NOMBRE)).strip()

            # Período YYYYMM (guardado como float → int → str)
            period_raw = ws.cell(r, COL_PERIOD).value
            period = str(int(float(period_raw))) if period_raw else ""

            amount = _cell_float(ws.cell(r, COL_SUMA)) - _cell_float(ws.cell(r, COL_RESTA))

            if rfc and re.match(r"^[A-Z]{3,4}\d{6}", rfc) and amount != 0:
                rows.append({
                    "rfc":    rfc,
                    "nombre": nombre,
                    "period": period,
                    "amount": round(amount, 2),
                })
        except Exception:
            continue
    return rows


# ── Google Drive ───────────────────────────────────────────────────────────────

def list_drive_files() -> list[dict]:
    url = (
        "https://www.googleapis.com/drive/v3/files"
        f"?q=%27{FOLDER_ID}%27+in+parents+and+trashed%3Dfalse"
        f"&fields=files(id,name)&pageSize=200&key={API_KEY}"
    )
    r = requests.get(url, timeout=30)
    r.raise_for_status()
    return r.json().get("files", [])


def download_file(file_id: str) -> bytes:
    url = f"https://www.googleapis.com/drive/v3/files/{file_id}?alt=media&key={API_KEY}"
    r = requests.get(url, timeout=60)
    r.raise_for_status()
    return r.content


# ── Lógica de omisos ──────────────────────────────────────────────────────────

def _segment(ref_avg: float) -> str:
    a = abs(ref_avg)
    if a >= ALTA_MIN:
        return "Alta"
    if a >= MEDIA_MIN:
        return "Media"
    return "Seguimiento"


def compute_month_data(raw: dict) -> dict:
    """
    raw: {YYYYMM: [{"rfc", "nombre", "amount"}, ...]}

    Para cada mes calcula:
    - real: suma de recaudación de contribuyentes presentes
    - omisos: contribuyentes con >= 2 meses de referencia previos que NO aparecen
    - segmentos Alta / Media / Seguimiento
    - proyección = real + estimado_Alta + estimado_Media
    """
    all_months = sorted(raw.keys())
    result = {}

    for i, mes in enumerate(all_months):
        current_rfcs = {row["rfc"] for row in raw[mes]}
        ref_months   = all_months[:i]  # meses anteriores disponibles

        # Acumular historial
        ref_count  = {}
        ref_total  = {}
        ref_nombre = {}
        for rm in ref_months:
            for row in raw[rm]:
                rfc = row["rfc"]
                ref_count[rfc]  = ref_count.get(rfc, 0) + 1
                ref_total[rfc]  = ref_total.get(rfc, 0.0) + row["amount"]
                ref_nombre[rfc] = row["nombre"]

        # Detectar omisos (>= 2 meses de referencia, ausentes en mes actual)
        omisos = []
        for rfc, cnt in ref_count.items():
            if cnt >= 2 and rfc not in current_rfcs:
                avg = ref_total[rfc] / cnt
                omisos.append({
                    "rfc":        rfc,
                    "nombre":     ref_nombre[rfc],
                    "ref_avg":    round(avg, 2),
                    "ref_months": cnt,
                    "segment":    _segment(avg),
                })

        alta        = [o for o in omisos if o["segment"] == "Alta"]
        media       = [o for o in omisos if o["segment"] == "Media"]
        seguimiento = [o for o in omisos if o["segment"] == "Seguimiento"]

        alta_est  = sum(abs(o["ref_avg"]) for o in alta)
        media_est = sum(abs(o["ref_avg"]) for o in media)

        real       = sum(row["amount"] for row in raw[mes])
        proyeccion = real + alta_est + media_est

        result[mes] = {
            "real":              round(real, 2),
            "omisos":            omisos,
            "alta_count":        len(alta),
            "media_count":       len(media),
            "seguimiento_count": len(seguimiento),
            "total_omisos":      len(omisos),
            "alta_est":          round(alta_est, 2),
            "media_est":         round(media_est, 2),
            "proyeccion":        round(proyeccion, 2),
            "contributors":      raw[mes],
        }
    return result


# ── Actualización del HTML ────────────────────────────────────────────────────

def update_html(computed: dict, html_path: str) -> None:
    with open(html_path, "r", encoding="utf-8") as f:
        html = f.read()

    # Cargar datos existentes para aplicar protección (solo actualiza si real > existente)
    existing: dict = {}
    m = re.search(r"let allData\s*=\s*(\{.*?\});", html, re.DOTALL)
    if m:
        try:
            existing = json.loads(m.group(1))
        except Exception:
            pass

    # Merge con protección: solo actualiza si el nuevo acumulado es MAYOR
    merged = dict(existing)
    for mes, data in computed.items():
        if mes not in merged or data["real"] > merged[mes].get("real", 0):
            merged[mes] = data

    # Reemplazar bloque allData
    new_data_js = "let allData = " + json.dumps(merged, ensure_ascii=False, separators=(",", ":")) + ";"
    html = re.sub(r"let allData\s*=\s*\{.*?\};", new_data_js, html, flags=re.DOTALL)

    # Actualizar timestamp
    ts = datetime.now().strftime("%d/%m/%Y %H:%M")
    html = re.sub(r"var lastUpdated\s*=\s*'[^']*';", f"var lastUpdated = '{ts}';", html)

    with open(html_path, "w", encoding="utf-8") as f:
        f.write(html)

    print(f"[OK] {html_path} actualizado | lastUpdated={ts}")


# ── Main ──────────────────────────────────────────────────────────────────────

def main() -> None:
    print("=" * 60)
    print("update_alcohol.py — AAFY 2026")
    print("=" * 60)

    if API_KEY == "YOUR_GOOGLE_DRIVE_API_KEY":
        print("[ERROR] Define la variable de entorno DRIVE_API_KEY")
        sys.exit(1)

    print("Listando archivos en Drive...")
    files = list_drive_files()
    print(f"  {len(files)} archivo(s) encontrado(s)")

    raw: dict[str, list] = {}

    for f in files:
        name_norm = _norm(f["name"])

        # Detectar mes en el nombre del archivo
        month_idx = None
        for idx, mes in enumerate(MESES_ES):
            if mes in name_norm:
                month_idx = idx + 1
                break

        if month_idx is None:
            print(f"  [skip] Sin mes detectado: {f['name']}")
            continue

        mes_key = f"{YEAR}{month_idx:02d}"
        print(f"  [proc] {f['name']}  →  {mes_key}")

        content = download_file(f["id"])
        try:
            wb = xlrd.open_workbook(file_contents=content)
            ws = wb.sheet_by_index(0)
        except Exception as e:
            print(f"         [ERROR] No se pudo abrir el archivo: {e}")
            continue

        rows = _parse_alcohol(ws)
        total_amount = sum(row["amount"] for row in rows)
        print(f"         Registros: {len(rows)}  |  Recaudación: ${total_amount:,.2f}")

        if mes_key not in raw:
            raw[mes_key] = []
        raw[mes_key].extend(rows)

    if not raw:
        print("\n[WARN] No se procesó ningún archivo. HTML sin cambios.")
        sys.exit(0)

    print("\nCalculando proyecciones y omisos...")
    computed = compute_month_data(raw)
    for mes, data in computed.items():
        print(f"  {mes}: real=${data['real']:,.0f}  omisos={data['total_omisos']}  proy=${data['proyeccion']:,.0f}")

    print("\nActualizando HTML...")
    update_html(computed, HTML_FILE)
    print("Listo.")


if __name__ == "__main__":
    main()
