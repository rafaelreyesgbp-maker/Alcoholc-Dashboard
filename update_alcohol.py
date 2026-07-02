#!/usr/bin/env python3
"""
update_alcohol.py
Descarga archivos XLS de Google Drive, calcula proyecciones de cierre y
actualiza index.html del Impuesto a la Venta Final de Bebidas con Contenido
Alcohólico (AAFY 2026).

El formato de allData en el HTML usa claves de mes (1-12) e incluye todos
los campos calculados que necesita renderDash() en el JavaScript.
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
    5_971_005,  # 1 Enero
    2_033_044,  # 2 Febrero
    2_357_586,  # 3 Marzo
    2_380_846,  # 4 Abril
    2_204_240,  # 5 Mayo
    2_495_661,  # 6 Junio
    2_249_735,  # 7 Julio
    2_324_560,  # 8 Agosto
    2_360_836,  # 9 Septiembre
    3_233_310,  # 10 Octubre
    3_540_960,  # 11 Noviembre
    4_050_214,  # 12 Diciembre
]

MONTH_LABELS = ["","Enero","Febrero","Marzo","Abril","Mayo","Junio",
                "Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"]

MESES_ES = ["enero","febrero","marzo","abril","mayo","junio",
            "julio","agosto","septiembre","octubre","noviembre","diciembre"]

# ── Columnas del archivo (base 0) ─────────────────────────────────────────────
ALC_RFC    = 0   # A
ALC_CONTRIB = 1  # B
ALC_PERIODO = 4  # E  (YYYYMM como float)
ALC_N      = 13  # N  (se suma)
ALC_I      = 7   # H  (se resta)  → Recaudación = N - H


# ── Utilidades ────────────────────────────────────────────────────────────────

def _norm(text: str) -> str:
    return "".join(c for c in unicodedata.normalize("NFD", text.lower())
                   if unicodedata.category(c) != "Mn")

def _cell_str(cell) -> str:
    v = cell.value
    if cell.ctype == 2 and v == int(v):
        return str(int(v))
    return str(v).strip()

def _cell_float(cell) -> float:
    try:
        return float(cell.value) if cell.ctype in (2, 3) else 0.0
    except Exception:
        return 0.0

def _format_period(p: str) -> str:
    labels = ["","Ene","Feb","Mar","Abr","May","Jun","Jul","Ago","Sep","Oct","Nov","Dic"]
    try:
        return labels[int(p[4:6])] + "-" + p[2:4]
    except Exception:
        return p

def _prev_period(p: str) -> str:
    y, m = int(p[:4]), int(p[4:])
    m -= 1
    if m == 0:
        m = 12; y -= 1
    return f"{y}{m:02d}"


# ── Parser del archivo de alcohol ─────────────────────────────────────────────

def _find_data_start(ws) -> int:
    for r in range(min(20, ws.nrows)):
        try:
            val = _norm(_cell_str(ws.cell(r, ALC_RFC)))
            if re.match(r"^[a-z]{3,4}\d{6}", val):
                return r
        except Exception:
            pass
    return 0

def _parse_alcohol(ws) -> list:
    start = _find_data_start(ws)
    rows = []
    for r in range(start, ws.nrows):
        try:
            rfc     = _cell_str(ws.cell(r, ALC_RFC)).upper().strip()
            contrib = _cell_str(ws.cell(r, ALC_CONTRIB)).strip()
            period_raw = ws.cell(r, ALC_PERIODO).value
            periodo = str(int(float(period_raw))) if period_raw else ""
            if len(periodo) != 6:
                continue
            n_val = _cell_float(ws.cell(r, ALC_N))
            i_val = _cell_float(ws.cell(r, ALC_I))
            rec   = n_val - i_val
            if rfc and re.match(r"^[A-Z]{3,4}\d{6}", rfc) and rec != 0:
                rows.append({"rfc": rfc, "contrib": contrib,
                             "periodo": periodo, "recaudacion": round(rec, 2)})
        except Exception:
            continue
    return rows


# ── Google Drive ───────────────────────────────────────────────────────────────

def _drive_list() -> list:
    url = (f"https://www.googleapis.com/drive/v3/files"
           f"?q=%27{FOLDER_ID}%27+in+parents+and+trashed%3Dfalse"
           f"&fields=files(id,name)&pageSize=200&key={API_KEY}")
    r = requests.get(url, timeout=30)
    r.raise_for_status()
    return r.json().get("files", [])

def _drive_download(file_id: str) -> bytes:
    url = f"https://www.googleapis.com/drive/v3/files/{file_id}?alt=media&key={API_KEY}"
    r = requests.get(url, timeout=60)
    r.raise_for_status()
    return r.content


# ── Lógica de omisos (equivalente a computeMonth() en el JS) ──────────────────

def _get_dominant(records: list):
    totals = {}
    for r in records:
        if r["periodo"]:
            totals[r["periodo"]] = totals.get(r["periodo"], 0) + r["recaudacion"]
    return max(totals, key=lambda k: totals[k]) if totals else None

def _get_missing_periods(paid_set: set, dominant: str, max_back: int, stop_before) -> list:
    out = []
    p = dominant
    while p not in paid_set:
        if stop_before and int(p) < int(stop_before):
            break
        out.append(p)
        p = _prev_period(p)
        if len(out) >= max_back:
            break
    return out

def compute_month(month_num: int, all_month_data: dict) -> dict:
    cur      = all_month_data.get(month_num, [])
    dominant = _get_dominant(cur)
    acumulado = sum(r["recaudacion"] for r in cur)

    ref_months = [m for m in range(max(1, month_num - 4), month_num)
                  if all_month_data.get(m)]
    n_ref = len(ref_months)

    # Periodos y nombres globales acumulados hasta este mes
    global_periods = {}
    global_contrib = {}
    for m in range(1, month_num + 1):
        for r in all_month_data.get(m, []):
            gp = global_periods.setdefault(r["rfc"], {})
            if r["periodo"] not in gp or r["recaudacion"] > gp[r["periodo"]]:
                gp[r["periodo"]] = r["recaudacion"]
            if r["rfc"] not in global_contrib and r["contrib"]:
                global_contrib[r["rfc"]] = r["contrib"]

    # Acumulado por mes de referencia
    rfc_ref_count   = {}
    rfc_ref_periods = {}
    rfc_contrib     = {}
    paid_2026_in_ref = {}

    for rm in ref_months:
        seen = set()
        for r in all_month_data.get(rm, []):
            rp = rfc_ref_periods.setdefault(r["rfc"], {})
            if r["periodo"] not in rp or r["recaudacion"] > rp[r["periodo"]]:
                rp[r["periodo"]] = r["recaudacion"]
            if str(r["periodo"]).startswith("2026"):
                paid_2026_in_ref.setdefault(r["rfc"], set()).add(str(r["periodo"]))
            if r["rfc"] not in seen:
                seen.add(r["rfc"])
                rfc_ref_count[r["rfc"]] = rfc_ref_count.get(r["rfc"], 0) + 1
            if r["rfc"] not in rfc_contrib and r["contrib"]:
                rfc_contrib[r["rfc"]] = r["contrib"]

    # Candidatos a omisos
    candidates = set()
    for rfc, cnt in rfc_ref_count.items():
        if cnt >= 2:
            candidates.add(rfc)
    for rfc, ps in paid_2026_in_ref.items():
        if len(ps) >= 2:
            candidates.add(rfc)

    omisos = []
    for rfc in candidates:
        cnt = rfc_ref_count.get(rfc, 0)
        p26 = paid_2026_in_ref.get(rfc, set())
        if cnt < 2 and len(p26) < 2:
            continue
        paid_set = set(global_periods.get(rfc, {}).keys())
        if not dominant or dominant in paid_set:
            continue
        has_2026 = any(p.startswith("2026") for p in paid_set)
        missing = (_get_missing_periods(paid_set, dominant, 12, None) if has_2026
                   else _get_missing_periods(paid_set, dominant, 12, "202601"))
        if not missing:
            continue
        ref_amounts = list(rfc_ref_periods.get(rfc, {}).values())
        if not ref_amounts:
            continue
        avg = sum(ref_amounts) / len(ref_amounts)
        if not has_2026:
            seg = "omisos_totales"
        elif cnt >= n_ref:
            seg = "alta"
        elif cnt >= 3:
            seg = "media"
        else:
            seg = "seguimiento"
        omisos.append({
            "rfc":      rfc,
            "contrib":  rfc_contrib.get(rfc, global_contrib.get(rfc, "")),
            "count":    cnt,
            "avg":      round(avg * len(missing)),
            "nMissing": len(missing),
            "pending":  [_format_period(p) for p in missing],
            "seg":      seg,
        })

    omisos.sort(key=lambda o: o["avg"], reverse=True)
    esperado = sum(o["avg"] for o in omisos if o["seg"] in ("alta", "media"))
    proyeccion = acumulado + esperado
    meta = MONTHLY_GOALS[month_num - 1]

    # Segmentos
    segments = {}
    for o in omisos:
        s = segments.setdefault(o["seg"], {"count": 0, "monto": 0, "omisos": []})
        s["count"] += 1
        s["monto"]  += o["avg"]
        s["omisos"].append(dict(o))
    for s in segments.values():
        s["monto"] = round(s["monto"])
        s["omisos"].sort(key=lambda o: o["avg"], reverse=True)

    return {
        "mes_label":       MONTH_LABELS[month_num],
        "mes_num":         month_num,
        "meta":            meta,
        "dominant_period": int(dominant) if dominant else 0,
        "ref_months":      ref_months,
        "acumulado_real":  round(acumulado),
        "total_omisos":    len(omisos),
        "total_esperado":  round(esperado),
        "proyeccion_cierre": round(proyeccion),
        "meta_cruzada":    proyeccion >= meta,
        "pct_acumulado":   acumulado / meta * 100 if meta else 0,
        "pct_proyeccion":  proyeccion / meta * 100 if meta else 0,
        "segmentos":       segments,
        "omisos":          omisos[:5000],
    }


# ── Actualización del HTML ────────────────────────────────────────────────────

def update_html(computed: dict, html_path: str) -> None:
    with open(html_path, "r", encoding="utf-8") as f:
        html = f.read()

    # Cargar datos existentes (filtrar claves inválidas como "202606")
    existing = {}
    m = re.search(r"let allData\s*=\s*(\{.*?\});", html, re.DOTALL)
    if m:
        try:
            raw_existing = json.loads(m.group(1))
            existing = {k: v for k, v in raw_existing.items()
                        if k.isdigit() and 1 <= int(k) <= 12}
        except Exception:
            pass

    # Merge con protección: solo actualiza si acumulado_real es MAYOR
    merged = dict(existing)
    for mes_key, data in computed.items():
        if (mes_key not in merged or
                data["acumulado_real"] > merged[mes_key].get("acumulado_real", 0)):
            merged[mes_key] = data

    new_js = "let allData = " + json.dumps(merged, ensure_ascii=False, separators=(",", ":")) + ";"
    html = re.sub(r"let allData\s*=\s*\{.*?\};", new_js, html, flags=re.DOTALL)

    ts = datetime.now().strftime("%d/%m/%Y %H:%M")
    html = re.sub(r"var lastUpdated\s*=\s*'[^']*';", f"var lastUpdated = '{ts}';", html)

    with open(html_path, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"[OK] {html_path} actualizado | {ts}")


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    print("=" * 60)
    print("update_alcohol.py — AAFY 2026")
    print("=" * 60)

    if API_KEY == "YOUR_GOOGLE_DRIVE_API_KEY":
        print("[ERROR] Define la variable de entorno DRIVE_API_KEY")
        sys.exit(1)

    print("Listando archivos en Drive...")
    files = _drive_list()
    print(f"  {len(files)} archivo(s) encontrado(s)")

    # all_month_data[month_num] = [registros]
    all_month_data: dict[int, list] = {}

    for f in files:
        name_norm = _norm(f["name"])
        month_idx = next((i for i, m in enumerate(MESES_ES) if m in name_norm), None)
        if month_idx is None:
            print(f"  [skip] Sin mes en nombre: {f['name']}")
            continue

        month_num = month_idx + 1
        print(f"  [proc] {f['name']}  →  mes {month_num} ({MONTH_LABELS[month_num]})")

        content = _drive_download(f["id"])
        try:
            wb = xlrd.open_workbook(file_contents=content)
            ws = wb.sheet_by_index(0)
        except Exception as e:
            print(f"         [ERROR] No se pudo abrir: {e}")
            continue

        rows = _parse_alcohol(ws)
        total = sum(r["recaudacion"] for r in rows)
        print(f"         Registros: {len(rows)}  |  Recaudación: ${total:,.2f}")

        all_month_data.setdefault(month_num, []).extend(rows)

    if not all_month_data:
        print("\n[WARN] No se procesó ningún archivo. HTML sin cambios.")
        sys.exit(0)

    print("\nCalculando proyecciones y omisos...")
    computed = {}
    for month_num in sorted(all_month_data.keys()):
        data = compute_month(month_num, all_month_data)
        computed[str(month_num)] = data   # ← clave = "6", no "202606"
        print(f"  {MONTH_LABELS[month_num]}: "
              f"real=${data['acumulado_real']:,.0f}  "
              f"omisos={data['total_omisos']}  "
              f"proy=${data['proyeccion_cierre']:,.0f}")

    print("\nActualizando HTML...")
    update_html(computed, HTML_FILE)
    print("Listo.")


if __name__ == "__main__":
    main()
