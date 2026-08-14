"""
build_data.py — Convierte la BBDD mensual de Turismo (Excel) al data.json
que consume el panel (index.html).

Uso:
    python build_data.py BBDD_Turismo.xlsx

Genera (o sobrescribe) data.json en la misma carpeta que este script.
"""
import sys
import json
import re
from datetime import datetime, timezone, timedelta
from pathlib import Path

import pandas as pd

AR_TZ = timezone(timedelta(hours=-3))  # Argentina (UTC-3, sin horario de verano)

SHEET_NAME = "Acumulado"

COL_ATENCION = "Cmo_calificaras_la_atencin_recibida_al_momento_de_solicitar_el_servicioDel_1_al_5_1__Muy_malo_5__Excelente"
COL_PRESTADOR = "Cmo_evaluaras_la_calidad_del_servicio_brindado_por_el_prestador_asignadoDel_1_al_5_1__Muy_malo_5__Excelente"

RESPONDIO_MAP = {"RESPONDIÓ": "Responde", "PARCIAL": "Parcial", "NO RESPONDIÓ": "No Responde"}
RECOMIENDA_MAP = {"RECOMIENDA": "Recomienda", "NEUTRAL": "Neutral", "NO RECOMIENDA": "No Recomienda"}


def to_score(value):
    if pd.isna(value):
        return None
    try:
        n = int(float(value))
    except (ValueError, TypeError):
        return None
    return n if 1 <= n <= 5 else None


def to_month_key(ts):
    if pd.isna(ts):
        return None
    return pd.Timestamp(ts).strftime("%Y-%m")


def limpio(v):
    if pd.isna(v):
        return None
    s = str(v).strip()
    return s if s else None


def build(input_path: Path, output_path: Path):
    df = pd.read_excel(input_path, sheet_name=SHEET_NAME)
    df.columns = [c.strip() for c in df.columns]

    casos = []
    for _, row in df.iterrows():
        respondio_raw = limpio(row.get("RESPONDIO?"))
        recomienda_raw = limpio(row.get("Recomienda Ok"))
        casos.append({
            "mes": to_month_key(row.get("Mes")),
            "campana": limpio(row.get("Campaña")),
            "atencion": to_score(row.get(COL_ATENCION)),
            "prestador": to_score(row.get(COL_PRESTADOR)),
            "recomienda": RECOMIENDA_MAP.get(recomienda_raw, recomienda_raw) if recomienda_raw else None,
            "respondio": RESPONDIO_MAP.get(respondio_raw, respondio_raw) if respondio_raw else None,
        })

    meses = sorted({c["mes"] for c in casos if c["mes"]})

    data = {
        "generadoEl": datetime.now(AR_TZ).strftime("%Y-%m-%d %H:%M"),
        "meses": meses,
        "casos": casos,
    }

    output_path.write_text(json.dumps(data, ensure_ascii=False, indent=0), encoding="utf-8")
    print(f"OK -> {output_path} ({len(casos)} casos, meses: {', '.join(meses)})")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Uso: python build_data.py <ruta al Excel de la BBDD>")
        sys.exit(1)

    input_path = Path(sys.argv[1])
    output_path = Path(__file__).parent / "data.json"
    build(input_path, output_path)
