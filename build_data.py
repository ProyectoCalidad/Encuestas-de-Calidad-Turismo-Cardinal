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
COL_SERVICIO_PRESTADO = "Servicio prestado"
COL_NOMBRE_PRESTADOR = "Prestador"

# Posición (0-indexada) de "Mes" y "Campaña" en la planilla. Se usan como respaldo
# porque en algunas exportaciones el encabezado de esas dos columnas llega roto
# (la celda trae un valor de dato en vez del nombre de columna).
POS_MES = 0
POS_CAMPANA = 3

NA_TEXTO = {"#N/A", "#N/D", "N/A", "0"}

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


def limpio_categoria(v):
    """Para Servicio prestado / Prestador: además de vacíos, descarta '#N/A' y similares."""
    s = limpio(v)
    if s is None or s.upper() in NA_TEXTO:
        return None
    return s


def valor_por_nombre_o_posicion(row, nombre, columnas, posicion):
    """Busca la columna por nombre; si no existe (encabezado roto), cae a la posición fija."""
    if nombre in columnas:
        return row.get(nombre)
    return row.iloc[posicion] if posicion < len(row) else None


def build(input_path: Path, output_path: Path):
    df = pd.read_excel(input_path, sheet_name=SHEET_NAME)
    df.columns = [c.strip() if isinstance(c, str) else c for c in df.columns]
    columnas = set(df.columns)

    casos = []
    for _, row in df.iterrows():
        respondio_raw = limpio(row.get("RESPONDIO?"))
        recomienda_raw = limpio(row.get("Recomienda Ok"))
        mes_raw = valor_por_nombre_o_posicion(row, "Mes", columnas, POS_MES)
        campana_raw = valor_por_nombre_o_posicion(row, "Campaña", columnas, POS_CAMPANA)
        casos.append({
            "mes": to_month_key(mes_raw),
            "campana": limpio(campana_raw),
            "atencion": to_score(row.get(COL_ATENCION)),
            "prestador": to_score(row.get(COL_PRESTADOR)),
            "recomienda": RECOMIENDA_MAP.get(recomienda_raw, recomienda_raw) if recomienda_raw else None,
            "respondio": RESPONDIO_MAP.get(respondio_raw, respondio_raw) if respondio_raw else None,
            "servicio": limpio_categoria(row.get(COL_SERVICIO_PRESTADO)),
            "prestadorNombre": limpio_categoria(row.get(COL_NOMBRE_PRESTADOR)),
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
