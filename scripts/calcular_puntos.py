"""
Porra Mundial 2026 - Calculador de puntos
==========================================
Lee los excels de cada participante (hoja Pool) y los resultados reales,
calcula puntos y genera el fichero web/data.json para el frontend.

Sistema de puntos:
  - Resultado exacto (ej. 2-1 = 2-1): 3 puntos
  - Ganador/empate correcto (ej. pronosticó 2-1 y fue 3-1): 1 punto
  - Fallo total: 0 puntos
"""

import json
import os
import re
import glob
from pathlib import Path
from datetime import datetime
import pandas as pd


# ─── RUTAS ───────────────────────────────────────────────────────────────────
BASE_DIR   = Path(__file__).parent.parent
EXCELS_DIR = BASE_DIR / "excels"
print(EXCELS_DIR)
RESULTS_FILE = BASE_DIR / "resultados_reales.json"
OUTPUT_FILE  = BASE_DIR / "data.json"

# ─── PUNTUACIÓN ──────────────────────────────────────────────────────────────
PTS_EXACTO   = 3
PTS_GANADOR  = 1
PTS_FALLO    = 0


def sign(local, visitante):
    """Devuelve '1', 'X' o '2' según el resultado."""
    if local > visitante:
        return "1"
    elif local < visitante:
        return "2"
    else:
        return "X"


def puntuar(pronostico_str, resultado_str):
    """
    pronostico_str: e.g. "1|2-1"   (signo|goles_local-goles_visitante)
    resultado_str:  e.g. "2-1"     (goles_local-goles_visitante) o None si no jugado
    """
    if not resultado_str:
        return None  # partido no jugado todavía

    # Parsear pronóstico
    m = re.match(r"([12X])\|(\d+)-(\d+)", pronostico_str.strip())
    if not m:
        return None
    p_signo = m.group(1)
    p_local, p_visit = int(m.group(2)), int(m.group(3))

    # Parsear resultado real
    m2 = re.match(r"(\d+)-(\d+)", resultado_str.strip())
    if not m2:
        return None
    r_local, r_visit = int(m2.group(1)), int(m2.group(2))

    r_signo = sign(r_local, r_visit)

    if p_local == r_local and p_visit == r_visit:
        return PTS_EXACTO
    elif p_signo == r_signo:
        return PTS_GANADOR
    else:
        return PTS_FALLO


def leer_pronosticos_excel(excel_path):
    """
    Lee la hoja Pool del excel y devuelve un dict:
      { "partido_key": "signo|gol_local-gol_visitante", ... }
    
    partido_key = "A1" para el primer partido del grupo A jornada 1, etc.
    En realidad usamos el nombre de partido: "México-Sudáfrica"
    """
    try:
        df = pd.read_excel(excel_path, sheet_name="Pool", header=None)
    except Exception as e:
        print(f"  ERROR leyendo {excel_path}: {e}")
        return {}, "DESCONOCIDO"

    pronosticos = {}
    nombre = "DESCONOCIDO"

    for _, row in df.iterrows():
        # Fila de nombre del participante: columna B tiene el nombre
        # En la hoja Pool: col B = etiqueta partido, col C = pronóstico
        vals = [str(v).strip() if pd.notna(v) else "" for v in row]

        # Detectar nombre (fila que empieza con algo parecido a un grupo)
        # El nombre está en la fila con "PARTICIPANTE" o al inicio
        # Buscamos la celda que contiene el nombre escrito por el usuario
        # En el excel de Analia: Home!B12 tiene el nombre
        # En Pool, la celda C3 tiene el nombre
        if vals[1] == "Nombre" and vals[2] and vals[2] != "":
            nombre = vals[2]

        # Líneas de pronóstico tienen formato: cols[1] = "A1\tMéxico-Sudáfrica", cols[2] = "1|3-1"
        if len(vals) >= 3 and re.match(r"[A-Z]\d", vals[0]) and "|" in vals[2]:
            # partido en col B (índice 1) o col A (índice 0)
            partido_raw = vals[1]  # ej "México-Sudáfrica"
            pronostico  = vals[2]  # ej "1|3-1"
            if partido_raw and pronostico:
                pronosticos[partido_raw] = pronostico

        elif len(vals) >= 3 and vals[1] and "|" in vals[2]:
            partido_raw = vals[1]
            pronostico  = vals[2]
            if re.search(r'\w+-\w+', partido_raw) and "|" in pronostico:
                pronosticos[partido_raw] = pronostico

    # Si no encontramos nombre, intentar desde Home
    if nombre == "DESCONOCIDO":
        try:
            df_home = pd.read_excel(excel_path, sheet_name="Home", header=None)
            for _, row in df_home.iterrows():
                vals = [str(v).strip() if pd.notna(v) else "" for v in row]
                if "Escribe tu nombre" in vals:
                    idx = vals.index("Escribe tu nombre")
                    if idx + 1 < len(vals) and vals[idx + 1]:
                        nombre = vals[idx + 1]
                        break
        except:
            pass

    # Último recurso: nombre del archivo
    if nombre == "DESCONOCIDO":
        nombre = Path(excel_path).stem.replace("Excel-Mundial-2026__", "").replace("_", " ")

    return pronosticos, nombre


def leer_pronosticos_excel_v2(excel_path):
    """Versión mejorada que lee la hoja Pool correctamente."""
    try:
        # Leer con openpyxl para mejor acceso a celdas
        import openpyxl
        wb = openpyxl.load_workbook(excel_path, data_only=True, read_only=True)
    except Exception as e:
        print(f"  ERROR abriendo {excel_path}: {e}")
        return {}, Path(excel_path).stem

    pronosticos = {}
    nombre = "DESCONOCIDO"

    # ── Obtener nombre desde Home ──
    if "Home" in wb.sheetnames:
        ws_home = wb["Home"]
        for row in ws_home.iter_rows(values_only=True):
            row_vals = [str(v).strip() if v is not None else "" for v in row]
            for i, v in enumerate(row_vals):
                if v == "Analia" or (i > 0 and row_vals[i-1] in ("Escribe tu nombre", "✍️ Escribe tu nombre")):
                    if v and v not in ("Escribe tu nombre", "✍️ Escribe tu nombre"):
                        nombre = v
                        break
            if nombre != "DESCONOCIDO":
                break

    # ── Leer pronósticos desde Pool ──
    if "Pool" in wb.sheetnames:
        ws_pool = wb["Pool"]
        for row in ws_pool.iter_rows(values_only=True):
            if row is None:
                continue
            # col B (índice 1) = partido, col C (índice 2) = pronóstico
            b = str(row[1]).strip() if row[1] is not None else ""
            c = str(row[2]).strip() if row[2] is not None else ""

            if b == "Nombre" and c:
                nombre = c

            # Pronósticos: "México-Sudáfrica" en B, "1|3-1" en C
            if b and c and "|" in c and "-" in b:
                # Limpiar clave
                key = b.strip()
                pronosticos[key] = c.strip()

    wb.close()

    # Último recurso nombre
    if nombre == "DESCONOCIDO":
        nombre = Path(excel_path).stem.replace("Excel-Mundial-2026__", "").replace("_", " ")

    return pronosticos, nombre


def cargar_resultados_reales():
    """Carga el JSON de resultados reales."""
    if not RESULTS_FILE.exists():
        print(f"AVISO: No existe {RESULTS_FILE}. Usando resultados vacíos.")
        return {}
    with open(RESULTS_FILE, encoding="utf-8") as f:
        return json.load(f)


def calcular_clasificacion():
    print("=" * 60)
    print("  PORRA MUNDIAL 2026 - Calculando clasificación")
    print("=" * 60)

    # Cargar resultados reales
    resultados_reales = cargar_resultados_reales()
    print(f"\n  Partidos con resultado: {len(resultados_reales)}")

    # Leer todos los excels
    excels = sorted(glob.glob(str(EXCELS_DIR / "*.xlsx")))
    if not excels:
        excels = sorted(glob.glob(str(EXCELS_DIR / "*.xls")))

    print(f"  Participantes encontrados: {len(excels)}\n")

    participantes = []

    for excel_path in excels:
        print(f"  Procesando: {Path(excel_path).name}")
        pronosticos, nombre = leer_pronosticos_excel_v2(excel_path)
        print(f"    Nombre: {nombre}  |  Pronósticos: {len(pronosticos)}")

        puntos_total   = 0
        exactos        = 0
        ganadores      = 0
        fallos         = 0
        partidos_eval  = 0
        detalle        = []

        for partido, resultado_real in resultados_reales.items():
            pronostico = pronosticos.get(partido)
            if pronostico is None:
                continue  # No encontrado en el excel

            pts = puntuar(pronostico, resultado_real)
            if pts is None:
                continue

            partidos_eval += 1
            puntos_total  += pts

            if pts == PTS_EXACTO:
                exactos += 1
            elif pts == PTS_GANADOR:
                ganadores += 1
            else:
                fallos += 1

            detalle.append({
                "partido": partido,
                "pronostico": pronostico,
                "resultado": resultado_real,
                "puntos": pts
            })

        participantes.append({
            "nombre": nombre,
            "puntos": puntos_total,
            "exactos": exactos,
            "ganadores": ganadores,
            "fallos": fallos,
            "partidos": partidos_eval,
            "detalle": detalle,
            "archivo": Path(excel_path).name
        })

    # Ordenar por puntos (desempate: exactos)
    participantes.sort(key=lambda x: (-x["puntos"], -x["exactos"]))

    # Añadir posición
    for i, p in enumerate(participantes):
        p["posicion"] = i + 1

    # Estadísticas generales
    total_partidos_jugados = len(resultados_reales)

    output = {
        "actualizado": datetime.now().strftime("%d/%m/%Y %H:%M"),
        "partidos_jugados": total_partidos_jugados,
        "participantes": participantes
    }

    # Guardar JSON
    OUTPUT_FILE.parent.mkdir(parents=True, exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2)

    print(f"\n  ✓ Clasificación guardada en: {OUTPUT_FILE}")
    print(f"\n{'─'*40}")
    print(f"  {'POS':<5} {'NOMBRE':<20} {'PTS':<6} {'✓':<5} {'~':<5} {'✗'}")
    print(f"{'─'*40}")
    for p in participantes:
        print(f"  {p['posicion']:<5} {p['nombre']:<20} {p['puntos']:<6} {p['exactos']:<5} {p['ganadores']:<5} {p['fallos']}")
    print(f"{'─'*40}")
    print(f"\n  Actualizado: {output['actualizado']}\n")

    return output


if __name__ == "__main__":
    calcular_clasificacion()
