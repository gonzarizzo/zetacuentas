#!/usr/bin/env python
import os
from functools import lru_cache
from datetime import datetime

import pandas as pd
import requests


# Archivo de entrada/salida
INPUT_FILE = "Detalle_Movimiento_Cuenta.xls"
OUTPUT_FILE = "brou_detalle_movimientos.xlsx"


# ======================
# Helpers de formato
# ======================

def parse_importe(value):
    """
    Convierte importes con formato tipo:
      - '1.234,56'
      - '-15,68'
      - '1234.56'
    a float.
    """
    if pd.isna(value):
        return 0.0

    s = str(value).strip()
    if not s:
        return 0.0

    sign = -1 if s.startswith("-") else 1
    s = s.replace("-", "")

    # Soporta formatos con . y , mezclados
    # Ej: "1.234,56" -> "1234.56"
    s = s.replace(".", "").replace(",", ".")

    try:
        return sign * float(s)
    except ValueError:
        return 0.0


def parse_fecha_texto(fecha_val):
    """
    Recibe una fecha como:
      - string 'DD/MM/AA' o 'DD/MM/AAAA'
      - objeto datetime / date
    y devuelve string 'DD/MM/AAAA' (texto).
    """
    if isinstance(fecha_val, (datetime, )):
        return fecha_val.strftime("%d/%m/%Y")

    fecha_str = str(fecha_val).strip()
    if not fecha_str:
        return ""

    for fmt in ("%d/%m/%y", "%d/%m/%Y", "%Y-%m-%d"):
        try:
            dt = datetime.strptime(fecha_str, fmt).date()
            return dt.strftime("%d/%m/%Y")
        except ValueError:
            continue

    # Si no pudimos parsear, devolvemos el original
    return fecha_str


# ======================
# Cotización USD/UYU
# ======================

@lru_cache(maxsize=None)
def get_usd_rate_uyu(fecha_texto):
    """
    Devuelve la cotización 1 USD en UYU para la fecha dada (DD/MM/AAAA),
    usando https://api.exchangerate.host/<fecha>?base=USD&symbols=UYU
    """
    dt = datetime.strptime(fecha_texto, "%d/%m/%Y").date()
    fecha_api = dt.isoformat()  # YYYY-MM-DD

    url = f"https://api.exchangerate.host/{fecha_api}"
    params = {
        "base": "USD",
        "symbols": "UYU",
    }

    try:
        resp = requests.get(url, params=params, timeout=10)
        resp.raise_for_status()
        data = resp.json()
        rate = data["rates"]["UYU"]
        return float(rate)
    except Exception as e:
        print(f"[ADVERTENCIA] No se pudo obtener cotización USD/UYU para {fecha_texto}: {e}")
        return 0.0


# ======================
# Moneda del archivo
# ======================

def detectar_moneda_detalle(path):
    """
    Detecta si el archivo Detalle_Movimiento_Cuenta.xls está en Pesos o Dólares.
    Busca palabras clave (PESOS / DÓLAR / DOLAR / USD) en las primeras filas.
    """
    df_raw = pd.read_excel(path, header=None, engine="xlrd")

    max_rows = min(15, len(df_raw))
    max_cols = min(10, df_raw.shape[1])

    moneda = "PESOS"

    for i in range(max_rows):
        for j in range(max_cols):
            val = df_raw.iat[i, j]
            if pd.isna(val):
                continue
            txt = str(val).upper()
            if any(x in txt for x in ["DÓLAR", "DOLAR", "USD"]):
                return "DOLARES"
            if "PESOS" in txt or "U$S" in txt:
                moneda = "PESOS"

    return moneda


# ======================
# Extracción de movimientos
# ======================

def extraer_movimientos(path):
    """
    Lee Detalle_Movimiento_Cuenta.xls asumiendo:
      - La información útil comienza en la fila 18 del archivo.
      - Fila 18 = encabezados (títulos de columnas).
      - Fila 19 en adelante = datos.

    Devuelve DataFrame con columnas:
      Fecha, Descripcion, Creditos, Debitos
    (sin cotización todavía).
    """
    # Leemos TODO sin encabezados
    df_raw = pd.read_excel(path, header=None, engine="xlrd")

    # Fila 18 en Excel = índice 17 (0-based)
    header_row_index = 17

    if header_row_index >= len(df_raw):
        raise ValueError("El archivo no tiene suficientes filas para que la fila 18 contenga encabezados.")

    # Encabezados = fila 18
    headers = df_raw.iloc[header_row_index].tolist()

    # Datos = filas a partir de la 19 (índice 18)
    df = df_raw.iloc[header_row_index + 1:].copy()
    df.columns = headers

    # Normalizamos nombres de columnas
    col_map = {col: str(col).strip().lower() for col in df.columns}

    # Buscar columna de fecha
    col_fecha = None
    for col, low in col_map.items():
        if "fecha" in low:
            col_fecha = col
            break

    # Buscar columna de descripción
    col_desc = None
    for col, low in col_map.items():
        if any(x in low for x in ["descripcion", "descripción", "detalle", "concepto"]):
            col_desc = col
            break

    # Buscar columna de importe único o débito/crédito separados
    col_importe = None
    col_debito = None
    col_credito = None

    for col, low in col_map.items():
        if "importe" in low or "monto" in low:
            col_importe = col
        if "debito" in low or "débito" in low:
            col_debito = col
        if "credito" in low or "crédito" in low:
            col_credito = col

    if col_fecha is None or col_desc is None:
        raise ValueError(
            "No se encontraron columnas de Fecha y/o Descripción en la sección de datos del archivo Detalle_Movimiento_Cuenta.xls"
        )

    if col_importe is None and not (col_debito and col_credito):
        raise ValueError(
            "No se encontró columna de importe ni columnas débito/crédito en la sección de datos del archivo Detalle_Movimiento_Cuenta.xls"
        )

    registros = []

    for _, row in df.iterrows():
        fecha_val = row.get(col_fecha, None)
        desc_val = row.get(col_desc, "")

        fecha_txt = parse_fecha_texto(fecha_val)
        desc_txt = str(desc_val).strip() if not pd.isna(desc_val) else ""

        # Fecha y descripción son obligatorias
        if not fecha_txt or not desc_txt:
            continue

        # Filtro de SALDO ANTERIOR / SALDO FINAL por si aparece en este formato también
        desc_upper = desc_txt.upper()
        if "SALDO ANTERIOR" in desc_upper or "SALDO FINAL" in desc_upper:
            continue

        if col_importe is not None:
            importe = parse_importe(row.get(col_importe, 0))
            if importe < 0:
                creditos = abs(importe)
                debitos = 0.0
            else:
                creditos = 0.0
                debitos = importe
        else:
            debitos = parse_importe(row.get(col_debito, 0)) if col_debito else 0.0
            creditos = parse_importe(row.get(col_credito, 0)) if col_credito else 0.0

        # Ignoramos filas totalmente en blanco
        if creditos == 0.0 and debitos == 0.0 and not desc_txt:
            continue

        registros.append(
            {
                "Fecha": fecha_txt,
                "Descripcion": desc_txt,
                "Creditos": creditos,
                "Debitos": debitos,
            }
        )

    if not registros:
        return pd.DataFrame(columns=["Fecha", "Descripcion", "Creditos", "Debitos"])

    return pd.DataFrame(registros)


# ======================
# Programa principal
# ======================

def main():
    if not os.path.exists(INPUT_FILE):
        print(f"❌ No se encontró el archivo {INPUT_FILE} en el directorio actual.")
        return

    print(f"📄 Procesando archivo: {INPUT_FILE}")

    moneda = detectar_moneda_detalle(INPUT_FILE)
    print(f"💱 Moneda detectada: {moneda}")

    df_mov = extraer_movimientos(INPUT_FILE)

    if df_mov.empty:
        print("⚠️ No se detectaron movimientos en el archivo.")
        return

    # Agregar columna de Cotizacion
    if moneda == "DOLARES":
        df_mov["Cotizacion"] = df_mov["Fecha"].apply(get_usd_rate_uyu)
    else:
        df_mov["Cotizacion"] = 0.0

    # Orden de columnas según especificación
    df_salida = df_mov[["Fecha", "Descripcion", "Creditos", "Debitos", "Cotizacion"]]

    df_salida.to_excel(OUTPUT_FILE, index=False)
    print(f"✅ Archivo de salida generado: {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
