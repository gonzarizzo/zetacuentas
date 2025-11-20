#!/usr/bin/env python
import glob
import os
from functools import lru_cache
from datetime import datetime

import pandas as pd
import requests


# Patrón de archivos de entrada
INPUT_PATTERN = "Estado_De_Cuenta*.xls"

# Archivos de salida
OUTPUT_PESOS = "itau_debito_pesos.xlsx"
OUTPUT_DOLARES = "itau_debito_dolares.xlsx"


# ======================
# Helpers de formato
# ======================

def parse_importe(value):
    if pd.isna(value):
        return 0.0

    s = str(value).strip()
    if not s:
        return 0.0

    sign = -1 if s.startswith("-") else 1
    s = s.replace("-", "")
    s = s.replace(".", "").replace(",", ".")
    try:
        return sign * float(s)
    except:
        return 0.0


def parse_fecha_texto(fecha_val):
    if isinstance(fecha_val, (datetime, )):
        return fecha_val.strftime("%d/%m/%Y")

    fecha_str = str(fecha_val).strip()
    if not fecha_str:
        return ""

    for fmt in ("%d/%m/%y", "%d/%m/%Y", "%Y-%m-%d"):
        try:
            dt = datetime.strptime(fecha_str, fmt).date()
            return dt.strftime("%d/%m/%Y")
        except:
            continue

    return fecha_str


# ======================
# Cotización USD/UYU
# ======================

@lru_cache(maxsize=None)
def get_usd_rate_uyu(fecha_texto):
    dt = datetime.strptime(fecha_texto, "%d/%m/%Y").date()
    fecha_api = dt.isoformat()

    url = f"https://api.exchangerate.host/{fecha_api}"
    params = {"base": "USD", "symbols": "UYU"}

    try:
        resp = requests.get(url, params=params, timeout=10)
        resp.raise_for_status()
        data = resp.json()
        return float(data["rates"]["UYU"])
    except Exception as e:
        print(f"[ADVERTENCIA] No se pudo obtener cotización USD/UYU para {fecha_texto}: {e}")
        return 0.0


# ======================
# Lectura de Excel Itaú
# ======================

def detectar_moneda(path):
    df = pd.read_excel(path, header=None, engine="xlrd")
    try:
        val = df.iloc[4, 5]
    except:
        val = ""

    val_str = str(val).upper()

    if "DÓLAR" in val_str or "DOLAR" in val_str or "USD" in val_str:
        return "DOLARES"
    return "PESOS"


def encontrar_fila_header(df_raw):
    for i in range(len(df_raw)):
        row = df_raw.iloc[i]
        for val in row:
            if isinstance(val, str) and "fecha" in val.strip().lower():
                return i
    return None


def extraer_movimientos_desde_archivo(path, moneda):
    df_raw = pd.read_excel(path, header=None, engine="xlrd")
    header_row = encontrar_fila_header(df_raw)

    if header_row is None:
        raise ValueError(f"No se encontró fila de encabezados en el archivo: {path}")

    headers = df_raw.iloc[header_row].tolist()
    df = df_raw.iloc[header_row + 1:].copy()
    df.columns = headers

    col_map = {col: str(col).strip().lower() for col in df.columns}

    col_fecha = next((col for col, low in col_map.items() if "fecha" in low), None)
    col_desc = next((col for col, low in col_map.items()
                     if any(x in low for x in ["descripcion", "descripción", "detalle", "concepto"])),
                    None)

    col_importe = next((col for col, low in col_map.items() if "importe" in low), None)
    col_debito = next((col for col, low in col_map.items() if "debito" in low or "débito" in low), None)
    col_credito = next((col for col, low in col_map.items() if "credito" in low or "crédito" in low), None)

    if col_fecha is None or col_desc is None:
        raise ValueError(f"No se encontraron columnas de Fecha/Descripción en el archivo: {path}")

    if col_importe is None and not (col_debito and col_credito):
        raise ValueError(
            f"No se encontró columna de importe ni columnas débito/crédito en el archivo: {path}"
        )

    registros = []

    for _, row in df.iterrows():
        fecha_val = row.get(col_fecha, None)
        fecha_txt = parse_fecha_texto(fecha_val)

        # ❌ Ignorar filas sin fecha (Saldo final)
        if not fecha_txt:
            continue

        desc_val = row.get(col_desc, "")
        desc_txt = str(desc_val).strip() if not pd.isna(desc_val) else ""
        desc_upper = desc_txt.upper()

        # ❌ NUEVO: eliminar SALDO ANTERIOR / SALDO FINAL
        if "SALDO ANTERIOR" in desc_upper or "SALDO FINAL" in desc_upper:
            continue

        if col_importe is not None:
            importe = parse_importe(row.get(col_importe, 0))
            creditos = abs(importe) if importe < 0 else 0.0
            debitos = importe if importe > 0 else 0.0
        else:
            debitos = parse_importe(row.get(col_debito, 0)) if col_debito else 0.0
            creditos = parse_importe(row.get(col_credito, 0)) if col_credito else 0.0

        if (creditos == 0.0 and debitos == 0.0) and not desc_txt:
            continue

        registros.append(
            {
                "Fecha": fecha_txt,
                "Descripcion": desc_txt,
                "Creditos": creditos,
                "Debitos": debitos,
            }
        )

    return pd.DataFrame(registros) if registros else pd.DataFrame(columns=["Fecha", "Descripcion", "Creditos", "Debitos"])


# ======================
# Programa principal
# ======================

def main():
    archivos = sorted(glob.glob(INPUT_PATTERN))

    if not archivos:
        print(f"No se encontraron archivos con el patrón: {INPUT_PATTERN}")
        return

    print("Archivos encontrados:")
    for a in archivos:
        print(" -", a)

    registros_pesos = []
    registros_dolares = []

    for path in archivos:
        print(f"\nProcesando archivo: {path}")
        moneda = detectar_moneda(path)
        print(f"  → Moneda detectada: {moneda}")

        df_mov = extraer_movimientos_desde_archivo(path, moneda)

        if df_mov.empty:
            print("  (Sin movimientos detectados, se omite)")
            continue

        if moneda == "DOLARES":
            registros_dolares.append(df_mov)
        else:
            registros_pesos.append(df_mov)

    if registros_pesos:
        df_pesos = pd.concat(registros_pesos, ignore_index=True)
        df_pesos["Cotizacion"] = 0.0
        df_pesos.to_excel(OUTPUT_PESOS, index=False)
        print(f"\n✅ Archivo de Pesos generado: {OUTPUT_PESOS}")

    if registros_dolares:
        df_dolares = pd.concat(registros_dolares, ignore_index=True)
        df_dolares["Cotizacion"] = df_dolares["Fecha"].apply(get_usd_rate_uyu)
        df_dolares.to_excel(OUTPUT_DOLARES, index=False)
        print(f"✅ Archivo de Dólares generado: {OUTPUT_DOLARES}")


if __name__ == "__main__":
    main()
