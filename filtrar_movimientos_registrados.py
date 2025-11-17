#!/usr/bin/env python3
"""
Filtra los movimientos que ya se encuentran cargados en comprobante.xlsx.

Para cada tipo de cuenta documentado en el comprobante, abre el archivo
correspondiente (movimientos_pesos, movimientos_dolares, etc.) y elimina
las filas cuyo trío (fecha, descripción, importe) ya figura en el comprobante.

Uso:
    conda run -n mi_entorno python filtrar_movimientos_registrados.py
"""
from __future__ import annotations

import unicodedata
from pathlib import Path
from typing import Dict, Iterable, Set, Tuple

import pandas as pd

# =========================
# Configuración base
# =========================

ACCOUNT_FILES = {
    "Crédito Itaú $": Path("movimientos_pesos.xlsx"),
    "Crédito Itaú U$S": Path("movimientos_dolares.xlsx"),
    "Débito BROU $": Path("brou_detalle_movimientos.xlsx"),
    "Débito Itaú $ Gonza": Path("itau_debito_pesos.xlsx"),
}

COMPROBANTE_CANDIDATES = [Path("comprobante.xlsx"), Path("cromprobante.xlsx")]

# =========================
# Funciones auxiliares
# =========================


def normalize_date(value) -> str | None:
    """Convierte cualquier fecha a string 'YYYY-MM-DD'."""
    if pd.isna(value):
        return None
    try:
        dt = pd.to_datetime(value, dayfirst=True, errors="coerce")
    except Exception:
        return None
    if pd.isna(dt):
        return None
    return dt.date().isoformat()


def normalize_description(value) -> str:
    """Pasa a mayúsculas, sin acentos y con espacios simples."""
    if pd.isna(value):
        return ""
    text = str(value).strip()
    if not text:
        return ""
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    return " ".join(text.upper().split())


def normalize_amount(value) -> float:
    """Convierte a float con 2 decimales."""
    if pd.isna(value):
        return 0.0
    try:
        return round(float(value), 2)
    except (TypeError, ValueError):
        return 0.0


def build_keyset(df: pd.DataFrame) -> Set[Tuple[str, str, float]]:
    """Devuelve set de tuplas (fecha, descripción, importe)."""
    return {
        (row.Fecha_norm, row.Descripcion_norm, row.Importe_norm)
        for row in df.itertuples(index=False)
        if row.Fecha_norm and row.Descripcion_norm
    }


def locate_comprobante() -> Path:
    """Busca el archivo de comprobante según las opciones disponibles."""
    for candidate in COMPROBANTE_CANDIDATES:
        if candidate.exists():
            return candidate
    raise FileNotFoundError(
        "No se encontró comprobante.xlsx (ni cromprobante.xlsx) en el directorio de trabajo."
    )


def load_comprobante_keys() -> Dict[str, Set[Tuple[str, str, float]]]:
    """
    Carga el comprobante y genera un diccionario:
        cuenta -> {(fecha, descripcion, importe)}
    """
    comprobante_path = locate_comprobante()
    df = pd.read_excel(comprobante_path, header=2)
    df = df.dropna(subset=["Fecha", "Descripción", "Cuenta", "Importe"])
    df["Fecha_norm"] = df["Fecha"].apply(normalize_date)
    df["Descripcion_norm"] = df["Descripción"].apply(normalize_description)
    df["Importe_norm"] = df["Importe"].apply(normalize_amount)

    keys: Dict[str, Set[Tuple[str, str, float]]] = {}
    for cuenta, group in df.groupby("Cuenta"):
        keys[cuenta] = build_keyset(group)
    return keys


def add_normalized_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Añade columnas normalizadas necesarias para el cruce."""
    df = df.copy()
    df["Fecha_norm"] = df["Fecha"].apply(normalize_date)
    df["Descripcion_norm"] = df["Descripcion"].apply(normalize_description)
    creditos = df.get("Creditos", 0).fillna(0).astype(float)
    debitos = df.get("Debitos", 0).fillna(0).astype(float)
    df["Importe_norm"] = (creditos - debitos).round(2)
    return df


def filter_dataframe(
    df: pd.DataFrame, keys: Set[Tuple[str, str, float]]
) -> Tuple[pd.DataFrame, int]:
    """Filtra filas presentes en keys y devuelve (df_filtrado, cantidad_eliminada)."""
    if df.empty:
        return df, 0
    df_norm = add_normalized_columns(df)
    mask = ~df_norm.apply(
        lambda row: (row["Fecha_norm"], row["Descripcion_norm"], row["Importe_norm"]) in keys,
        axis=1,
    )
    removed = int((~mask).sum())
    filtered = df.loc[mask].copy()
    return filtered, removed


def process_files():
    keys_by_account = load_comprobante_keys()
    total_removed = 0

    for cuenta, path in ACCOUNT_FILES.items():
        if not path.exists():
            print(f"[ADVERTENCIA] No se encontró {path}. Se omite la cuenta '{cuenta}'.")
            continue
        keys = keys_by_account.get(cuenta, set())
        if not keys:
            print(f"[INFO] El comprobante no tiene registros para '{cuenta}'. Nada que filtrar.")
            continue

        df = pd.read_excel(path)
        filtered, removed = filter_dataframe(df, keys)
        if removed:
            filtered.to_excel(path, index=False)
            print(f"[OK] {path}: se eliminaron {removed} fila(s) coincidentes.")
        else:
            print(f"[INFO] {path}: no se encontraron coincidencias.")
        total_removed += removed

    print(f"\nTotal de filas eliminadas: {total_removed}")


if __name__ == "__main__":
    process_files()
