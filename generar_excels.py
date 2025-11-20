import pandas as pd
import requests
from datetime import datetime
from functools import lru_cache


# Nombres de archivo (podés cambiarlos si querés)
INPUT_CSV = "movimientos.csv"
OUTPUT_XLSX_PESOS = "movimientos_pesos.xlsx"
OUTPUT_XLSX_DOLARES = "movimientos_dolares.xlsx"


def parse_importe(value):
    """Convierte importes con formato '1.234,56' o '-15,68' a float."""
    if pd.isna(value):
        return 0.0

    s = str(value).strip()
    if not s:
        return 0.0

    # Mantener el signo si lo hubiera
    sign = -1 if s.startswith("-") else 1
    # Sacamos el signo del string para limpiar el resto
    s = s.replace("-", "")

    # El CSV viene con miles con punto y decimales con coma
    # Ej: "1.300,00" -> "1300.00"
    s = s.replace(".", "").replace(",", ".")

    try:
        return sign * float(s)
    except ValueError:
        # Si algo raro viene en el campo, lo tratamos como 0
        return 0.0


def parse_fecha_texto(fecha_str):
    """
    Recibe 'DD/MM/AA' o 'DD/MM/AAAA' y devuelve 'DD/MM/AAAA' (string).
    Esta columna se va a exportar como texto al Excel.
    """
    fecha_str = str(fecha_str).strip()
    if not fecha_str:
        return ""

    for fmt in ("%d/%m/%y", "%d/%m/%Y"):
        try:
            dt = datetime.strptime(fecha_str, fmt).date()
            return dt.strftime("%d/%m/%Y")
        except ValueError:
            continue

    # Si no se puede parsear, devolvemos el original
    return fecha_str


EXCHANGE_RATE_API_KEY = "112085d534849519e2ad9806"


@lru_cache(maxsize=None)
def get_usd_rate_uyu(fecha_texto):
    """
    Obtiene la cotización USD->UYU (1 USD en Pesos Uruguayos) usando la
    última tasa publicada por la API (fecha_texto se mantiene por compatibilidad).
    """
    url = f"https://v6.exchangerate-api.com/v6/{EXCHANGE_RATE_API_KEY}/latest/USD"

    try:
        resp = requests.get(url, timeout=10)
        resp.raise_for_status()
        data = resp.json()
        rate = data["conversion_rates"]["UYU"]
        return float(rate)
    except Exception as e:
        print(f"[ADVERTENCIA] No se pudo obtener cotización USD/UYU para {fecha_texto}: {e}")
        # En caso de error, devolvemos None para forzar ingreso manual
        return None


def solicitar_cotizacion_manual():
    """
    Solicita al usuario una cotización USD/UYU.
    Devuelve None si el usuario decide no ingresar un valor.
    """
    while True:
        valor = input("Ingrese la cotización USD/UYU a usar (Enter para cancelar): ").strip()
        if valor == "":
            return None
        valor = valor.replace(",", ".")
        try:
            return float(valor)
        except ValueError:
            print("Valor inválido. Ingrese un número (ejemplo: 39.5).")


def construir_tabla(df):
    """
    A partir del DataFrame original del CSV construye la tabla con:
    Fecha, Descripcion, Creditos, Debitos, Cotizacion
    """
    df = df.copy()

    # Filtramos filas con fecha y nombre válidos (son obligatorios)
    df = df.dropna(subset=["Fecha", "Nombre"])

    # Fecha en texto DD/MM/AAAA
    df["Fecha_texto"] = df["Fecha"].apply(parse_fecha_texto)

    # Importe numérico
    df["Importe_num"] = df["Importe"].apply(parse_importe)

    # Créditos (ingresos) y Débitos (egresos)
    df["Creditos"] = df["Importe_num"].apply(lambda x: abs(x) if x < 0 else 0.0)
    df["Debitos"] = df["Importe_num"].apply(lambda x: x if x > 0 else 0.0)

    # Armamos la tabla final en el orden pedido
    tabla = pd.DataFrame()
    tabla["Fecha"] = df["Fecha_texto"]              # Columna 1: texto DD/MM/AAAA
    tabla["Descripcion"] = df["Nombre"].astype(str) # Columna 2: texto
    tabla["Creditos"] = df["Creditos"]              # Columna 3: numérico
    tabla["Debitos"] = df["Debitos"]                # Columna 4: numérico
    # Columna 5 (Cotizacion) se completa después según la moneda

    return tabla, df


def main():
    # Leer CSV original (con el mismo formato que el que adjuntaste)
    df = pd.read_csv(INPUT_CSV, encoding="latin1")
    # Eliminar movimientos que sean recibos de pago
    mask_recibo = df["Nombre"].astype(str).str.contains("RECIBO DE PAGO", na=False, regex=False)
    df = df[~mask_recibo]
    # Separar por moneda
    df_pesos = df[df["Moneda"] == "Pesos"].copy()
    df_dolares = df[df["Moneda"] == "Dólares"].copy()

    # ========= Archivo de Pesos =========
    tabla_pesos, _ = construir_tabla(df_pesos)
    # Moneda nacional: cotización 0
    tabla_pesos["Cotizacion"] = 0.0
    # Exportar a Excel
    tabla_pesos.to_excel(OUTPUT_XLSX_PESOS, index=False)
    print(f"Archivo generado: {OUTPUT_XLSX_PESOS}")

    # ========= Archivo de Dólares =========
    if not df_dolares.empty:
        tabla_dolares, _ = construir_tabla(df_dolares)
        # Pedimos la cotización a la API una sola vez
        cotizacion_api = get_usd_rate_uyu(tabla_dolares["Fecha"].iloc[0])
        if cotizacion_api is None:
            print("La cotización automática falló. Ingrese un valor manual.")
            cotizacion_manual = solicitar_cotizacion_manual()
            if cotizacion_manual is None:
                print("No se ingresó cotización manual; se usará 0.0.")
                cotizacion_manual = 0.0
            tabla_dolares["Cotizacion"] = cotizacion_manual
        else:
            tabla_dolares["Cotizacion"] = cotizacion_api
        tabla_dolares.to_excel(OUTPUT_XLSX_DOLARES, index=False)
        print(f"Archivo generado: {OUTPUT_XLSX_DOLARES}")
    else:
        print("No se encontraron movimientos en dólares; no se genera archivo en USD.")


if __name__ == "__main__":
    main()
