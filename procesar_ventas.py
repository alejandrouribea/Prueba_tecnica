import sys
from pathlib import Path

import pandas as pd

ARCHIVO_ENTRADA = "datos_ventas.xlsx"
ARCHIVO_SALIDA = "resumen_ventas.xlsx"
ANIO_OBJETIVO = 2023
COLUMNAS_REQUERIDAS = [
    "ID_Venta",
    "Fecha",
    "Producto",
    "Cantidad",
    "Precio_Unitario",
    "Total_Venta",
    "Vendedor",
]


def cargar_datos(ruta: str) -> pd.DataFrame:
    path = Path(ruta)
    if not path.exists():
        raise FileNotFoundError(f"No se encontró el archivo de entrada: {ruta}")
    df = pd.read_excel(path)
    faltantes = [c for c in COLUMNAS_REQUERIDAS if c not in df.columns]
    if faltantes:
        raise ValueError(f"Faltan columnas requeridas en el Excel: {faltantes}")
    return df


def convertir_fecha(df: pd.DataFrame) -> pd.DataFrame:
    df["Fecha"] = pd.to_datetime(df["Fecha"], errors="coerce")
    invalidas = df["Fecha"].isna().sum()
    if invalidas:
        print(f"[Fecha] Se descartaron {invalidas} filas con fecha inválida o no interpretable.")
        df = df.dropna(subset=["Fecha"])
    return df


def normalizar_vendedor(df: pd.DataFrame) -> pd.DataFrame:
    vendedor = df["Vendedor"].apply(lambda v: str(v).strip() if pd.notna(v) else "")
    vacio = vendedor == ""
    if vacio.any():
        print(f"[Vendedor] Se etiquetaron {vacio.sum()} filas sin vendedor como 'Sin Vendedor'.")
    vendedor[vacio] = "Sin Vendedor"
    df["Vendedor"] = vendedor
    return df


def normalizar_numero(valor):
    if pd.isna(valor):
        return None
    if isinstance(valor, (int, float)):
        return valor

    texto = str(valor).strip()
    if texto == "":
        return None

    comas = texto.count(",")
    puntos = texto.count(".")

    if comas and puntos:
        if texto.rfind(",") > texto.rfind("."):
            texto = texto.replace(".", "").replace(",", ".")
        else:
            texto = texto.replace(",", "")
    elif comas:
        texto = texto.replace(",", ".")
    elif puntos > 1:
        texto = texto.replace(".", "")
    elif puntos == 1:
        entero, _, decimales = texto.rpartition(".")
        if len(decimales) == 3:
            texto = entero + decimales

    return texto


def validar_numericos(df: pd.DataFrame) -> pd.DataFrame:
    for columna in ("Cantidad", "Precio_Unitario"):
        df[columna] = pd.to_numeric(df[columna].map(normalizar_numero), errors="coerce")

    antes = len(df)
    df = df.dropna(subset=["Cantidad", "Precio_Unitario"])
    no_numericas = antes - len(df)
    if no_numericas:
        print(f"[Cantidad/Precio_Unitario] Se descartaron {no_numericas} filas con valores no numéricos.")

    antes = len(df)
    df = df[(df["Cantidad"] > 0) & (df["Precio_Unitario"] > 0)]
    no_positivas = antes - len(df)
    if no_positivas:
        print(f"[Cantidad/Precio_Unitario] Se descartaron {no_positivas} filas con valores menores o iguales a 0.")

    return df


def completar_total_venta(df: pd.DataFrame) -> pd.DataFrame:
    original_nulo = df["Total_Venta"].isna()
    normalizado = pd.to_numeric(df["Total_Venta"].map(normalizar_numero), errors="coerce")

    no_parseable = (~original_nulo) & normalizado.isna()
    if no_parseable.any():
        print(f"[Total_Venta] Advertencia: {no_parseable.sum()} filas tenían un valor no numérico y se recalcularán.")

    df["Total_Venta"] = normalizado
    calculado = df["Cantidad"] * df["Precio_Unitario"]

    a_completar = df["Total_Venta"].isna()
    if a_completar.any():
        print(f"[Total_Venta] Se completaron {a_completar.sum()} valores faltantes (Cantidad x Precio_Unitario).")
    df.loc[a_completar, "Total_Venta"] = calculado[a_completar]

    inconsistentes = (~a_completar) & (df["Total_Venta"] != calculado)
    if inconsistentes.any():
        print(f"[Total_Venta] Advertencia: {inconsistentes.sum()} filas no coinciden con Cantidad x Precio_Unitario.")

    return df


def filtrar_por_año(df: pd.DataFrame, año: int) -> pd.DataFrame:
    filtrado = df[df["Fecha"].dt.year == año].copy()
    descartadas = len(df) - len(filtrado)
    if descartadas:
        print(f"[Filtro año] Se descartaron {descartadas} filas fuera del año {año}.")
    if filtrado.empty:
        raise ValueError(f"No quedan filas para el año {año} después de aplicar el filtro.")
    return filtrado


def agregar_columna_mes(df: pd.DataFrame) -> pd.DataFrame:
    df["Mes"] = df["Fecha"].dt.month
    return df


def calcular_resumen_por_vendedor(df: pd.DataFrame) -> pd.DataFrame:
    return (
        df.groupby("Vendedor", as_index=False)["Total_Venta"]
        .sum()
        .sort_values("Total_Venta", ascending=False)
        .reset_index(drop=True)
    )


def calcular_resumen_mensual(df: pd.DataFrame) -> pd.DataFrame:
    return (
        df.groupby("Mes", as_index=False)["Total_Venta"]
        .sum()
        .sort_values("Mes")
        .reset_index(drop=True)
    )


def exportar_resumen(resumen_vendedor: pd.DataFrame, resumen_mensual: pd.DataFrame, ruta_salida: str) -> None:
    with pd.ExcelWriter(ruta_salida, engine="openpyxl") as writer:
        resumen_vendedor.to_excel(writer, sheet_name="Resumen_Ventas", index=False)
        resumen_mensual.to_excel(writer, sheet_name="Ventas_Mensuales", index=False)


def main() -> None:
    try:
        df = cargar_datos(ARCHIVO_ENTRADA)
        df = convertir_fecha(df)
        df = normalizar_vendedor(df)
        df = validar_numericos(df)
        df = completar_total_venta(df)
        df = filtrar_por_año(df, ANIO_OBJETIVO)
        df = agregar_columna_mes(df)

        resumen_vendedor = calcular_resumen_por_vendedor(df)
        resumen_mensual = calcular_resumen_mensual(df)

        exportar_resumen(resumen_vendedor, resumen_mensual, ARCHIVO_SALIDA)

        print(f"\nArchivo generado: {ARCHIVO_SALIDA}")
        print("\nResumen_Ventas:")
        print(resumen_vendedor.to_string(index=False))
        print("\nVentas_Mensuales:")
        print(resumen_mensual.to_string(index=False))

    except (FileNotFoundError, ValueError) as error:
        print(f"Error: {error}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
