import pandas as pd
import os

# Archivo de entrada
input_file = "data/estadoObra.xlsx"

# Archivo de salida
output_file = "data/estadoObra_filtrado.csv"

# Columnas requeridas
columnas_requeridas = [
    "descSucursal",
    "descProyecto",
    "Actividad",
    "tipoRestriccion",
    "fechaRegistro"
]

try:

    # Leer hoja específica
    df = pd.read_excel(input_file, sheet_name="Restricciones")

    # Limpiar espacios en nombres de columnas
    df.columns = df.columns.str.strip()

    # Verificar columnas existentes
    columnas_faltantes = [c for c in columnas_requeridas if c not in df.columns]

    if columnas_faltantes:
        raise ValueError(f"Faltan columnas en el Excel: {columnas_faltantes}")

    # Mantener solo columnas requeridas
    df = df[columnas_requeridas]

    # Filtrar por sucursal
    df = df[df["descSucursal"].str.contains("BOGOTA ", na=False)]

    # Convertir fecha si es necesario
    df["fechaRegistro"] = pd.to_datetime(df["fechaRegistro"], errors="coerce")

    # Guardar resultado
    df.to_csv(output_file, index=False)

    print("Archivo procesado correctamente")

except Exception as e:
    print(f"Error procesando el archivo: {e}")
    raise
