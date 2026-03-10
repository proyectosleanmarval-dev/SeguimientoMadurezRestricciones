import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime

# Archivos
input_file = "data/estadoObra.xlsx"
csv_output = "data/estadoObra_filtrado.csv"
grafico_output = "data/grafico_restricciones_mes_proyecto.png"

# Columnas requeridas
columnas_requeridas = [
    "descSucursal",
    "descProyecto",
    "Actividad",
    "tipoRestriccion",
    "fechaRegistro"
]

try:

    # Leer hoja
    df = pd.read_excel(input_file, sheet_name="Restricciones")

    # Limpiar nombres de columnas
    df.columns = df.columns.str.strip()

    # Validar columnas
    columnas_faltantes = [c for c in columnas_requeridas if c not in df.columns]
    if columnas_faltantes:
        raise ValueError(f"Faltan columnas: {columnas_faltantes}")

    # Mantener columnas necesarias
    df = df[columnas_requeridas]

    # Filtrar sucursal
    df = df[df["descSucursal"].str.contains("BOGOTA ", na=False)]

    # Convertir fecha
    df["fechaRegistro"] = pd.to_datetime(df["fechaRegistro"], errors="coerce")

    # Guardar CSV filtrado
    df.to_csv(csv_output, index=False)

    print("CSV generado")

    # -----------------------------
    # FILTRO MES ACTUAL
    # -----------------------------

    hoy = datetime.now()

    df_mes = df[
        (df["fechaRegistro"].dt.month == hoy.month) &
        (df["fechaRegistro"].dt.year == hoy.year)
    ]

    # -----------------------------
    # AGRUPAR PROYECTO + RESTRICCION
    # -----------------------------

    tabla = (
        df_mes
        .groupby(["descProyecto", "tipoRestriccion"])
        .size()
        .unstack(fill_value=0)
    )

    # -----------------------------
    # GRAFICO
    # -----------------------------

    plt.figure(figsize=(12,8))

    tabla.plot(
        kind="barh",
        stacked=False
    )

    plt.title("Restricciones registradas este mes por proyecto")
    plt.xlabel("Número de registros")
    plt.ylabel("Proyecto")

    plt.tight_layout()

    plt.savefig(grafico_output)

    print("Gráfico generado")

except Exception as e:
    print(f"Error procesando archivo: {e}")
    raise
