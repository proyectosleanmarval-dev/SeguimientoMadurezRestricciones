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

    # Leer Excel
    df = pd.read_excel(input_file, sheet_name="Restricciones")

    # Limpiar columnas
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
    # FILTRAR MES ACTUAL
    # -----------------------------

    hoy = datetime.now()

    df_mes = df[
        (df["fechaRegistro"].dt.month == hoy.month) &
        (df["fechaRegistro"].dt.year == hoy.year)
    ]

    # -----------------------------
    # TABLA PROYECTO vs RESTRICCION
    # -----------------------------

    tabla = (
        df_mes
        .groupby(["descProyecto", "tipoRestriccion"])
        .size()
        .unstack(fill_value=0)
    )

    # Ordenar proyectos por total
    tabla["Total"] = tabla.sum(axis=1)
    tabla = tabla.sort_values("Total")
    tabla = tabla.drop(columns="Total")

    # -----------------------------
    # GRAFICO APILADO
    # -----------------------------

    fig, ax = plt.subplots(figsize=(14, 10))

    tabla.plot(
        kind="barh",
        stacked=True,
        ax=ax
    )

    ax.set_title(
        "Restricciones registradas en el mes actual por proyecto",
        fontsize=16,
        pad=20
    )

    ax.set_xlabel("Número de registros")
    ax.set_ylabel("Proyecto")

    plt.legend(
        title="Tipo de restricción",
        bbox_to_anchor=(1.02, 1),
        loc="upper left"
    )

    plt.subplots_adjust(top=0.90, right=0.75)

    plt.tight_layout()

    plt.savefig(grafico_output, dpi=300)

    print("Gráfico generado")

except Exception as e:
    print(f"Error procesando archivo: {e}")
    raise
