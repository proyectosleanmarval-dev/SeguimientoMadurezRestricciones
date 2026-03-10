import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime

# Archivos
input_file = "data/estadoObra.xlsx"
csv_output = "data/estadoObra_filtrado.csv"
grafico_output = "data/grafico_restricciones_3_meses.png"

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

    # Guardar CSV
    df.to_csv(csv_output, index=False)

    print("CSV generado")

    # --------------------------
    # FILTRAR ULTIMOS 3 MESES
    # --------------------------

    hoy = datetime.now()

    df["mes"] = df["fechaRegistro"].dt.to_period("M")

    mes_actual = hoy.strftime("%Y-%m")
    mes_anterior = (pd.Period(mes_actual) - 1).strftime("%Y-%m")
    mes_antepasado = (pd.Period(mes_actual) - 2).strftime("%Y-%m")

    meses_interes = [mes_antepasado, mes_anterior, mes_actual]

    df_3m = df[df["mes"].astype(str).isin(meses_interes)]

    # --------------------------
    # AGRUPAR
    # --------------------------

    tabla = (
        df_3m
        .groupby(["descProyecto", "mes", "tipoRestriccion"])
        .size()
        .unstack(fill_value=0)
    )

    # Ordenar
    tabla = tabla.sort_index()

    # --------------------------
    # ETIQUETAS PARA EL EJE Y
    # --------------------------

    tabla.index = [f"{proj} - {mes}" for proj, mes in tabla.index]

    # --------------------------
    # GRAFICO
    # --------------------------

    fig, ax = plt.subplots(figsize=(14, 12))

    tabla.plot(
        kind="barh",
        stacked=True,
        ax=ax
    )

    ax.set_title(
        "Restricciones por proyecto en los últimos 3 meses",
        fontsize=16,
        pad=20
    )

    ax.set_xlabel("Número de registros")
    ax.set_ylabel("Proyecto y mes")

    plt.legend(
        title="Tipo de restricción",
        bbox_to_anchor=(1.02, 1),
        loc="upper left"
    )

    plt.subplots_adjust(top=0.9, right=0.75)

    plt.tight_layout()

    plt.savefig(grafico_output, dpi=300)

    print("Gráfico generado")

except Exception as e:
    print(f"Error procesando archivo: {e}")
    raise
