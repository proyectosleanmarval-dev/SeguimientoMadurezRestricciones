import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime

# -----------------------------
# ARCHIVOS
# -----------------------------

input_file = "data/estadoObra.xlsx"
csv_output = "data/estadoObra_filtrado.csv"
grafico_output = "data/grafico_restricciones_3_meses.png"
proyectos_sin_output = "data/proyectos_sin_registros.csv"

# -----------------------------
# COLUMNAS REQUERIDAS
# -----------------------------

columnas_requeridas = [
    "descSucursal",
    "descProyecto",
    "Actividad",
    "tipoRestriccion",
    "fechaRegistro"
]

# -----------------------------
# LISTADO COMPLETO PROYECTOS BOGOTA
# -----------------------------

proyectos_bogota = [
"Abadia San Rafael",
"Aragon",
"Baviera",
"Burdeos Ciudad La Salle",
"Burgos Castilla Reservado",
"Cadiz",
"Chamonix Ciudad La Salle",
"El Campo",
"El Castell Iberia Reservado",
"Gallet Ciudad La Salle",
"Izola Zentral-citizzen",
"La Almeria Alsacia Reservado",
"La Cabrera",
"La Cruz",
"La Palma",
"La Peña",
"La Scala",
"La Terra Alsacia Reservado",
"Linz E1",
"Loira Ciudad La Salle",
"Lorca",
"Lyon 2 Ciudad La Salle",
"Lyon Ciudad La Salle",
"Metropoli 30",
"Montpellier Ciudad La Salle",
"Paseo De Sevilla",
"Paseo San Rafael",
"Peñazul El Poblado",
"Peñon De Alicante",
"Provenza Prestige",
"Saint Michel Ciudad La Salle",
"San Lucas La Quinta",
"San Mateo - La Quinta",
"San Sebastian La Quinta",
"San Simon La Quinta",
"Urb Externo Ciudad La Salle",
"Urbanismo Externo Lote 5 Banca",
"Urbanismo Tres Quebradas",
"Verdania Bosques De Guaymaral",
"Zuri - Zentral"
]

try:

    # -----------------------------
    # LEER EXCEL
    # -----------------------------

    df = pd.read_excel(input_file, sheet_name="Restricciones")

    # limpiar nombres columnas
    df.columns = df.columns.str.strip()

    # validar columnas
    columnas_faltantes = [c for c in columnas_requeridas if c not in df.columns]

    if columnas_faltantes:
        raise ValueError(f"Faltan columnas en el Excel: {columnas_faltantes}")

    # mantener columnas necesarias
    df = df[columnas_requeridas]

    # -----------------------------
    # FILTRAR SUCURSAL BOGOTA
    # -----------------------------

    df = df[df["descSucursal"].str.contains("BOGOTA ", na=False)]

    # -----------------------------
    # CONVERTIR FECHA
    # -----------------------------

    df["fechaRegistro"] = pd.to_datetime(df["fechaRegistro"], errors="coerce")

    # -----------------------------
    # GUARDAR CSV FILTRADO
    # -----------------------------

    df.to_csv(csv_output, index=False)

    print("CSV filtrado generado")

    # -----------------------------
    # IDENTIFICAR PROYECTOS SIN REGISTROS
    # -----------------------------

    proyectos_excel = set(df["descProyecto"].dropna().unique())

    proyectos_sin_registros = sorted(
        set(proyectos_bogota) - proyectos_excel
    )

    df_sin = pd.DataFrame({
        "proyecto_sin_registros": proyectos_sin_registros
    })

    df_sin.to_csv(proyectos_sin_output, index=False)

    print("Listado de proyectos sin registros generado")

    # -----------------------------
    # PREPARAR DATOS PARA GRAFICO
    # -----------------------------

    df["mes"] = df["fechaRegistro"].dt.to_period("M")

    hoy = datetime.now()

    mes_actual = pd.Period(hoy.strftime("%Y-%m"))
    mes_anterior = mes_actual - 1
    mes_antepasado = mes_actual - 2

    meses_interes = [mes_antepasado, mes_anterior, mes_actual]

    df_3m = df[df["mes"].isin(meses_interes)]

    # -----------------------------
    # AGRUPAR PROYECTO / MES / RESTRICCION
    # -----------------------------

    tabla = (
        df_3m
        .groupby(["descProyecto", "mes", "tipoRestriccion"])
        .size()
        .unstack(fill_value=0)
    )

    tabla = tabla.sort_index()

    # etiquetas eje Y
    tabla.index = [
        f"{proyecto} - {str(mes)}"
        for proyecto, mes in tabla.index
    ]

    # -----------------------------
    # GENERAR GRAFICO
    # -----------------------------

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

    plt.subplots_adjust(
        top=0.9,
        right=0.75
    )

    plt.tight_layout()

    plt.savefig(
        grafico_output,
        dpi=300
    )

    print("Gráfico generado correctamente")

except Exception as e:

    print(f"Error procesando el archivo: {e}")
    raise
