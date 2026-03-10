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
# LISTADO COMPLETO PROYECTOS BOGOTA
# -----------------------------

proyectos_bogota = [
"ABADIA SAN RAFAEL","ARAGON","BAVIERA","BURDEOS CIUDAD LA SALLE",
"BURGOS CASTILLA RESERVADO","CADIZ","CHAMONIX CIUDAD LA SALLE",
"EL CAMPO","EL CASTELL IBERIA RESERVADO","GALLET CIUDAD LA SALLE",
"IZOLA ZENTRAL-CITIZZEN","LA ALMERIA ALSACIA RESERVADO","LA CABRERA",
"LA CRUZ","LA PALMA","LA PEÑA","LA SCALA","LA TERRA ALSACIA RESERVADO",
"LINZ E1","LOIRA CIUDAD LA SALLE","LORCA","LYON 2 CIUDAD LA SALLE",
"LYON CIUDAD LA SALLE","METROPOLI 30","MONTPELLIER CIUDAD LA SALLE",
"PASEO DE SEVILLA","PASEO SAN RAFAEL","PEÑAZUL EL POBLADO",
"PEÑON DE ALICANTE","PROVENZA PRESTIGE","SAINT MICHEL CIUDAD LA SALLE",
"SAN LUCAS LA QUINTA","SAN MATEO - LA QUINTA","SAN SEBASTIAN LA QUINTA",
"SAN SIMON LA QUINTA","URB EXTERNO CIUDAD LA SALLE",
"URBANISMO EXTERNO LOTE 5 BANCA","URBANISMO TRES QUEBRADAS",
"VERDANIA BOSQUES DE GUAYMARAL","ZURI - ZENTRAL"
]

try:

    # -----------------------------
    # LEER EXCEL
    # -----------------------------

    df = pd.read_excel(input_file, header=0)

    # limpiar encabezados
    df.columns = (
        df.columns
        .astype(str)
        .str.strip()
        .str.replace("\n","")
        .str.replace("\r","")
    )

    print("Columnas detectadas:", list(df.columns))

    columnas_requeridas = [
        "descSucursal",
        "descProyecto",
        "Actividad",
        "tipoRestriccion",
        "fechaRegistro"
    ]

    for col in columnas_requeridas:
        if col not in df.columns:
            raise Exception(f"Falta la columna {col} en el Excel")

    df = df[columnas_requeridas]

    # -----------------------------
    # FILTRAR SUCURSAL BOGOTA
    # -----------------------------

    df = df[
        df["descSucursal"]
        .astype(str)
        .str.contains("BOGOTA", case=False, na=False)
    ]

    # -----------------------------
    # NORMALIZAR PROYECTOS
    # -----------------------------

    df["descProyecto"] = (
        df["descProyecto"]
        .astype(str)
        .str.upper()
        .str.strip()
        .str.replace(r"\s+", " ", regex=True)
    )

    # -----------------------------
    # FECHA
    # -----------------------------

    df["fechaRegistro"] = pd.to_datetime(
        df["fechaRegistro"],
        errors="coerce"
    )

    # -----------------------------
    # GUARDAR CSV
    # -----------------------------

    df.to_csv(csv_output, index=False)

    print("CSV filtrado generado")

    # -----------------------------
    # PROYECTOS SIN REGISTROS
    # -----------------------------

    proyectos_excel = set(df["descProyecto"].dropna().unique())

    proyectos_sin_registros = sorted(
        set(proyectos_bogota) - proyectos_excel
    )

    pd.DataFrame({
        "proyecto_sin_registros": proyectos_sin_registros
    }).to_csv(proyectos_sin_output, index=False)

    print("Listado proyectos sin registros generado")

    # -----------------------------
    # CALCULAR MESES
    # -----------------------------

    df["mes"] = df["fechaRegistro"].dt.to_period("M")

    hoy = datetime.now()

    mes_actual = pd.Period(hoy.strftime("%Y-%m"))
    mes_anterior = mes_actual - 1
    mes_antepasado = mes_actual - 2

    meses_interes = [
        mes_antepasado,
        mes_anterior,
        mes_actual
    ]

    # -----------------------------
    # BASE COMPLETA PROYECTOS × MESES
    # -----------------------------

    base_completa = pd.MultiIndex.from_product(
        [proyectos_bogota, meses_interes],
        names=["descProyecto","mes"]
    )

    df_3m = df[df["mes"].isin(meses_interes)]

    tabla = (
        df_3m
        .groupby(["descProyecto","mes","tipoRestriccion"])
        .size()
        .unstack(fill_value=0)
    )

    tabla = tabla.reindex(base_completa, fill_value=0)

    # convertir a float para permitir valores pequeños
    tabla = tabla.astype(float)

    # -----------------------------
    # INDICADOR SIN REGISTROS
    # -----------------------------

    tabla["SIN REGISTROS"] = 0

    sin_datos = tabla.sum(axis=1) == 0

    tabla.loc[sin_datos,"SIN REGISTROS"] = 0.01

    # etiquetas
    tabla.index = [
        f"{proyecto} - {mes}"
        for proyecto, mes in tabla.index
    ]

    # -----------------------------
    # GRAFICO
    # -----------------------------

    colores = list(plt.cm.tab20.colors)

    if "SIN REGISTROS" in tabla.columns:
        colores = colores[:len(tabla.columns)-1] + ["red"]

    fig, ax = plt.subplots(figsize=(15,18))

    tabla.plot(
        kind="barh",
        stacked=True,
        ax=ax,
        color=colores
    )

    ax.set_title(
        "Restricciones por proyecto - últimos 3 meses",
        fontsize=16
    )

    ax.set_xlabel("Número de registros")
    ax.set_ylabel("Proyecto / Mes")

    plt.legend(
        title="Tipo restricción",
        bbox_to_anchor=(1.02,1),
        loc="upper left"
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
