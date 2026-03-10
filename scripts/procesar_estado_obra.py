import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime
import os

archivo_excel = "data/estadoObra.xlsx"

# --------------------------------
# Verificar hojas disponibles
# --------------------------------
xls = pd.ExcelFile(archivo_excel, engine="openpyxl")

print("Hojas disponibles:", xls.sheet_names)

if "Restricciones" not in xls.sheet_names:
    raise Exception(f"La hoja 'Restricciones' no existe. Hojas encontradas: {xls.sheet_names}")

# --------------------------------
# Leer hoja correcta
# --------------------------------
df = pd.read_excel(
    xls,
    sheet_name="Restricciones"
)

# limpiar nombres de columnas
df.columns = df.columns.str.strip()

print("Columnas detectadas:", list(df.columns))

# --------------------------------
# Validar columnas necesarias
# --------------------------------
columnas_requeridas = [
    "descSucursal",
    "descProyecto",
    "tipoRestriccion",
    "fechaRegistro"
]

for col in columnas_requeridas:
    if col not in df.columns:
        raise Exception(f"Falta la columna {col} en la hoja Restricciones")

# --------------------------------
# Normalizar datos
# --------------------------------
df["fechaRegistro"] = pd.to_datetime(df["fechaRegistro"], errors="coerce")

df["descProyecto"] = df["descProyecto"].astype(str).str.upper()

# lista completa de proyectos
proyectos = sorted(df["descProyecto"].dropna().unique())

# categorías
categorias = sorted(df["tipoRestriccion"].dropna().unique())

# --------------------------------
# Definir meses
# --------------------------------
hoy = datetime.today()

mes_actual = hoy.month
anio_actual = hoy.year

mes_pasado = mes_actual - 1 if mes_actual > 1 else 12
anio_mes_pasado = anio_actual if mes_actual > 1 else anio_actual - 1

mes_antepasado = mes_actual - 2
anio_mes_antepasado = anio_actual

if mes_antepasado <= 0:
    mes_antepasado += 12
    anio_mes_antepasado -= 1


def filtrar_mes(df, mes, anio):

    return df[
        (df["fechaRegistro"].dt.month == mes) &
        (df["fechaRegistro"].dt.year == anio)
    ]


df_mes_actual = filtrar_mes(df, mes_actual, anio_actual)
df_mes_pasado = filtrar_mes(df, mes_pasado, anio_mes_pasado)
df_mes_antepasado = filtrar_mes(df, mes_antepasado, anio_mes_antepasado)

# --------------------------------
# Preparar tabla de conteos
# --------------------------------
def preparar(df_mes):

    data = []

    for proyecto in proyectos:

        df_p = df_mes[df_mes["descProyecto"] == proyecto]

        conteos = []

        for cat in categorias:
            conteos.append((df_p["tipoRestriccion"] == cat).sum())

        data.append(conteos)

    return pd.DataFrame(data, columns=categorias, index=proyectos)


tabla_actual = preparar(df_mes_actual)
tabla_pasado = preparar(df_mes_pasado)
tabla_antepasado = preparar(df_mes_antepasado)

# --------------------------------
# CONFIGURACIÓN DEL GRÁFICO
# --------------------------------

# altura grande para evitar solapamiento
altura = max(12, len(proyectos) * 1.2)

fig, ax = plt.subplots(figsize=(18, altura))

y = range(len(proyectos))

offset = 0.38
bar_height = 0.32

bottom_actual = [0]*len(proyectos)
bottom_pasado = [0]*len(proyectos)
bottom_antepasado = [0]*len(proyectos)

# --------------------------------
# Dibujar barras apiladas
# --------------------------------
for cat in categorias:

    vals_actual = tabla_actual[cat].values
    vals_pasado = tabla_pasado[cat].values
    vals_antepasado = tabla_antepasado[cat].values

    ax.barh(
        [i-offset for i in y],
        vals_actual,
        height=bar_height,
        left=bottom_actual
    )

    ax.barh(
        y,
        vals_pasado,
        height=bar_height,
        left=bottom_pasado
    )

    ax.barh(
        [i+offset for i in y],
        vals_antepasado,
        height=bar_height,
        left=bottom_antepasado
    )

    bottom_actual = [a+b for a,b in zip(bottom_actual, vals_actual)]
    bottom_pasado = [a+b for a,b in zip(bottom_pasado, vals_pasado)]
    bottom_antepasado = [a+b for a,b in zip(bottom_antepasado, vals_antepasado)]

# --------------------------------
# Barras mínimas si no hay datos
# --------------------------------
for i in range(len(proyectos)):

    if bottom_actual[i] == 0:
        ax.barh(i-offset, 0.05, height=bar_height, color="red")

    if bottom_pasado[i] == 0:
        ax.barh(i, 0.05, height=bar_height, color="red")

    if bottom_antepasado[i] == 0:
        ax.barh(i+offset, 0.05, height=bar_height, color="red")

# --------------------------------
# Etiquetas
# --------------------------------
ax.set_yticks(list(y))
ax.set_yticklabels(proyectos, fontsize=8)

# alineación de texto
for label in ax.get_yticklabels():
    label.set_horizontalalignment('right')

ax.set_title(
    "Restricciones registradas por proyecto\nÚltimos 3 meses",
    fontsize=18,
    pad=30
)

ax.set_xlabel("Cantidad de restricciones", fontsize=12)

# ampliar margen izquierdo para nombres largos
plt.subplots_adjust(left=0.35)

plt.tight_layout()

# --------------------------------
# Guardar gráfico
# --------------------------------
os.makedirs("data", exist_ok=True)

plt.savefig(
    "data/grafico_restricciones_mes_proyecto.png",
    dpi=300,
    bbox_inches="tight"
)

print("Gráfico generado correctamente")
