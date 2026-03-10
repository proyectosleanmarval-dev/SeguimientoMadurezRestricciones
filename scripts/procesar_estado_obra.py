import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime
import os
import calendar
import numpy as np

archivo_excel = "data/estadoObra.xlsx"

# -----------------------------
# Leer archivo
# -----------------------------
xls = pd.ExcelFile(archivo_excel, engine="openpyxl")

if "Restricciones" not in xls.sheet_names:
    raise Exception("La hoja 'Restricciones' no existe")

df = pd.read_excel(xls, sheet_name="Restricciones")
df.columns = df.columns.str.strip()

# -----------------------------
# Validar columnas
# -----------------------------
columnas_requeridas = [
    "descSucursal",
    "descProyecto",
    "tipoRestriccion",
    "fechaRegistro"
]

for col in columnas_requeridas:
    if col not in df.columns:
        raise Exception(f"Falta la columna {col}")

# -----------------------------
# Normalizar
# -----------------------------
df["fechaRegistro"] = pd.to_datetime(df["fechaRegistro"], errors="coerce")
df["descProyecto"] = df["descProyecto"].astype(str).str.upper()
df["descSucursal"] = df["descSucursal"].astype(str).str.strip().str.upper()

# -----------------------------
# FILTRO BOGOTA
# -----------------------------
df = df[df["descSucursal"] == "BOGOTA"]

# -----------------------------
# listas base
# -----------------------------
proyectos = sorted(df["descProyecto"].dropna().unique())
categorias = sorted(df["tipoRestriccion"].dropna().unique())

# colores - paleta más suave y distinguishable
colores = plt.cm.Set3.colors  # Cambio a una paleta más suave
color_categoria = {cat: colores[i % len(colores)] for i, cat in enumerate(categorias)}

# -----------------------------
# meses
# -----------------------------
hoy = datetime.today()

mes_actual = hoy.month
anio_actual = hoy.year

mes_pasado = mes_actual-1 if mes_actual > 1 else 12
anio_pasado = anio_actual if mes_actual > 1 else anio_actual-1

mes_ante = mes_actual-2
anio_ante = anio_actual

if mes_ante <= 0:
    mes_ante += 12
    anio_ante -= 1

mes_actual_txt = calendar.month_abbr[mes_actual].upper()
mes_pasado_txt = calendar.month_abbr[mes_pasado].upper()
mes_ante_txt = calendar.month_abbr[mes_ante].upper()

# -----------------------------
# filtro meses
# -----------------------------
def filtrar(df, mes, anio):
    return df[
        (df["fechaRegistro"].dt.month == mes) &
        (df["fechaRegistro"].dt.year == anio)
    ]

df_actual = filtrar(df, mes_actual, anio_actual)
df_pasado = filtrar(df, mes_pasado, anio_pasado)
df_ante = filtrar(df, mes_ante, anio_ante)

# -----------------------------
# preparar tabla
# -----------------------------
def preparar(df_mes):

    data = []

    for proyecto in proyectos:

        df_p = df_mes[df_mes["descProyecto"] == proyecto]

        conteos = []
        for cat in categorias:
            conteos.append((df_p["tipoRestriccion"] == cat).sum())

        data.append(conteos)

    return pd.DataFrame(data, columns=categorias, index=proyectos)

t_actual = preparar(df_actual)
t_pasado = preparar(df_pasado)
t_ante = preparar(df_ante)

# -----------------------------
# gráfico
# -----------------------------
altura = max(10, len(proyectos) * 1.5)  # Ajuste de altura

fig, ax = plt.subplots(figsize=(22, altura))

y = np.arange(len(proyectos))

# Ajuste de offsets para evitar superposición de etiquetas
offset = 0.35  # Reducido ligeramente
bar_height = 0.25  # Reducido para más espacio

bottom_actual = [0]*len(proyectos)
bottom_pasado = [0]*len(proyectos)
bottom_ante = [0]*len(proyectos)

# -----------------------------
# barras apiladas con etiquetas de categoría
# -----------------------------
for cat in categorias:
    color = color_categoria[cat]

    v1 = t_actual[cat].values
    v2 = t_pasado[cat].values
    v3 = t_ante[cat].values

    bars1 = ax.barh([i-offset for i in y], v1, height=bar_height,
                    left=bottom_actual, color=color, edgecolor='white', linewidth=0.5)

    bars2 = ax.barh(y, v2, height=bar_height,
                    left=bottom_pasado, color=color, edgecolor='white', linewidth=0.5)

    bars3 = ax.barh([i+offset for i in y], v3, height=bar_height,
                    left=bottom_ante, color=color, edgecolor='white', linewidth=0.5)

    # Etiquetas dentro de segmentos: número + abreviatura de categoría
    for bar, val in zip(bars1, v1):
        if val > 0:
            # Mostrar número y abreviatura de categoría
            cat_abbr = ''.join([word[0] for word in cat.split()])[:3]  # Abreviatura de 3 letras
            ax.text(bar.get_x() + bar.get_width()/2,
                    bar.get_y() + bar.get_height()/2,
                    f'{int(val)}\n{cat_abbr}', 
                    ha='center', va='center', fontsize=6, linespacing=0.8,
                    bbox=dict(boxstyle="round,pad=0.2", facecolor='white', alpha=0.7, edgecolor='none'))

    for bar, val in zip(bars2, v2):
        if val > 0:
            cat_abbr = ''.join([word[0] for word in cat.split()])[:3]
            ax.text(bar.get_x() + bar.get_width()/2,
                    bar.get_y() + bar.get_height()/2,
                    f'{int(val)}\n{cat_abbr}', 
                    ha='center', va='center', fontsize=6, linespacing=0.8,
                    bbox=dict(boxstyle="round,pad=0.2", facecolor='white', alpha=0.7, edgecolor='none'))

    for bar, val in zip(bars3, v3):
        if val > 0:
            cat_abbr = ''.join([word[0] for word in cat.split()])[:3]
            ax.text(bar.get_x() + bar.get_width()/2,
                    bar.get_y() + bar.get_height()/2,
                    f'{int(val)}\n{cat_abbr}', 
                    ha='center', va='center', fontsize=6, linespacing=0.8,
                    bbox=dict(boxstyle="round,pad=0.2", facecolor='white', alpha=0.7, edgecolor='none'))

    bottom_actual = [a+b for a,b in zip(bottom_actual, v1)]
    bottom_pasado = [a+b for a,b in zip(bottom_pasado, v2)]
    bottom_ante = [a+b for a,b in zip(bottom_ante, v3)]

# -----------------------------
# títulos de mes en cada barra (ahora en los extremos para evitar superposición)
# -----------------------------
for i in range(len(proyectos)):
    total1 = t_actual.iloc[i].sum()
    total2 = t_pasado.iloc[i].sum()
    total3 = t_ante.iloc[i].sum()
    
    max_width = max(total1, total2, total3)
    
    # Posicionar etiquetas de mes en los extremos de las barras
    if total1 > 0:
        ax.text(max_width + 0.5, i-offset, mes_actual_txt,
                ha='left', va='center', fontsize=8, fontweight='bold',
                bbox=dict(boxstyle="round,pad=0.2", facecolor='lightblue', alpha=0.8))

    if total2 > 0:
        ax.text(max_width + 0.5, i, mes_pasado_txt,
                ha='left', va='center', fontsize=8, fontweight='bold',
                bbox=dict(boxstyle="round,pad=0.2", facecolor='lightgreen', alpha=0.8))

    if total3 > 0:
        ax.text(max_width + 0.5, i+offset, mes_ante_txt,
                ha='left', va='center', fontsize=8, fontweight='bold',
                bbox=dict(boxstyle="round,pad=0.2", facecolor='lightcoral', alpha=0.8))

# -----------------------------
# separadores proyecto (más sutiles)
# -----------------------------
for i in range(len(proyectos)):
    ax.axhline(i+0.5, color='gray', linewidth=0.2, linestyle='--', alpha=0.5)

# -----------------------------
# etiquetas proyectos con mejor formato
# -----------------------------
ax.set_yticks(list(y))
ax.set_yticklabels(proyectos, fontsize=9, fontweight='bold')

# Ajustar límites del eje x para dar espacio a las etiquetas de mes
max_total = max(t_actual.sum(axis=1).max(), 
                t_pasado.sum(axis=1).max(), 
                t_ante.sum(axis=1).max())
ax.set_xlim(0, max_total + 3)

plt.subplots_adjust(left=0.35)  # Más espacio para nombres largos

# -----------------------------
# título y subtítulo
# -----------------------------
ax.set_title(
    "RESTRICCIONES POR PROYECTO - BOGOTÁ\nComparación últimos 3 meses",
    fontsize=16,
    fontweight='bold',
    pad=20
)

ax.set_xlabel("Cantidad de restricciones", fontsize=11, fontweight='bold')

# Eliminar grid vertical que pueda causar ruido
ax.grid(axis='x', alpha=0.2)

# -----------------------------
# Guardar
# -----------------------------
os.makedirs("data", exist_ok=True)

plt.savefig(
    "data/grafico_restricciones_mes_proyecto_mejorado.png",
    dpi=300,
    bbox_inches="tight",
    facecolor='white',
    edgecolor='none'
)

print("Gráfico generado correctamente con las mejoras visuales")
