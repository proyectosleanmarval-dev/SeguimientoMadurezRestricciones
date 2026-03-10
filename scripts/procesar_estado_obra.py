import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime
import os

archivo_excel = "data/estadoObra.xlsx"

# -----------------------------
# Leer archivo
# -----------------------------
xls = pd.ExcelFile(archivo_excel, engine="openpyxl")

if "Restricciones" not in xls.sheet_names:
    raise Exception(f"La hoja 'Restricciones' no existe. Hojas encontradas: {xls.sheet_names}")

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
# Normalizar datos
# -----------------------------
df["fechaRegistro"] = pd.to_datetime(df["fechaRegistro"], errors="coerce")

df["descProyecto"] = df["descProyecto"].astype(str).str.upper()

df["descSucursal"] = df["descSucursal"].astype(str).str.strip().str.upper()

# -----------------------------
# FILTRO SUCURSAL BOGOTA
# -----------------------------
df = df[df["descSucursal"] == "BOGOTA"]

# -----------------------------
# listas base
# -----------------------------
proyectos = sorted(df["descProyecto"].dropna().unique())
categorias = sorted(df["tipoRestriccion"].dropna().unique())

# -----------------------------
# meses
# -----------------------------
hoy = datetime.today()

mes_actual = hoy.month
anio_actual = hoy.year

mes_pasado = mes_actual-1 if mes_actual > 1 else 12
anio_pasado = anio_actual if mes_actual > 1 else anio_actual-1

mes_antepasado = mes_actual-2
anio_antepasado = anio_actual

if mes_antepasado <= 0:
    mes_antepasado += 12
    anio_antepasado -= 1

nombre_mes = {
1:"ENE",2:"FEB",3:"MAR",4:"ABR",5:"MAY",6:"JUN",
7:"JUL",8:"AGO",9:"SEP",10:"OCT",11:"NOV",12:"DIC"
}

# -----------------------------
# filtrar meses
# -----------------------------
def filtrar(df, mes, anio):
    return df[(df["fechaRegistro"].dt.month==mes) &
              (df["fechaRegistro"].dt.year==anio)]

df_actual = filtrar(df,mes_actual,anio_actual)
df_pasado = filtrar(df,mes_pasado,anio_pasado)
df_ante = filtrar(df,mes_antepasado,anio_antepasado)

# -----------------------------
# preparar tablas
# -----------------------------
def preparar(df_mes):

    data=[]

    for proyecto in proyectos:

        df_p=df_mes[df_mes["descProyecto"]==proyecto]

        conteos=[]

        for cat in categorias:
            conteos.append((df_p["tipoRestriccion"]==cat).sum())

        data.append(conteos)

    return pd.DataFrame(data,columns=categorias,index=proyectos)

t_actual=preparar(df_actual)
t_pasado=preparar(df_pasado)
t_ante=preparar(df_ante)

# -----------------------------
# gráfico
# -----------------------------
altura=max(12,len(proyectos)*1.2)

fig,ax=plt.subplots(figsize=(20,altura))

y=range(len(proyectos))

offset=0.4
bar_height=0.32

bottom_actual=[0]*len(proyectos)
bottom_pasado=[0]*len(proyectos)
bottom_ante=[0]*len(proyectos)

for cat in categorias:

    v1=t_actual[cat].values
    v2=t_pasado[cat].values
    v3=t_ante[cat].values

    ax.barh([i-offset for i in y],v1,height=bar_height,left=bottom_actual)
    ax.barh(y,v2,height=bar_height,left=bottom_pasado)
    ax.barh([i+offset for i in y],v3,height=bar_height,left=bottom_ante)

    for i,val in enumerate(v1):
        if val>0:
            ax.text(bottom_actual[i]+val/2,i-offset,str(val),
                    ha='center',va='center',fontsize=8,color="white")

    for i,val in enumerate(v2):
        if val>0:
            ax.text(bottom_pasado[i]+val/2,i,str(val),
                    ha='center',va='center',fontsize=8,color="white")

    for i,val in enumerate(v3):
        if val>0:
            ax.text(bottom_ante[i]+val/2,i+offset,str(val),
                    ha='center',va='center',fontsize=8,color="white")

    bottom_actual=[a+b for a,b in zip(bottom_actual,v1)]
    bottom_pasado=[a+b for a,b in zip(bottom_pasado,v2)]
    bottom_ante=[a+b for a,b in zip(bottom_ante,v3)]

# -----------------------------
# líneas separadoras
# -----------------------------
for i in range(len(proyectos)):
    ax.axhline(i+0.6,color='gray',linewidth=0.3)

ax.set_yticks(list(y))
ax.set_yticklabels(proyectos,fontsize=8)

for label in ax.get_yticklabels():
    label.set_horizontalalignment('right')

plt.subplots_adjust(left=0.35)

ax.set_title(
"Restricciones por Proyecto - Sucursal Bogotá\nÚltimos 3 Meses",
fontsize=18,
pad=30
)

ax.set_xlabel("Cantidad de restricciones")

plt.tight_layout()

os.makedirs("data",exist_ok=True)

plt.savefig(
"data/grafico_restricciones_mes_proyecto.png",
dpi=300,
bbox_inches="tight"
)

print("Gráfico generado correctamente (Sucursal Bogotá)")
