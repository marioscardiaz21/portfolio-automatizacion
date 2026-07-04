import pandas as pd
import re
import os
from sqlalchemy import create_engine, text
from urllib.parse import quote_plus
from datetime import date

# ==============================
# RUTA ORIGEN
# ==============================
ruta = r"\\RUTA_PROYECTO\CLARO\IVR\BASE IVR"

archivos = [
    os.path.join(ruta, f)
    for f in os.listdir(ruta)
    if f.endswith(".xlsx") and f.startswith("IVR INTERACTIVO")
]

print(f"📂 Archivos encontrados: {len(archivos)}")

# ==============================
# FUNCIÓN FECHA
# ==============================
def extraer_fecha_archivo(nombre_archivo, anio=2026):
    match = re.search(r"(\d{2})\.(\d{2})", nombre_archivo)
    if match:
        dia, mes = match.groups()
        return date(anio, int(mes), int(dia))
    return None

# ==============================
# LECTURA
# ==============================
dfs = []

for archivo in archivos:
    nombre_archivo = os.path.basename(archivo)
    fecha_carga = extraer_fecha_archivo(nombre_archivo)

    try:
        df = pd.read_excel(archivo, dtype=str)
        df["archivo_origen"] = nombre_archivo
        df["fecha_carga"] = fecha_carga

        dfs.append(df)
        print(f"✅ Cargado: {nombre_archivo}")

    except Exception as e:
        print(f"❌ Error en {nombre_archivo}: {e}")

if not dfs:
    print("⚠️ No hay archivos válidos")
    exit()

df_final = pd.concat(dfs, ignore_index=True)

print(f"📊 Total registros: {len(df_final)}")

# ==============================
# LIMPIEZAS
# ==============================
if "MONTO" in df_final.columns:
    df_final["MONTO"] = (
        df_final["MONTO"]
        .str.replace(",", ".", regex=False)
        .astype(float)
    )

if "FECHA" in df_final.columns:
    df_final["FECHA"] = pd.to_datetime(
        df_final["FECHA"],
        format="%d/%m/%Y",
        errors="coerce"
    )

# ==============================
# CONEXIÓN SQL
# ==============================
import os
usuario = os.environ["DB_USER"]
password = quote_plus(os.environ["DB_PASSWORD"])
servidor = os.environ["DB_SERVER"]

engine = create_engine(
    f"mssql+pyodbc://{usuario}:{password}@{servidor}/IBR_REPORTE"
    "?driver=ODBC+Driver+17+for+SQL+Server"
)
# ==============================
# TRUNCAR + INSERTAR
# ==============================
with engine.connect() as conn:
    try:
        conn.execute(text("TRUNCATE TABLE CLP_IVR_INTERACTIVO_CONSOLIDADO"))
        conn.commit()
        print("🧹 Tabla truncada")
    except:
        conn.execute(text("DELETE FROM CLP_IVR_INTERACTIVO_CONSOLIDADO"))
        conn.commit()

df_final.to_sql(
    "CLP_IVR_INTERACTIVO_CONSOLIDADO",
    engine,
    if_exists="append",
    index=False,
    chunksize=2000
)

print("🚀 Carga completada")

# ==============================
# EJECUTAR SP
# ==============================
with engine.connect() as conn:
    try:
        conn.execute(text("EXEC CLP_SP_KPI_IVR_ARAFAT"))
        conn.commit()
        print("⚙️ SP ejecutado correctamente")
    except Exception as e:
        print(f"❌ Error SP: {e}")
        exit()

# ==============================
# CONSULTA KPI
# ==============================
query = """
SELECT 
    CAST(FECHA_EJECUCION AS DATE) AS FECHA_EJECUCION,
    'IBR' AS EMPRESA,
    SERVICIO AS TECNOLOGIA,
    UPPER(ETAPA) AS ETAPA,
    KPI,
    VALOR
FROM CLP_KPI_IVR_ARAFAT
WHERE FECHA_EJECUCION = CAST(GETDATE() AS DATE)
"""

df = pd.read_sql(query, engine)

print(f"📊 Registros KPI: {len(df)}")

# ==============================
# MAPEO KPI
# ==============================
mapeo_kpi = {
    "CLIENTES_BENEFICIO": "Clientes Con beneficio",
    "CLIENTES_GESTIONADOS_IVR": "Clientes gestionados con IVR",
    "CLIENTES_CONTACTADOS": "Clientes Contactados",
    "CLIENTES_INFORMADOS": "Clientes Informados",
    "CLIENTES_INTERESADOS": "Clientes Interesados",
    "CLIENTES_ATENDIDOS": "Clientes Atendidos",
    "CLIENTES_ABANDONADOS": "Clientes con Abandono",
    "CLIENTES_RECUPERADOS": "Clientes Recuperados",
    "CUENTAS_RECUPERADOS": "Cuentas Recuperadas"
}

df["KPI"] = df["KPI"].map(mapeo_kpi)

# ==============================
# PIVOT
# ==============================
df_pivot = df.pivot_table(
    index=["FECHA_EJECUCION", "EMPRESA", "TECNOLOGIA", "ETAPA"],
    columns="KPI",
    values="VALOR",
    aggfunc="sum"
).reset_index()

# ==============================
# ORDEN
# ==============================
orden = [
    "FECHA_EJECUCION", "EMPRESA", "TECNOLOGIA", "ETAPA",
    "Clientes Con beneficio",
    "Clientes gestionados con IVR",
    "Clientes Contactados",
    "Clientes Informados",
    "Clientes Interesados",
    "Clientes Atendidos",
    "Clientes con Abandono",
    "Clientes Recuperados",
    "Cuentas Recuperadas"
]

df_pivot = df_pivot.reindex(columns=orden)

# ==============================
# RENOMBRE FINAL
# ==============================
df_pivot.columns = [
    "Fecha", "Empresa", "Tecnología", "Etapa",
    "Clientes Con beneficio",
    "Clientes gestionados con IVR",
    "Clientes Contactados",
    "Clientes Informados",
    "Clientes Interesados",
    "Clientes Atendidos",
    "Clientes con Abandono",
    "Clientes Recuperados",
    "Cuentas Recuperadas"
]

# ==============================
# FORMATO FECHA
# ==============================
df_pivot["Fecha"] = pd.to_datetime(df_pivot["Fecha"], errors="coerce")

# ==============================
# CONVERTIR A NUMÉRICO
# ==============================
columnas_numericas = df_pivot.columns[4:]

for col in columnas_numericas:
    df_pivot[col] = pd.to_numeric(df_pivot[col], errors="coerce").fillna(0)

# ==============================
# EXPORTAR EXCEL
# ==============================
ruta_salida = r= r"\\RUTA_PROYECTO\IVR\IVR_KPI_ORDENADO.xlsx"

with pd.ExcelWriter(ruta_salida, engine="xlsxwriter") as writer:
    df_pivot.to_excel(writer, index=False, sheet_name="KPI")

    workbook = writer.book
    worksheet = writer.sheets["KPI"]

    formato_fecha = workbook.add_format({'num_format': 'dd/mm/yyyy'})
    worksheet.set_column(0, 0, 15, formato_fecha)

print("📁 Excel generado correctamente 🚀")
