import streamlit as st
import pandas as pd
import plotly.express as px
from sqlalchemy import create_engine
import os

# ====================================
# CONFIGURACIÓN
# ====================================

st.set_page_config(
    page_title="Dashboard Ventas",
    page_icon="📊",
    layout="wide"
)

# ====================================
# ESTILO FUTURISTA
# ====================================

st.markdown("""
<style>

.main {
    background-color: #0E1117;
}

section[data-testid="stSidebar"] {
    background-color: #161B22;
}

h1, h2, h3 {
    color: white;
}

.stMetric {
    background-color: #161B22;
    padding: 15px;
    border-radius: 15px;
    border: 1px solid #30363D;
}

</style>
""", unsafe_allow_html=True)

# ====================================
# SQLITE
# ====================================

engine = create_engine("sqlite:///ventas.db")

# ====================================
# CREAR BASE AUTOMÁTICAMENTE
# ====================================

if not os.path.exists("ventas.db"):

    archivo_2024 = "VENTA 2024.xlsx"
    archivo_2025 = "VENTA 2025.xlsx"

    ventas_2024 = pd.read_excel(archivo_2024)
    ventas_2025 = pd.read_excel(archivo_2025)

    df_total = pd.concat(
        [ventas_2024, ventas_2025],
        ignore_index=True
    )

    # LIMPIAR COLUMNAS
    df_total.columns = df_total.columns.str.strip()

    # CAMBIAR NOMBRES SI ES NECESARIO
    columnas = df_total.columns.tolist()

    # AJUSTA ESTOS NOMBRES SI TU EXCEL ES DIFERENTE
    if 'fecha' not in columnas:
        df_total.rename(columns={columnas[0]: 'fecha'}, inplace=True)

    if 'cantidad' not in columnas:
        df_total.rename(columns={columnas[1]: 'cantidad'}, inplace=True)

    # FECHAS
    df_total['fecha'] = pd.to_datetime(df_total['fecha'])

    df_total['anio'] = df_total['fecha'].dt.year
    df_total['mes'] = df_total['fecha'].dt.month
    df_total['dia'] = df_total['fecha'].dt.day
    df_total['fecha_dia'] = df_total['fecha'].dt.date

    # SQLITE
    df_total.to_sql(
        'ventas',
        engine,
        if_exists='replace',
        index=False
    )

# ====================================
# LEER SQLITE
# ====================================

df = pd.read_sql("SELECT * FROM ventas", engine)

# ====================================
# FECHAS
# ====================================

df['fecha'] = pd.to_datetime(df['fecha'])

# ====================================
# SIDEBAR
# ====================================

st.sidebar.title("⚡ Filtros")

anios = sorted(df['anio'].unique())

anio = st.sidebar.multiselect(
    "Selecciona Año",
    anios,
    default=anios
)

# ====================================
# FILTRO
# ====================================

filtro = df[df['anio'].isin(anio)]

# ====================================
# TITULO
# ====================================

st.title("🚀 Dashboard Futurista de Ventas")
st.caption("Comparativo de unidades vendidas")

# ====================================
# KPIs
# ====================================

ventas_totales = int(filtro['cantidad'].sum())

col1, col2 = st.columns(2)

col1.metric(
    "🛒 Unidades Vendidas",
    f"{ventas_totales:,}"
)

col2.metric(
    "📅 Registros",
    len(filtro)
)

# ====================================
# AÑO
# ====================================

st.subheader("📈 Ventas por Año")

ventas_anio = (
    filtro.groupby('anio')['cantidad']
    .sum()
    .reset_index()
)

fig_anio = px.bar(
    ventas_anio,
    x='anio',
    y='cantidad',
    text_auto=True,
    template='plotly_dark'
)

st.plotly_chart(fig_anio, use_container_width=True)

# ====================================
# MES
# ====================================

st.subheader("📅 Ventas por Mes")

ventas_mes = (
    filtro.groupby(['anio', 'mes'])['cantidad']
    .sum()
    .reset_index()
)

fig_mes = px.line(
    ventas_mes,
    x='mes',
    y='cantidad',
    color='anio',
    markers=True,
    template='plotly_dark'
)

st.plotly_chart(fig_mes, use_container_width=True)

# ====================================
# DÍA
# ====================================

st.subheader("📊 Ventas Diarias")

ventas_dia = (
    filtro.groupby('fecha_dia')['cantidad']
    .sum()
    .reset_index()
)

fig_dia = px.area(
    ventas_dia,
    x='fecha_dia',
    y='cantidad',
    template='plotly_dark'
)

st.plotly_chart(fig_dia, use_container_width=True)

# ====================================
# TABLA
# ====================================

st.subheader("📋 Detalle")

st.dataframe(
    filtro,
    use_container_width=True
)