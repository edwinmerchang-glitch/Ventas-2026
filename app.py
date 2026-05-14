import streamlit as st
import pandas as pd
import plotly.express as px

# ======================================
# CONFIG
# ======================================

st.set_page_config(
    page_title="Dashboard Ventas",
    page_icon="🚀",
    layout="wide"
)

# ======================================
# ESTILOS
# ======================================

st.markdown("""
<style>

.main {
    background-color: #050816;
}

h1, h2, h3 {
    color: white;
}

section[data-testid="stSidebar"] {
    background-color: #111827;
}

.kpi {
    background: linear-gradient(145deg,#0F172A,#111827);
    border: 1px solid #00F5FF;
    border-radius: 20px;
    padding: 20px;
    text-align: center;
    margin-bottom: 15px;
}

.kpi h3 {
    color: white;
    font-size: 16px;
}

.kpi h1 {
    color: #00F5FF;
    font-size: 35px;
}

</style>
""", unsafe_allow_html=True)

# ======================================
# LEER ARCHIVOS EXCEL
# ======================================

ventas_2024 = pd.read_excel("VENTA 2024.xlsx")
ventas_2025 = pd.read_excel("VENTA 2025.xlsx")

df = pd.concat(
    [ventas_2024, ventas_2025],
    ignore_index=True
)

# ======================================
# LIMPIAR COLUMNAS
# ======================================

df.columns = df.columns.str.strip().str.lower()

# ======================================
# AJUSTAR COLUMNAS
# ======================================

columnas = df.columns.tolist()

if 'fecha' not in columnas:
    df.rename(columns={columnas[0]: 'fecha'}, inplace=True)

if 'cantidad' not in columnas:
    df.rename(columns={columnas[1]: 'cantidad'}, inplace=True)

# ======================================
# FECHAS
# ======================================

df['fecha'] = pd.to_datetime(
    df['fecha'],
    errors='coerce'
)

df = df.dropna(subset=['fecha'])

df['cantidad'] = pd.to_numeric(
    df['cantidad'],
    errors='coerce'
).fillna(0)

df['anio'] = df['fecha'].dt.year
df['mes'] = df['fecha'].dt.month
df['dia'] = df['fecha'].dt.day

# ======================================
# FILTROS
# ======================================

st.sidebar.title("⚡ Filtros")

anios = sorted(df['anio'].unique())

anio = st.sidebar.multiselect(
    "Selecciona Año",
    anios,
    default=anios
)

filtro = df[df['anio'].isin(anio)]

# ======================================
# TITULO
# ======================================

st.title("🚀 Dashboard Ejecutivo Ventas")
st.caption("Comparativo 2024 vs 2025")

# ======================================
# KPIS
# ======================================

ventas_total = int(filtro['cantidad'].sum())

ventas_2024_total = int(
    filtro[filtro['anio'] == 2024]['cantidad'].sum()
)

ventas_2025_total = int(
    filtro[filtro['anio'] == 2025]['cantidad'].sum()
)

col1, col2, col3 = st.columns(3)

with col1:
    st.markdown(f"""
    <div class="kpi">
        <h3>UNIDADES</h3>
        <h1>{ventas_total:,}</h1>
    </div>
    """, unsafe_allow_html=True)

with col2:
    st.markdown(f"""
    <div class="kpi">
        <h3>VENTAS 2024</h3>
        <h1>{ventas_2024_total:,}</h1>
    </div>
    """, unsafe_allow_html=True)

with col3:
    st.markdown(f"""
    <div class="kpi">
        <h3>VENTAS 2025</h3>
        <h1>{ventas_2025_total:,}</h1>
    </div>
    """, unsafe_allow_html=True)

# ======================================
# GRAFICO AÑO
# ======================================

st.subheader("📈 Comparativo por Año")

ventas_anio = (
    filtro.groupby('anio')['cantidad']
    .sum()
    .reset_index()
)

fig1 = px.bar(
    ventas_anio,
    x='anio',
    y='cantidad',
    text_auto=True,
    template='plotly_dark'
)

st.plotly_chart(fig1, width='stretch')

# ======================================
# GRAFICO MES
# ======================================

st.subheader("📅 Ventas por Mes")

ventas_mes = (
    filtro.groupby(['anio', 'mes'])['cantidad']
    .sum()
    .reset_index()
)

fig2 = px.line(
    ventas_mes,
    x='mes',
    y='cantidad',
    color='anio',
    markers=True,
    template='plotly_dark'
)

st.plotly_chart(fig2, width='stretch')

# ======================================
# TABLA
# ======================================

st.subheader("📋 Detalle")

st.dataframe(
    filtro,
    width='stretch',
    height=500
)

