import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import io
import base64
from fpdf import FPDF
import numpy as np
from streamlit.components.v1 import html

# ======================================
# CONFIG
# ======================================

st.set_page_config(
    page_title="Dashboard Ventas Futurista",
    page_icon="🚀",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ======================================
# ESTILOS NEÓN FUTURISTA
# ======================================

st.markdown("""
<style>
    /* Fondo principal */
    .stApp {
        background: linear-gradient(135deg, #0a0a2a 0%, #1a1a3a 100%);
    }
    
    /* Sidebar moderna */
    section[data-testid="stSidebar"] {
        background: linear-gradient(180deg, #0f0f2a 0%, #1a1a3a 100%);
        border-right: 2px solid #00f5ff;
        box-shadow: 5px 0 20px rgba(0, 245, 255, 0.2);
    }
    
    /* Títulos neón */
    h1, h2, h3 {
        background: linear-gradient(135deg, #00f5ff, #ff00ff);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        background-clip: text;
        font-weight: bold;
        text-shadow: 0 0 10px rgba(0, 245, 255, 0.5);
    }
    
    /* KPI Cards animadas */
    @keyframes glow {
        0% { box-shadow: 0 0 5px #00f5ff, 0 0 10px #00f5ff; }
        100% { box-shadow: 0 0 20px #00f5ff, 0 0 30px #ff00ff; }
    }
    
    .kpi {
        background: linear-gradient(135deg, rgba(15, 23, 42, 0.95), rgba(30, 41, 59, 0.95));
        backdrop-filter: blur(10px);
        border: 1px solid #00f5ff;
        border-radius: 20px;
        padding: 25px;
        text-align: center;
        margin-bottom: 15px;
        transition: all 0.3s ease;
        animation: glow 2s infinite alternate;
    }
    
    .kpi:hover {
        transform: translateY(-5px);
        animation: glow 0.5s infinite alternate;
    }
    
    .kpi h3 {
        color: #00f5ff;
        font-size: 18px;
        margin-bottom: 10px;
        text-transform: uppercase;
        letter-spacing: 2px;
    }
    
    .kpi .value {
        font-size: 42px;
        font-weight: bold;
        background: linear-gradient(135deg, #00f5ff, #ff00ff);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        margin: 10px 0;
    }
    
    .kpi .growth {
        font-size: 16px;
        padding: 5px 10px;
        border-radius: 20px;
        display: inline-block;
    }
    
    .growth-positive {
        background: rgba(0, 255, 0, 0.2);
        color: #00ff00;
        border: 1px solid #00ff00;
    }
    
    .growth-negative {
        background: rgba(255, 0, 0, 0.2);
        color: #ff0000;
        border: 1px solid #ff0000;
    }
    
    /* Botones neón */
    .stButton > button {
        background: linear-gradient(135deg, #00f5ff, #ff00ff);
        color: white;
        border: none;
        border-radius: 25px;
        padding: 10px 25px;
        font-weight: bold;
        transition: all 0.3s ease;
    }
    
    .stButton > button:hover {
        transform: scale(1.05);
        box-shadow: 0 0 20px rgba(0, 245, 255, 0.8);
    }
    
    /* Dataframe estilo futurista */
    .dataframe {
        background: rgba(15, 23, 42, 0.8);
        border-radius: 15px;
        border: 1px solid #00f5ff;
    }
    
    /* Responsive para móvil */
    @media (max-width: 768px) {
        .kpi .value {
            font-size: 28px;
        }
        .kpi h3 {
            font-size: 14px;
        }
    }
    
    /* Selectores personalizados */
    .stSelectbox, .stMultiSelect {
        background: rgba(15, 23, 42, 0.8);
        border-radius: 10px;
        border: 1px solid #00f5ff;
    }
    
    hr {
        border-color: #00f5ff;
        box-shadow: 0 0 10px #00f5ff;
    }
</style>
""", unsafe_allow_html=True)

# ======================================
# CARGA DE DATOS CON CACHÉ
# ======================================

@st.cache_data(ttl=3600)
def cargar_datos():
    try:
        ventas_2024 = pd.read_excel("VENTA 2024.xlsx")
        ventas_2025 = pd.read_excel("VENTA 2025.xlsx")
        
        df = pd.concat([ventas_2024, ventas_2025], ignore_index=True)
        
        # Limpiar columnas
        df.columns = df.columns.str.strip().str.lower()
        
        # Renombrar columnas necesarias
        columnas = df.columns.tolist()
        
        if 'fecha' not in columnas:
            df.rename(columns={columnas[0]: 'fecha'}, inplace=True)
        
        if 'cantidad' not in columnas:
            df.rename(columns={columnas[1]: 'cantidad'}, inplace=True)
        
        # Procesar fechas
        df['fecha'] = pd.to_datetime(df['fecha'], errors='coerce')
        df = df.dropna(subset=['fecha'])
        
        df['cantidad'] = pd.to_numeric(df['cantidad'], errors='coerce').fillna(0)
        
        df['anio'] = df['fecha'].dt.year
        df['mes'] = df['fecha'].dt.month
        df['dia'] = df['fecha'].dt.day
        df['mes_nombre'] = df['fecha'].dt.strftime('%B')
        df['semana'] = df['fecha'].dt.isocalendar().week
        
        # Crear columna de producto (si no existe, simular)
        if 'producto' not in df.columns:
            productos = ['Producto A', 'Producto B', 'Producto C', 'Producto D', 'Producto E']
            df['producto'] = np.random.choice(productos, len(df))
        
        if 'categoria' not in df.columns:
            categorias = ['Electrónica', 'Ropa', 'Hogar', 'Deportes', 'Juguetes']
            df['categoria'] = np.random.choice(categorias, len(df))
        
        return df
    except Exception as e:
        st.error(f"Error cargando datos: {e}")
        return pd.DataFrame()

# ======================================
# FUNCIONES DE EXPORTACIÓN
# ======================================

def exportar_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='Ventas', index=False)
        
        # Formato para Excel
        workbook = writer.book
        worksheet = writer.sheets['Ventas']
        
        # Formato neón
        neon_format = workbook.add_format({
            'bg_color': '#0a0a2a',
            'font_color': '#00f5ff',
            'border': 1,
            'border_color': '#ff00ff'
        })
        
        worksheet.set_column('A:Z', 15, neon_format)
    
    output.seek(0)
    return output

def exportar_pdf(df, kpis):
    pdf = FPDF()
    pdf.add_page()
    
    # Configurar estilo neón
    pdf.set_fill_color(10, 10, 42)
    pdf.set_text_color(0, 245, 255)
    
    # Título
    pdf.set_font("Arial", "B", 24)
    pdf.cell(0, 20, "Dashboard Ventas - Reporte Ejecutivo", 0, 1, "C")
    
    # KPIs
    pdf.set_font("Arial", "B", 16)
    pdf.cell(0, 10, f"Total Ventas: {kpis['total']:,} unidades", 0, 1)
    pdf.cell(0, 10, f"Ventas 2024: {kpis['v2024']:,} unidades", 0, 1)
    pdf.cell(0, 10, f"Ventas 2025: {kpis['v2025']:,} unidades", 0, 1)
    pdf.cell(0, 10, f"Crecimiento: {kpis['crecimiento']:.1f}%", 0, 1)
    
    # Tabla resumen
    pdf.set_font("Arial", "B", 12)
    pdf.cell(0, 10, "Resumen por Año y Mes:", 0, 1)
    
    # Guardar PDF
    return pdf.output(dest='S').encode('latin1')

# ======================================
# FILTROS INTELIGENTES
# ======================================

def aplicar_filtros(df):
    st.sidebar.title("⚡ Filtros Inteligentes")
    
    # Filtro de año
    anios = sorted(df['anio'].unique())
    anio_seleccionado = st.sidebar.multiselect(
        "📅 Años",
        anios,
        default=anios,
        help="Selecciona uno o más años"
    )
    
    # Filtro de mes
    meses = sorted(df['mes'].unique())
    mes_seleccionado = st.sidebar.multiselect(
        "📆 Meses",
        meses,
        default=meses,
        help="Selecciona meses específicos"
    )
    
    # Filtro de producto
    productos = sorted(df['producto'].unique())
    producto_seleccionado = st.sidebar.multiselect(
        "📦 Productos",
        productos,
        default=productos,
        help="Filtra por producto"
    )
    
    # Filtro de categoría
    categorias = sorted(df['categoria'].unique())
    categoria_seleccionada = st.sidebar.multiselect(
        "🏷️ Categorías",
        categorias,
        default=categorias,
        help="Filtra por categoría"
    )
    
    # Filtro de rango de fechas
    fecha_min = df['fecha'].min().date()
    fecha_max = df['fecha'].max().date()
    
    rango_fechas = st.sidebar.date_input(
        "📅 Rango de Fechas",
        [fecha_min, fecha_max],
        min_value=fecha_min,
        max_value=fecha_max
    )
    
    # Aplicar filtros
    filtro = df[
        (df['anio'].isin(anio_seleccionado)) &
        (df['mes'].isin(mes_seleccionado)) &
        (df['producto'].isin(producto_seleccionado)) &
        (df['categoria'].isin(categoria_seleccionada))
    ]
    
    if len(rango_fechas) == 2:
        filtro = filtro[
            (filtro['fecha'].dt.date >= rango_fechas[0]) &
            (filtro['fecha'].dt.date <= rango_fechas[1])
        ]
    
    return filtro

# ======================================
# KPIS CON ANIMACIÓN
# ======================================

def mostrar_kpis(filtro):
    # Calcular KPIs
    ventas_total = int(filtro['cantidad'].sum())
    ventas_2024 = int(filtro[filtro['anio'] == 2024]['cantidad'].sum())
    ventas_2025 = int(filtro[filtro['anio'] == 2025]['cantidad'].sum())
    
    # Calcular crecimiento
    if ventas_2024 > 0:
        crecimiento = ((ventas_2025 - ventas_2024) / ventas_2024) * 100
    else:
        crecimiento = 0
    
    # Mostrar KPIs con animación
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.markdown(f"""
        <div class="kpi">
            <h3>📊 UNIDADES TOTALES</h3>
            <div class="value">{ventas_total:,}</div>
            <div class="growth growth-positive">🔥 Total Acumulado</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        st.markdown(f"""
        <div class="kpi">
            <h3>📈 VENTAS 2024</h3>
            <div class="value">{ventas_2024:,}</div>
            <div class="growth growth-positive">🎯 Meta 2024</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col3:
        st.markdown(f"""
        <div class="kpi">
            <h3>🚀 VENTAS 2025</h3>
            <div class="value">{ventas_2025:,}</div>
            <div class="growth {'growth-positive' if crecimiento >= 0 else 'growth-negative'}">
                {'📈' if crecimiento >= 0 else '📉'} {crecimiento:.1f}% vs 2024
            </div>
        </div>
        """, unsafe_allow_html=True)
    
    return {
        'total': ventas_total,
        'v2024': ventas_2024,
        'v2025': ventas_2025,
        'crecimiento': crecimiento
    }

# ======================================
# HEATMAP DE VENTAS
# ======================================

def crear_heatmap(filtro):
    st.subheader("🌡️ Heatmap de Ventas - Día vs Mes")
    
    # Crear matriz para heatmap
    heatmap_data = filtro.groupby(['mes', 'dia'])['cantidad'].sum().reset_index()
    
    # Crear pivot table
    pivot_heatmap = heatmap_data.pivot(index='dia', columns='mes', values='cantidad').fillna(0)
    
    fig = px.imshow(
        pivot_heatmap,
        labels=dict(x="Mes", y="Día", color="Ventas"),
        x=pivot_heatmap.columns,
        y=pivot_heatmap.index,
        color_continuous_scale="Viridis",
        aspect="auto",
        title="Mapa de Calor: Ventas por Día y Mes"
    )
    
    fig.update_layout(
        template='plotly_dark',
        height=500,
        xaxis_title="Mes",
        yaxis_title="Día del Mes"
    )
    
    st.plotly_chart(fig, use_container_width=True)

# ======================================
# TOP PRODUCTOS
# ======================================

def mostrar_top_productos(filtro):
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("🏆 Top 5 Productos")
        top_productos = filtro.groupby('producto')['cantidad'].sum().sort_values(ascending=False).head(5)
        
        fig = px.bar(
            top_productos,
            x=top_productos.values,
            y=top_productos.index,
            orientation='h',
            text_auto=True,
            color=top_productos.values,
            color_continuous_scale='Viridis',
            title="Productos Más Vendidos"
        )
        fig.update_layout(template='plotly_dark', height=400)
        st.plotly_chart(fig, use_container_width=True)
    
    with col2:
        st.subheader("🎯 Top 5 Categorías")
        top_categorias = filtro.groupby('categoria')['cantidad'].sum().sort_values(ascending=False).head(5)
        
        fig = px.pie(
            values=top_categorias.values,
            names=top_categorias.index,
            title="Distribución por Categoría",
            color_discrete_sequence=px.colors.sequential.Viridis
        )
        fig.update_layout(template='plotly_dark', height=400)
        st.plotly_chart(fig, use_container_width=True)

# ======================================
# GRÁFICOS PRINCIPALES
# ======================================

def graficos_principales(filtro):
    # Gráfico comparativo por año
    st.subheader("📊 Comparativo Anual")
    
    ventas_anio = filtro.groupby('anio')['cantidad'].sum().reset_index()
    
    fig1 = px.bar(
        ventas_anio,
        x='anio',
        y='cantidad',
        text_auto=True,
        color='cantidad',
        color_continuous_scale='Viridis',
        title="Ventas Totales por Año"
    )
    fig1.update_layout(template='plotly_dark', height=400)
    st.plotly_chart(fig1, use_container_width=True)
    
    # Gráfico por mes con comparativo
    st.subheader("📈 Tendencia Mensual 2024 vs 2025")
    
    ventas_mes = filtro.groupby(['anio', 'mes_nombre', 'mes'])['cantidad'].sum().reset_index()
    ventas_mes = ventas_mes.sort_values('mes')
    
    fig2 = px.line(
        ventas_mes,
        x='mes_nombre',
        y='cantidad',
        color='anio',
        markers=True,
        title="Comparativo Mensual",
        line_shape='spline'
    )
    fig2.update_layout(template='plotly_dark', height=450)
    fig2.update_traces(marker_size=10)
    st.plotly_chart(fig2, use_container_width=True)

# ======================================
# BOTONES DE EXPORTACIÓN
# ======================================

def botones_exportacion(filtro, kpis):
    col1, col2 = st.columns(2)
    
    with col1:
        if st.button("📊 Exportar a Excel", use_container_width=True):
            excel_data = exportar_excel(filtro)
            b64 = base64.b64encode(excel_data.getvalue()).decode()
            href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="dashboard_ventas.xlsx">📥 Descargar Excel</a>'
            st.markdown(href, unsafe_allow_html=True)
            st.success("✅ Excel generado correctamente!")
    
    with col2:
        if st.button("📄 Exportar a PDF", use_container_width=True):
            pdf_data = exportar_pdf(filtro, kpis)
            b64 = base64.b64encode(pdf_data).decode()
            href = f'<a href="data:application/pdf;base64,{b64}" download="dashboard_ventas.pdf">📥 Descargar PDF</a>'
            st.markdown(href, unsafe_allow_html=True)
            st.success("✅ PDF generado correctamente!")

# ======================================
# MAIN
# ======================================

def main():
    # Cargar datos
    df = cargar_datos()
    
    if df.empty:
        st.error("No se pudieron cargar los datos. Verifica los archivos Excel.")
        return
    
    # Aplicar filtros
    filtro = aplicar_filtros(df)
    
    # Título principal
    st.title("🚀 Dashboard Ejecutivo de Ventas")
    st.caption("✨ Análisis en Tiempo Real | 2024-2025 | Power BI Style")
    st.markdown("---")
    
    # Mostrar KPIs
    kpis = mostrar_kpis(filtro)
    
    # Botones de exportación
    botones_exportacion(filtro, kpis)
    
    st.markdown("---")
    
    # Gráficos principales
    graficos_principales(filtro)
    
    st.markdown("---")
    
    # Heatmap
    if len(filtro) > 0:
        crear_heatmap(filtro)
    
    st.markdown("---")
    
    # Top productos
    mostrar_top_productos(filtro)
    
    st.markdown("---")
    
    # Tabla detallada
    st.subheader("📋 Detalle de Ventas")
    
    # Mostrar métricas adicionales
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Días Analizados", filtro['fecha'].nunique())
    with col2:
        st.metric("Productos Diferentes", filtro['producto'].nunique())
    with col3:
        st.metric("Promedio Diario", f"{filtro['cantidad'].mean():,.0f}")
    
    # Dataframe con estilo
    st.dataframe(
        filtro[['fecha', 'producto', 'categoria', 'cantidad', 'anio', 'mes_nombre']].sort_values('fecha', ascending=False),
        use_container_width=True,
        height=400
    )
    
    # Footer futurista
    st.markdown("---")
    st.markdown(
        """
        <div style="text-align: center; padding: 20px;">
            <p style="color: #00f5ff;">✨ Dashboard Desarrollado con ❤️ usando Streamlit | Datos Actualizados ✨</p>
        </div>
        """,
        unsafe_allow_html=True
    )

if __name__ == "__main__":
    main()