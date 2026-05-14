import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import io
import base64
from fpdf import FPDF
import numpy as np
from pathlib import Path

# ======================================
# CONFIG
# ======================================

st.set_page_config(
    page_title="Dashboard Ventas Corporativo",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ======================================
# ESTILOS CORPORATIVOS PROFESIONALES
# ======================================

st.markdown("""
<style>
    /* Fondo corporativo limpio */
    .stApp {
        background: linear-gradient(135deg, #f5f7fa 0%, #ffffff 100%);
    }
    
    /* Sidebar corporativa */
    section[data-testid="stSidebar"] {
        background: linear-gradient(180deg, #1e293b 0%, #0f172a 100%);
        border-right: 1px solid #e2e8f0;
    }
    
    /* Títulos corporativos */
    h1 {
        color: #1e293b;
        font-size: 2rem;
        font-weight: 700;
        border-bottom: 3px solid #3b82f6;
        padding-bottom: 0.5rem;
        display: inline-block;
    }
    
    h2, h3 {
        color: #334155;
        font-weight: 600;
    }
    
    h2 {
        border-left: 4px solid #3b82f6;
        padding-left: 1rem;
        margin-top: 1rem;
    }
    
    /* KPI Cards corporativas */
    .kpi {
        background: white;
        border-radius: 12px;
        padding: 1.2rem;
        text-align: center;
        box-shadow: 0 1px 3px rgba(0,0,0,0.12), 0 1px 2px rgba(0,0,0,0.08);
        transition: all 0.3s ease;
        border-top: 4px solid #3b82f6;
    }
    
    .kpi:hover {
        transform: translateY(-2px);
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }
    
    .kpi h3 {
        color: #64748b;
        font-size: 0.75rem;
        text-transform: uppercase;
        letter-spacing: 1px;
        margin-bottom: 0.5rem;
    }
    
    .kpi .value {
        font-size: 1.5rem;
        font-weight: bold;
        color: #1e293b;
        margin: 0.5rem 0;
    }
    
    .kpi .growth {
        font-size: 0.75rem;
        padding: 0.2rem 0.4rem;
        border-radius: 20px;
        display: inline-block;
    }
    
    .growth-positive {
        background: #dcfce7;
        color: #166534;
    }
    
    .growth-negative {
        background: #fee2e2;
        color: #991b1b;
    }
    
    /* Botones corporativos */
    .stButton > button {
        background: #3b82f6;
        color: white;
        border: none;
        border-radius: 8px;
        padding: 0.4rem 0.8rem;
        font-weight: 500;
        transition: all 0.2s ease;
        width: 100%;
    }
    
    .stButton > button:hover {
        background: #2563eb;
        transform: translateY(-1px);
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
    }
    
    /* Dataframe estilo corporativo */
    .stDataFrame {
        border-radius: 8px;
        border: 1px solid #e2e8f0;
    }
    
    /* Selectores corporativos */
    .stSelectbox label, .stMultiSelect label {
        color: #ffffff !important;
        font-weight: 500;
        font-size: 0.8rem;
    }
    
    /* Sidebar text color */
    section[data-testid="stSidebar"] .stMarkdown {
        color: #e2e8f0;
    }
    
    section[data-testid="stSidebar"] h3 {
        color: #ffffff !important;
        font-size: 0.9rem;
    }
    
    section[data-testid="stSidebar"] h4 {
        color: #ffffff !important;
        font-size: 0.8rem;
    }
    
    /* Footer */
    .footer {
        text-align: center;
        padding: 1rem;
        color: #94a3b8;
        border-top: 1px solid #e2e8f0;
        margin-top: 1rem;
        font-size: 0.8rem;
    }
    
    /* Responsive */
    @media (max-width: 768px) {
        .kpi .value {
            font-size: 1.2rem;
        }
        h1 {
            font-size: 1.5rem;
        }
        .kpi {
            padding: 0.8rem;
        }
    }
</style>
""", unsafe_allow_html=True)

# ======================================
# FUNCIONES DE LIMPIEZA DE DATOS
# ======================================

@st.cache_data(ttl=3600)
def limpiar_dataframe(df):
    """Limpia el dataframe de valores erróneos"""
    
    # Limpiar columna de año - asegurar que sea numérico
    if 'anio' in df.columns:
        df['anio'] = pd.to_numeric(df['anio'], errors='coerce')
    
    # Limpiar columna de mes - asegurar que sea numérico
    if 'mes' in df.columns:
        df['mes'] = pd.to_numeric(df['mes'], errors='coerce')
    
    # Limpiar columna de cantidad
    if 'cantidad' in df.columns:
        df['cantidad'] = pd.to_numeric(df['cantidad'], errors='coerce').fillna(0)
    
    # Eliminar filas con fechas inválidas
    if 'fecha' in df.columns:
        df['fecha'] = pd.to_datetime(df['fecha'], errors='coerce')
        df = df.dropna(subset=['fecha'])
        
        # Recalcular año y mes si es necesario
        df['anio'] = df['fecha'].dt.year
        df['mes'] = df['fecha'].dt.month
        df['mes_nombre'] = df['fecha'].dt.strftime('%B')
        df['dia'] = df['fecha'].dt.day
        df['trimestre'] = df['fecha'].dt.quarter
    
    # Limpiar valores nulos en columnas importantes
    df['producto'] = df['producto'].fillna('No especificado').astype(str)
    df['marca'] = df['marca'].fillna('No especificada').astype(str)
    df['proveedor'] = df['proveedor'].fillna('No especificado').astype(str)
    df['categoria'] = df['categoria'].fillna('General').astype(str)
    
    # Eliminar filas donde el año sea inválido (NaN o fuera de rango)
    df = df.dropna(subset=['anio'])
    df = df[(df['anio'] >= 2000) & (df['anio'] <= 2030)]
    
    return df

@st.cache_data(ttl=3600)
def cargar_excel(file):
    """Carga archivo Excel y procesa los datos"""
    try:
        df = pd.read_excel(file)
        
        # Limpiar columnas
        df.columns = df.columns.str.strip().str.lower()
        
        # Identificar columnas necesarias
        columnas = df.columns.tolist()
        
        # Buscar columna de fecha
        fecha_col = None
        for col in columnas:
            if 'fecha' in col or 'date' in col:
                fecha_col = col
                break
        
        if fecha_col is None:
            fecha_col = columnas[0]
        
        # Buscar columna de cantidad/ventas
        cantidad_col = None
        for col in columnas:
            if 'cantidad' in col or 'venta' in col or 'monto' in col:
                cantidad_col = col
                break
        
        if cantidad_col is None:
            cantidad_col = columnas[1] if len(columnas) > 1 else columnas[0]
        
        # Renombrar columnas
        df.rename(columns={fecha_col: 'fecha', cantidad_col: 'cantidad'}, inplace=True)
        
        # Buscar o crear columna de producto
        if 'producto' not in df.columns:
            for col in columnas:
                if 'producto' in col or 'item' in col or 'articulo' in col:
                    df.rename(columns={col: 'producto'}, inplace=True)
                    break
            else:
                df['producto'] = 'Producto General'
        
        # Buscar o crear columna de marca
        if 'marca' not in df.columns:
            for col in columnas:
                if 'marca' in col or 'brand' in col:
                    df.rename(columns={col: 'marca'}, inplace=True)
                    break
            else:
                df['marca'] = 'General'
        
        # Buscar o crear columna de proveedor
        if 'proveedor' not in df.columns:
            for col in columnas:
                if 'proveedor' in col or 'supplier' in col:
                    df.rename(columns={col: 'proveedor'}, inplace=True)
                    break
            else:
                df['proveedor'] = 'General'
        
        # Buscar columna de categoría si existe
        if 'categoria' not in df.columns:
            for col in columnas:
                if 'categoria' in col or 'category' in col:
                    df.rename(columns={col: 'categoria'}, inplace=True)
                    break
            else:
                df['categoria'] = 'General'
        
        # Aplicar limpieza de datos
        df = limpiar_dataframe(df)
        
        return df, True
    except Exception as e:
        return None, False

@st.cache_data(ttl=3600)
def cargar_datos_ejemplo():
    """Genera datos de ejemplo para demostración con marcas y proveedores"""
    np.random.seed(42)
    
    fechas = pd.date_range('2024-01-01', '2025-12-31', freq='D')
    productos = ['Producto A', 'Producto B', 'Producto C', 'Producto D', 'Producto E']
    marcas = ['HAIKO NATURAL S.A.S.', 'Marca B', 'Marca C', 'Marca D']
    proveedores = ['HAIKO NATURAL S.A.S.', 'Proveedor X', 'Proveedor Y', 'Proveedor Z']
    categorias = ['Electrónica', 'Ropa', 'Hogar', 'Deportes', 'Juguetes']
    
    data = []
    for fecha in fechas:
        for _ in range(np.random.randint(1, 3)):  # Reducido para mejor rendimiento
            data.append({
                'fecha': fecha,
                'producto': np.random.choice(productos),
                'marca': np.random.choice(marcas),
                'proveedor': np.random.choice(proveedores),
                'categoria': np.random.choice(categorias),
                'cantidad': np.random.randint(1, 50)
            })
    
    df = pd.DataFrame(data)
    df = limpiar_dataframe(df)
    
    return df

# ======================================
# FUNCIONES DE EXPORTACIÓN
# ======================================

def exportar_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # Limitar a 10000 filas para evitar problemas de rendimiento
        df_to_export = df.head(10000)
        df_to_export.to_excel(writer, sheet_name='Ventas', index=False)
    
    output.seek(0)
    return output

def exportar_pdf(df, kpis, filtros_aplicados):
    pdf = FPDF()
    pdf.add_page()
    
    # Configurar estilo corporativo
    pdf.set_fill_color(59, 130, 246)
    pdf.set_text_color(30, 41, 59)
    
    # Título
    pdf.set_font("Arial", "B", 20)
    pdf.cell(0, 15, "Reporte Ejecutivo de Ventas", 0, 1, "C")
    pdf.set_font("Arial", "", 9)
    pdf.cell(0, 8, f"Generado: {datetime.now().strftime('%Y-%m-%d %H:%M')}", 0, 1, "C")
    pdf.ln(5)
    
    # Filtros aplicados (resumidos)
    pdf.set_font("Arial", "B", 10)
    pdf.cell(0, 8, "Filtros Aplicados:", 0, 1)
    pdf.set_font("Arial", "", 8)
    
    filtros_resumidos = {
        'Fechas': filtros_aplicados.get('Rango de fechas', 'Todos')[:50],
        'Años': filtros_aplicados.get('Años', 'Todos')[:50],
        'Marcas': filtros_aplicados.get('Marcas', 'Todas')[:50]
    }
    
    for key, value in filtros_resumidos.items():
        pdf.cell(0, 5, f"{key}: {value}", 0, 1)
    
    pdf.ln(8)
    
    # KPIs
    pdf.set_font("Arial", "B", 12)
    pdf.cell(0, 8, "Indicadores Clave", 0, 1)
    pdf.set_font("Arial", "", 10)
    pdf.cell(0, 7, f"Total Ventas: {kpis['total']:,} unidades", 0, 1)
    pdf.cell(0, 7, f"Ventas 2024: {kpis['v2024']:,} unidades", 0, 1)
    pdf.cell(0, 7, f"Ventas 2025: {kpis['v2025']:,} unidades", 0, 1)
    pdf.cell(0, 7, f"Crecimiento: {kpis['crecimiento']:.1f}%", 0, 1)
    
    return pdf.output(dest='S').encode('latin1')

# ======================================
# FILTROS AVANZADOS OPTIMIZADOS
# ======================================

def aplicar_filtros(df):
    st.sidebar.markdown("### 📊 Panel de Control")
    st.sidebar.markdown("---")
    
    # ===== SECCIÓN FECHAS =====
    st.sidebar.markdown("#### 📅 Filtros de Fecha")
    
    # Obtener fechas mínima y máxima
    fecha_min = df['fecha'].min().date()
    fecha_max = df['fecha'].max().date()
    
    # Selector de rango de fechas
    rango_fechas = st.sidebar.date_input(
        "📆 Rango de Fechas",
        [fecha_min, fecha_max],
        min_value=fecha_min,
        max_value=fecha_max,
        key="rango_fechas"
    )
    
    # Filtro de año - con manejo de valores vacíos
    anios = sorted(df['anio'].dropna().unique())
    anios = [a for a in anios if isinstance(a, (int, float)) and a >= 2000]
    
    default_anios = anios if anios else []
    anio_seleccionado = st.sidebar.multiselect(
        "📅 Años",
        options=anios,
        default=default_anios,
        key="anios"
    )
    
    # Filtro de mes
    nombres_meses = {
        1: 'Enero', 2: 'Febrero', 3: 'Marzo', 4: 'Abril',
        5: 'Mayo', 6: 'Junio', 7: 'Julio', 8: 'Agosto',
        9: 'Septiembre', 10: 'Octubre', 11: 'Noviembre', 12: 'Diciembre'
    }
    
    meses_disponibles = sorted([m for m in df['mes'].dropna().unique() if m in nombres_meses])
    meses_nombres = [f"{m} - {nombres_meses[m]}" for m in meses_disponibles]
    
    default_meses = meses_nombres if meses_nombres else []
    meses_seleccionados = st.sidebar.multiselect(
        "📆 Meses",
        options=meses_nombres,
        default=default_meses,
        key="meses"
    )
    
    # Convertir selección de meses
    meses_filtro = []
    for m in meses_seleccionados:
        try:
            meses_filtro.append(int(m.split(' - ')[0]))
        except:
            pass
    
    st.sidebar.markdown("---")
    
    # ===== SECCIÓN MARCAS =====
    st.sidebar.markdown("#### 🏷️ Filtros de Marca")
    
    # Obtener marcas únicas (excluyendo valores no deseados)
    marcas_disponibles = sorted([m for m in df['marca'].unique() if m not in ['No especificada', 'General', 'ZUZUS']])
    default_marcas = marcas_disponibles if marcas_disponibles else []
    
    marcas_seleccionadas = st.sidebar.multiselect(
        "🎯 Marcas",
        options=marcas_disponibles,
        default=default_marcas,
        key="marcas",
        help="Filtrar por marca específica"
    )
    
    st.sidebar.markdown("---")
    
    # ===== SECCIÓN PROVEEDORES =====
    st.sidebar.markdown("#### 🏭 Filtros de Proveedor")
    
    proveedores_disponibles = sorted([p for p in df['proveedor'].unique() if p not in ['No especificado', 'General']])
    default_proveedores = proveedores_disponibles if proveedores_disponibles else []
    
    proveedores_seleccionados = st.sidebar.multiselect(
        "📦 Proveedores",
        options=proveedores_disponibles,
        default=default_proveedores,
        key="proveedores",
        help="Filtrar por proveedor específico"
    )
    
    st.sidebar.markdown("---")
    
    # ===== SECCIÓN PRODUCTOS =====
    st.sidebar.markdown("#### 📦 Filtros de Producto")
    
    productos_disponibles = sorted([p for p in df['producto'].unique() if p not in ['No especificado', 'General']])[:50]  # Limitar a 50 para rendimiento
    default_productos = productos_disponibles if productos_disponibles else []
    
    producto_seleccionado = st.sidebar.multiselect(
        "📦 Productos",
        options=productos_disponibles,
        default=default_productos,
        key="productos"
    )
    
    # Filtro de categoría
    categorias_disponibles = sorted([c for c in df['categoria'].unique() if c not in ['General']])
    default_categorias = categorias_disponibles if categorias_disponibles else []
    
    categoria_seleccionada = st.sidebar.multiselect(
        "🏷️ Categorías",
        options=categorias_disponibles,
        default=default_categorias,
        key="categorias"
    )
    
    st.sidebar.markdown("---")
    
    # Aplicar filtros de manera eficiente
    filtro = df.copy()
    
    # Aplicar filtro de rango de fechas
    if len(rango_fechas) == 2:
        mask = (filtro['fecha'].dt.date >= rango_fechas[0]) & (filtro['fecha'].dt.date <= rango_fechas[1])
        filtro = filtro[mask]
    
    # Aplicar filtro de años
    if anio_seleccionado:
        filtro = filtro[filtro['anio'].isin(anio_seleccionado)]
    
    # Aplicar filtro de meses
    if meses_filtro:
        filtro = filtro[filtro['mes'].isin(meses_filtro)]
    
    # Aplicar filtro de marcas
    if marcas_seleccionadas:
        filtro = filtro[filtro['marca'].isin(marcas_seleccionadas)]
    
    # Aplicar filtro de proveedores
    if proveedores_seleccionados:
        filtro = filtro[filtro['proveedor'].isin(proveedores_seleccionados)]
    
    # Aplicar filtro de productos
    if producto_seleccionado:
        filtro = filtro[filtro['producto'].isin(producto_seleccionado)]
    
    # Aplicar filtro de categorías
    if categoria_seleccionada:
        filtro = filtro[filtro['categoria'].isin(categoria_seleccionada)]
    
    # Mostrar estadísticas de filtros en sidebar
    with st.sidebar.expander("📈 Estadísticas", expanded=False):
        st.metric("Registros", f"{len(filtro):,}")
        st.metric("Ventas totales", f"{int(filtro['cantidad'].sum()):,}")
    
    # Guardar filtros aplicados para exportación
    filtros_dict = {
        'Rango de fechas': f"{rango_fechas[0]} a {rango_fechas[1]}" if len(rango_fechas) == 2 else "Todos",
        'Años': ', '.join([str(a) for a in anio_seleccionado]) if anio_seleccionado else "Todos",
        'Marcas': ', '.join([str(m) for m in marcas_seleccionadas[:3]]) + ('...' if len(marcas_seleccionadas) > 3 else '') if marcas_seleccionadas else "Todas",
        'Proveedores': ', '.join([str(p) for p in proveedores_seleccionados[:3]]) + ('...' if len(proveedores_seleccionados) > 3 else '') if proveedores_seleccionados else "Todos",
    }
    
    return filtro, filtros_dict

# ======================================
# KPIS OPTIMIZADOS
# ======================================

def mostrar_kpis(filtro):
    # Calcular KPIs de forma eficiente
    ventas_total = int(filtro['cantidad'].sum())
    
    ventas_2024 = 0
    ventas_2025 = 0
    
    if 2024 in filtro['anio'].values:
        ventas_2024 = int(filtro[filtro['anio'] == 2024]['cantidad'].sum())
    if 2025 in filtro['anio'].values:
        ventas_2025 = int(filtro[filtro['anio'] == 2025]['cantidad'].sum())
    
    if ventas_2024 > 0:
        crecimiento = ((ventas_2025 - ventas_2024) / ventas_2024) * 100
    else:
        crecimiento = 0
    
    # KPIs adicionales
    marcas_unicas = filtro['marca'].nunique()
    proveedores_unicos = filtro['proveedor'].nunique()
    
    col1, col2, col3, col4, col5 = st.columns(5)
    
    with col1:
        st.markdown(f"""
        <div class="kpi">
            <h3>📊 Total</h3>
            <div class="value">{ventas_total:,}</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        st.markdown(f"""
        <div class="kpi">
            <h3>📈 2024</h3>
            <div class="value">{ventas_2024:,}</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col3:
        st.markdown(f"""
        <div class="kpi">
            <h3>🚀 2025</h3>
            <div class="value">{ventas_2025:,}</div>
            <div class="growth {'growth-positive' if crecimiento >= 0 else 'growth-negative'}">
                {crecimiento:.1f}%
            </div>
        </div>
        """, unsafe_allow_html=True)
    
    with col4:
        st.markdown(f"""
        <div class="kpi">
            <h3>🏷️ Marcas</h3>
            <div class="value">{marcas_unicas}</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col5:
        st.markdown(f"""
        <div class="kpi">
            <h3>🏭 Proveedores</h3>
            <div class="value">{proveedores_unicos}</div>
        </div>
        """, unsafe_allow_html=True)
    
    return {
        'total': ventas_total,
        'v2024': ventas_2024,
        'v2025': ventas_2025,
        'crecimiento': crecimiento
    }

# ======================================
# GRÁFICOS OPTIMIZADOS
# ======================================

@st.cache_data(ttl=300)
def preparar_datos_graficos(filtro):
    """Prepara los datos para los gráficos de forma cacheada"""
    ventas_anio = filtro.groupby('anio')['cantidad'].sum().reset_index()
    ventas_mes = filtro.groupby(['anio', 'mes_nombre', 'mes'])['cantidad'].sum().reset_index()
    ventas_marca = filtro.groupby('marca')['cantidad'].sum().sort_values(ascending=False).head(10)
    ventas_proveedor = filtro.groupby('proveedor')['cantidad'].sum().sort_values(ascending=False).head(10)
    top_productos = filtro.groupby('producto')['cantidad'].sum().sort_values(ascending=False).head(5)
    top_categorias = filtro.groupby('categoria')['cantidad'].sum()
    
    return ventas_anio, ventas_mes, ventas_marca, ventas_proveedor, top_productos, top_categorias

def graficos_principales(filtro):
    ventas_anio, ventas_mes, ventas_marca, ventas_proveedor, top_productos, top_categorias = preparar_datos_graficos(filtro)
    
    # Gráfico comparativo por año
    st.markdown("### 📊 Comparativo Anual")
    
    if len(ventas_anio) > 0:
        fig1 = px.bar(
            ventas_anio,
            x='anio',
            y='cantidad',
            text_auto=True,
            color='cantidad',
            color_continuous_scale='Blues',
            title="Ventas Totales por Año"
        )
        fig1.update_layout(
            template='plotly_white',
            height=350,
            showlegend=False,
            xaxis_title="Año",
            yaxis_title="Unidades"
        )
        st.plotly_chart(fig1, use_container_width=True)
    
    # Gráfico por mes
    st.markdown("### 📈 Tendencia Mensual")
    
    if len(ventas_mes) > 0:
        ventas_mes = ventas_mes.sort_values('mes')
        fig2 = px.line(
            ventas_mes,
            x='mes_nombre',
            y='cantidad',
            color='anio',
            markers=True,
            title="Comparativo Mensual",
            color_discrete_sequence=['#3b82f6', '#10b981']
        )
        fig2.update_layout(
            template='plotly_white',
            height=350,
            xaxis_title="Mes",
            yaxis_title="Unidades"
        )
        fig2.update_traces(marker_size=6)
        st.plotly_chart(fig2, use_container_width=True)
    
    # Dos columnas para gráficos
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("### 🏷️ Top Marcas")
        if len(ventas_marca) > 0:
            fig3 = px.bar(
                ventas_marca,
                x=ventas_marca.values,
                y=ventas_marca.index,
                orientation='h',
                text_auto=True,
                color=ventas_marca.values,
                color_continuous_scale='Blues',
                title="Top 10 Marcas"
            )
            fig3.update_layout(template='plotly_white', height=350)
            st.plotly_chart(fig3, use_container_width=True)
    
    with col2:
        st.markdown("### 🏭 Top Proveedores")
        if len(ventas_proveedor) > 0:
            fig4 = px.bar(
                ventas_proveedor,
                x=ventas_proveedor.values,
                y=ventas_proveedor.index,
                orientation='h',
                text_auto=True,
                color=ventas_proveedor.values,
                color_continuous_scale='Greens',
                title="Top 10 Proveedores"
            )
            fig4.update_layout(template='plotly_white', height=350)
            st.plotly_chart(fig4, use_container_width=True)

def mostrar_top_productos(filtro):
    _, _, _, _, top_productos, top_categorias = preparar_datos_graficos(filtro)
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("### 🏆 Top 5 Productos")
        if len(top_productos) > 0:
            fig = px.bar(
                top_productos,
                x=top_productos.values,
                y=top_productos.index,
                orientation='h',
                text_auto=True,
                color=top_productos.values,
                color_continuous_scale='Blues'
            )
            fig.update_layout(template='plotly_white', height=350)
            st.plotly_chart(fig, use_container_width=True)
    
    with col2:
        st.markdown("### 🎯 Por Categoría")
        top_categorias = top_categorias[top_categorias.index != 'General']
        if len(top_categorias) > 0:
            fig = px.pie(
                values=top_categorias.values,
                names=top_categorias.index,
                color_discrete_sequence=px.colors.sequential.Blues_r
            )
            fig.update_layout(template='plotly_white', height=350)
            st.plotly_chart(fig, use_container_width=True)

# ======================================
# BOTONES DE EXPORTACIÓN
# ======================================

def botones_exportacion(filtro, kpis, filtros_aplicados):
    col1, col2 = st.columns(2)
    
    with col1:
        if st.button("📊 Exportar Excel", use_container_width=True):
            with st.spinner("Generando Excel..."):
                excel_data = exportar_excel(filtro)
                b64 = base64.b64encode(excel_data.getvalue()).decode()
                href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="reporte_ventas.xlsx">📥 Descargar</a>'
                st.markdown(href, unsafe_allow_html=True)
                st.success("✅ Listo!")
    
    with col2:
        if st.button("📄 Exportar PDF", use_container_width=True):
            with st.spinner("Generando PDF..."):
                pdf_data = exportar_pdf(filtro, kpis, filtros_aplicados)
                b64 = base64.b64encode(pdf_data).decode()
                href = f'<a href="data:application/pdf;base64,{b64}" download="reporte_ventas.pdf">📥 Descargar</a>'
                st.markdown(href, unsafe_allow_html=True)
                st.success("✅ Listo!")

# ======================================
# MENÚ PRINCIPAL
# ======================================

def main():
    # Inicializar estado de sesión
    if 'datos_cargados' not in st.session_state:
        st.session_state.datos_cargados = False
        st.session_state.df = None
    
    # Sidebar - Menú de carga
    with st.sidebar:
        st.markdown("### 📁 Datos")
        
        archivo_2024 = st.file_uploader("📂 Ventas 2024", type=['xlsx', 'xls'], key="2024")
        archivo_2025 = st.file_uploader("📂 Ventas 2025", type=['xlsx', 'xls'], key="2025")
        
        col1, col2 = st.columns(2)
        with col1:
            if st.button("🔄 Cargar", use_container_width=True):
                if archivo_2024 and archivo_2025:
                    with st.spinner("Cargando..."):
                        df1, success1 = cargar_excel(archivo_2024)
                        df2, success2 = cargar_excel(archivo_2025)
                        
                        if success1 and success2:
                            st.session_state.df = pd.concat([df1, df2], ignore_index=True)
                            st.session_state.df = limpiar_dataframe(st.session_state.df)
                            st.session_state.datos_cargados = True
                            st.success("✅ Listo!")
                        else:
                            st.error("Error al cargar")
                else:
                    st.warning("Selecciona ambos archivos")
        
        with col2:
            if st.button("📊 Ejemplo", use_container_width=True):
                with st.spinner("Cargando ejemplo..."):
                    st.session_state.df = cargar_datos_ejemplo()
                    st.session_state.datos_cargados = True
                    st.success("✅ Listo!")
        
        st.markdown("---")
        
        if st.session_state.datos_cargados:
            st.info(f"📊 {len(st.session_state.df):,} registros")
    
    # Main content
    if not st.session_state.datos_cargados:
        st.markdown("""
        <div style="text-align: center; padding: 2rem;">
            <h1>📊 Dashboard de Ventas</h1>
            <p style="color: #64748b;">Carga tus archivos Excel en el menú lateral</p>
        </div>
        """, unsafe_allow_html=True)
    else:
        # Aplicar filtros
        filtro, filtros_aplicados = aplicar_filtros(st.session_state.df)
        
        if len(filtro) == 0:
            st.warning("⚠️ No hay datos con estos filtros")
            return
        
        # Título
        st.markdown("# 📊 Dashboard de Ventas")
        st.caption(f"📅 {datetime.now().strftime('%Y-%m-%d %H:%M')}")
        
        # Mostrar filtros activos
        with st.expander("🔍 Filtros activos", expanded=False):
            for key, value in filtros_aplicados.items():
                st.caption(f"**{key}:** {value}")
        
        st.markdown("---")
        
        # KPIs
        kpis = mostrar_kpis(filtro)
        
        # Exportar
        botones_exportacion(filtro, kpis, filtros_aplicados)
        
        st.markdown("---")
        
        # Gráficos
        graficos_principales(filtro)
        
        st.markdown("---")
        
        # Top productos
        mostrar_top_productos(filtro)
        
        st.markdown("---")
        
        # Tabla detallada (limitada para rendimiento)
        st.markdown("### 📋 Detalle")
        
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("📅 Días", filtro['fecha'].nunique())
        with col2:
            st.metric("📦 Productos", filtro['producto'].nunique())
        with col3:
            st.metric("⭐ Promedio", f"{filtro['cantidad'].mean():.0f}")
        
        # Mostrar solo las primeras 1000 filas para rendimiento
        columnas_mostrar = ['fecha', 'producto', 'marca', 'proveedor', 'cantidad', 'anio', 'mes_nombre']
        columnas_disponibles = [col for col in columnas_mostrar if col in filtro.columns]
        
        df_mostrar = filtro[columnas_disponibles].head(1000).copy()
        
        st.dataframe(
            df_mostrar.sort_values('fecha', ascending=False),
            use_container_width=True,
            height=400
        )
        
        # Footer
        st.markdown("""
        <div class="footer">
            Dashboard de Ventas | Streamlit
        </div>
        """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()