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
        font-size: 2.5rem;
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
        padding: 1.5rem;
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
        font-size: 0.875rem;
        text-transform: uppercase;
        letter-spacing: 1px;
        margin-bottom: 0.5rem;
    }
    
    .kpi .value {
        font-size: 2rem;
        font-weight: bold;
        color: #1e293b;
        margin: 0.5rem 0;
    }
    
    .kpi .growth {
        font-size: 0.875rem;
        padding: 0.25rem 0.5rem;
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
        padding: 0.5rem 1rem;
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
    }
    
    /* Sidebar text color */
    section[data-testid="stSidebar"] .stMarkdown {
        color: #e2e8f0;
    }
    
    section[data-testid="stSidebar"] h3 {
        color: #ffffff !important;
    }
    
    /* Cards de métricas */
    .metric-card {
        background: white;
        border-radius: 8px;
        padding: 1rem;
        box-shadow: 0 1px 2px rgba(0,0,0,0.05);
        text-align: center;
    }
    
    /* Footer */
    .footer {
        text-align: center;
        padding: 1.5rem;
        color: #94a3b8;
        border-top: 1px solid #e2e8f0;
        margin-top: 2rem;
    }
    
    /* Mensajes de éxito/error */
    .stAlert {
        border-radius: 8px;
    }
    
    /* Filtros en sidebar */
    .filter-section {
        margin-bottom: 1rem;
    }
    
    /* Responsive */
    @media (max-width: 768px) {
        .kpi .value {
            font-size: 1.5rem;
        }
        h1 {
            font-size: 1.75rem;
        }
    }
</style>
""", unsafe_allow_html=True)

# ======================================
# FUNCIONES DE CARGA DE ARCHIVOS
# ======================================

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
        
        # Procesar fechas
        df['fecha'] = pd.to_datetime(df['fecha'], errors='coerce')
        df = df.dropna(subset=['fecha'])
        
        # Procesar cantidades
        df['cantidad'] = pd.to_numeric(df['cantidad'], errors='coerce').fillna(0)
        
        # Crear columnas adicionales
        df['anio'] = df['fecha'].dt.year
        df['mes'] = df['fecha'].dt.month
        df['dia'] = df['fecha'].dt.day
        df['mes_nombre'] = df['fecha'].dt.strftime('%B')
        df['semana'] = df['fecha'].dt.isocalendar().week
        df['trimestre'] = df['fecha'].dt.quarter
        
        # Buscar o crear columna de producto
        if 'producto' not in df.columns:
            for col in columnas:
                if 'producto' in col or 'item' in col or 'articulo' in col:
                    df.rename(columns={col: 'producto'}, inplace=True)
                    break
            else:
                df['producto'] = 'Producto General'
        
        # Buscar o crear columna de categoría/marca
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
        
        return df, True, "✅ Datos cargados correctamente"
    except Exception as e:
        return None, False, f"❌ Error al cargar archivo: {str(e)}"

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
        for _ in range(np.random.randint(1, 5)):
            data.append({
                'fecha': fecha,
                'producto': np.random.choice(productos),
                'marca': np.random.choice(marcas),
                'proveedor': np.random.choice(proveedores),
                'categoria': np.random.choice(categorias),
                'cantidad': np.random.randint(1, 100)
            })
    
    df = pd.DataFrame(data)
    df['anio'] = df['fecha'].dt.year
    df['mes'] = df['fecha'].dt.month
    df['dia'] = df['fecha'].dt.day
    df['mes_nombre'] = df['fecha'].dt.strftime('%B')
    df['semana'] = df['fecha'].dt.isocalendar().week
    df['trimestre'] = df['fecha'].dt.quarter
    
    return df

# ======================================
# FUNCIONES DE EXPORTACIÓN
# ======================================

def exportar_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='Ventas', index=False)
        
        workbook = writer.book
        worksheet = writer.sheets['Ventas']
        
        # Formato corporativo
        header_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'top',
            'fg_color': '#3b82f6',
            'font_color': 'white',
            'border': 1
        })
        
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)
            worksheet.set_column(col_num, col_num, 15)
    
    output.seek(0)
    return output

def exportar_pdf(df, kpis, filtros_aplicados):
    pdf = FPDF()
    pdf.add_page()
    
    # Configurar estilo corporativo
    pdf.set_fill_color(59, 130, 246)
    pdf.set_text_color(30, 41, 59)
    
    # Logo y título
    pdf.set_font("Arial", "B", 24)
    pdf.cell(0, 20, "Reporte Ejecutivo de Ventas", 0, 1, "C")
    pdf.set_font("Arial", "", 10)
    pdf.cell(0, 10, f"Generado: {datetime.now().strftime('%Y-%m-%d %H:%M')}", 0, 1, "C")
    pdf.ln(5)
    
    # Filtros aplicados
    pdf.set_font("Arial", "B", 12)
    pdf.cell(0, 10, "Filtros Aplicados:", 0, 1)
    pdf.set_font("Arial", "", 10)
    for key, value in filtros_aplicados.items():
        pdf.cell(0, 6, f"{key}: {value}", 0, 1)
    
    pdf.ln(10)
    
    # KPIs
    pdf.set_font("Arial", "B", 14)
    pdf.cell(0, 10, "Indicadores Clave", 0, 1)
    pdf.set_font("Arial", "", 12)
    pdf.cell(0, 8, f"Total Ventas: {kpis['total']:,} unidades", 0, 1)
    pdf.cell(0, 8, f"Ventas 2024: {kpis['v2024']:,} unidades", 0, 1)
    pdf.cell(0, 8, f"Ventas 2025: {kpis['v2025']:,} unidades", 0, 1)
    pdf.cell(0, 8, f"Crecimiento: {kpis['crecimiento']:.1f}%", 0, 1)
    
    return pdf.output(dest='S').encode('latin1')

# ======================================
# FILTROS AVANZADOS
# ======================================

def aplicar_filtros(df):
    st.sidebar.markdown("### 📊 Panel de Control")
    st.sidebar.markdown("---")
    
    # ===== SECCIÓN FECHAS =====
    st.sidebar.markdown("#### 📅 Filtros de Fecha")
    
    # Rango de fechas
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
    
    # Filtro de año
    anios = sorted(df['anio'].unique())
    anio_seleccionado = st.sidebar.multiselect(
        "📅 Años",
        anios,
        default=anios,
        key="anios"
    )
    
    # Filtro de mes
    meses = sorted(df['mes'].unique())
    nombres_meses = {
        1: 'Enero', 2: 'Febrero', 3: 'Marzo', 4: 'Abril',
        5: 'Mayo', 6: 'Junio', 7: 'Julio', 8: 'Agosto',
        9: 'Septiembre', 10: 'Octubre', 11: 'Noviembre', 12: 'Diciembre'
    }
    meses_nombres = [f"{m} - {nombres_meses[m]}" for m in meses]
    meses_seleccionados = st.sidebar.multiselect(
        "📆 Meses",
        meses_nombres,
        default=meses_nombres,
        key="meses"
    )
    
    # Convertir selección de meses
    meses_filtro = [int(m.split(' - ')[0]) for m in meses_seleccionados]
    
    st.sidebar.markdown("---")
    
    # ===== SECCIÓN MARCAS =====
    st.sidebar.markdown("#### 🏷️ Filtros de Marca")
    
    # Obtener marcas únicas
    marcas = sorted(df['marca'].unique())
    marcas_seleccionadas = st.sidebar.multiselect(
        "🎯 Marcas",
        marcas,
        default=marcas,
        key="marcas",
        help="Filtrar por marca específica"
    )
    
    st.sidebar.markdown("---")
    
    # ===== SECCIÓN PROVEEDORES =====
    st.sidebar.markdown("#### 🏭 Filtros de Proveedor")
    
    # Obtener proveedores únicos
    proveedores = sorted(df['proveedor'].unique())
    proveedores_seleccionados = st.sidebar.multiselect(
        "📦 Proveedores",
        proveedores,
        default=proveedores,
        key="proveedores",
        help="Filtrar por proveedor específico"
    )
    
    st.sidebar.markdown("---")
    
    # ===== SECCIÓN PRODUCTOS =====
    st.sidebar.markdown("#### 📦 Filtros de Producto")
    
    # Filtro de producto
    productos = sorted(df['producto'].unique())
    producto_seleccionado = st.sidebar.multiselect(
        "📦 Productos",
        productos,
        default=productos,
        key="productos"
    )
    
    # Filtro de categoría
    categorias = sorted(df['categoria'].unique())
    categoria_seleccionada = st.sidebar.multiselect(
        "🏷️ Categorías",
        categorias,
        default=categorias,
        key="categorias"
    )
    
    st.sidebar.markdown("---")
    
    # Mostrar resumen de filtros
    st.sidebar.markdown("#### 📊 Resumen de Filtros")
    
    # Aplicar todos los filtros
    filtro = df.copy()
    
    # Aplicar filtro de rango de fechas
    if len(rango_fechas) == 2:
        filtro = filtro[
            (filtro['fecha'].dt.date >= rango_fechas[0]) &
            (filtro['fecha'].dt.date <= rango_fechas[1])
        ]
    
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
    with st.sidebar.expander("📈 Estadísticas de Filtros", expanded=False):
        st.metric("Registros filtrados", f"{len(filtro):,}")
        st.metric("Ventas totales", f"{int(filtro['cantidad'].sum()):,}")
        st.metric("Marcas seleccionadas", len(marcas_seleccionadas))
        st.metric("Proveedores seleccionados", len(proveedores_seleccionados))
    
    # Guardar filtros aplicados para exportación - Versión corregida
    filtros_dict = {
        'Rango de fechas': f"{rango_fechas[0]} a {rango_fechas[1]}" if len(rango_fechas) == 2 else "Todos",
        'Años': ', '.join([str(a) for a in anio_seleccionado]) if anio_seleccionado else "Todos",
        'Meses': ', '.join([nombres_meses[m] for m in meses_filtro]) if meses_filtro else "Todos",
        'Marcas': ', '.join([str(m) for m in marcas_seleccionadas]) if marcas_seleccionadas else "Todas",
        'Proveedores': ', '.join([str(p) for p in proveedores_seleccionados]) if proveedores_seleccionados else "Todos",
        'Productos': ', '.join([str(prod) for prod in producto_seleccionado]) if producto_seleccionado else "Todos",
        'Categorías': ', '.join([str(cat) for cat in categoria_seleccionada]) if categoria_seleccionada else "Todas"
    }
    
    return filtro, filtros_dict

# ======================================
# KPIS
# ======================================

def mostrar_kpis(filtro):
    ventas_total = int(filtro['cantidad'].sum())
    ventas_2024 = int(filtro[filtro['anio'] == 2024]['cantidad'].sum())
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
            <h3>📊 Total Unidades</h3>
            <div class="value">{ventas_total:,}</div>
            <div class="growth growth-positive">Período seleccionado</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        st.markdown(f"""
        <div class="kpi">
            <h3>📈 Ventas 2024</h3>
            <div class="value">{ventas_2024:,}</div>
            <div class="growth growth-positive">Año base</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col3:
        st.markdown(f"""
        <div class="kpi">
            <h3>🚀 Ventas 2025</h3>
            <div class="value">{ventas_2025:,}</div>
            <div class="growth {'growth-positive' if crecimiento >= 0 else 'growth-negative'}">
                {'📈' if crecimiento >= 0 else '📉'} {crecimiento:.1f}% vs 2024
            </div>
        </div>
        """, unsafe_allow_html=True)
    
    with col4:
        st.markdown(f"""
        <div class="kpi">
            <h3>🏷️ Marcas</h3>
            <div class="value">{marcas_unicas}</div>
            <div class="growth growth-positive">Activas</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col5:
        st.markdown(f"""
        <div class="kpi">
            <h3>🏭 Proveedores</h3>
            <div class="value">{proveedores_unicos}</div>
            <div class="growth growth-positive">Activos</div>
        </div>
        """, unsafe_allow_html=True)
    
    return {
        'total': ventas_total,
        'v2024': ventas_2024,
        'v2025': ventas_2025,
        'crecimiento': crecimiento,
        'marcas': marcas_unicas,
        'proveedores': proveedores_unicos
    }

# ======================================
# GRÁFICOS
# ======================================

def graficos_principales(filtro):
    # Gráfico comparativo por año
    st.markdown("### 📊 Comparativo Anual")
    
    ventas_anio = filtro.groupby('anio')['cantidad'].sum().reset_index()
    
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
        height=400,
        showlegend=False,
        xaxis_title="Año",
        yaxis_title="Unidades Vendidas"
    )
    st.plotly_chart(fig1, use_container_width=True)
    
    # Gráfico por mes
    st.markdown("### 📈 Tendencia Mensual Comparativa")
    
    ventas_mes = filtro.groupby(['anio', 'mes_nombre', 'mes'])['cantidad'].sum().reset_index()
    ventas_mes = ventas_mes.sort_values('mes')
    
    fig2 = px.line(
        ventas_mes,
        x='mes_nombre',
        y='cantidad',
        color='anio',
        markers=True,
        title="Comparativo Mensual 2024 vs 2025",
        color_discrete_sequence=['#3b82f6', '#10b981']
    )
    fig2.update_layout(
        template='plotly_white',
        height=450,
        xaxis_title="Mes",
        yaxis_title="Unidades Vendidas"
    )
    fig2.update_traces(marker_size=8)
    st.plotly_chart(fig2, use_container_width=True)
    
    # Gráfico por marca
    st.markdown("### 🏷️ Ventas por Marca")
    
    ventas_marca = filtro.groupby('marca')['cantidad'].sum().sort_values(ascending=False).head(10)
    
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
    fig3.update_layout(
        template='plotly_white',
        height=400,
        xaxis_title="Unidades Vendidas",
        yaxis_title="Marca"
    )
    st.plotly_chart(fig3, use_container_width=True)
    
    # Gráfico por proveedor
    st.markdown("### 🏭 Ventas por Proveedor")
    
    ventas_proveedor = filtro.groupby('proveedor')['cantidad'].sum().sort_values(ascending=False).head(10)
    
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
    fig4.update_layout(
        template='plotly_white',
        height=400,
        xaxis_title="Unidades Vendidas",
        yaxis_title="Proveedor"
    )
    st.plotly_chart(fig4, use_container_width=True)

def mostrar_top_productos(filtro):
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("### 🏆 Top 5 Productos")
        top_productos = filtro.groupby('producto')['cantidad'].sum().sort_values(ascending=False).head(5)
        
        fig = px.bar(
            top_productos,
            x=top_productos.values,
            y=top_productos.index,
            orientation='h',
            text_auto=True,
            color=top_productos.values,
            color_continuous_scale='Blues',
            title="Productos Más Vendidos"
        )
        fig.update_layout(template='plotly_white', height=400)
        st.plotly_chart(fig, use_container_width=True)
    
    with col2:
        st.markdown("### 🎯 Distribución por Categoría")
        top_categorias = filtro.groupby('categoria')['cantidad'].sum()
        
        fig = px.pie(
            values=top_categorias.values,
            names=top_categorias.index,
            title="Ventas por Categoría",
            color_discrete_sequence=px.colors.sequential.Blues_r
        )
        fig.update_layout(template='plotly_white', height=400)
        st.plotly_chart(fig, use_container_width=True)

# ======================================
# BOTONES DE EXPORTACIÓN
# ======================================

def botones_exportacion(filtro, kpis, filtros_aplicados):
    col1, col2 = st.columns(2)
    
    with col1:
        if st.button("📊 Exportar a Excel", use_container_width=True):
            excel_data = exportar_excel(filtro)
            b64 = base64.b64encode(excel_data.getvalue()).decode()
            href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="reporte_ventas.xlsx">📥 Descargar Excel</a>'
            st.markdown(href, unsafe_allow_html=True)
            st.success("✅ Reporte exportado a Excel")
    
    with col2:
        if st.button("📄 Exportar a PDF", use_container_width=True):
            pdf_data = exportar_pdf(filtro, kpis, filtros_aplicados)
            b64 = base64.b64encode(pdf_data).decode()
            href = f'<a href="data:application/pdf;base64,{b64}" download="reporte_ventas.pdf">📥 Descargar PDF</a>'
            st.markdown(href, unsafe_allow_html=True)
            st.success("✅ Reporte exportado a PDF")

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
        st.markdown("### 📁 Menú de Datos")
        
        # Opción 1: Cargar archivos
        st.markdown("#### Cargar archivos Excel")
        archivo_2024 = st.file_uploader("📂 Ventas 2024", type=['xlsx', 'xls'], key="2024")
        archivo_2025 = st.file_uploader("📂 Ventas 2025", type=['xlsx', 'xls'], key="2025")
        
        col1, col2 = st.columns(2)
        with col1:
            if st.button("🔄 Cargar datos", use_container_width=True):
                if archivo_2024 and archivo_2025:
                    df1, success1, msg1 = cargar_excel(archivo_2024)
                    df2, success2, msg2 = cargar_excel(archivo_2025)
                    
                    if success1 and success2:
                        st.session_state.df = pd.concat([df1, df2], ignore_index=True)
                        st.session_state.datos_cargados = True
                        st.success("✅ Datos cargados correctamente")
                    else:
                        st.error("Error al cargar los archivos")
                else:
                    st.warning("⚠️ Selecciona ambos archivos")
        
        with col2:
            if st.button("📊 Datos ejemplo", use_container_width=True):
                st.session_state.df = cargar_datos_ejemplo()
                st.session_state.datos_cargados = True
                st.success("✅ Datos de ejemplo cargados")
        
        st.markdown("---")
        
        # Mostrar estado de datos
        if st.session_state.datos_cargados:
            st.info(f"📊 Datos: {len(st.session_state.df):,} registros")
            
            # Mostrar columnas disponibles
            with st.expander("📋 Columnas disponibles"):
                for col in st.session_state.df.columns:
                    st.caption(f"• {col}")
    
    # Main content
    if not st.session_state.datos_cargados:
        # Pantalla de bienvenida
        st.markdown("""
        <div style="text-align: center; padding: 3rem;">
            <h1>📊 Dashboard Corporativo de Ventas</h1>
            <p style="color: #64748b; font-size: 1.1rem; margin-top: 1rem;">
                Sistema avanzado de análisis de ventas con filtros por marca y proveedor
            </p>
            <div style="margin-top: 2rem;">
                <p style="color: #3b82f6;">Para comenzar, carga tus archivos Excel en el menú lateral</p>
            </div>
        </div>
        """, unsafe_allow_html=True)
        
        # Mostrar instrucciones
        with st.expander("📖 Instrucciones", expanded=False):
            st.markdown("""
            ### Cómo usar el dashboard:
            
            1. **Cargar datos:** Usa el menú lateral para cargar archivos Excel
            2. **Formato esperado:** Los archivos deben tener columnas de fecha y cantidad
            3. **Filtros disponibles:** 
               - 📅 Fechas (rango, años, meses)
               - 🏷️ Marcas
               - 🏭 Proveedores
               - 📦 Productos
               - 🏷️ Categorías
            4. **Datos de ejemplo:** Puedes probar con los datos de ejemplo
            5. **Exportar:** Genera reportes en Excel o PDF
            
            ### Columnas recomendadas:
            - Fecha (fecha, date)
            - Cantidad/Ventas (cantidad, venta, monto)
            - Marca (marca, brand)
            - Proveedor (proveedor, supplier)
            - Producto (opcional)
            - Categoría (opcional)
            """)
    else:
        # Aplicar filtros
        filtro, filtros_aplicados = aplicar_filtros(st.session_state.df)
        
        if len(filtro) == 0:
            st.warning("⚠️ No hay datos con los filtros seleccionados. Por favor, ajusta los filtros.")
            return
        
        # Título principal
        st.markdown("# 📊 Dashboard Corporativo de Ventas")
        st.caption(f"📅 Actualizado: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        
        # Mostrar filtros activos
        with st.expander("🔍 Filtros activos", expanded=False):
            col1, col2 = st.columns(2)
            with col1:
                st.markdown("**📅 Fechas:**")
                st.caption(f"• {filtros_aplicados['Rango de fechas']}")
                st.caption(f"• Años: {filtros_aplicados['Años']}")
                st.caption(f"• Meses: {filtros_aplicados['Meses']}")
            with col2:
                st.markdown("**🏷️ Comerciales:**")
                st.caption(f"• Marcas: {filtros_aplicados['Marcas']}")
                st.caption(f"• Proveedores: {filtros_aplicados['Proveedores']}")
                st.caption(f"• Productos: {filtros_aplicados['Productos']}")
        
        st.markdown("---")
        
        # Mostrar KPIs
        kpis = mostrar_kpis(filtro)
        
        # Botones de exportación
        botones_exportacion(filtro, kpis, filtros_aplicados)
        
        st.markdown("---")
        
        # Gráficos principales
        graficos_principales(filtro)
        
        st.markdown("---")
        
        # Top productos
        mostrar_top_productos(filtro)
        
        st.markdown("---")
        
        # Tabla detallada
        st.markdown("### 📋 Detalle de Ventas")
        
        # Métricas adicionales
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("📅 Días analizados", filtro['fecha'].nunique())
        with col2:
            st.metric("📦 Productos únicos", filtro['producto'].nunique())
        with col3:
            st.metric("⭐ Promedio diario", f"{filtro['cantidad'].mean():,.0f}")
        
        # Mostrar dataframe con todas las columnas relevantes
        columnas_mostrar = ['fecha', 'producto', 'marca', 'proveedor', 'categoria', 'cantidad', 'anio', 'mes_nombre']
        columnas_disponibles = [col for col in columnas_mostrar if col in filtro.columns]
        
        st.dataframe(
            filtro[columnas_disponibles].sort_values('fecha', ascending=False),
            use_container_width=True,
            height=400
        )
        
        # Footer
        st.markdown("""
        <div class="footer">
            <p>Dashboard Corporativo de Ventas | Desarrollado con Streamlit | Filtros por Marca y Proveedor</p>
        </div>
        """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()