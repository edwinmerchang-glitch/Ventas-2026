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
    
    /* Estilo para tabla */
    .dataframe {
        font-size: 0.9rem;
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
    }
</style>
""", unsafe_allow_html=True)

# ======================================
# FUNCIONES DE LIMPIEZA DE DATOS
# ======================================

@st.cache_data(ttl=3600)
def limpiar_dataframe(df):
    """Limpia el dataframe de valores erróneos"""
    
    # Limpiar columna de año
    if 'anio' in df.columns:
        df['anio'] = pd.to_numeric(df['anio'], errors='coerce')
    
    # Limpiar columna de cantidad
    if 'cantidad' in df.columns:
        df['cantidad'] = pd.to_numeric(df['cantidad'], errors='coerce').fillna(0)
    
    # Eliminar filas con fechas inválidas
    if 'fecha' in df.columns:
        df['fecha'] = pd.to_datetime(df['fecha'], errors='coerce')
        df = df.dropna(subset=['fecha'])
        
        # Recalcular año y mes
        df['anio'] = df['fecha'].dt.year
        df['mes'] = df['fecha'].dt.month
        df['mes_nombre'] = df['fecha'].dt.strftime('%B')
        df['dia'] = df['fecha'].dt.day
        df['trimestre'] = df['fecha'].dt.quarter
    
    # Limpiar valores nulos
    df['producto'] = df['producto'].fillna('No especificado').astype(str)
    df['marca'] = df['marca'].fillna('No especificada').astype(str)
    df['proveedor'] = df['proveedor'].fillna('No especificado').astype(str)
    
    # Eliminar filas donde el año sea inválido
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
                if 'producto' in col or 'item' in col or 'codigo' in col:
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
        
        # Aplicar limpieza de datos
        df = limpiar_dataframe(df)
        
        return df, True
    except Exception as e:
        return None, False

@st.cache_data(ttl=3600)
def cargar_datos_ejemplo():
    """Genera datos de ejemplo para demostración"""
    np.random.seed(42)
    
    # Crear datos similares a la imagen
    productos = [
        "019962797569 - CREATINE HEALTHY S.MONOHYD.3000MG X300G",
        "0300651481228 - SYSTANE COMPLETE GOTAS.OFT.0.6X10ML",
        "0300653610794 - SOLUC.DESINFEC.OPTI-FREE P.LENTES X90ML",
        "039800088635 - BATERIA ENERGIZER BOTON DE LITIO CR2032",
        "054402330821 - BRONCE.AUST.GOLD SPR.GEL FPS30 X237ML",
        "0700083725400 - COLAGENO BELFAN HIDROLIZADO VLLA.X600G",
        "0715134328936 - PROBIOTICO JARROW EPS CAP.X30",
        "0715134328974 - JARRO-DOPHILUS EPS CAP.X60",
        "0715134329087 - B-FORMULA CAP.X100",
        "0715134329148 - OMEGA 3 CARLSON LIQUIDO NORUEGO X500ML",
        "1000025250700 - REP BAND BANDA ELAST.#2 NARANJAX1.5MT",
        "1000025250717 - REP BAND BANDA FI AST.#3 VFRDFX1.5MT"
    ]
    
    marcas = ["COLGATE", "GENOMMA LAB", "HALEON", "ENERGIZER", "BELFAN", "JARROW"]
    proveedores = ["COLGATE PALMOLIVE CIA", "GENOMMA LAB COLOMBIA LTDA", "HALEON COLOMBIA S.A.S"]
    
    data = []
    
    # Generar datos para 2024
    for producto in productos:
        ventas_2024 = np.random.randint(1, 40)
        for _ in range(ventas_2024):
            fecha = pd.Timestamp(f"2024-{np.random.randint(1, 13)}-{np.random.randint(1, 28)}")
            data.append({
                'fecha': fecha,
                'producto': producto,
                'marca': np.random.choice(marcas),
                'proveedor': np.random.choice(proveedores),
                'cantidad': 1
            })
    
    # Generar datos para 2025
    for producto in productos:
        ventas_2025 = np.random.randint(1, 10)
        for _ in range(ventas_2025):
            fecha = pd.Timestamp(f"2025-{np.random.randint(1, 13)}-{np.random.randint(1, 28)}")
            data.append({
                'fecha': fecha,
                'producto': producto,
                'marca': np.random.choice(marcas),
                'proveedor': np.random.choice(proveedores),
                'cantidad': 1
            })
    
    df = pd.DataFrame(data)
    df = limpiar_dataframe(df)
    
    return df

# ======================================
# FUNCIÓN PARA CREAR TABLA COMPARATIVA
# ======================================

@st.cache_data(ttl=300)
def crear_tabla_comparativa(df, top_n=None):
    """Crea tabla comparativa de productos similar a Power BI"""
    
    # Agrupar por producto y año
    ventas_producto = df.groupby(['producto', 'anio'])['cantidad'].sum().reset_index()
    
    # Pivotar para tener 2024 y 2025 como columnas
    tabla_pivot = ventas_producto.pivot(index='producto', columns='anio', values='cantidad').fillna(0)
    
    # Renombrar columnas
    tabla_pivot.columns = [f'ventas_{int(col)}' for col in tabla_pivot.columns]
    
    # Asegurar que existen ambas columnas
    if 'ventas_2024' not in tabla_pivot.columns:
        tabla_pivot['ventas_2024'] = 0
    if 'ventas_2025' not in tabla_pivot.columns:
        tabla_pivot['ventas_2025'] = 0
    
    # Calcular diferencia y variación
    tabla_pivot['diferencia'] = tabla_pivot['ventas_2025'] - tabla_pivot['ventas_2024']
    tabla_pivot['variacion_porcentaje'] = np.where(
        tabla_pivot['ventas_2024'] > 0,
        (tabla_pivot['diferencia'] / tabla_pivot['ventas_2024']) * 100,
        0
    )
    
    # Reset index para tener producto como columna
    tabla_comparativa = tabla_pivot.reset_index()
    
    # Renombrar columnas para mejor presentación
    tabla_comparativa.columns = ['producto', 'ventas_2024', 'ventas_2025', 'diferencia', 'variacion_porcentaje']
    
    # Ordenar por ventas 2024 (mayor a menor)
    tabla_comparativa = tabla_comparativa.sort_values('ventas_2024', ascending=False)
    
    # Filtrar productos con ventas > 0
    tabla_comparativa = tabla_comparativa[(tabla_comparativa['ventas_2024'] > 0) | (tabla_comparativa['ventas_2025'] > 0)]
    
    # Limitar a top N si se especifica
    if top_n and top_n != "Todos":
        tabla_comparativa = tabla_comparativa.head(top_n)
    
    # Agregar fila de total
    total_ventas_2024 = tabla_comparativa['ventas_2024'].sum()
    total_ventas_2025 = tabla_comparativa['ventas_2025'].sum()
    total_diferencia = total_ventas_2025 - total_ventas_2024
    total_variacion = ((total_ventas_2025 - total_ventas_2024) / total_ventas_2024 * 100) if total_ventas_2024 > 0 else 0
    
    total_row = pd.DataFrame({
        'producto': ['**TOTAL**'],
        'ventas_2024': [total_ventas_2024],
        'ventas_2025': [total_ventas_2025],
        'diferencia': [total_diferencia],
        'variacion_porcentaje': [total_variacion]
    })
    
    tabla_comparativa = pd.concat([tabla_comparativa, total_row], ignore_index=True)
    
    return tabla_comparativa

# ======================================
# FUNCIONES DE EXPORTACIÓN
# ======================================

def exportar_excel(df, tabla_comparativa):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # Hoja de detalle de ventas
        df.head(10000).to_excel(writer, sheet_name='Detalle_Ventas', index=False)
        
        # Hoja de tabla comparativa
        tabla_comparativa.to_excel(writer, sheet_name='Comparativo_Productos', index=False)
        
        # Dar formato
        workbook = writer.book
        header_format = workbook.add_format({
            'bold': True,
            'bg_color': '#3b82f6',
            'font_color': 'white'
        })
        
        for sheet_name in writer.sheets:
            worksheet = writer.sheets[sheet_name]
            for col_num, value in enumerate(tabla_comparativa.columns):
                worksheet.write(0, col_num, value, header_format)
                worksheet.set_column(col_num, col_num, 20)
    
    output.seek(0)
    return output

def exportar_pdf(tabla_comparativa, kpis, filtros_aplicados):
    pdf = FPDF()
    pdf.add_page()
    
    # Configurar estilo
    pdf.set_fill_color(59, 130, 246)
    pdf.set_text_color(30, 41, 59)
    
    # Título
    pdf.set_font("Arial", "B", 20)
    pdf.cell(0, 15, "Reporte Comparativo de Ventas por Producto", 0, 1, "C")
    pdf.set_font("Arial", "", 9)
    pdf.cell(0, 8, f"Generado: {datetime.now().strftime('%Y-%m-%d %H:%M')}", 0, 1, "C")
    pdf.ln(5)
    
    # KPIs
    pdf.set_font("Arial", "B", 12)
    pdf.cell(0, 8, "Indicadores Clave", 0, 1)
    pdf.set_font("Arial", "", 10)
    pdf.cell(0, 7, f"Total Ventas 2024: {kpis['v2024']:,} unidades", 0, 1)
    pdf.cell(0, 7, f"Total Ventas 2025: {kpis['v2025']:,} unidades", 0, 1)
    pdf.cell(0, 7, f"Crecimiento: {kpis['crecimiento']:.1f}%", 0, 1)
    
    pdf.ln(8)
    
    # Tabla comparativa (top 20 para PDF)
    pdf.set_font("Arial", "B", 9)
    pdf.cell(80, 8, "Producto", 1, 0, 'C')
    pdf.cell(25, 8, "Ventas 2024", 1, 0, 'C')
    pdf.cell(25, 8, "Ventas 2025", 1, 0, 'C')
    pdf.cell(25, 8, "Diferencia", 1, 0, 'C')
    pdf.cell(30, 8, "Variación %", 1, 1, 'C')
    
    pdf.set_font("Arial", "", 8)
    for _, row in tabla_comparativa.head(25).iterrows():
        producto = row['producto'][:50] if row['producto'] != '**TOTAL**' else row['producto']
        pdf.cell(80, 7, producto, 1, 0)
        pdf.cell(25, 7, f"{int(row['ventas_2024']):,}", 1, 0, 'R')
        pdf.cell(25, 7, f"{int(row['ventas_2025']):,}", 1, 0, 'R')
        pdf.cell(25, 7, f"{int(row['diferencia']):,}", 1, 0, 'R')
        
        # Color para variación
        variacion = row['variacion_porcentaje']
        variacion_str = f"{variacion:.1f}%"
        pdf.cell(30, 7, variacion_str, 1, 1, 'R')
    
    return pdf.output(dest='S').encode('latin1')

# ======================================
# FILTROS AVANZADOS
# ======================================

def aplicar_filtros(df):
    st.sidebar.markdown("### 📊 Panel de Control")
    st.sidebar.markdown("---")
    
    # Filtro de año
    anios = sorted(df['anio'].dropna().unique())
    anios = [a for a in anios if isinstance(a, (int, float)) and a >= 2000]
    
    anio_seleccionado = st.sidebar.multiselect(
        "📅 Años",
        options=anios,
        default=anios if anios else [],
        key="anios"
    )
    
    # Filtro de marca
    marcas_disponibles = sorted([m for m in df['marca'].unique() if m not in ['No especificada', 'General']])
    marcas_seleccionadas = st.sidebar.multiselect(
        "🏷️ Marcas",
        options=marcas_disponibles,
        default=marcas_disponibles if marcas_disponibles else [],
        key="marcas"
    )
    
    # Filtro de proveedor
    proveedores_disponibles = sorted([p for p in df['proveedor'].unique() if p not in ['No especificado', 'General']])
    proveedores_seleccionados = st.sidebar.multiselect(
        "🏭 Proveedores",
        options=proveedores_disponibles,
        default=proveedores_disponibles if proveedores_disponibles else [],
        key="proveedores"
    )
    
    st.sidebar.markdown("---")
    
    # Aplicar filtros
    filtro = df.copy()
    
    if anio_seleccionado:
        filtro = filtro[filtro['anio'].isin(anio_seleccionado)]
    
    if marcas_seleccionadas:
        filtro = filtro[filtro['marca'].isin(marcas_seleccionadas)]
    
    if proveedores_seleccionados:
        filtro = filtro[filtro['proveedor'].isin(proveedores_seleccionados)]
    
    # Estadísticas
    with st.sidebar.expander("📈 Estadísticas", expanded=False):
        st.metric("Registros", f"{len(filtro):,}")
        st.metric("Productos", filtro['producto'].nunique())
        st.metric("Ventas totales", f"{int(filtro['cantidad'].sum()):,}")
    
    filtros_dict = {
        'Años': ', '.join([str(a) for a in anio_seleccionado]) if anio_seleccionado else "Todos",
        'Marcas': ', '.join(marcas_seleccionadas[:3]) + ('...' if len(marcas_seleccionadas) > 3 else '') if marcas_seleccionadas else "Todas",
        'Proveedores': ', '.join(proveedores_seleccionados[:3]) + ('...' if len(proveedores_seleccionados) > 3 else '') if proveedores_seleccionados else "Todos",
    }
    
    return filtro, filtros_dict

# ======================================
# KPIS
# ======================================

def mostrar_kpis(filtro):
    ventas_total = int(filtro['cantidad'].sum())
    ventas_2024 = int(filtro[filtro['anio'] == 2024]['cantidad'].sum()) if 2024 in filtro['anio'].values else 0
    ventas_2025 = int(filtro[filtro['anio'] == 2025]['cantidad'].sum()) if 2025 in filtro['anio'].values else 0
    
    if ventas_2024 > 0:
        crecimiento = ((ventas_2025 - ventas_2024) / ventas_2024) * 100
    else:
        crecimiento = 0
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.markdown(f"""
        <div class="kpi">
            <h3>📊 Total Unidades</h3>
            <div class="value">{ventas_total:,}</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        st.markdown(f"""
        <div class="kpi">
            <h3>📈 Ventas 2024</h3>
            <div class="value">{ventas_2024:,}</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col3:
        st.markdown(f"""
        <div class="kpi">
            <h3>🚀 Ventas 2025</h3>
            <div class="value">{ventas_2025:,}</div>
            <div class="growth {'growth-positive' if crecimiento >= 0 else 'growth-negative'}">
                {crecimiento:.1f}%
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
# TABLA COMPARATIVA ESTILO POWER BI
# ======================================

def mostrar_tabla_comparativa(tabla_comparativa):
    st.markdown("### 📊 Comparativo de Ventas por Producto")
    st.markdown("*Análisis 2024 vs 2025 - Estilo Power BI*")
    
    # Selector de cantidad de filas
    col1, col2 = st.columns([3, 1])
    with col2:
        top_n = st.selectbox("Mostrar top", [20, 50, 100, "Todos"], index=0)
    
    # Filtrar tabla según selección
    if top_n != "Todos":
        # Excluir total temporalmente para el filtro
        sin_total = tabla_comparativa[tabla_comparativa['producto'] != '**TOTAL**'].copy()
        sin_total_filtrado = sin_total.head(top_n)
        total_row = tabla_comparativa[tabla_comparativa['producto'] == '**TOTAL**'].copy()
        tabla_mostrar = pd.concat([sin_total_filtrado, total_row], ignore_index=True)
    else:
        tabla_mostrar = tabla_comparativa.copy()
    
    # Formatear la tabla para mostrar
    tabla_formateada = tabla_mostrar.copy()
    tabla_formateada['ventas_2024'] = tabla_formateada['ventas_2024'].apply(lambda x: f"{int(x):,}")
    tabla_formateada['ventas_2025'] = tabla_formateada['ventas_2025'].apply(lambda x: f"{int(x):,}")
    tabla_formateada['diferencia'] = tabla_formateada['diferencia'].apply(lambda x: f"{int(x):,}")
    tabla_formateada['variacion_porcentaje'] = tabla_formateada['variacion_porcentaje'].apply(
        lambda x: f"{x:.1f}%"
    )
    
    # Renombrar columnas para mejor visualización
    tabla_formateada.columns = ['Producto', 'Ventas 2024', 'Ventas 2025', 'Diferencia', 'Variación %']
    
    # Aplicar formato condicional con colores usando map (nueva forma)
    def color_variacion_series(val):
        if isinstance(val, str) and '%' in val and val != '0.0%':
            try:
                num = float(val.replace('%', ''))
                if num < 0:
                    return 'color: #dc2626; font-weight: bold'
                elif num > 0:
                    return 'color: #10b981; font-weight: bold'
            except:
                pass
        return ''
    
    # Aplicar estilo usando map (método actualizado)
    styled_df = tabla_formateada.style.map(color_variacion_series, subset=['Variación %'])
    
    # Configurar formato de números
    styled_df = styled_df.format({
        'Ventas 2024': lambda x: x,
        'Ventas 2025': lambda x: x,
        'Diferencia': lambda x: x,
        'Variación %': lambda x: x
    })
    
    # Mostrar tabla
    st.dataframe(
        styled_df,
        use_container_width=True,
        height=500,
        hide_index=True
    )
    
    # Mostrar resumen
    total_row_data = tabla_mostrar[tabla_mostrar['Producto'] == '**TOTAL**']
    if len(total_row_data) > 0:
        st.info(f"📊 **Resumen General:** Total ventas 2024: {total_row_data['Ventas 2024'].values[0]} | "
               f"Total ventas 2025: {total_row_data['Ventas 2025'].values[0]} | "
               f"Variación total: {total_row_data['Variación %'].values[0]}")

# ======================================
# GRÁFICO DE BARRAS COMPARATIVO
# ======================================

def mostrar_grafico_comparativo(tabla_comparativa):
    st.markdown("### 📈 Top Productos con Mayor Volumen de Ventas")
    
    # Excluir total y tomar top 10
    top_productos = tabla_comparativa[
        (tabla_comparativa['producto'] != '**TOTAL**') & 
        (tabla_comparativa['ventas_2024'] > 0)
    ].head(10).copy()
    
    if len(top_productos) > 0:
        fig = go.Figure()
        
        # Acortar nombres de productos para mejor visualización
        productos_short = top_productos['producto'].apply(lambda x: x[:40] + '...' if len(x) > 40 else x)
        
        # Barras para 2024
        fig.add_trace(go.Bar(
            name='Ventas 2024',
            x=productos_short,
            y=top_productos['ventas_2024'],
            marker_color='#3b82f6',
            text=top_productos['ventas_2024'],
            textposition='outside'
        ))
        
        # Barras para 2025
        fig.add_trace(go.Bar(
            name='Ventas 2025',
            x=productos_short,
            y=top_productos['ventas_2025'],
            marker_color='#ef4444',
            text=top_productos['ventas_2025'],
            textposition='outside'
        ))
        
        fig.update_layout(
            title="Comparativo de Ventas - Top 10 Productos",
            template='plotly_white',
            height=500,
            barmode='group',
            xaxis_title="Producto",
            yaxis_title="Unidades Vendidas",
            xaxis_tickangle=-45
        )
        
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("No hay datos suficientes para mostrar el gráfico comparativo")

# ======================================
# BOTONES DE EXPORTACIÓN
# ======================================

def botones_exportacion(df, tabla_comparativa, kpis, filtros_aplicados):
    col1, col2 = st.columns(2)
    
    with col1:
        if st.button("📊 Exportar a Excel", use_container_width=True):
            with st.spinner("Generando Excel..."):
                excel_data = exportar_excel(df, tabla_comparativa)
                b64 = base64.b64encode(excel_data.getvalue()).decode()
                href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="reporte_comparativo.xlsx">📥 Descargar Excel</a>'
                st.markdown(href, unsafe_allow_html=True)
                st.success("✅ Reporte generado!")
    
    with col2:
        if st.button("📄 Exportar a PDF", use_container_width=True):
            with st.spinner("Generando PDF..."):
                pdf_data = exportar_pdf(tabla_comparativa, kpis, filtros_aplicados)
                b64 = base64.b64encode(pdf_data).decode()
                href = f'<a href="data:application/pdf;base64,{b64}" download="reporte_comparativo.pdf">📥 Descargar PDF</a>'
                st.markdown(href, unsafe_allow_html=True)
                st.success("✅ Reporte generado!")

# ======================================
# MENÚ PRINCIPAL
# ======================================

def main():
    # Inicializar estado de sesión
    if 'datos_cargados' not in st.session_state:
        st.session_state.datos_cargados = False
        st.session_state.df = None
    
    # Sidebar
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
                            st.success("✅ Datos cargados!")
                        else:
                            st.error("Error al cargar")
                else:
                    st.warning("Selecciona ambos archivos")
        
        with col2:
            if st.button("📊 Ejemplo", use_container_width=True):
                with st.spinner("Cargando ejemplo..."):
                    st.session_state.df = cargar_datos_ejemplo()
                    st.session_state.datos_cargados = True
                    st.success("✅ Datos de ejemplo cargados!")
        
        st.markdown("---")
        
        if st.session_state.datos_cargados:
            st.info(f"📊 {len(st.session_state.df):,} registros | {st.session_state.df['producto'].nunique()} productos")
    
    # Main content
    if not st.session_state.datos_cargados:
        st.markdown("""
        <div style="text-align: center; padding: 3rem;">
            <h1>📊 Dashboard Comparativo de Ventas</h1>
            <p style="color: #64748b; font-size: 1.1rem;">
                Análisis 2024 vs 2025 por producto - Estilo Power BI
            </p>
            <div style="margin-top: 2rem;">
                <p style="color: #3b82f6;">Carga tus archivos Excel en el menú lateral</p>
            </div>
        </div>
        """, unsafe_allow_html=True)
        
        with st.expander("📖 Instrucciones", expanded=False):
            st.markdown("""
            ### Cómo usar:
            1. Carga los archivos Excel de ventas 2024 y 2025
            2. El dashboard generará automáticamente la tabla comparativa
            3. Puedes filtrar por marca o proveedor
            4. Exporta los resultados a Excel o PDF
            
            ### Columnas necesarias:
            - Fecha (fecha, date)
            - Producto (producto, codigo, item)
            - Cantidad (cantidad, venta)
            - Marca (opcional)
            - Proveedor (opcional)
            """)
    else:
        # Aplicar filtros
        filtro, filtros_aplicados = aplicar_filtros(st.session_state.df)
        
        if len(filtro) == 0:
            st.warning("⚠️ No hay datos con los filtros seleccionados")
            return
        
        # Crear tabla comparativa
        tabla_comparativa = crear_tabla_comparativa(filtro, top_n=None)
        
        # Título
        st.markdown("# 📊 Dashboard Comparativo de Ventas")
        st.caption(f"📅 Actualizado: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        
        # Mostrar filtros activos
        with st.expander("🔍 Filtros activos", expanded=False):
            for key, value in filtros_aplicados.items():
                st.caption(f"**{key}:** {value}")
        
        st.markdown("---")
        
        # KPIs
        kpis = mostrar_kpis(filtro)
        
        # Exportar
        botones_exportacion(filtro, tabla_comparativa, kpis, filtros_aplicados)
        
        st.markdown("---")
        
        # Tabla comparativa principal
        mostrar_tabla_comparativa(tabla_comparativa)
        
        st.markdown("---")
        
        # Gráfico comparativo
        mostrar_grafico_comparativo(tabla_comparativa)
        
        st.markdown("---")
        
        # Footer
        st.markdown("""
        <div class="footer">
            Dashboard Comparativo de Ventas | Análisis 2024 vs 2025 | Desarrollado con Streamlit
        </div>
        """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()