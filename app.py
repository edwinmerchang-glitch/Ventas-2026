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
# INICIALIZACIÓN DE ESTADO DE SESIÓN
# ======================================

def init_session_state():
    """Inicializa todas las variables de estado de sesión"""
    if 'datos_cargados' not in st.session_state:
        st.session_state.datos_cargados = False
    if 'df_combinado' not in st.session_state:
        st.session_state.df_combinado = None
    if 'df_2024' not in st.session_state:
        st.session_state.df_2024 = None
    if 'df_2025' not in st.session_state:
        st.session_state.df_2025 = None

# ======================================
# ESTILOS CORPORATIVOS PROFESIONALES
# ======================================

st.markdown("""
<style>
    .stApp {
        background: linear-gradient(135deg, #f5f7fa 0%, #ffffff 100%);
    }
    
    section[data-testid="stSidebar"] {
        background: linear-gradient(180deg, #1e293b 0%, #0f172a 100%);
        border-right: 1px solid #e2e8f0;
    }
    
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
    
    .footer {
        text-align: center;
        padding: 1rem;
        color: #94a3b8;
        border-top: 1px solid #e2e8f0;
        margin-top: 1rem;
        font-size: 0.8rem;
    }
    
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
    if df is None or len(df) == 0:
        return df
    
    # Limpiar columna de año
    if 'anio' in df.columns:
        df['anio'] = pd.to_numeric(df['anio'], errors='coerce')
    
    # Limpiar columna de cantidad
    if 'cantidad' in df.columns:
        df['cantidad'] = pd.to_numeric(df['cantidad'], errors='coerce').fillna(0)
    
    # Procesar fechas
    if 'fecha' in df.columns:
        df['fecha'] = pd.to_datetime(df['fecha'], errors='coerce')
        df = df.dropna(subset=['fecha'])
        df['anio'] = df['fecha'].dt.year
        df['mes'] = df['fecha'].dt.month
        df['mes_nombre'] = df['fecha'].dt.strftime('%B')
        df['dia'] = df['fecha'].dt.day
        df['trimestre'] = df['fecha'].dt.quarter
    
    # Eliminar filas con año inválido
    if 'anio' in df.columns:
        df = df.dropna(subset=['anio'])
        df = df[(df['anio'] >= 2000) & (df['anio'] <= 2030)]
    
    return df

@st.cache_data(ttl=3600)
def cargar_excel(file):
    """Carga archivo Excel y procesa los datos"""
    try:
        if file is None:
            return None, False
        df = pd.read_excel(file)
        df.columns = df.columns.str.strip().str.lower()
        return df, True
    except Exception as e:
        return None, False

# ======================================
# FUNCIÓN PARA CREAR TABLA COMPARATIVA
# ======================================

@st.cache_data(ttl=300)
def crear_tabla_comparativa(df, columna_producto, top_n=None):
    """Crea tabla comparativa de productos"""
    
    if df is None or len(df) == 0:
        return pd.DataFrame()
    
    # Verificar que la columna de producto existe
    if columna_producto not in df.columns:
        st.error(f"La columna '{columna_producto}' no existe en los datos")
        return pd.DataFrame()
    
    # Verificar que existe la columna de cantidad
    if 'cantidad' not in df.columns:
        st.error("La columna 'cantidad' no existe en los datos")
        return pd.DataFrame()
    
    # Verificar que existe la columna de año
    if 'anio' not in df.columns:
        st.error("La columna 'anio' no existe en los datos")
        return pd.DataFrame()
    
    # Agrupar por producto y año
    ventas_producto = df.groupby([columna_producto, 'anio'])['cantidad'].sum().reset_index()
    
    # Pivotar para tener 2024 y 2025 como columnas
    tabla_pivot = ventas_producto.pivot(index=columna_producto, columns='anio', values='cantidad').fillna(0)
    
    # Asegurar que existen ambas columnas
    if 2024 not in tabla_pivot.columns:
        tabla_pivot[2024] = 0
    if 2025 not in tabla_pivot.columns:
        tabla_pivot[2025] = 0
    
    # Renombrar columnas
    tabla_pivot.columns = ['ventas_2024', 'ventas_2025']
    
    # Calcular diferencia y variación
    tabla_pivot['diferencia'] = tabla_pivot['ventas_2025'] - tabla_pivot['ventas_2024']
    tabla_pivot['variacion_porcentaje'] = np.where(
        tabla_pivot['ventas_2024'] > 0,
        (tabla_pivot['diferencia'] / tabla_pivot['ventas_2024']) * 100,
        0
    )
    
    # Reset index
    tabla_comparativa = tabla_pivot.reset_index()
    tabla_comparativa = tabla_comparativa.rename(columns={columna_producto: 'producto'})
    
    # Filtrar productos con ventas > 0
    tabla_comparativa = tabla_comparativa[(tabla_comparativa['ventas_2024'] > 0) | (tabla_comparativa['ventas_2025'] > 0)]
    
    # Ordenar por ventas 2024
    tabla_comparativa = tabla_comparativa.sort_values('ventas_2024', ascending=False)
    
    # Limitar a top N
    if top_n and top_n != "Todos" and isinstance(top_n, int):
        tabla_comparativa = tabla_comparativa.head(top_n)
    
    # Agregar fila de total
    if len(tabla_comparativa) > 0:
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
        if df is not None and len(df) > 0:
            df.head(10000).to_excel(writer, sheet_name='Detalle_Ventas', index=False)
        if tabla_comparativa is not None and len(tabla_comparativa) > 0:
            tabla_comparativa.to_excel(writer, sheet_name='Comparativo_Productos', index=False)
    
    output.seek(0)
    return output

def exportar_pdf(tabla_comparativa, kpis):
    pdf = FPDF()
    pdf.add_page()
    
    pdf.set_fill_color(59, 130, 246)
    pdf.set_text_color(30, 41, 59)
    
    pdf.set_font("Arial", "B", 20)
    pdf.cell(0, 15, "Reporte Comparativo de Ventas por Producto", 0, 1, "C")
    pdf.set_font("Arial", "", 9)
    pdf.cell(0, 8, f"Generado: {datetime.now().strftime('%Y-%m-%d %H:%M')}", 0, 1, "C")
    pdf.ln(5)
    
    pdf.set_font("Arial", "B", 12)
    pdf.cell(0, 8, "Indicadores Clave", 0, 1)
    pdf.set_font("Arial", "", 10)
    pdf.cell(0, 7, f"Total Ventas 2024: {kpis['v2024']:,} unidades", 0, 1)
    pdf.cell(0, 7, f"Total Ventas 2025: {kpis['v2025']:,} unidades", 0, 1)
    pdf.cell(0, 7, f"Crecimiento: {kpis['crecimiento']:.1f}%", 0, 1)
    
    pdf.ln(8)
    
    if tabla_comparativa is not None and len(tabla_comparativa) > 0:
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
            pdf.cell(30, 7, f"{row['variacion_porcentaje']:.1f}%", 1, 1, 'R')
    
    return pdf.output(dest='S').encode('latin1')

# ======================================
# FILTROS
# ======================================

def aplicar_filtros(df):
    st.sidebar.markdown("### 📊 Panel de Control")
    st.sidebar.markdown("---")
    
    if df is None or len(df) == 0:
        return df, {}
    
    # Filtro de año
    if 'anio' in df.columns:
        anios = sorted(df['anio'].dropna().unique())
        anios = [a for a in anios if isinstance(a, (int, float)) and a >= 2000]
        
        anio_seleccionado = st.sidebar.multiselect(
            "📅 Años",
            options=anios,
            default=anios if anios else [],
            key="anios"
        )
    else:
        anio_seleccionado = []
    
    # Filtro de marca (si existe)
    if 'marca' in df.columns:
        marcas_disponibles = sorted([str(m) for m in df['marca'].unique() if str(m) not in ['No especificada', 'General', 'nan', 'None']])
        if marcas_disponibles:
            marcas_seleccionadas = st.sidebar.multiselect(
                "🏷️ Marcas",
                options=marcas_disponibles,
                default=marcas_disponibles if marcas_disponibles else [],
                key="marcas"
            )
        else:
            marcas_seleccionadas = []
    else:
        marcas_seleccionadas = []
    
    st.sidebar.markdown("---")
    
    # Aplicar filtros
    filtro = df.copy()
    
    if anio_seleccionado:
        filtro = filtro[filtro['anio'].isin(anio_seleccionado)]
    
    if marcas_seleccionadas:
        filtro = filtro[filtro['marca'].isin(marcas_seleccionadas)]
    
    with st.sidebar.expander("📈 Estadísticas", expanded=False):
        st.metric("Registros", f"{len(filtro):,}")
        if 'cantidad' in filtro.columns:
            st.metric("Ventas totales", f"{int(filtro['cantidad'].sum()):,}")
    
    filtros_dict = {
        'Años': ', '.join([str(a) for a in anio_seleccionado]) if anio_seleccionado else "Todos",
        'Marcas': ', '.join(marcas_seleccionadas[:3]) + ('...' if len(marcas_seleccionadas) > 3 else '') if marcas_seleccionadas else "Todas",
    }
    
    return filtro, filtros_dict

# ======================================
# KPIS
# ======================================

def mostrar_kpis(filtro):
    if filtro is None or len(filtro) == 0:
        return {'total': 0, 'v2024': 0, 'v2025': 0, 'crecimiento': 0}
    
    ventas_total = int(filtro['cantidad'].sum()) if 'cantidad' in filtro.columns else 0
    ventas_2024 = int(filtro[(filtro['anio'] == 2024)]['cantidad'].sum()) if 'anio' in filtro.columns and 2024 in filtro['anio'].values else 0
    ventas_2025 = int(filtro[(filtro['anio'] == 2025)]['cantidad'].sum()) if 'anio' in filtro.columns and 2025 in filtro['anio'].values else 0
    
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
# TABLA COMPARATIVA
# ======================================

def mostrar_tabla_comparativa(tabla_comparativa):
    st.markdown("### 📊 Comparativo de Ventas por Producto")
    st.markdown("*Análisis 2024 vs 2025 - Desglose por producto individual*")
    
    if tabla_comparativa is None or len(tabla_comparativa) == 0:
        st.warning("No hay datos para mostrar")
        return
    
    # Selector de cantidad de filas
    col1, col2 = st.columns([3, 1])
    with col2:
        top_n = st.selectbox("Mostrar top", ["20", "50", "100", "Todos"], index=0)
    
    # Filtrar tabla
    if top_n != "Todos":
        try:
            n = int(top_n)
            sin_total = tabla_comparativa[tabla_comparativa['producto'] != '**TOTAL**'].copy()
            sin_total_filtrado = sin_total.head(n)
            total_row = tabla_comparativa[tabla_comparativa['producto'] == '**TOTAL**'].copy()
            tabla_mostrar = pd.concat([sin_total_filtrado, total_row], ignore_index=True)
        except:
            tabla_mostrar = tabla_comparativa.copy()
    else:
        tabla_mostrar = tabla_comparativa.copy()
    
    # Crear copia para display
    tabla_display = tabla_mostrar.copy()
    
    # Formatear valores
    tabla_display['ventas_2024'] = tabla_display['ventas_2024'].apply(lambda x: f"{int(x):,}")
    tabla_display['ventas_2025'] = tabla_display['ventas_2025'].apply(lambda x: f"{int(x):,}")
    tabla_display['diferencia'] = tabla_display['diferencia'].apply(lambda x: f"{int(x):,}")
    tabla_display['variacion_porcentaje'] = tabla_display['variacion_porcentaje'].apply(
        lambda x: f"{x:.1f}%"
    )
    
    # Renombrar columnas
    tabla_display.columns = ['Producto', 'Ventas 2024', 'Ventas 2025', 'Diferencia', 'Variación %']
    
    # Colorear variaciones
    def color_variacion(val):
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
    
    styled_df = tabla_display.style.map(color_variacion, subset=['Variación %'])
    
    st.dataframe(
        styled_df,
        use_container_width=True,
        height=500,
        hide_index=True
    )
    
    # Resumen
    total_row_data = tabla_comparativa[tabla_comparativa['producto'] == '**TOTAL**']
    if len(total_row_data) > 0:
        total_2024 = int(total_row_data['ventas_2024'].iloc[0])
        total_2025 = int(total_row_data['ventas_2025'].iloc[0])
        total_var = total_row_data['variacion_porcentaje'].iloc[0]
        
        st.info(f"📊 **Resumen:** {len(tabla_comparativa)-1} productos | "
               f"Total 2024: {total_2024:,} | "
               f"Total 2025: {total_2025:,} | "
               f"Variación: {total_var:.1f}%")

# ======================================
# GRÁFICO
# ======================================

def mostrar_grafico_comparativo(tabla_comparativa):
    st.markdown("### 📈 Top 10 Productos - Comparativo 2024 vs 2025")
    
    if tabla_comparativa is None or len(tabla_comparativa) == 0:
        st.info("No hay datos suficientes")
        return
    
    top_productos = tabla_comparativa[
        (tabla_comparativa['producto'] != '**TOTAL**') & 
        (tabla_comparativa['ventas_2024'] > 0)
    ].head(10).copy()
    
    if len(top_productos) > 0:
        fig = go.Figure()
        
        productos_short = top_productos['producto'].apply(lambda x: x[:40] + '...' if len(x) > 40 else x)
        
        fig.add_trace(go.Bar(
            name='Ventas 2024',
            x=productos_short,
            y=top_productos['ventas_2024'],
            marker_color='#3b82f6',
            text=top_productos['ventas_2024'],
            textposition='outside'
        ))
        
        fig.add_trace(go.Bar(
            name='Ventas 2025',
            x=productos_short,
            y=top_productos['ventas_2025'],
            marker_color='#ef4444',
            text=top_productos['ventas_2025'],
            textposition='outside'
        ))
        
        fig.update_layout(
            title="Top 10 Productos más vendidos",
            template='plotly_white',
            height=500,
            barmode='group',
            xaxis_title="Producto",
            yaxis_title="Unidades Vendidas",
            xaxis_tickangle=-45
        )
        
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("No hay datos suficientes para mostrar el gráfico")

# ======================================
# BOTONES DE EXPORTACIÓN
# ======================================

def botones_exportacion(df, tabla_comparativa, kpis):
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
                pdf_data = exportar_pdf(tabla_comparativa, kpis)
                b64 = base64.b64encode(pdf_data).decode()
                href = f'<a href="data:application/pdf;base64,{b64}" download="reporte_comparativo.pdf">📥 Descargar PDF</a>'
                st.markdown(href, unsafe_allow_html=True)
                st.success("✅ Reporte generado!")

# ======================================
# MENÚ PRINCIPAL
# ======================================

def main():
    # Inicializar estado de sesión
    init_session_state()
    
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
                        
                        if success1 and success2 and df1 is not None and df2 is not None:
                            st.session_state.df_combinado = pd.concat([df1, df2], ignore_index=True)
                            st.session_state.datos_cargados = True
                            st.success("✅ Datos cargados!")
                        else:
                            st.error("Error al cargar los archivos")
                else:
                    st.warning("Selecciona ambos archivos")
        
        with col2:
            if st.button("📊 Ejemplo", use_container_width=True):
                with st.spinner("Cargando ejemplo..."):
                    # Crear datos de ejemplo con múltiples productos
                    np.random.seed(42)
                    productos_ejemplo = [f"Producto_{i}" for i in range(1, 51)]
                    fechas = pd.date_range('2024-01-01', '2025-12-31', freq='D')
                    data = []
                    for fecha in fechas:
                        for _ in range(np.random.randint(5, 20)):
                            data.append({
                                'fecha': fecha,
                                'producto': np.random.choice(productos_ejemplo),
                                'cantidad': np.random.randint(1, 10),
                                'marca': np.random.choice(['Marca A', 'Marca B', 'Marca C'])
                            })
                    df = pd.DataFrame(data)
                    df = limpiar_dataframe(df)
                    st.session_state.df_combinado = df
                    st.session_state.datos_cargados = True
                    st.success("✅ Datos de ejemplo cargados!")
        
        st.markdown("---")
        
        if st.session_state.datos_cargados and st.session_state.df_combinado is not None:
            st.info(f"📊 {len(st.session_state.df_combinado):,} registros")
            
            # Mostrar columnas disponibles
            with st.expander("📋 Columnas disponibles"):
                for col in st.session_state.df_combinado.columns:
                    st.caption(f"• {col}")
    
    # Main content
    if not st.session_state.datos_cargados or st.session_state.df_combinado is None:
        st.markdown("""
        <div style="text-align: center; padding: 3rem;">
            <h1>📊 Dashboard Comparativo de Ventas</h1>
            <p style="color: #64748b; font-size: 1.1rem;">
                Análisis 2024 vs 2025 por producto individual
            </p>
            <div style="margin-top: 2rem;">
                <p style="color: #3b82f6;">Carga tus archivos Excel en el menú lateral</p>
            </div>
        </div>
        """, unsafe_allow_html=True)
        
        with st.expander("📖 Instrucciones", expanded=False):
            st.markdown("""
            ### Columnas necesarias en tus archivos Excel:
            
            **Obligatorias:**
            - **Fecha** (fecha, date, Fecha) - Para identificar el año de la venta
            - **Cantidad** (cantidad, venta, monto) - Número de unidades vendidas
            - **Producto** (producto, codigo, item, SKU) - Para identificar cada producto individual
            
            **Opcionales:**
            - Marca (marca, brand)
            - Proveedor (proveedor, supplier)
            
            ### Consejos:
            1. Asegúrate de que la columna de producto tenga valores únicos por producto
            2. Si tu columna de producto se llama diferente, podrás seleccionarla manualmente
            3. Los datos se agruparán automáticamente por producto y año
            """)
    else:
        # Limpiar datos
        df = limpiar_dataframe(st.session_state.df_combinado)
        
        if df is None or len(df) == 0:
            st.error("No hay datos válidos después de la limpieza")
            return
        
        # Selección de columna de producto
        st.sidebar.markdown("---")
        st.sidebar.markdown("### 🏷️ Configuración")
        
        # Detectar posibles columnas de producto
        columnas_excluir = ['fecha', 'cantidad', 'anio', 'mes', 'dia', 'mes_nombre', 'trimestre', 'semana']
        columnas_texto = [col for col in df.columns if col not in columnas_excluir]
        
        # Buscar columna de producto por defecto
        columna_producto_default = None
        for col in columnas_texto:
            if 'producto' in col.lower() or 'codigo' in col.lower() or 'item' in col.lower() or 'sku' in col.lower():
                columna_producto_default = col
                break
        
        if columna_producto_default is None and len(columnas_texto) > 0:
            columna_producto_default = columnas_texto[0]
        
        if columna_producto_default is None:
            st.error("No se encontró ninguna columna para identificar productos. Asegúrate de tener una columna con nombres de productos.")
            return
        
        columna_producto = st.sidebar.selectbox(
            "📦 Selecciona la columna que identifica los productos",
            options=columnas_texto,
            index=columnas_texto.index(columna_producto_default) if columna_producto_default in columnas_texto else 0
        )
        
        # Aplicar filtros
        filtro, filtros_aplicados = aplicar_filtros(df)
        
        if len(filtro) == 0:
            st.warning("⚠️ No hay datos con los filtros seleccionados")
            return
        
        # Crear tabla comparativa por producto
        tabla_comparativa = crear_tabla_comparativa(filtro, columna_producto, top_n=None)
        
        if tabla_comparativa is None or len(tabla_comparativa) == 0:
            st.warning(f"No se pudieron generar datos. Verifica que la columna '{columna_producto}' tenga valores válidos")
            return
        
        # Título
        st.markdown("# 📊 Dashboard Comparativo de Ventas")
        st.caption(f"📅 Actualizado: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        st.caption(f"📦 Analizando {len(tabla_comparativa)-1} productos individuales desde la columna: **{columna_producto}**")
        
        # Mostrar filtros activos
        with st.expander("🔍 Filtros activos", expanded=False):
            for key, value in filtros_aplicados.items():
                st.caption(f"**{key}:** {value}")
        
        st.markdown("---")
        
        # KPIs
        kpis = mostrar_kpis(filtro)
        
        # Exportar
        botones_exportacion(filtro, tabla_comparativa, kpis)
        
        st.markdown("---")
        
        # Tabla comparativa
        mostrar_tabla_comparativa(tabla_comparativa)
        
        st.markdown("---")
        
        # Gráfico
        mostrar_grafico_comparativo(tabla_comparativa)
        
        st.markdown("---")
        
        # Vista previa de datos
        with st.expander("🔍 Vista previa de los datos cargados", expanded=False):
            st.dataframe(filtro.head(100), use_container_width=True)
        
        # Footer
        st.markdown("""
        <div class="footer">
            Dashboard Comparativo de Ventas | Análisis 2024 vs 2025 por producto | Desarrollado con Streamlit
        </div>
        """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()