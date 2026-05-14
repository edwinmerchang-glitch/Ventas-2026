import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta
import io
import base64
from fpdf import FPDF
import numpy as np
from pathlib import Path
import calendar

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
    
    section[data-testid="stSidebar"] .stMarkdown {
        color: #e2e8f0 !important;
    }
    
    section[data-testid="stSidebar"] label {
        color: #ffffff !important;
        font-weight: 500 !important;
    }
    
    section[data-testid="stSidebar"] h1,
    section[data-testid="stSidebar"] h2,
    section[data-testid="stSidebar"] h3,
    section[data-testid="stSidebar"] h4 {
        color: #ffffff !important;
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
    
    /* Estilo para el calendario */
    .calendar-container {
        background: white;
        border-radius: 10px;
        padding: 10px;
        margin: 10px 0;
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
    """Limpia el dataframe de valores erróneos y asegura que exista la columna año"""
    if df is None or len(df) == 0:
        return df
    
    # Limpiar columna de cantidad
    if 'cantidad' in df.columns:
        df['cantidad'] = pd.to_numeric(df['cantidad'], errors='coerce').fillna(0)
    else:
        df['cantidad'] = 1
    
    # Procesar fechas
    if 'fecha' in df.columns:
        df['fecha'] = pd.to_datetime(df['fecha'], errors='coerce')
        df = df.dropna(subset=['fecha'])
        df['anio'] = df['fecha'].dt.year
        df['mes'] = df['fecha'].dt.month
        df['mes_nombre'] = df['fecha'].dt.strftime('%B')
        df['dia'] = df['fecha'].dt.day
        df['dia_semana'] = df['fecha'].dt.dayofweek
        df['nombre_dia'] = df['fecha'].dt.day_name()
        df['semana'] = df['fecha'].dt.isocalendar().week
        df['trimestre'] = df['fecha'].dt.quarter
        df['anio_mes'] = df['fecha'].dt.strftime('%Y-%m')
        df['fecha_str'] = df['fecha'].dt.strftime('%Y-%m-%d')
    
    # Eliminar filas con año inválido
    if 'anio' in df.columns:
        df = df.dropna(subset=['anio'])
        df = df[(df['anio'] >= 2000) & (df['anio'] <= 2030)]
        df['anio'] = df['anio'].astype(int)
    
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
    
    if columna_producto not in df.columns:
        st.error(f"❌ La columna '{columna_producto}' no existe en los datos")
        return pd.DataFrame()
    
    if 'cantidad' not in df.columns:
        st.error("❌ La columna 'cantidad' no existe en los datos")
        return pd.DataFrame()
    
    if 'anio' not in df.columns:
        st.error("❌ La columna 'anio' no existe en los datos")
        return pd.DataFrame()
    
    # Agrupar por producto y año
    ventas_producto = df.groupby([columna_producto, 'anio'])['cantidad'].sum().reset_index()
    
    # Pivotar
    tabla_pivot = ventas_producto.pivot(index=columna_producto, columns='anio', values='cantidad').fillna(0)
    
    # Asegurar columnas
    if 2024 not in tabla_pivot.columns:
        tabla_pivot[2024] = 0
    if 2025 not in tabla_pivot.columns:
        tabla_pivot[2025] = 0
    
    tabla_pivot.columns = ['ventas_2024', 'ventas_2025']
    tabla_pivot['diferencia'] = tabla_pivot['ventas_2025'] - tabla_pivot['ventas_2024']
    tabla_pivot['variacion_porcentaje'] = np.where(
        tabla_pivot['ventas_2024'] > 0,
        (tabla_pivot['diferencia'] / tabla_pivot['ventas_2024']) * 100,
        0
    )
    
    tabla_comparativa = tabla_pivot.reset_index()
    tabla_comparativa = tabla_comparativa.rename(columns={columna_producto: 'producto'})
    tabla_comparativa = tabla_comparativa[(tabla_comparativa['ventas_2024'] > 0) | (tabla_comparativa['ventas_2025'] > 0)]
    tabla_comparativa = tabla_comparativa.sort_values('ventas_2024', ascending=False)
    
    if top_n and top_n != "Todos" and isinstance(top_n, int):
        tabla_comparativa = tabla_comparativa.head(top_n)
    
    # Agregar total
    if len(tabla_comparativa) > 0:
        total_row = pd.DataFrame({
            'producto': ['**TOTAL**'],
            'ventas_2024': [tabla_comparativa['ventas_2024'].sum()],
            'ventas_2025': [tabla_comparativa['ventas_2025'].sum()],
            'diferencia': [tabla_comparativa['diferencia'].sum()],
            'variacion_porcentaje': [
                ((tabla_comparativa['ventas_2025'].sum() - tabla_comparativa['ventas_2024'].sum()) / 
                 tabla_comparativa['ventas_2024'].sum() * 100) if tabla_comparativa['ventas_2024'].sum() > 0 else 0
            ]
        })
        tabla_comparativa = pd.concat([tabla_comparativa, total_row], ignore_index=True)
    
    return tabla_comparativa

# ======================================
# FILTROS AVANZADOS CON CALENDARIO
# ======================================

def aplicar_filtros_avanzados(df):
    """Aplica filtros avanzados incluyendo calendario, meses y días"""
    st.sidebar.markdown("### 📊 Panel de Control")
    st.sidebar.markdown("---")
    
    if df is None or len(df) == 0:
        return df, {}
    
    # Mostrar diagnóstico
    with st.sidebar.expander("🔍 Diagnóstico", expanded=False):
        st.caption(f"Total registros: {len(df):,}")
        if 'fecha' in df.columns:
            st.caption(f"Desde: {df['fecha'].min().date()}")
            st.caption(f"Hasta: {df['fecha'].max().date()}")
    
    # ===== FILTRO DE AÑOS =====
    st.sidebar.markdown("### 📅 Filtros de Tiempo")
    
    años_disponibles = sorted(df['anio'].unique()) if 'anio' in df.columns else []
    if años_disponibles:
        años_seleccionados = st.sidebar.multiselect(
            "📅 Años",
            options=años_disponibles,
            default=años_disponibles,
            key="anios_filtro"
        )
    else:
        años_seleccionados = []
    
    # ===== CALENDARIO - RANGO DE FECHAS =====
    st.sidebar.markdown("#### 📆 Calendario - Rango de Fechas")
    
    if 'fecha' in df.columns:
        fecha_min = df['fecha'].min().date()
        fecha_max = df['fecha'].max().date()
        
        # Selector de rango de fechas
        rango_fechas = st.sidebar.date_input(
            "Selecciona rango de fechas",
            [fecha_min, fecha_max],
            min_value=fecha_min,
            max_value=fecha_max,
            key="rango_fechas"
        )
        
        # Checkbox para usar rango de fechas
        usar_rango_fechas = st.sidebar.checkbox("Activar rango de fechas personalizado", value=False)
    else:
        usar_rango_fechas = False
        rango_fechas = []
    
    # ===== FILTRO DE MESES =====
    st.sidebar.markdown("#### 📆 Meses")
    
    nombres_meses = {
        1: 'Enero', 2: 'Febrero', 3: 'Marzo', 4: 'Abril',
        5: 'Mayo', 6: 'Junio', 7: 'Julio', 8: 'Agosto',
        9: 'Septiembre', 10: 'Octubre', 11: 'Noviembre', 12: 'Diciembre'
    }
    
    meses_disponibles = sorted(df['mes'].unique()) if 'mes' in df.columns else []
    meses_nombres = [f"{m} - {nombres_meses.get(m, '')}" for m in meses_disponibles if m in nombres_meses]
    
    meses_seleccionados = st.sidebar.multiselect(
        "Selecciona meses",
        options=meses_nombres,
        default=meses_nombres,
        key="meses_filtro"
    )
    
    meses_filtro = [int(m.split(' - ')[0]) for m in meses_seleccionados]
    
    # ===== FILTRO DE DÍAS =====
    st.sidebar.markdown("#### 📆 Días del mes")
    
    # Opción para seleccionar días específicos o rangos
    tipo_filtro_dia = st.sidebar.radio(
        "Tipo de filtro por día",
        ["Todos los días", "Días específicos", "Rango de días"],
        key="tipo_filtro_dia"
    )
    
    dias_filtro = []
    
    if tipo_filtro_dia == "Días específicos":
        dias_disponibles = list(range(1, 32))
        dias_seleccionados = st.sidebar.multiselect(
            "Selecciona días del mes",
            options=dias_disponibles,
            default=[],
            key="dias_filtro"
        )
        dias_filtro = dias_seleccionados
    elif tipo_filtro_dia == "Rango de días":
        col1, col2 = st.sidebar.columns(2)
        with col1:
            dia_inicio = st.number_input("Día inicio", min_value=1, max_value=31, value=1, key="dia_inicio")
        with col2:
            dia_fin = st.number_input("Día fin", min_value=1, max_value=31, value=31, key="dia_fin")
        dias_filtro = list(range(dia_inicio, dia_fin + 1))
    
    # ===== FILTRO DE MARCAS =====
    st.sidebar.markdown("---")
    st.sidebar.markdown("### 🏷️ Filtros Comerciales")
    
    if 'marca' in df.columns:
        marcas_disponibles = sorted([str(m) for m in df['marca'].unique() if str(m) not in ['No especificada', 'General', 'nan', 'None']])
        if marcas_disponibles:
            # Agregar opción "Todas"
            opciones_marcas = ["Todas"] + marcas_disponibles
            seleccion_marcas = st.sidebar.multiselect(
                "🏷️ Marcas",
                options=opciones_marcas,
                default=["Todas"],
                key="marcas_filtro"
            )
            
            if "Todas" in seleccion_marcas:
                marcas_seleccionadas = marcas_disponibles
            else:
                marcas_seleccionadas = [m for m in seleccion_marcas if m != "Todas"]
        else:
            marcas_seleccionadas = []
    else:
        marcas_seleccionadas = []
    
    st.sidebar.markdown("---")
    
    # ===== APLICAR FILTROS =====
    filtro = df.copy()
    
    # Filtrar por años
    if años_seleccionados:
        filtro = filtro[filtro['anio'].isin(años_seleccionados)]
    
    # Filtrar por rango de fechas
    if usar_rango_fechas and len(rango_fechas) == 2 and 'fecha' in filtro.columns:
        filtro = filtro[
            (filtro['fecha'].dt.date >= rango_fechas[0]) &
            (filtro['fecha'].dt.date <= rango_fechas[1])
        ]
    
    # Filtrar por meses
    if meses_filtro and 'mes' in filtro.columns:
        filtro = filtro[filtro['mes'].isin(meses_filtro)]
    
    # Filtrar por días
    if dias_filtro and 'dia' in filtro.columns:
        filtro = filtro[filtro['dia'].isin(dias_filtro)]
    
    # Filtrar por marcas
    if marcas_seleccionadas and 'marca' in filtro.columns:
        filtro = filtro[filtro['marca'].isin(marcas_seleccionadas)]
    
    # Mostrar resumen de filtros
    with st.sidebar.expander("📈 Resumen de filtros", expanded=False):
        st.metric("Registros filtrados", f"{len(filtro):,}")
        if 'cantidad' in filtro.columns:
            st.metric("Ventas totales", f"{int(filtro['cantidad'].sum()):,}")
        if 'anio' in filtro.columns:
            años_filtro = sorted(filtro['anio'].unique())
            st.metric("Años", ', '.join(map(str, años_filtro)) if años_filtro else "Ninguno")
        if meses_filtro:
            st.metric("Meses", ', '.join([nombres_meses.get(m, str(m)) for m in meses_filtro]))
        if dias_filtro and len(dias_filtro) <= 10:
            st.metric("Días", ', '.join(map(str, dias_filtro)))
        elif dias_filtro:
            st.metric("Días", f"{min(dias_filtro)} - {max(dias_filtro)}")
    
    # Guardar filtros para exportación
    filtros_dict = {
        'Años': ', '.join(map(str, años_seleccionados)) if años_seleccionados else "Todos",
        'Rango fechas': f"{rango_fechas[0]} a {rango_fechas[1]}" if usar_rango_fechas and len(rango_fechas) == 2 else "No aplica",
        'Meses': ', '.join([nombres_meses.get(m, str(m)) for m in meses_filtro]) if meses_filtro else "Todos",
        'Días': ', '.join(map(str, dias_filtro[:10])) + ('...' if len(dias_filtro) > 10 else '') if dias_filtro else "Todos",
        'Marcas': ', '.join(marcas_seleccionadas[:3]) + ('...' if len(marcas_seleccionadas) > 3 else '') if marcas_seleccionadas else "Todas",
    }
    
    return filtro, filtros_dict

# ======================================
# GRÁFICO DE VENTAS POR MES (COMPARATIVO)
# ======================================

def mostrar_grafico_mensual(filtro):
    """Muestra gráfico de ventas por mes comparativo"""
    if filtro is None or len(filtro) == 0:
        return
    
    st.markdown("### 📈 Ventas por Mes - Comparativo 2024 vs 2025")
    
    if 'anio' in filtro.columns and 'mes' in filtro.columns and 'cantidad' in filtro.columns:
        ventas_mensuales = filtro.groupby(['anio', 'mes'])['cantidad'].sum().reset_index()
        
        nombres_meses = {
            1: 'Ene', 2: 'Feb', 3: 'Mar', 4: 'Abr',
            5: 'May', 6: 'Jun', 7: 'Jul', 8: 'Ago',
            9: 'Sep', 10: 'Oct', 11: 'Nov', 12: 'Dic'
        }
        ventas_mensuales['mes_nombre'] = ventas_mensuales['mes'].map(nombres_meses)
        
        fig = px.line(
            ventas_mensuales,
            x='mes_nombre',
            y='cantidad',
            color='anio',
            markers=True,
            title="Comparativo Mensual de Ventas",
            color_discrete_sequence=['#3b82f6', '#10b981']
        )
        fig.update_layout(
            template='plotly_white',
            height=400,
            xaxis_title="Mes",
            yaxis_title="Unidades Vendidas"
        )
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("No hay datos suficientes para el gráfico mensual")

def mostrar_grafico_diario(filtro):
    """Muestra gráfico de ventas por día del mes"""
    if filtro is None or len(filtro) == 0:
        return
    
    st.markdown("### 📊 Ventas por Día del Mes")
    
    if 'dia' in filtro.columns and 'cantidad' in filtro.columns:
        # Filtrar solo 2024 y 2025
        df_2024 = filtro[filtro['anio'] == 2024] if 2024 in filtro['anio'].values else pd.DataFrame()
        df_2025 = filtro[filtro['anio'] == 2025] if 2025 in filtro['anio'].values else pd.DataFrame()
        
        if len(df_2024) > 0 or len(df_2025) > 0:
            fig = go.Figure()
            
            if len(df_2024) > 0:
                ventas_dia_2024 = df_2024.groupby('dia')['cantidad'].sum().reset_index()
                fig.add_trace(go.Scatter(
                    x=ventas_dia_2024['dia'],
                    y=ventas_dia_2024['cantidad'],
                    mode='lines+markers',
                    name='2024',
                    line=dict(color='#3b82f6', width=2),
                    marker=dict(size=8)
                ))
            
            if len(df_2025) > 0:
                ventas_dia_2025 = df_2025.groupby('dia')['cantidad'].sum().reset_index()
                fig.add_trace(go.Scatter(
                    x=ventas_dia_2025['dia'],
                    y=ventas_dia_2025['cantidad'],
                    mode='lines+markers',
                    name='2025',
                    line=dict(color='#10b981', width=2),
                    marker=dict(size=8)
                ))
            
            fig.update_layout(
                title="Ventas por Día del Mes",
                template='plotly_white',
                height=400,
                xaxis_title="Día del Mes",
                yaxis_title="Unidades Vendidas",
                xaxis=dict(tickmode='linear', dtick=2)
            )
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("No hay datos para 2024 o 2025")
    else:
        st.info("No hay datos suficientes para el gráfico diario")

# ======================================
# FUNCIONES DE EXPORTACIÓN
# ======================================

def exportar_excel(df, tabla_comparativa, filtros_aplicados):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # Hoja de resumen de filtros
        resumen_df = pd.DataFrame(list(filtros_aplicados.items()), columns=['Filtro', 'Valor'])
        resumen_df.to_excel(writer, sheet_name='Filtros_Aplicados', index=False)
        
        if df is not None and len(df) > 0:
            df.head(10000).to_excel(writer, sheet_name='Detalle_Ventas', index=False)
        if tabla_comparativa is not None and len(tabla_comparativa) > 0:
            tabla_comparativa.to_excel(writer, sheet_name='Comparativo_Productos', index=False)
    
    output.seek(0)
    return output

def exportar_pdf(tabla_comparativa, kpis, filtros_aplicados):
    pdf = FPDF()
    pdf.add_page()
    
    pdf.set_fill_color(59, 130, 246)
    pdf.set_text_color(30, 41, 59)
    
    pdf.set_font("Arial", "B", 20)
    pdf.cell(0, 15, "Reporte Comparativo de Ventas por Producto", 0, 1, "C")
    pdf.set_font("Arial", "", 9)
    pdf.cell(0, 8, f"Generado: {datetime.now().strftime('%Y-%m-%d %H:%M')}", 0, 1, "C")
    pdf.ln(5)
    
    # Filtros aplicados
    pdf.set_font("Arial", "B", 10)
    pdf.cell(0, 8, "Filtros aplicados:", 0, 1)
    pdf.set_font("Arial", "", 8)
    for key, value in filtros_aplicados.items():
        pdf.cell(0, 5, f"{key}: {value}", 0, 1)
    
    pdf.ln(5)
    
    # KPIs
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
# KPIS
# ======================================

def mostrar_kpis(filtro):
    if filtro is None or len(filtro) == 0:
        return {'total': 0, 'v2024': 0, 'v2025': 0, 'crecimiento': 0}
    
    ventas_total = int(filtro['cantidad'].sum()) if 'cantidad' in filtro.columns else 0
    
    ventas_2024 = 0
    ventas_2025 = 0
    
    if 'anio' in filtro.columns:
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
# TABLA COMPARATIVA
# ======================================

def mostrar_tabla_comparativa(tabla_comparativa):
    st.markdown("### 📊 Comparativo de Ventas por Producto")
    st.markdown("*Análisis 2024 vs 2025 - Desglose por producto individual*")
    
    if tabla_comparativa is None or len(tabla_comparativa) == 0:
        st.warning("No hay datos para mostrar")
        return
    
    col1, col2 = st.columns([3, 1])
    with col2:
        top_n = st.selectbox("Mostrar top", ["20", "50", "100", "Todos"], index=0)
    
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
    
    tabla_display = tabla_mostrar.copy()
    tabla_display['ventas_2024'] = tabla_display['ventas_2024'].apply(lambda x: f"{int(x):,}")
    tabla_display['ventas_2025'] = tabla_display['ventas_2025'].apply(lambda x: f"{int(x):,}")
    tabla_display['diferencia'] = tabla_display['diferencia'].apply(lambda x: f"{int(x):,}")
    tabla_display['variacion_porcentaje'] = tabla_display['variacion_porcentaje'].apply(lambda x: f"{x:.1f}%")
    
    tabla_display.columns = ['Producto', 'Ventas 2024', 'Ventas 2025', 'Diferencia', 'Variación %']
    
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
    
    st.dataframe(styled_df, use_container_width=True, height=500, hide_index=True)
    
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
# TOP PRODUCTOS GRÁFICO
# ======================================

def mostrar_top_productos_grafico(tabla_comparativa):
    st.markdown("### 🏆 Top 10 Productos más Vendidos")
    
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

def botones_exportacion(df, tabla_comparativa, kpis, filtros_aplicados):
    col1, col2 = st.columns(2)
    
    with col1:
        if st.button("📊 Exportar a Excel", use_container_width=True):
            with st.spinner("Generando Excel..."):
                excel_data = exportar_excel(df, tabla_comparativa, filtros_aplicados)
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
    init_session_state()
    
    # Sidebar
    with st.sidebar:
        st.markdown("### 📁 Datos")
        
        archivo_2024 = st.file_uploader("📂 Año Anterior", type=['xlsx', 'xls'], key="2024")
        archivo_2025 = st.file_uploader("📂 Año Actual", type=['xlsx', 'xls'], key="2025")
        
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
                            st.success("✅ Datos cargados correctamente!")
                        else:
                            st.error("❌ Error al cargar los archivos")
                else:
                    st.warning("⚠️ Selecciona ambos archivos")
        
        with col2:
            if st.button("📊 Ejemplo", use_container_width=True):
                with st.spinner("Cargando ejemplo..."):
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
            
            with st.expander("📋 Columnas disponibles"):
                for col in st.session_state.df_combinado.columns:
                    st.caption(f"• {col}")
    
    # Main content
    if not st.session_state.datos_cargados or st.session_state.df_combinado is None:
        st.markdown("""
        <div style="text-align: center; padding: 3rem;">
            <h1>📊 Dashboard Comparativo de Ventas</h1>
            <p style="color: #64748b; font-size: 1.1rem;">
                Análisis 2024 vs 2025 por producto con filtros avanzados de tiempo
            </p>
            <div style="margin-top: 2rem;">
                <p style="color: #3b82f6;">Carga tus archivos Excel en el menú lateral</p>
            </div>
        </div>
        """, unsafe_allow_html=True)
        
        with st.expander("📖 Instrucciones", expanded=False):
            st.markdown("""
            ### Filtros disponibles:
            
            **Filtros de tiempo:**
            - 📅 **Años**: Selecciona 2024, 2025 o ambos
            - 📆 **Calendario**: Rango de fechas personalizado
            - 📆 **Meses**: Selecciona meses específicos
            - 📆 **Días**: Filtra por días específicos o rango de días
            
            **Filtros comerciales:**
            - 🏷️ **Marcas**: Filtra por una o múltiples marcas
            
            **Análisis:**
            - Comparativo 2024 vs 2025 por producto
            - Gráficos mensuales y diarios
            - Exportación a Excel y PDF
            """)
    else:
        # Limpiar datos
        df = limpiar_dataframe(st.session_state.df_combinado)
        
        if df is None or len(df) == 0:
            st.error("❌ No hay datos válidos después de la limpieza")
            return
        
        # Aplicar filtros avanzados
        filtro, filtros_aplicados = aplicar_filtros_avanzados(df)
        
        if len(filtro) == 0:
            st.warning("⚠️ No hay datos con los filtros seleccionados")
            return
        
        # Selección de columna de producto
        st.sidebar.markdown("---")
        st.sidebar.markdown("### 🏷️ Configuración")
        
        columnas_excluir = ['fecha', 'cantidad', 'anio', 'mes', 'dia', 'mes_nombre', 'dia_semana', 
                           'nombre_dia', 'semana', 'trimestre', 'anio_mes', 'fecha_str']
        columnas_texto = [col for col in filtro.columns if col not in columnas_excluir]
        
        if len(columnas_texto) == 0:
            st.error("❌ No se encontró ninguna columna para identificar productos")
            return
        
        st.sidebar.markdown("**Selecciona la columna que identifica los productos:**")
        columna_producto = st.sidebar.selectbox(
            "Columnas disponibles:",
            options=columnas_texto,
            index=0,
            help="Elige la columna que contiene los nombres o códigos de los productos"
        )
        
        productos_unicos = filtro[columna_producto].nunique()
        st.sidebar.success(f"📦 **{productos_unicos}** productos únicos")
        
        # Crear tabla comparativa
        tabla_comparativa = crear_tabla_comparativa(filtro, columna_producto, top_n=None)
        
        if tabla_comparativa is None or len(tabla_comparativa) == 0:
            st.warning(f"❌ No se pudieron generar datos")
            return
        
        # Título
        st.markdown("# 📊 Dashboard Comparativo de Ventas")
        st.caption(f"📅 Actualizado: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        st.caption(f"📦 Analizando **{len(tabla_comparativa)-1}** productos")
        
        # Mostrar filtros activos
        with st.expander("🔍 Filtros activos", expanded=False):
            for key, value in filtros_aplicados.items():
                st.caption(f"**{key}:** {value}")
        
        st.markdown("---")
        
        # KPIs
        kpis = mostrar_kpis(filtro)
        
        # Botones exportación
        botones_exportacion(filtro, tabla_comparativa, kpis, filtros_aplicados)
        
        st.markdown("---")
        
        # Gráficos de tiempo
        mostrar_grafico_mensual(filtro)
        st.markdown("---")
        mostrar_grafico_diario(filtro)
        
        st.markdown("---")
        
        # Tabla comparativa
        mostrar_tabla_comparativa(tabla_comparativa)
        
        st.markdown("---")
        
        # Top productos gráfico
        mostrar_top_productos_grafico(tabla_comparativa)
        
        st.markdown("---")
        
        # Vista previa
        with st.expander("🔍 Vista previa de los datos", expanded=False):
            st.dataframe(filtro.head(100), use_container_width=True)
        
        st.markdown("""
        <div class="footer">
            Dashboard Comparativo de Ventas | Filtros por año, mes, día y calendario | Desarrollado con Streamlit
        </div>
        """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()