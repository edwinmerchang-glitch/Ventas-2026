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
    
    # Asegurar que la columna de producto sea string
    if 'producto' in df.columns:
        df['producto'] = df['producto'].fillna('No especificado').astype(str)
    
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
# FUNCIÓN PARA CREAR TABLA COMPARATIVA (CORREGIDA)
# ======================================

@st.cache_data(ttl=300)
def crear_tabla_comparativa(df, columna_producto, año1=None, año2=None, top_n=None):
    """Crea tabla comparativa de productos entre dos años seleccionados"""
    
    if df is None or len(df) == 0:
        return pd.DataFrame(), []
    
    if columna_producto not in df.columns:
        st.error(f"❌ La columna '{columna_producto}' no existe en los datos")
        return pd.DataFrame(), []
    
    if 'cantidad' not in df.columns:
        st.error("❌ La columna 'cantidad' no existe en los datos")
        return pd.DataFrame(), []
    
    if 'anio' not in df.columns:
        st.error("❌ La columna 'anio' no existe en los datos")
        return pd.DataFrame(), []
    
    # Asegurar que la columna de producto sea string
    df[columna_producto] = df[columna_producto].fillna('No especificado').astype(str)
    
    # Obtener años disponibles
    años_disponibles = sorted(df['anio'].unique())
    
    # Si no se especifican años, usar los dos primeros años disponibles
    if año1 is None and len(años_disponibles) >= 1:
        año1 = años_disponibles[0]
    if año2 is None and len(años_disponibles) >= 2:
        año2 = años_disponibles[1]
    
    # Si solo hay un año, duplicarlo para comparación
    if año2 is None and len(años_disponibles) == 1:
        año2 = año1
    
    # Filtrar datos por los años seleccionados
    df_filtrado = df[df['anio'].isin([año1, año2])]
    
    # Agrupar por producto y año
    ventas_producto = df_filtrado.groupby([columna_producto, 'anio'])['cantidad'].sum().reset_index()
    
    # Pivotar para tener los años como columnas
    tabla_pivot = ventas_producto.pivot(index=columna_producto, columns='anio', values='cantidad').fillna(0)
    
    # Renombrar columnas dinámicamente
    nuevas_columnas = []
    for col in tabla_pivot.columns:
        nuevas_columnas.append(f'ventas_{int(col)}')
    tabla_pivot.columns = nuevas_columnas
    
    # Asegurar que ambas columnas existan
    col_año1 = f'ventas_{año1}'
    col_año2 = f'ventas_{año2}'
    
    if col_año1 not in tabla_pivot.columns:
        tabla_pivot[col_año1] = 0
    if col_año2 not in tabla_pivot.columns:
        tabla_pivot[col_año2] = 0
    
    # Calcular diferencia y variación
    tabla_pivot['diferencia'] = tabla_pivot[col_año2] - tabla_pivot[col_año1]
    tabla_pivot['variacion_porcentaje'] = np.where(
        tabla_pivot[col_año1] > 0,
        (tabla_pivot['diferencia'] / tabla_pivot[col_año1]) * 100,
        0
    )
    
    # Reset index
    tabla_comparativa = tabla_pivot.reset_index()
    tabla_comparativa = tabla_comparativa.rename(columns={columna_producto: 'producto'})
    
    # Filtrar productos con ventas > 0
    tabla_comparativa = tabla_comparativa[(tabla_comparativa[col_año1] > 0) | (tabla_comparativa[col_año2] > 0)]
    
    # Ordenar por ventas del primer año
    tabla_comparativa = tabla_comparativa.sort_values(col_año1, ascending=False)
    
    # Limitar a top N
    if top_n and top_n != "Todos" and isinstance(top_n, int):
        tabla_comparativa = tabla_comparativa.head(top_n)
    
    # Agregar fila de total
    if len(tabla_comparativa) > 0:
        total_ventas_año1 = tabla_comparativa[col_año1].sum()
        total_ventas_año2 = tabla_comparativa[col_año2].sum()
        total_diferencia = total_ventas_año2 - total_ventas_año1
        total_variacion = ((total_ventas_año2 - total_ventas_año1) / total_ventas_año1 * 100) if total_ventas_año1 > 0 else 0
        
        total_row = pd.DataFrame({
            'producto': ['**TOTAL**'],
            col_año1: [total_ventas_año1],
            col_año2: [total_ventas_año2],
            'diferencia': [total_diferencia],
            'variacion_porcentaje': [total_variacion]
        })
        
        tabla_comparativa = pd.concat([tabla_comparativa, total_row], ignore_index=True)
    
    return tabla_comparativa, [año1, año2]

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
    
    # ===== FILTRO DE AÑOS PARA COMPARACIÓN =====
    st.sidebar.markdown("### 📅 Comparar Años")
    
    años_disponibles = sorted(df['anio'].unique()) if 'anio' in df.columns else []
    
    if len(años_disponibles) >= 2:
        col1, col2 = st.sidebar.columns(2)
        with col1:
            año_comparar_1 = st.selectbox(
                "Año base",
                options=años_disponibles,
                index=0,
                key="año_base"
            )
        with col2:
            año_comparar_2 = st.selectbox(
                "Año comparar",
                options=años_disponibles,
                index=min(1, len(años_disponibles)-1),
                key="año_comparar"
            )
    else:
        año_comparar_1 = años_disponibles[0] if años_disponibles else None
        año_comparar_2 = años_disponibles[0] if años_disponibles else None
        st.sidebar.info(f"Solo se encontró el año {año_comparar_1}")
    
    st.sidebar.markdown("---")
    
    # ===== CALENDARIO - RANGO DE FECHAS =====
    st.sidebar.markdown("#### 📆 Filtros Adicionales")
    
    # Checkbox para activar filtros adicionales
    usar_filtros_avanzados = st.sidebar.checkbox("Activar filtros avanzados (meses/días)", value=False)
    
    meses_filtro = []
    dias_filtro = []
    
    if usar_filtros_avanzados and 'fecha' in df.columns:
        # Filtro de meses
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
        
        # Filtro de días
        st.sidebar.markdown("#### 📆 Días del mes")
        
        tipo_filtro_dia = st.sidebar.radio(
            "Tipo de filtro por día",
            ["Todos los días", "Días específicos", "Rango de días"],
            key="tipo_filtro_dia"
        )
        
        if tipo_filtro_dia == "Días específicos":
            dias_seleccionados = st.sidebar.multiselect(
                "Selecciona días del mes",
                options=list(range(1, 32)),
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
    
    # Filtrar por años (solo los años seleccionados para comparación)
    filtro = filtro[filtro['anio'].isin([año_comparar_1, año_comparar_2])]
    
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
        st.metric("Años comparados", f"{año_comparar_1} vs {año_comparar_2}")
        if meses_filtro:
            nombres_meses = {1: 'Ene', 2: 'Feb', 3: 'Mar', 4: 'Abr', 5: 'May', 6: 'Jun',
                           7: 'Jul', 8: 'Ago', 9: 'Sep', 10: 'Oct', 11: 'Nov', 12: 'Dic'}
            st.metric("Meses", ', '.join([nombres_meses.get(m, str(m)) for m in meses_filtro]))
    
    # Guardar filtros para exportación
    filtros_dict = {
        'Año base': año_comparar_1,
        'Año comparar': año_comparar_2,
        'Meses': ', '.join([str(m) for m in meses_filtro]) if meses_filtro else "Todos",
        'Días': ', '.join(map(str, dias_filtro[:10])) + ('...' if len(dias_filtro) > 10 else '') if dias_filtro else "Todos",
        'Marcas': ', '.join(marcas_seleccionadas[:3]) + ('...' if len(marcas_seleccionadas) > 3 else '') if marcas_seleccionadas else "Todas",
    }
    
    return filtro, filtros_dict, año_comparar_1, año_comparar_2

# ======================================
# GRÁFICO DE VENTAS POR MES
# ======================================

def mostrar_grafico_mensual(filtro, año1, año2):
    """Muestra gráfico de ventas por mes comparativo"""
    if filtro is None or len(filtro) == 0:
        return
    
    st.markdown(f"### 📈 Ventas por Mes - Comparativo {año1} vs {año2}")
    
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
            title=f"Comparativo Mensual de Ventas - {año1} vs {año2}",
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

def mostrar_grafico_diario(filtro, año1, año2):
    """Muestra gráfico de ventas por día del mes"""
    if filtro is None or len(filtro) == 0:
        return
    
    st.markdown(f"### 📊 Ventas por Día del Mes - {año1} vs {año2}")
    
    if 'dia' in filtro.columns and 'cantidad' in filtro.columns:
        # Filtrar los dos años
        df_año1 = filtro[filtro['anio'] == año1] if año1 in filtro['anio'].values else pd.DataFrame()
        df_año2 = filtro[filtro['anio'] == año2] if año2 in filtro['anio'].values else pd.DataFrame()
        
        if len(df_año1) > 0 or len(df_año2) > 0:
            fig = go.Figure()
            
            if len(df_año1) > 0:
                ventas_dia_año1 = df_año1.groupby('dia')['cantidad'].sum().reset_index()
                fig.add_trace(go.Scatter(
                    x=ventas_dia_año1['dia'],
                    y=ventas_dia_año1['cantidad'],
                    mode='lines+markers',
                    name=str(año1),
                    line=dict(color='#3b82f6', width=2),
                    marker=dict(size=8)
                ))
            
            if len(df_año2) > 0:
                ventas_dia_año2 = df_año2.groupby('dia')['cantidad'].sum().reset_index()
                fig.add_trace(go.Scatter(
                    x=ventas_dia_año2['dia'],
                    y=ventas_dia_año2['cantidad'],
                    mode='lines+markers',
                    name=str(año2),
                    line=dict(color='#10b981', width=2),
                    marker=dict(size=8)
                ))
            
            fig.update_layout(
                title=f"Ventas por Día del Mes - {año1} vs {año2}",
                template='plotly_white',
                height=400,
                xaxis_title="Día del Mes",
                yaxis_title="Unidades Vendidas",
                xaxis=dict(tickmode='linear', dtick=2)
            )
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info(f"No hay datos para {año1} o {año2}")
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

def exportar_pdf(tabla_comparativa, kpis, filtros_aplicados, año1, año2):
    pdf = FPDF()
    pdf.add_page()
    
    pdf.set_fill_color(59, 130, 246)
    pdf.set_text_color(30, 41, 59)
    
    pdf.set_font("Arial", "B", 20)
    pdf.cell(0, 15, f"Reporte Comparativo de Ventas {año1} vs {año2}", 0, 1, "C")
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
    pdf.cell(0, 7, f"Total Ventas {año1}: {kpis['v_año1']:,} unidades", 0, 1)
    pdf.cell(0, 7, f"Total Ventas {año2}: {kpis['v_año2']:,} unidades", 0, 1)
    pdf.cell(0, 7, f"Crecimiento: {kpis['crecimiento']:.1f}%", 0, 1)
    
    pdf.ln(8)
    
    if tabla_comparativa is not None and len(tabla_comparativa) > 0:
        pdf.set_font("Arial", "B", 9)
        pdf.cell(70, 8, "Producto", 1, 0, 'C')
        pdf.cell(25, 8, f"{año1}", 1, 0, 'C')
        pdf.cell(25, 8, f"{año2}", 1, 0, 'C')
        pdf.cell(25, 8, "Diferencia", 1, 0, 'C')
        pdf.cell(35, 8, "Variación %", 1, 1, 'C')
        
        pdf.set_font("Arial", "", 8)
        col_año1 = f'ventas_{año1}'
        col_año2 = f'ventas_{año2}'
        
        for _, row in tabla_comparativa.head(25).iterrows():
            producto = str(row['producto'])[:45] if row['producto'] != '**TOTAL**' else row['producto']
            pdf.cell(70, 7, producto, 1, 0)
            pdf.cell(25, 7, f"{int(row[col_año1]):,}", 1, 0, 'R')
            pdf.cell(25, 7, f"{int(row[col_año2]):,}", 1, 0, 'R')
            pdf.cell(25, 7, f"{int(row['diferencia']):,}", 1, 0, 'R')
            pdf.cell(35, 7, f"{row['variacion_porcentaje']:.1f}%", 1, 1, 'R')
    
    return pdf.output(dest='S').encode('latin1')

# ======================================
# KPIS
# ======================================

def mostrar_kpis(filtro, año1, año2):
    if filtro is None or len(filtro) == 0:
        return {'total': 0, 'v_año1': 0, 'v_año2': 0, 'crecimiento': 0}
    
    ventas_total = int(filtro['cantidad'].sum()) if 'cantidad' in filtro.columns else 0
    
    ventas_año1 = 0
    ventas_año2 = 0
    
    if 'anio' in filtro.columns:
        ventas_año1 = int(filtro[filtro['anio'] == año1]['cantidad'].sum()) if año1 in filtro['anio'].values else 0
        ventas_año2 = int(filtro[filtro['anio'] == año2]['cantidad'].sum()) if año2 in filtro['anio'].values else 0
    
    if ventas_año1 > 0:
        crecimiento = ((ventas_año2 - ventas_año1) / ventas_año1) * 100
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
            <h3>📈 Ventas {año1}</h3>
            <div class="value">{ventas_año1:,}</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col3:
        st.markdown(f"""
        <div class="kpi">
            <h3>🚀 Ventas {año2}</h3>
            <div class="value">{ventas_año2:,}</div>
            <div class="growth {'growth-positive' if crecimiento >= 0 else 'growth-negative'}">
                {crecimiento:.1f}%
            </div>
        </div>
        """, unsafe_allow_html=True)
    
    return {
        'total': ventas_total,
        'v_año1': ventas_año1,
        'v_año2': ventas_año2,
        'crecimiento': crecimiento
    }

# ======================================
# FUNCIÓN PARA FORMATEAR PRODUCTOS (CORREGIDA)
# ======================================

def formatear_producto(valor):
    """Formatea un valor de producto a string seguro"""
    if pd.isna(valor):
        return "No especificado"
    try:
        texto = str(valor)
        if len(texto) > 40:
            return texto[:40] + '...'
        return texto
    except:
        return "Producto"

# ======================================
# TABLA COMPARATIVA
# ======================================

def mostrar_tabla_comparativa(tabla_comparativa, año1, año2):
    st.markdown(f"### 📊 Comparativo de Ventas por Producto - {año1} vs {año2}")
    st.markdown(f"*Análisis comparativo entre {año1} y {año2} - Desglose por producto individual*")
    
    if tabla_comparativa is None or len(tabla_comparativa) == 0:
        st.warning("No hay datos para mostrar")
        return
    
    # Identificar nombres de columnas dinámicamente
    col_año1 = f'ventas_{año1}'
    col_año2 = f'ventas_{año2}'
    
    if col_año1 not in tabla_comparativa.columns:
        st.error(f"No se encontró la columna {col_año1}")
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
    tabla_display[col_año1] = tabla_display[col_año1].apply(lambda x: f"{int(x):,}")
    tabla_display[col_año2] = tabla_display[col_año2].apply(lambda x: f"{int(x):,}")
    tabla_display['diferencia'] = tabla_display['diferencia'].apply(lambda x: f"{int(x):,}")
    tabla_display['variacion_porcentaje'] = tabla_display['variacion_porcentaje'].apply(lambda x: f"{x:.1f}%")
    
    tabla_display.columns = ['Producto', f'Ventas {año1}', f'Ventas {año2}', 'Diferencia', 'Variación %']
    
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
        total_año1 = int(total_row_data[col_año1].iloc[0])
        total_año2 = int(total_row_data[col_año2].iloc[0])
        total_var = total_row_data['variacion_porcentaje'].iloc[0]
        
        st.info(f"📊 **Resumen:** {len(tabla_comparativa)-1} productos | "
               f"Total {año1}: {total_año1:,} | "
               f"Total {año2}: {total_año2:,} | "
               f"Variación: {total_var:.1f}%")

# ======================================
# TOP PRODUCTOS GRÁFICO (CORREGIDO)
# ======================================

def mostrar_top_productos_grafico(tabla_comparativa, año1, año2):
    st.markdown(f"### 🏆 Top 10 Productos más Vendidos - {año1} vs {año2}")
    
    if tabla_comparativa is None or len(tabla_comparativa) == 0:
        st.info("No hay datos suficientes")
        return
    
    col_año1 = f'ventas_{año1}'
    col_año2 = f'ventas_{año2}'
    
    if col_año1 not in tabla_comparativa.columns:
        st.info(f"No hay datos para {año1}")
        return
    
    top_productos = tabla_comparativa[
        (tabla_comparativa['producto'] != '**TOTAL**') & 
        (tabla_comparativa[col_año1] > 0)
    ].head(10).copy()
    
    if len(top_productos) > 0:
        fig = go.Figure()
        
        # Convertir productos a string de forma segura
        productos_short = []
        for producto in top_productos['producto']:
            try:
                texto = str(producto)
                if len(texto) > 40:
                    productos_short.append(texto[:40] + '...')
                else:
                    productos_short.append(texto)
            except:
                productos_short.append("Producto")
        
        fig.add_trace(go.Bar(
            name=str(año1),
            x=productos_short,
            y=top_productos[col_año1],
            marker_color='#3b82f6',
            text=top_productos[col_año1],
            textposition='outside'
        ))
        
        fig.add_trace(go.Bar(
            name=str(año2),
            x=productos_short,
            y=top_productos[col_año2],
            marker_color='#ef4444',
            text=top_productos[col_año2],
            textposition='outside'
        ))
        
        fig.update_layout(
            title=f"Top 10 Productos más vendidos - {año1} vs {año2}",
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

def botones_exportacion(df, tabla_comparativa, kpis, filtros_aplicados, año1, año2):
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
                pdf_data = exportar_pdf(tabla_comparativa, kpis, filtros_aplicados, año1, año2)
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
                    fechas = pd.date_range('2024-01-01', '2026-12-31', freq='D')
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
                Análisis comparativo entre años por producto con filtros avanzados
            </p>
            <div style="margin-top: 2rem;">
                <p style="color: #3b82f6;">Carga tus archivos Excel en el menú lateral</p>
            </div>
        </div>
        """, unsafe_allow_html=True)
        
        with st.expander("📖 Instrucciones", expanded=False):
            st.markdown("""
            ### Comparación entre años:
            
            **Selecciona los años a comparar:**
            - Puedes comparar cualquier par de años disponibles (ej: 2024 vs 2025, 2025 vs 2026)
            
            **Filtros disponibles:**
            - 📅 **Años**: Selecciona qué años quieres comparar
            - 📆 **Meses**: Filtra por meses específicos
            - 📆 **Días**: Filtra por días específicos o rango de días
            - 🏷️ **Marcas**: Filtra por una o múltiples marcas
            
            **Análisis:**
            - Tabla comparativa producto por producto
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
        filtro, filtros_aplicados, año1, año2 = aplicar_filtros_avanzados(df)
        
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
        
        # Crear tabla comparativa con los años seleccionados
        tabla_comparativa, años_usados = crear_tabla_comparativa(filtro, columna_producto, año1, año2, top_n=None)
        
        if tabla_comparativa is None or len(tabla_comparativa) == 0:
            st.warning(f"❌ No se pudieron generar datos para la comparación {año1} vs {año2}")
            return
        
        # Título
        st.markdown(f"# 📊 Dashboard Comparativo de Ventas - {año1} vs {año2}")
        st.caption(f"📅 Actualizado: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        st.caption(f"📦 Analizando **{len(tabla_comparativa)-1}** productos")
        
        # Mostrar filtros activos
        with st.expander("🔍 Filtros activos", expanded=False):
            for key, value in filtros_aplicados.items():
                st.caption(f"**{key}:** {value}")
        
        st.markdown("---")
        
        # KPIs
        kpis = mostrar_kpis(filtro, año1, año2)
        
        # Botones exportación
        botones_exportacion(filtro, tabla_comparativa, kpis, filtros_aplicados, año1, año2)
        
        st.markdown("---")
        
        # Gráficos de tiempo
        mostrar_grafico_mensual(filtro, año1, año2)
        st.markdown("---")
        mostrar_grafico_diario(filtro, año1, año2)
        
        st.markdown("---")
        
        # Tabla comparativa
        mostrar_tabla_comparativa(tabla_comparativa, año1, año2)
        
        st.markdown("---")
        
        # Top productos gráfico
        mostrar_top_productos_grafico(tabla_comparativa, año1, año2)
        
        st.markdown("---")
        
        # Vista previa
        with st.expander("🔍 Vista previa de los datos", expanded=False):
            st.dataframe(filtro.head(100), use_container_width=True)
        
        st.markdown(f"""
        <div class="footer">
            Dashboard Comparativo de Ventas | Análisis {año1} vs {año2} | Filtros por mes, día y marca
        </div>
        """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()