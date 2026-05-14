import streamlit as st
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

fig_dia.update_layout(
    height=500,
    xaxis_title="Fecha",
    yaxis_title="Unidades Vendidas"
)

st.plotly_chart(fig_dia, use_container_width=True)

# ====================================
# TOP MARCAS
# ====================================

st.markdown("---")
st.subheader("🔥 Top 10 Marcas")

marcas_top = (
    filtro.groupby('MARCA')['cantidad']
    .sum()
    .sort_values(ascending=False)
    .head(10)
    .reset_index()
)

fig_marcas = px.bar(
    marcas_top,
    x='cantidad',
    y='MARCA',
    orientation='h',
    text_auto=True,
    template='plotly_dark'
)

fig_marcas.update_layout(height=500)

st.plotly_chart(fig_marcas, use_container_width=True)

# ====================================
# TABLA DETALLE
# ====================================

st.markdown("---")
st.subheader("📋 Detalle de Ventas")

st.dataframe(
    filtro[[
        'fecha',
        'articulo',
        'nombre',
        'cantidad',
        'MARCA',
        'Mundo'
    ]],
    use_container_width=True,
    height=500
)