import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from sqlalchemy import create_engine
from io import BytesIO
from fpdf import FPDF
import numpy as np

# =========================================
# CONFIG
# =========================================

st.set_page_config(
    page_title="Power BI Ventas",
    page_icon="🚀",
    layout="wide",
    initial_sidebar_state="expanded"
)

# =========================================
# ESTILO FUTURISTA
# =========================================

st.markdown("""
<style>

html, body, [class*="css"] {
    font-family: 'Segoe UI';
}

.main {
    background-color: #050816;
}

section[data-testid="stSidebar"] {
    background: linear-gradient(180deg,#0B1023,#111827);
}

.kpi-card {
    background: linear-gradient(145deg,#0F172A,#111827);
    border: 1px solid #00F5FF;
    border-radius: 20px;
    padding: 25px;
    box-shadow: 0 0 20px rgba(0,245,255,0.3);
    transition: 0.3s;
}

.kpi-card:hover {
    transform: translateY(-5px);
    box-shadow: 0 0 35px rgba(0,245,255,0.6);
}

.big-number {
    font-size: 38px;
    font-weight: bold;
    color: #00F5FF;
}

.small-text {
    color: white;
    font-size: 15px;
}

h1, h2, h3 {
    color: white;
}

</style>
""", unsafe_allow_html=True)