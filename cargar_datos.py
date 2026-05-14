import pandas as pd
from database import engine

# =========================
# CARGAR ARCHIVOS EXCEL
# =========================

archivo_2024 = "VENTA 2024.xlsx"
archivo_2025 = "VENTA 2025.xlsx"

print("Cargando archivos...")

# Leer archivos
ventas_2024 = pd.read_excel(archivo_2024)
ventas_2025 = pd.read_excel(archivo_2025)

# Unir dataframes
ventas = pd.concat([ventas_2024, ventas_2025], ignore_index=True)

# =========================
# LIMPIEZA DE DATOS
# =========================

ventas['fecha'] = pd.to_datetime(ventas['fecha'])
ventas['cantidad'] = ventas['cantidad'].fillna(0)

ventas['anio'] = ventas['fecha'].dt.year
ventas['mes'] = ventas['fecha'].dt.month
ventas['mes_nombre'] = ventas['fecha'].dt.strftime('%B')
ventas['dia'] = ventas['fecha'].dt.day
ventas['fecha_dia'] = ventas['fecha'].dt.date

# =========================
# GUARDAR SQLITE
# =========================

ventas.to_sql('ventas', engine, if_exists='replace', index=False)

print("Base de datos creada correctamente")