# utils/data_processor.py

import pandas as pd
import io
from datetime import datetime

def load_data(file):
    """Cargar datos desde un archivo Excel"""
    if file is not None:
        return pd.read_excel(file)
    return None

def process_matriculados(file):
    """Procesar el archivo de matriculados"""
    df = pd.read_excel(file, sheet_name="matriculados")
    
    # Convertir columnas de fecha a datetime
    date_columns = ["Fecha ingreso", "Fecha matrícula"]
    for col in date_columns:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors='coerce')
    
    return df

def process_leads(file):
    """Procesar el archivo de leads activos"""
    df = pd.read_excel(file, sheet_name="leads_activos")
    
    # Convertir columnas de fecha a datetime
    if "Fecha ingreso" in df.columns:
        df["Fecha ingreso"] = pd.to_datetime(df["Fecha ingreso"], errors='coerce')
    
    return df

def process_planificacion(file):
    """Procesar el archivo de planificación"""
    # Leer cada pestaña del archivo
    df_plan_mensual = pd.read_excel(file, sheet_name="plan_mensual")
    df_inversion = pd.read_excel(file, sheet_name="inversion_acumulada")
    df_calendario = pd.read_excel(file, sheet_name="calendario_convocatoria")
    
    # Convertir columnas de fecha a datetime
    if "Fecha" in df_inversion.columns:
        df_inversion["Fecha"] = pd.to_datetime(df_inversion["Fecha"], errors='coerce')
    
    date_columns = ["Fecha inicio", "Fecha fin"]
    for col in date_columns:
        if col in df_calendario.columns:
            df_calendario[col] = pd.to_datetime(df_calendario[col], errors='coerce')
    
    return df_plan_mensual, df_inversion, df_calendario 