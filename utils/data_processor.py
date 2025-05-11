# utils/data_processor.py

import pandas as pd
import io
from datetime import datetime
import traceback

def load_data(file):
    """Cargar datos desde un archivo Excel"""
    if file is not None:
        return pd.read_excel(file)
    return None

def process_matriculados(file):
    """Procesar el archivo de matriculados"""
    try:
        # Primero intentamos con la hoja 'matriculados'
        df = pd.read_excel(file, sheet_name="matriculados")
    except Exception as e:
        # Si falla, mostrar las hojas disponibles y usar la primera
        excel_file = pd.ExcelFile(file)
        sheet_names = excel_file.sheet_names
        
        if not sheet_names:
            raise ValueError(f"El archivo no contiene hojas de cálculo: {file.name}")
        
        # Usar la primera hoja disponible
        sheet_name = sheet_names[0]
        print(f"Hoja 'matriculados' no encontrada. Usando primera hoja: '{sheet_name}'")
        df = pd.read_excel(file, sheet_name=sheet_name)
    
    # Convertir columnas de fecha a datetime
    date_columns = ["Fecha ingreso", "Fecha matrícula"]
    for col in date_columns:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors='coerce')
    
    return df

def process_leads(file):
    """Procesar el archivo de leads activos"""
    try:
        # Primero intentamos con la hoja 'leads_activos'
        df = pd.read_excel(file, sheet_name="leads_activos")
    except Exception as e:
        # Si falla, mostrar las hojas disponibles y usar la primera
        excel_file = pd.ExcelFile(file)
        sheet_names = excel_file.sheet_names
        
        if not sheet_names:
            raise ValueError(f"El archivo no contiene hojas de cálculo: {file.name}")
        
        # Usar la primera hoja disponible
        sheet_name = sheet_names[0]
        print(f"Hoja 'leads_activos' no encontrada. Usando primera hoja: '{sheet_name}'")
        df = pd.read_excel(file, sheet_name=sheet_name)
    
    # Convertir columnas de fecha a datetime
    if "Fecha ingreso" in df.columns:
        df["Fecha ingreso"] = pd.to_datetime(df["Fecha ingreso"], errors='coerce')
    
    return df

def process_planificacion(file):
    """Procesar el archivo de planificación"""
    excel_file = pd.ExcelFile(file)
    sheet_names = excel_file.sheet_names
    
    if not sheet_names:
        raise ValueError(f"El archivo de planificación no contiene hojas de cálculo")
    
    required_sheets = ["plan_mensual", "inversion_acumulada", "calendario_convocatoria"]
    missing_sheets = [sheet for sheet in required_sheets if sheet not in sheet_names]
    
    if missing_sheets:
        # Si falta alguna hoja, mostrar mensaje
        print(f"Hojas faltantes en archivo de planificación: {missing_sheets}")
        print(f"Hojas disponibles: {sheet_names}")
    
    # Leer cada pestaña del archivo o usar un DataFrame vacío si no existe
    try:
        df_plan_mensual = pd.read_excel(file, sheet_name="plan_mensual")
    except Exception:
        print("Error al leer la hoja 'plan_mensual', usando DataFrame vacío")
        df_plan_mensual = pd.DataFrame(columns=["Marca", "Canal", "Presupuesto total mes", "CPL estimado", "Leads estimados"])
    
    try:
        df_inversion = pd.read_excel(file, sheet_name="inversion_acumulada")
    except Exception:
        print("Error al leer la hoja 'inversion_acumulada', usando DataFrame vacío")
        df_inversion = pd.DataFrame(columns=["Fecha", "Marca", "Canal", "Inversión acumulada", "CPL estimado"])
    
    try:
        df_calendario = pd.read_excel(file, sheet_name="calendario_convocatoria")
    except Exception:
        print("Error al leer la hoja 'calendario_convocatoria', usando DataFrame vacío")
        df_calendario = pd.DataFrame(columns=["Marca", "Programa", "Fecha inicio", "Fecha fin", "Tipo"])
    
    # Convertir columnas de fecha a datetime
    if "Fecha" in df_inversion.columns:
        df_inversion["Fecha"] = pd.to_datetime(df_inversion["Fecha"], errors='coerce')
    
    date_columns = ["Fecha inicio", "Fecha fin"]
    for col in date_columns:
        if col in df_calendario.columns:
            df_calendario[col] = pd.to_datetime(df_calendario[col], errors='coerce')
    
    return df_plan_mensual, df_inversion, df_calendario 