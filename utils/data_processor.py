# utils/data_processor.py

import pandas as pd
import io
import re
import numpy as np
from datetime import datetime
import traceback

def normalize_program_name(name):
    """Normaliza nombres de programas para facilitar coincidencias"""
    if pd.isna(name) or name == '':
        return ''
    
    # Convertir a string si es necesario
    if not isinstance(name, str):
        name = str(name)
    
    # Eliminar espacios extra y convertir a título
    normalized = re.sub(r'\s+', ' ', name.strip()).title()
    
    # Unificar palabras comunes para variaciones de un mismo programa
    replacements = {
        r'\bAdmon\b': 'Administración',
        r'\bAdmin\b': 'Administración',
        r'\bDcho\b': 'Derecho',
        r'\bInf\b': 'Informática',
        r'\bTec\b': 'Tecnología',
        r'\bEcon\b': 'Economía',
        r'\bMkt\b': 'Marketing',
        r'\bBusiness\b': 'Empresas',
        r'\bPed\b': 'Pedagogía',
        r'\bPsic\b': 'Psicología',
        r'\bComun\b': 'Comunicación',
        r'\bFisio\b': 'Fisioterapia',
    }
    
    for pattern, replacement in replacements.items():
        normalized = re.sub(pattern, replacement, normalized, flags=re.IGNORECASE)
    
    return normalized

def validate_dataframe(df, required_columns, source_name):
    """Valida que un DataFrame tenga las columnas requeridas"""
    if df.empty:
        print(f"Advertencia: El DataFrame de {source_name} está vacío")
        return False
    
    missing_columns = [col for col in required_columns if col not in df.columns]
    if missing_columns:
        print(f"Advertencia: Faltan columnas en {source_name}: {missing_columns}")
        print(f"Columnas disponibles: {list(df.columns)}")
        return False
    
    return True

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
    
    # Validar estructura mínima requerida
    required_columns = ["ID lead", "Marca", "Programa"]
    if not validate_dataframe(df, required_columns, "matriculados"):
        # Crear columnas faltantes si es necesario
        for col in required_columns:
            if col not in df.columns:
                df[col] = np.nan
    
    # Convertir columnas de fecha a datetime
    date_columns = ["Fecha ingreso", "Fecha matrícula"]
    for col in date_columns:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors='coerce')
    
    # Normalizar nombres de programas
    if 'Programa' in df.columns:
        df['Programa'] = df['Programa'].apply(normalize_program_name)
        # Eliminar filas con programa vacío después de normalización
        df = df[df['Programa'] != '']
    
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
    
    # Validar estructura mínima requerida
    required_columns = ["ID lead", "Marca", "Programa"]
    if not validate_dataframe(df, required_columns, "leads activos"):
        # Crear columnas faltantes si es necesario
        for col in required_columns:
            if col not in df.columns:
                df[col] = np.nan
    
    # Convertir columnas de fecha a datetime
    if "Fecha ingreso" in df.columns:
        df["Fecha ingreso"] = pd.to_datetime(df["Fecha ingreso"], errors='coerce')
    
    # Normalizar nombres de programas
    if 'Programa' in df.columns:
        df['Programa'] = df['Programa'].apply(normalize_program_name)
        # Eliminar filas con programa vacío después de normalización
        df = df[df['Programa'] != '']
    
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
        # Validar columnas mínimas requeridas
        required_columns = ["Marca", "Canal", "Presupuesto total mes"]
        if not validate_dataframe(df_plan_mensual, required_columns, "plan_mensual"):
            df_plan_mensual = pd.DataFrame(columns=["Marca", "Canal", "Presupuesto total mes", "CPL estimado", "Leads estimados"])
    except Exception as e:
        print(f"Error al leer la hoja 'plan_mensual': {str(e)}")
        df_plan_mensual = pd.DataFrame(columns=["Marca", "Canal", "Presupuesto total mes", "CPL estimado", "Leads estimados"])
    
    try:
        df_inversion = pd.read_excel(file, sheet_name="inversion_acumulada")
        # Validar columnas mínimas requeridas
        required_columns = ["Fecha", "Marca", "Canal", "Inversión acumulada"]
        if not validate_dataframe(df_inversion, required_columns, "inversion_acumulada"):
            df_inversion = pd.DataFrame(columns=["Fecha", "Marca", "Canal", "Inversión acumulada", "CPL estimado"])
    except Exception as e:
        print(f"Error al leer la hoja 'inversion_acumulada': {str(e)}")
        df_inversion = pd.DataFrame(columns=["Fecha", "Marca", "Canal", "Inversión acumulada", "CPL estimado"])
    
    try:
        df_calendario = pd.read_excel(file, sheet_name="calendario_convocatoria")
        # Validar columnas mínimas requeridas
        required_columns = ["Marca", "Programa", "Fecha inicio", "Fecha fin"]
        if not validate_dataframe(df_calendario, required_columns, "calendario_convocatoria"):
            df_calendario = pd.DataFrame(columns=["Marca", "Programa", "Fecha inicio", "Fecha fin", "Tipo"])
    except Exception as e:
        print(f"Error al leer la hoja 'calendario_convocatoria': {str(e)}")
        df_calendario = pd.DataFrame(columns=["Marca", "Programa", "Fecha inicio", "Fecha fin", "Tipo"])
    
    # Convertir columnas de fecha a datetime
    if "Fecha" in df_inversion.columns:
        df_inversion["Fecha"] = pd.to_datetime(df_inversion["Fecha"], errors='coerce')
    
    date_columns = ["Fecha inicio", "Fecha fin"]
    for col in date_columns:
        if col in df_calendario.columns:
            df_calendario[col] = pd.to_datetime(df_calendario[col], errors='coerce')
    
    # Normalizar nombres de programas en el calendario
    if 'Programa' in df_calendario.columns:
        # Preservar 'Todos los programas' sin normalizar
        todos_programas_mask = df_calendario['Programa'] == 'Todos los programas'
        
        # Normalizar el resto
        df_calendario.loc[~todos_programas_mask, 'Programa'] = df_calendario.loc[~todos_programas_mask, 'Programa'].apply(normalize_program_name)
        
        # Eliminar filas con programa vacío después de normalización, excepto 'Todos los programas'
        df_calendario = df_calendario[(df_calendario['Programa'] != '') | todos_programas_mask]
    
    return df_plan_mensual, df_inversion, df_calendario 