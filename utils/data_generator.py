# utils/data_generator.py

import pandas as pd
import numpy as np
import os
from datetime import datetime, timedelta
from faker import Faker

fake = Faker('es_ES')

def generate_sample_data():
    """
    Genera archivos de muestra para probar la aplicación
    """
    os.makedirs('sample_data', exist_ok=True)
    
    # Definir marcas y programas
    marcas = {
        "GRADO": ["Ingeniería Civil", "Derecho", "Medicina", "Arquitectura", "Psicología", 
                 "Administración de Empresas", "Comunicación Social", "Economía", "Ingeniería Industrial"],
        "POSGRADO": ["MBA", "Maestría en Derecho Corporativo", "Maestría en Finanzas", 
                    "Especialización en Recursos Humanos", "Doctorado en Ciencias"],
        "ADVANCE": ["MBA Ejecutivo", "Diploma en Marketing Digital", "Programa Ejecutivo en Gestión de Proyectos"],
        "WIZARD": ["Inglés Intensivo", "Business English", "Preparación TOEFL"],
        "AJA": ["Curso de Emprendimiento", "Innovación y Transformación Digital"],
        "UNISUD": ["Maestría Internacional", "Doble Titulación Europea"]
    }
    
    # Generar fechas
    today = datetime.now()
    start_date = today - timedelta(days=90)  # 3 meses atrás
    end_date = today + timedelta(days=90)    # 3 meses adelante
    
    # 1. Generar matriculados.xlsx para cada marca
    for marca, programas in marcas.items():
        # Generar entre 50 y 200 matrículas por marca
        num_matriculas = np.random.randint(50, 200)
        
        matriculados_data = []
        
        for _ in range(num_matriculas):
            programa = np.random.choice(programas)
            fecha_ingreso = fake.date_time_between(start_date=start_date, end_date=today)
            fecha_matricula = fake.date_time_between(start_date=fecha_ingreso, end_date=today)
            
            matriculados_data.append({
                "ID lead": fake.uuid4(),
                "Fecha ingreso": fecha_ingreso,
                "Fecha matrícula": fecha_matricula,
                "Marca": marca,
                "Programa": programa
            })
        
        df_matriculados = pd.DataFrame(matriculados_data)
        
        # Guardar archivo para esta marca
        filename = f"sample_data/matriculados_{marca.lower()}.xlsx"
        with pd.ExcelWriter(filename, engine='openpyxl') as writer:
            df_matriculados.to_excel(writer, sheet_name="matriculados", index=False)
        
        # 2. Generar leads_activos.xlsx para cada marca
        num_leads = np.random.randint(200, 500)  # Más leads que matrículas
        
        leads_data = []
        estados = ["Contactado", "En seguimiento", "Interesado", "Postulante", "En proceso"]
        
        for _ in range(num_leads):
            programa = np.random.choice(programas)
            fecha_ingreso = fake.date_time_between(start_date=start_date, end_date=today)
            
            leads_data.append({
                "ID lead": fake.uuid4(),
                "Fecha ingreso": fecha_ingreso,
                "Estado actual": np.random.choice(estados),
                "Marca": marca,
                "Programa": programa
            })
        
        df_leads = pd.DataFrame(leads_data)
        
        # Guardar archivo para esta marca
        filename = f"sample_data/leads_activos_{marca.lower()}.xlsx"
        with pd.ExcelWriter(filename, engine='openpyxl') as writer:
            df_leads.to_excel(writer, sheet_name="leads_activos", index=False)
    
    # 3. Generar planificacion.xlsx para todas las marcas
    # 3.1 plan_mensual
    plan_mensual_data = []
    canales = ["Facebook", "Instagram", "Google", "LinkedIn", "Email", "Orgánico"]
    
    for marca in marcas.keys():
        for canal in canales:
            presupuesto = np.random.randint(1000, 10000)
            cpl_estimado = np.random.randint(10, 50)
            leads_estimados = presupuesto / cpl_estimado
            
            plan_mensual_data.append({
                "Marca": marca,
                "Canal": canal,
                "Presupuesto total mes": presupuesto,
                "CPL estimado": cpl_estimado,
                "Leads estimados": int(leads_estimados)
            })
    
    df_plan_mensual = pd.DataFrame(plan_mensual_data)
    
    # 3.2 inversion_acumulada
    inversion_acumulada_data = []
    
    for marca in marcas.keys():
        for canal in canales:
            for day in range(1, 31):  # Datos diarios por un mes
                fecha = today - timedelta(days=30-day)
                inversion = np.random.randint(100, 500)
                cpl = np.random.randint(10, 50)
                
                inversion_acumulada_data.append({
                    "Fecha": fecha,
                    "Marca": marca,
                    "Canal": canal,
                    "Inversión acumulada": inversion,
                    "CPL estimado": cpl
                })
    
    df_inversion_acumulada = pd.DataFrame(inversion_acumulada_data)
    
    # 3.3 calendario_convocatoria
    calendario_data = []
    
    for marca, programas in marcas.items():
        for programa in programas:
            # Para GRADO: convocatorias fijas
            if marca == "GRADO":
                fecha_inicio = today - timedelta(days=np.random.randint(30, 60))
                fecha_fin = today + timedelta(days=np.random.randint(30, 90))
                tipo = "Convocatoria"
            # Para otros: cohortes variables
            else:
                fecha_inicio = today - timedelta(days=np.random.randint(15, 45))
                fecha_fin = today + timedelta(days=np.random.randint(15, 75))
                tipo = "Cohorte" if marca in ["POSGRADO", "ADVANCE"] else "Convocatoria"
            
            calendario_data.append({
                "Marca": marca,
                "Programa": programa,
                "Fecha inicio": fecha_inicio,
                "Fecha fin": fecha_fin,
                "Tipo": tipo
            })
    
    df_calendario = pd.DataFrame(calendario_data)
    
    # Guardar archivo de planificación con las tres pestañas
    filename = "sample_data/planificacion.xlsx"
    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        df_plan_mensual.to_excel(writer, sheet_name="plan_mensual", index=False)
        df_inversion_acumulada.to_excel(writer, sheet_name="inversion_acumulada", index=False)
        df_calendario.to_excel(writer, sheet_name="calendario_convocatoria", index=False)

    return True

if __name__ == "__main__":
    generate_sample_data()
    print("Datos de ejemplo generados correctamente en la carpeta 'sample_data'.") 