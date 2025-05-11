# utils/data_generator.py

import pandas as pd
import numpy as np
import os
from datetime import datetime, timedelta
from faker import Faker

fake = Faker('es_ES')

def generate_sample_data():
    """
    Genera archivos de muestra para probar la aplicación con datos más realistas
    """
    os.makedirs('sample_data', exist_ok=True)
    
    # Definir marcas y programas con datos más realistas para Colombia
    marcas = {
        "GRADO": [
            "Ingeniería Civil", "Derecho", "Medicina", "Administración de Empresas", 
            "Psicología", "Contaduría Pública", "Comunicación Social", "Ingeniería Industrial",
            "Ingeniería de Sistemas", "Arquitectura", "Economía", "Enfermería", 
            "Licenciatura en Inglés", "Marketing Digital", "Diseño Gráfico"
        ],
        "POSGRADO": [
            "MBA Ejecutivo", "Maestría en Derecho Corporativo", "Maestría en Finanzas", 
            "Especialización en Recursos Humanos", "Doctorado en Ingeniería", 
            "Maestría en Gestión de Proyectos", "Especialización en Marketing Digital",
            "Maestría en Psicología Clínica", "Doctorado en Ciencias Económicas"
        ],
        "ADVANCE": [
            "MBA Flexible", "Programa Ejecutivo en Transformación Digital", 
            "Programa Ejecutivo en Gestión de Proyectos", "Diploma en Marketing Digital",
            "Programa de Alta Dirección"
        ],
        "WIZARD": [
            "Inglés Intensivo Nivel 1-5", "Business English", "Preparación TOEFL/IELTS",
            "Inglés Conversacional", "Inglés para Negocios"
        ],
        "AJA": [
            "Emprendimiento Avanzado", "Innovación Empresarial", "Transformación Digital",
            "Liderazgo Estratégico", "Data Analytics para Negocios"
        ],
        "UNISUD": [
            "Maestría Internacional en Desarrollo Sostenible", "Doble Titulación Europea en Negocios",
            "Programa Internacional de Liderazgo", "MBA Global", "Maestría en Gestión del Talento"
        ]
    }
    
    # Configuración más realista para GRADO (principal enfoque)
    today = datetime.now()
    
    # Para GRADO: definir fechas de convocatoria realistas
    grado_inicio = today - timedelta(days=45)  # Convocatoria comenzó hace 45 días
    grado_fin = today + timedelta(days=75)     # Termina en 75 días
    
    # Configurar inversión y conversiones realistas para GRADO
    grado_config = {
        'leads_totales': 2500,         # Total de leads para GRADO
        'conversion_promedio': 5.2,    # Tasa de conversión promedio (%)
        'variacion_conversion': 2.5,   # Variación en la tasa de conversión entre programas
        'inversion_total': 125000,     # Inversión total en la campaña (COP en miles)
        'cpl_promedio': 50,            # Costo por lead promedio (COP en miles)
        'remarketing_pct': 30,         # Porcentaje de leads de remarketing
    }
    
    # Configuración para las otras marcas (menos detallada)
    otras_config = {
        'POSGRADO': {'leads': 1200, 'conversion': 3.8, 'inversion': 85000},
        'ADVANCE': {'leads': 800, 'conversion': 4.5, 'inversion': 60000},
        'WIZARD': {'leads': 450, 'conversion': 8.2, 'inversion': 25000},
        'AJA': {'leads': 350, 'conversion': 6.5, 'inversion': 20000},
        'UNISUD': {'leads': 600, 'conversion': 3.2, 'inversion': 45000}
    }
    
    # Distribución realista de inversión por canales
    canales = {
        "Meta Ads": 0.35,              # 35% de la inversión
        "Google Ads": 0.30,            # 30% de la inversión
        "LinkedIn": 0.15,              # 15% de la inversión
        "Email Marketing": 0.10,       # 10% de la inversión
        "Remarketing": 0.08,           # 8% de la inversión
        "Orgánico": 0.02               # 2% de la inversión
    }
    
    # 1. Generar datos detallados para GRADO
    # Distribuir leads entre programas (algunos programas son más populares)
    grado_programas_leads = {}
    leads_restantes = grado_config['leads_totales']
    
    # Asignar leads de forma que algunos programas sean más populares
    popularidad = np.random.dirichlet(np.ones(len(marcas["GRADO"]))*2, size=1)[0]
    popularidad = popularidad / np.sum(popularidad) # Normalizar
    
    for i, programa in enumerate(marcas["GRADO"]):
        leads_programa = int(grado_config['leads_totales'] * popularidad[i])
        grado_programas_leads[programa] = leads_programa
    
    # Asegurarnos que el total sea correcto
    total_asignado = sum(grado_programas_leads.values())
    if total_asignado != grado_config['leads_totales']:
        # Ajustar el programa con más leads
        programa_max = max(grado_programas_leads, key=grado_programas_leads.get)
        grado_programas_leads[programa_max] += (grado_config['leads_totales'] - total_asignado)
    
    # Distribuir conversión entre programas (algunos programas convierten mejor)
    grado_programas_conversion = {}
    for programa in marcas["GRADO"]:
        # Generar tasa de conversión con variación
        tasa_base = grado_config['conversion_promedio']
        variacion = np.random.uniform(-grado_config['variacion_conversion'], grado_config['variacion_conversion'])
        tasa_conversion = max(1.0, tasa_base + variacion)  # Mínimo 1%
        grado_programas_conversion[programa] = tasa_conversion / 100  # Convertir a decimal
    
    # 1. Generar matriculados.xlsx para GRADO
    matriculados_data = []
    
    # Para cada programa
    for programa, leads in grado_programas_leads.items():
        # Calcular matrículas para este programa
        matriculas = int(leads * grado_programas_conversion[programa])
        
        # Generar matrículas en proporción a los leads
        for _ in range(matriculas):
            # Decidir si es lead nuevo o de remarketing
            es_remarketing = np.random.random() < (grado_config['remarketing_pct'] / 100)
            
            if es_remarketing:
                # Para remarketing, fecha de ingreso anterior al inicio de la convocatoria
                dias_antes = np.random.randint(10, 180)  # Entre 10 días y 6 meses antes
                fecha_ingreso = grado_inicio - timedelta(days=dias_antes)
            else:
                # Para leads nuevos, fecha de ingreso después del inicio de la convocatoria
                dias_despues = np.random.randint(0, (today - grado_inicio).days)
                fecha_ingreso = grado_inicio + timedelta(days=dias_despues)
            
            # Fecha de matrícula siempre después de la fecha de ingreso
            dias_hasta_matricula = np.random.randint(3, 30)  # Entre 3 y 30 días
            fecha_matricula = fecha_ingreso + timedelta(days=dias_hasta_matricula)
            
            # Limitar a fecha actual
            fecha_matricula = min(fecha_matricula, today)
            
            matriculados_data.append({
                "ID lead": fake.uuid4(),
                "Fecha ingreso": fecha_ingreso,
                "Fecha matrícula": fecha_matricula,
                "Marca": "GRADO",
                "Programa": programa
            })
    
    # Guardar archivo para GRADO
    df_matriculados_grado = pd.DataFrame(matriculados_data)
    filename = f"sample_data/matriculados_grado.xlsx"
    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        df_matriculados_grado.to_excel(writer, sheet_name="matriculados", index=False)
        print(f"Archivo {filename} generado con {len(matriculados_data)} matrículas reales")
    
    # 2. Generar leads_activos.xlsx para GRADO
    leads_data = []
    
    # Cantidad de leads que no se han matriculado aún
    leads_sin_matricula = grado_config['leads_totales'] - len(matriculados_data)
    
    # IDs de leads ya matriculados (para evitar duplicados)
    ids_matriculados = set(item["ID lead"] for item in matriculados_data)
    
    # Estados de los leads
    estados = ["Contactado", "Interesado", "En proceso", "Calificado", "Evaluando opciones"]
    
    # Para cada programa
    for programa, leads in grado_programas_leads.items():
        # Calcular cuántos leads sin matrícula para este programa
        matriculas_programa = sum(1 for item in matriculados_data if item["Programa"] == programa)
        leads_sin_matricula_programa = leads - matriculas_programa
        
        # Generar leads sin matrícula
        for _ in range(leads_sin_matricula_programa):
            # Similar lógica para remarketing vs nuevos
            es_remarketing = np.random.random() < (grado_config['remarketing_pct'] / 100)
            
            if es_remarketing:
                dias_antes = np.random.randint(10, 180)
                fecha_ingreso = grado_inicio - timedelta(days=dias_antes)
            else:
                dias_despues = np.random.randint(0, (today - grado_inicio).days)
                fecha_ingreso = grado_inicio + timedelta(days=dias_despues)
            
            # Generar ID único
            lead_id = fake.uuid4()
            while lead_id in ids_matriculados:
                lead_id = fake.uuid4()
            
            leads_data.append({
                "ID lead": lead_id,
                "Fecha ingreso": fecha_ingreso,
                "Estado actual": np.random.choice(estados, p=[0.15, 0.30, 0.25, 0.20, 0.10]),
                "Marca": "GRADO",
                "Programa": programa
            })
    
    # Agregar los leads matriculados (ya en contacto)
    for item in matriculados_data:
        leads_data.append({
            "ID lead": item["ID lead"],
            "Fecha ingreso": item["Fecha ingreso"],
            "Estado actual": "Matriculado",
            "Marca": "GRADO",
            "Programa": item["Programa"]
        })
    
    # Guardar archivo
    df_leads_grado = pd.DataFrame(leads_data)
    filename = f"sample_data/leads_activos_grado.xlsx"
    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        df_leads_grado.to_excel(writer, sheet_name="leads_activos", index=False)
        print(f"Archivo {filename} generado con {len(leads_data)} leads activos reales")
    
    # 3. Generar planificacion.xlsx con datos realistas
    # 3.1 plan_mensual - Distribución por canales
    plan_mensual_data = []
    
    # Para GRADO, distribución por canales realista
    for canal, proporcion in canales.items():
        inversion_canal = grado_config['inversion_total'] * proporcion
        cpl_estimado = grado_config['cpl_promedio'] * (0.9 + 0.2 * np.random.random())  # Variación del 10% arriba o abajo
        leads_estimados = inversion_canal / cpl_estimado
        
        plan_mensual_data.append({
            "Marca": "GRADO",
            "Canal": canal,
            "Presupuesto total mes": int(inversion_canal),
            "CPL estimado": round(cpl_estimado, 2),
            "Leads estimados": int(leads_estimados)
        })
    
    # Para otras marcas, menos detallado
    for marca, config in otras_config.items():
        for canal, proporcion in canales.items():
            inversion_canal = config['inversion'] * proporcion
            cpl_canal = (config['inversion'] / config['leads']) * (0.9 + 0.2 * np.random.random())
            leads_canal = inversion_canal / cpl_canal
            
            plan_mensual_data.append({
                "Marca": marca,
                "Canal": canal,
                "Presupuesto total mes": int(inversion_canal),
                "CPL estimado": round(cpl_canal, 2),
                "Leads estimados": int(leads_canal)
            })
    
    # 3.2 inversion_acumulada - Datos diarios
    inversion_acumulada_data = []
    
    # Distribución realista de la inversión en el tiempo (curva S)
    dias_transcurridos = (today - grado_inicio).days
    total_dias = (grado_fin - grado_inicio).days
    
    # Porcentaje del tiempo transcurrido
    pct_tiempo = min(1.0, max(0.0, dias_transcurridos / total_dias))
    
    # Usando función logística para modelar curva S de inversión
    def curva_s(x):
        return 1 / (1 + np.exp(-10 * (x - 0.5)))
    
    # Porcentaje de la inversión que debería estar gastada
    pct_inversion = curva_s(pct_tiempo)
    
    # Inversión acumulada hasta hoy
    inversion_acumulada_grado = grado_config['inversion_total'] * pct_inversion
    
    # Distribuir la inversión diaria con variación
    for dia in range(1, dias_transcurridos + 1):
        fecha = grado_inicio + timedelta(days=dia)
        pct_dia = dia / total_dias
        pct_inversion_dia = curva_s(pct_dia)
        
        inversion_acumulada_dia = grado_config['inversion_total'] * pct_inversion_dia
        
        # Distribución por canales para este día
        for canal, proporcion in canales.items():
            inversion_canal = inversion_acumulada_dia * proporcion * (0.9 + 0.2 * np.random.random())  # Agregar variación
            
            inversion_acumulada_data.append({
                "Fecha": fecha,
                "Marca": "GRADO",
                "Canal": canal,
                "Inversión acumulada": int(inversion_canal),
                "CPL estimado": round(grado_config['cpl_promedio'] * (0.9 + 0.2 * np.random.random()), 2)
            })
    
    # Para otras marcas, datos más simples
    for marca, config in otras_config.items():
        for dia in range(1, dias_transcurridos + 1):
            fecha = grado_inicio + timedelta(days=dia)
            pct_dia = dia / total_dias
            pct_inversion_dia = min(1.0, pct_dia * 1.1)  # Lineal con ligera sobreinversión
            
            inversion_acumulada_dia = config['inversion'] * pct_inversion_dia
            
            for canal, proporcion in canales.items():
                inversion_canal = inversion_acumulada_dia * proporcion * (0.9 + 0.2 * np.random.random())
                
                inversion_acumulada_data.append({
                    "Fecha": fecha,
                    "Marca": marca,
                    "Canal": canal,
                    "Inversión acumulada": int(inversion_canal),
                    "CPL estimado": round((config['inversion'] / config['leads']) * (0.9 + 0.2 * np.random.random()), 2)
                })
    
    # 3.3 calendario_convocatoria - Fechas de convocatoria
    calendario_data = []
    
    # GRADO - Una única convocatoria para todos los programas
    calendario_data.append({
        "Marca": "GRADO",
        "Programa": "Todos los programas",
        "Fecha inicio": grado_inicio,
        "Fecha fin": grado_fin,
        "Tipo": "Convocatoria"
    })
    
    # UNISUD - Similar a GRADO
    unisud_inicio = today - timedelta(days=30)
    unisud_fin = today + timedelta(days=90)
    
    calendario_data.append({
        "Marca": "UNISUD",
        "Programa": "Todos los programas",
        "Fecha inicio": unisud_inicio,
        "Fecha fin": unisud_fin,
        "Tipo": "Convocatoria"
    })
    
    # Otras marcas - Por programa
    for marca, programas in marcas.items():
        if marca not in ["GRADO", "UNISUD"]:
            for programa in programas:
                # Para cohortes variables
                dias_inicio = np.random.randint(5, 60)
                duracion = np.random.randint(30, 120)
                
                fecha_inicio = today - timedelta(days=dias_inicio)
                fecha_fin = fecha_inicio + timedelta(days=duracion)
                
                tipo = "Cohorte" if marca in ["POSGRADO", "ADVANCE"] else "Programa"
                
                calendario_data.append({
                    "Marca": marca,
                    "Programa": programa,
                    "Fecha inicio": fecha_inicio,
                    "Fecha fin": fecha_fin,
                    "Tipo": tipo
                })
    
    # Guardar archivo de planificación con las tres pestañas
    df_plan_mensual = pd.DataFrame(plan_mensual_data)
    df_inversion_acumulada = pd.DataFrame(inversion_acumulada_data)
    df_calendario = pd.DataFrame(calendario_data)
    
    filename = "sample_data/planificacion.xlsx"
    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        df_plan_mensual.to_excel(writer, sheet_name="plan_mensual", index=False)
        df_inversion_acumulada.to_excel(writer, sheet_name="inversion_acumulada", index=False)
        df_calendario.to_excel(writer, sheet_name="calendario_convocatoria", index=False)
        print(f"Archivo {filename} generado con datos realistas")
    
    # 4. Generación simplificada para otras marcas
    for marca, config in otras_config.items():
        # 4.1 Matriculados
        matriculados_data = []
        matriculas = int(config['leads'] * (config['conversion'] / 100))
        
        for i in range(matriculas):
            fecha_ingreso = today - timedelta(days=np.random.randint(15, 90))
            fecha_matricula = fecha_ingreso + timedelta(days=np.random.randint(1, 30))
            programa = np.random.choice(marcas[marca])
            
            matriculados_data.append({
                "ID lead": fake.uuid4(),
                "Fecha ingreso": fecha_ingreso,
                "Fecha matrícula": fecha_matricula,
                "Marca": marca,
                "Programa": programa
            })
        
        filename = f"sample_data/matriculados_{marca.lower()}.xlsx"
        with pd.ExcelWriter(filename, engine='openpyxl') as writer:
            pd.DataFrame(matriculados_data).to_excel(writer, sheet_name="matriculados", index=False)
        
        # 4.2 Leads activos
        leads_data = []
        
        # Los que ya se matricularon
        ids_matriculados = set(item["ID lead"] for item in matriculados_data)
        for item in matriculados_data:
            leads_data.append({
                "ID lead": item["ID lead"],
                "Fecha ingreso": item["Fecha ingreso"],
                "Estado actual": "Matriculado",
                "Marca": marca,
                "Programa": item["Programa"]
            })
        
        # Los que aún no se matriculan
        leads_pendientes = config['leads'] - matriculas
        for i in range(leads_pendientes):
            fecha_ingreso = today - timedelta(days=np.random.randint(1, 90))
            programa = np.random.choice(marcas[marca])
            estado = np.random.choice(["Contactado", "Interesado", "En proceso", "Calificado", "Evaluando opciones"])
            
            lead_id = fake.uuid4()
            while lead_id in ids_matriculados:
                lead_id = fake.uuid4()
            
            leads_data.append({
                "ID lead": lead_id,
                "Fecha ingreso": fecha_ingreso,
                "Estado actual": estado,
                "Marca": marca,
                "Programa": programa
            })
        
        filename = f"sample_data/leads_activos_{marca.lower()}.xlsx"
        with pd.ExcelWriter(filename, engine='openpyxl') as writer:
            pd.DataFrame(leads_data).to_excel(writer, sheet_name="leads_activos", index=False)

    print("Datos de ejemplo realistas generados correctamente en la carpeta 'sample_data'.")
    return True

if __name__ == "__main__":
    generate_sample_data()
    print("Datos de ejemplo generados correctamente en la carpeta 'sample_data'.") 