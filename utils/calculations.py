# utils/calculations.py

import pandas as pd
import numpy as np
from datetime import datetime
from scipy import stats

def calculate_metrics(df_matriculados, df_leads, df_calendario, df_inversion, marca, objetivo_matriculas=100):
    """Calcular métricas para el reporte estratégico"""
    metrics = {}
    
    # Fecha actual para cálculos
    now = datetime.now()
    
    # 1. Tiempo transcurrido
    # El cálculo solo se aplica a marcas con convocatorias (GRADO y UNISUD)
    if marca in ["GRADO", "UNISUD"]:
        # Calculamos el promedio ponderado del tiempo transcurrido
        # Para la nueva estructura, usamos el mismo valor para todos los programas
        if not df_calendario.empty:
            # Si hay una fila con 'Todos los programas', usar esa
            if 'Todos los programas' in df_calendario['Programa'].values:
                calendario_convocatoria = df_calendario[df_calendario['Programa'] == 'Todos los programas']
            else:
                # De lo contrario, usar la primera fila
                calendario_convocatoria = df_calendario.iloc[0:1]
            
            # Usar la única fila para calcular el tiempo transcurrido
            fecha_inicio = calendario_convocatoria['Fecha inicio'].iloc[0]
            fecha_fin = calendario_convocatoria['Fecha fin'].iloc[0]
            
            if pd.notna(fecha_inicio) and pd.notna(fecha_fin):
                duracion_total = (fecha_fin - fecha_inicio).total_seconds() / (24 * 3600)  # convertir a días
                transcurrido = (now - fecha_inicio).total_seconds() / (24 * 3600)  # convertir a días
                
                if duracion_total > 0:
                    metrics['tiempo_transcurrido'] = min(100, max(0, (transcurrido / duracion_total) * 100))
                else:
                    metrics['tiempo_transcurrido'] = 0
            else:
                metrics['tiempo_transcurrido'] = 0
        else:
            metrics['tiempo_transcurrido'] = 0
    else:
        # Para marcas sin convocatorias, este cálculo no es relevante
        metrics['tiempo_transcurrido'] = None
    
    # 2. Leads acumulados
    metrics['leads_acumulados'] = df_leads[df_leads['Marca'] == marca].shape[0]
    
    # 3. Matrículas acumuladas
    metrics['matriculas_acumuladas'] = df_matriculados[df_matriculados['Marca'] == marca].shape[0]
    
    # 4. Objetivo de matrículas (valor configurado por el usuario)
    metrics['objetivo_matriculas'] = objetivo_matriculas
    
    # 5. Tasa de conversión
    if metrics['leads_acumulados'] > 0:
        metrics['tasa_conversion'] = (metrics['matriculas_acumuladas'] / metrics['leads_acumulados']) * 100
    else:
        metrics['tasa_conversion'] = 0
    
    # 6. Composición de matrículas (nuevos vs remarketing)
    matriculas_marca = df_matriculados[df_matriculados['Marca'] == marca]
    matriculas_nuevas = 0
    matriculas_remarketing = 0
    
    # Obtener la fecha de inicio para esta marca
    fecha_inicio_marca = None
    if not df_calendario.empty:
        # Si hay una fila con 'Todos los programas', usar esa
        if 'Todos los programas' in df_calendario['Programa'].values:
            fecha_inicio_df = df_calendario[df_calendario['Programa'] == 'Todos los programas']['Fecha inicio']
            if not fecha_inicio_df.empty:
                fecha_inicio_marca = fecha_inicio_df.iloc[0]
    
    # Inicializar contador para programas procesados
    programas_procesados = set()
    
    for _, matricula in matriculas_marca.iterrows():
        programa = matricula['Programa']
        fecha_ingreso = matricula['Fecha ingreso']
        
        programas_procesados.add(programa)
        
        # Usar la fecha de inicio de la marca para todos los programas
        if fecha_inicio_marca is not None and pd.notna(fecha_ingreso):
            if fecha_ingreso >= fecha_inicio_marca:
                matriculas_nuevas += 1
            else:
                matriculas_remarketing += 1
        else:
            # Si no tenemos fecha de inicio, intentar con la fecha específica del programa
            calendario_programa = df_calendario[(df_calendario['Marca'] == marca) & 
                                              (df_calendario['Programa'] == programa)]
            
            if not calendario_programa.empty and pd.notna(fecha_ingreso):
                fecha_inicio = calendario_programa['Fecha inicio'].iloc[0]
                
                if pd.notna(fecha_inicio):
                    if fecha_ingreso >= fecha_inicio:
                        matriculas_nuevas += 1
                    else:
                        matriculas_remarketing += 1
            else:
                # Si no hay fecha específica, consideramos como lead nuevo (es lo más común)
                matriculas_nuevas += 1
    
    total_matriculas = matriculas_nuevas + matriculas_remarketing
    
    if total_matriculas > 0:
        metrics['pct_matriculas_nuevos'] = (matriculas_nuevas / total_matriculas) * 100
        metrics['pct_matriculas_remarketing'] = (matriculas_remarketing / total_matriculas) * 100
    else:
        metrics['pct_matriculas_nuevos'] = 0
        metrics['pct_matriculas_remarketing'] = 0
    
    # Agregar información sobre programas únicos
    metrics['programas_procesados'] = len(programas_procesados)
    
    # 7. Inversión acumulada
    metrics['inversion_acumulada'] = df_inversion['Inversión acumulada'].sum()
    
    # 8. CPL promedio
    if metrics['leads_acumulados'] > 0:
        metrics['cpl_promedio'] = metrics['inversion_acumulada'] / metrics['leads_acumulados']
    else:
        metrics['cpl_promedio'] = 0
    
    return metrics

def project_results(metrics, df_inversion, marca, num_simulations=10000):
    """Proyectar resultados futuros usando simulación Monte Carlo"""
    projections = {}
    
    # Parámetros base
    inversion_total = 10000  # Este valor debería calcularse o extraerse de los datos
    inversion_restante = max(0, inversion_total - metrics['inversion_acumulada'])
    
    # Parámetros históricos (valores medios)
    tasa_conversion_media = metrics['tasa_conversion'] / 100  # Convertir a decimal
    cpl_medio = metrics['cpl_promedio']
    
    # Inicializar arrays para resultados de simulación
    matriculas_simuladas = np.zeros(num_simulations)
    leads_simulados = np.zeros(num_simulations)
    
    # Ejecutar simulación Monte Carlo
    for i in range(num_simulations):
        # 1. Simular CPL con distribución normal (±15% alrededor de la media)
        cpl_std_dev = cpl_medio * 0.15
        cpl_simulado = np.random.normal(cpl_medio, cpl_std_dev)
        cpl_simulado = max(1, cpl_simulado)  # Asegurar CPL positivo
        
        # 2. Simular leads generados con la inversión restante
        leads_simulados[i] = inversion_restante / cpl_simulado
        
        # 3. Simular tasa de conversión con distribución beta
        # La distribución beta es adecuada para tasas/proporciones (valores entre 0 y 1)
        # Calculamos alpha y beta para centrar la distribución alrededor de nuestra tasa histórica
        if 0 < tasa_conversion_media < 1:
            # Varianza deseada (ajustable según la confianza en los datos históricos)
            var_deseada = (tasa_conversion_media * 0.3)**2
            
            # Cálculo de parámetros alpha y beta para distribución beta
            total = tasa_conversion_media * (1 - tasa_conversion_media) / var_deseada - 1
            alpha = tasa_conversion_media * total
            beta = (1 - tasa_conversion_media) * total
            
            # Simular tasa de conversión
            tasa_simulada = np.random.beta(max(0.1, alpha), max(0.1, beta))
        else:
            # Fallback a una distribución normal truncada si la tasa está en los extremos
            tasa_simulada = np.random.normal(tasa_conversion_media, 0.02)
            tasa_simulada = max(0.001, min(0.999, tasa_simulada))
        
        # 4. Calcular matrículas esperadas para esta simulación
        matriculas_simuladas[i] = leads_simulados[i] * tasa_simulada
    
    # Calcular estadísticas de la simulación
    projections['leads_proyectados'] = int(np.mean(leads_simulados))
    projections['leads_proyectados_std'] = np.std(leads_simulados)
    
    # Redondear matrículas al ser números enteros
    matriculas_mean = np.mean(matriculas_simuladas)
    matriculas_std = np.std(matriculas_simuladas)
    
    # Intervalos de confianza del 90%
    percentiles = np.percentile(matriculas_simuladas, [5, 25, 50, 75, 95])
    
    projections['matriculas_proyectadas_min'] = int(percentiles[0])  # P5
    projections['matriculas_proyectadas_q1'] = int(percentiles[1])   # P25
    projections['matriculas_proyectadas_median'] = int(percentiles[2])  # P50
    projections['matriculas_proyectadas_q3'] = int(percentiles[3])   # P75
    projections['matriculas_proyectadas_max'] = int(percentiles[4])  # P95
    
    # Media y desviación estándar
    projections['matriculas_proyectadas_mean'] = int(matriculas_mean)
    projections['matriculas_proyectadas_std'] = matriculas_std
    
    # Probabilidad de alcanzar diferentes niveles de objetivos
    if metrics['objetivo_matriculas'] > 0:
        matriculas_totales = metrics['matriculas_acumuladas'] + matriculas_simuladas
        
        # Probabilidad de alcanzar diferentes porcentajes del objetivo
        umbrales = [0.8, 0.9, 1.0, 1.1, 1.2]
        for umbral in umbrales:
            meta_ajustada = metrics['objetivo_matriculas'] * umbral
            prob = np.mean(matriculas_totales >= meta_ajustada) * 100
            projections[f'prob_meta_{int(umbral*100)}'] = prob
        
        # Porcentaje de cumplimiento proyectado (basado en la media)
        projections['pct_cumplimiento_proyectado'] = ((metrics['matriculas_acumuladas'] + matriculas_mean) / 
                                                     metrics['objetivo_matriculas']) * 100
    else:
        for umbral in [0.8, 0.9, 1.0, 1.1, 1.2]:
            projections[f'prob_meta_{int(umbral*100)}'] = 0
        projections['pct_cumplimiento_proyectado'] = 0
    
    # Guardar los datos de la simulación para posibles visualizaciones
    projections['simulacion_matriculas'] = matriculas_simuladas.tolist()
    
    return projections

def analyze_programs(df_matriculados, df_leads, df_calendario):
    """Analizar programas para identificar los mejores y con oportunidades"""
    result = {}
    
    # Mejorado: Crear un conjunto de todos los programas únicos presentes en todos los datos disponibles
    programas_marca = set()
    
    # Agregar programas de los datos de matriculados
    if not df_matriculados.empty and 'Programa' in df_matriculados.columns:
        for programa in df_matriculados['Programa'].unique():
            if pd.notna(programa) and programa != '':
                programas_marca.add(programa)
    
    # Agregar programas de los datos de leads
    if not df_leads.empty and 'Programa' in df_leads.columns:
        for programa in df_leads['Programa'].unique():
            if pd.notna(programa) and programa != '':
                programas_marca.add(programa)
    
    # Agregar programas del calendario si está disponible
    if not df_calendario.empty and 'Programa' in df_calendario.columns:
        for programa in df_calendario['Programa'].unique():
            if pd.notna(programa) and programa != '' and programa != 'Todos los programas':
                programas_marca.add(programa)
    
    # Convertir a lista para facilidad de uso
    programas_marca = list(programas_marca)
    
    # Crear DataFrame para análisis de programas
    programas = []
    
    for programa in programas_marca:
        # Contar leads y matrículas por programa, manejando casos donde podría no existir
        leads = 0
        if not df_leads.empty and 'Programa' in df_leads.columns:
            leads = df_leads[df_leads['Programa'] == programa].shape[0]
            
        matriculas = 0
        if not df_matriculados.empty and 'Programa' in df_matriculados.columns:
            matriculas = df_matriculados[df_matriculados['Programa'] == programa].shape[0]
        
        # Evitar programas que no tienen datos (0 leads y 0 matrículas)
        if leads == 0 and matriculas == 0:
            continue
        
        # Calcular tasa de conversión
        tasa_conversion = 0
        if leads > 0:
            tasa_conversion = (matriculas / leads) * 100
        
        programas.append({
            'Programa': programa,
            'Leads': leads,
            'Matrículas': matriculas,
            'Tasa Conversión (%)': round(tasa_conversion, 2)
        })
    
    # Crear DataFrame con los datos de los programas
    df_programas = pd.DataFrame(programas)
    
    # Manejo de DataFrames vacíos
    if df_programas.empty:
        empty_df = pd.DataFrame(columns=['Programa', 'Leads', 'Matrículas', 'Tasa Conversión (%)', 'Clasificación'])
        result['tabla_completa'] = empty_df
        result['top_matriculas'] = empty_df
        result['menor_conversion'] = empty_df
        return result
    
    # Clasificar automáticamente los programas
    # Añadir columna de clasificación
    df_programas['Clasificación'] = ''
    
    # Top 5 programas con más matrículas
    if len(df_programas) > 0:
        top_matriculas = df_programas.nlargest(min(5, len(df_programas)), 'Matrículas')['Programa'].tolist()
        df_programas.loc[df_programas['Programa'].isin(top_matriculas), 'Clasificación'] = 'Top 5 Matrículas'
        
        # Programas con baja conversión (menos del 5% pero con más de 10 leads)
        baja_conversion = df_programas[
            (df_programas['Tasa Conversión (%)'] < 5) & 
            (df_programas['Leads'] > 10) &
            (~df_programas['Programa'].isin(top_matriculas))
        ]['Programa'].tolist()
        df_programas.loc[df_programas['Programa'].isin(baja_conversion), 'Clasificación'] = 'Baja Conversión'
        
        # Oportunidades (programas con alta conversión pero pocos leads)
        oportunidades = df_programas[
            (df_programas['Tasa Conversión (%)'] > 15) & 
            (df_programas['Leads'] < 20) &
            (~df_programas['Programa'].isin(top_matriculas)) &
            (~df_programas['Programa'].isin(baja_conversion))
        ]['Programa'].tolist()
        df_programas.loc[df_programas['Programa'].isin(oportunidades), 'Clasificación'] = 'Oportunidad'
    
    # Ordenar por número de matrículas (descendente)
    df_programas = df_programas.sort_values('Matrículas', ascending=False)
    
    # Preparar resultados
    result['tabla_completa'] = df_programas
    
    # Crear copias para evitar problemas de referencias
    if len(df_programas) > 0:
        result['top_matriculas'] = df_programas.nlargest(min(5, len(df_programas)), 'Matrículas').copy()
        result['menor_conversion'] = df_programas.nsmallest(min(5, len(df_programas)), 'Tasa Conversión (%)').copy()
    else:
        result['top_matriculas'] = pd.DataFrame(columns=df_programas.columns)
        result['menor_conversion'] = pd.DataFrame(columns=df_programas.columns)
    
    return result 