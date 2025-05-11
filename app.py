import streamlit as st
import pandas as pd
import numpy as np
import os
import matplotlib.pyplot as plt
import io
from datetime import datetime, timedelta
from utils.data_processor import load_data, process_matriculados, process_leads, process_planificacion
from utils.calculations import calculate_metrics, project_results, analyze_programs
from utils.report_generator import generate_excel, generate_pdf, generate_pptx

# Configuración de la página
st.set_page_config(
    page_title="Reporte Status Semanal",
    page_icon="📊",
    layout="wide"
)

# Título y descripción
st.title("Reporte Status Semanal")
st.markdown("""
Esta aplicación genera reportes estratégicos semanales por marca y programa educativo,
adaptados a los diferentes modelos como GRADO y POSGRADO.
""")

# Carga de archivos
st.header("Carga de Archivos")

with st.expander("Instrucciones", expanded=False):
    st.markdown("""
    ### Archivos requeridos:
    Debe cargar los archivos para la marca específica que desea analizar:
    1. **matriculados.xlsx**: Contiene la información de los matriculados de la marca seleccionada
    2. **leads_activos.xlsx**: Contiene la información de los leads activos de la marca seleccionada
    3. **planificacion.xlsx**: (Opcional) Contiene la planificación mensual, inversión acumulada y calendario de convocatorias.
       También puede introducir estos datos manualmente en la interfaz.
    """)

# Selector de marca
marcas = ["GRADO", "POSGRADO", "ADVANCE", "WIZARD", "AJA", "UNISUD"]
selected_marca = st.selectbox("Seleccionar Marca para el Reporte:", marcas)

st.write(f"Cargue los archivos correspondientes a la marca: **{selected_marca}**")

col1, col2, col3 = st.columns(3)

with col1:
    matriculados_file = st.file_uploader(f"Subir archivo de matriculados - {selected_marca}", type=["xlsx"])
    
with col2:
    leads_file = st.file_uploader(f"Subir archivo de leads activos - {selected_marca}", type=["xlsx"])
    
with col3:
    planificacion_file = st.file_uploader(f"Subir archivo de planificación - {selected_marca} (opcional)", type=["xlsx"])

# Configuración adicional
st.header("Configuración del Reporte")

# Configurar objetivo de matrículas
objetivo_matriculas = st.number_input(
    "Objetivo de Matrículas", 
    min_value=1, 
    value=100, 
    help="Establece el objetivo de matrículas para esta marca y período"
)

# Opción para ingresar datos de planificación manualmente
usar_planificacion_manual = not planificacion_file and st.checkbox("Ingresar datos de planificación manualmente", value=not planificacion_file)

# Datos de planificación manual
df_plan_mensual_manual = None
df_inversion_manual = None 
df_calendario_manual = None

if usar_planificacion_manual:
    st.subheader("Datos de Planificación Manual")
    
    with st.expander("Plan Mensual", expanded=True):
        st.write("Introduzca la planificación mensual para la marca seleccionada")
        
        # Definir canales comunes
        canales_default = ["Facebook", "Google", "Instagram", "TikTok", "Email Marketing"]
        
        # Crear formulario para plan mensual
        plan_data = []
        
        for i, canal in enumerate(canales_default):
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                canal_nombre = st.text_input(f"Canal #{i+1}", value=canal)
            with col2:
                presupuesto = st.number_input(f"Presupuesto {canal}", min_value=0, value=1000*(i+1))
            with col3:
                cpl = st.number_input(f"CPL {canal}", min_value=0.0, value=float(5+i*2))
            with col4:
                leads_estimados = presupuesto / max(0.1, cpl)
                st.text(f"Leads estimados: {int(leads_estimados)}")
            
            plan_data.append({
                "Marca": selected_marca,
                "Canal": canal_nombre,
                "Presupuesto total mes": presupuesto,
                "CPL estimado": cpl,
                "Leads estimados": int(leads_estimados)
            })
        
        # Crear DataFrame
        df_plan_mensual_manual = pd.DataFrame(plan_data)
        
        # Mostrar vista previa
        st.write("Vista previa del plan mensual:")
        st.dataframe(df_plan_mensual_manual)
    
    with st.expander("Inversión Acumulada", expanded=True):
        st.write("Introduzca la inversión acumulada por canal")
        
        inversion_data = []
        
        # Obtener canales del plan mensual
        canales = df_plan_mensual_manual["Canal"].unique().tolist()
        
        for i, canal in enumerate(canales):
            col1, col2, col3 = st.columns(3)
            with col1:
                st.text(f"Canal: {canal}")
            with col2:
                inversion = st.number_input(f"Inversión {canal}", min_value=0, value=int(500*(i+1)))
            with col3:
                cpl_actual = st.number_input(f"CPL actual {canal}", min_value=0.0, value=float(4+i*1.5))
            
            inversion_data.append({
                "Fecha": datetime.now(),
                "Marca": selected_marca,
                "Canal": canal,
                "Inversión acumulada": inversion,
                "CPL estimado": cpl_actual
            })
        
        # Crear DataFrame
        df_inversion_manual = pd.DataFrame(inversion_data)
        
        # Mostrar vista previa
        st.write("Vista previa de inversión acumulada:")
        st.dataframe(df_inversion_manual)
    
    with st.expander("Calendario de Convocatoria", expanded=True):
        st.write("Introduzca fechas de convocatoria para los diferentes programas")
        
        # Para GRADO y UNISUD, usar las fechas ya configuradas
        if selected_marca in ["GRADO", "UNISUD"]:
            st.info(f"Para {selected_marca}, se utilizarán las fechas configuradas en 'Calendario de Convocatoria'")
            # El calendario se creará más adelante con las fechas configuradas
            
            df_calendario_manual = None
        else:
            # Para otras marcas, permitir configurar convocatorias específicas
            calendario_data = []
            
            # Número de programas a configurar
            num_programas = st.number_input("Número de programas a configurar", min_value=1, max_value=10, value=3)
            
            for i in range(num_programas):
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    programa = st.text_input(f"Programa #{i+1}", value=f"Programa {i+1}")
                with col2:
                    fecha_inicio_prog = st.date_input(f"Inicio {programa}", value=datetime.now() - timedelta(days=15))
                with col3:
                    fecha_fin_prog = st.date_input(f"Fin {programa}", value=datetime.now() + timedelta(days=45))
                with col4:
                    tipo = st.selectbox(f"Tipo {programa}", options=["Convocatoria", "Cohorte"], index=0)
                
                calendario_data.append({
                    "Marca": selected_marca,
                    "Programa": programa,
                    "Fecha inicio": datetime.combine(fecha_inicio_prog, datetime.min.time()),
                    "Fecha fin": datetime.combine(fecha_fin_prog, datetime.min.time()),
                    "Tipo": tipo
                })
            
            # Crear DataFrame
            df_calendario_manual = pd.DataFrame(calendario_data)
            
            # Mostrar vista previa
            st.write("Vista previa de calendario:")
            st.dataframe(df_calendario_manual)

# Sólo para GRADO y UNISUD: configurar fechas de convocatoria
if selected_marca in ["GRADO", "UNISUD"]:
    st.subheader(f"Calendario de Convocatoria para {selected_marca}")
    st.write("Esta marca se organiza por convocatorias que incluyen múltiples programas.")
    
    col1, col2 = st.columns(2)
    
    today = datetime.now()
    default_start = today - timedelta(days=30)
    default_end = today + timedelta(days=60)
    
    with col1:
        fecha_inicio = st.date_input(
            "Fecha de inicio de la convocatoria", 
            value=default_start,
            help="Fecha en que inició o iniciará la convocatoria"
        )
    
    with col2:
        fecha_fin = st.date_input(
            "Fecha de fin de la convocatoria", 
            value=default_end,
            help="Fecha en que finaliza o finalizará la convocatoria"
        )
    
    # Calcular y mostrar tiempo transcurrido
    if fecha_inicio and fecha_fin:
        duracion_total = (fecha_fin - fecha_inicio).days
        transcurrido = (datetime.now().date() - fecha_inicio).days
        
        if duracion_total > 0:
            pct_transcurrido = min(100, max(0, (transcurrido / duracion_total) * 100))
            
            # Barra de progreso más destacada
            st.markdown(f"### Tiempo transcurrido: {pct_transcurrido:.1f}%")
            st.progress(pct_transcurrido / 100)
            
            # Agregar contexto adicional
            st.write(f"Convocatoria: día {transcurrido} de {duracion_total} ({(fecha_fin - today.date()).days} días restantes)")
        else:
            st.error("La fecha de fin debe ser posterior a la fecha de inicio.")
else:
    # Para marcas que no usan convocatorias
    fecha_inicio = None
    fecha_fin = None
    st.info(f"La marca {selected_marca} no se organiza por convocatorias con fechas fijas.")

# Generación de reportes
if st.button("Generar Reporte") and matriculados_file and leads_file and (planificacion_file or usar_planificacion_manual):
    # Mostrar indicador de progreso
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    try:
        # Procesar datos
        status_text.text("Cargando datos...")
        progress_bar.progress(10)
        
        df_matriculados = process_matriculados(matriculados_file)
        progress_bar.progress(30)
        status_text.text("Procesando leads...")
        
        df_leads = process_leads(leads_file)
        progress_bar.progress(50)
        status_text.text("Procesando planificación...")
        
        # Usar datos de planificación del archivo o de la entrada manual
        if planificacion_file:
            df_plan_mensual, df_inversion, df_calendario = process_planificacion(planificacion_file)
        else:
            # Usar datos ingresados manualmente
            df_plan_mensual = df_plan_mensual_manual
            df_inversion = df_inversion_manual
            
            # Para GRADO y UNISUD, crear calendario a partir de las fechas configuradas
            if selected_marca in ["GRADO", "UNISUD"] and fecha_inicio and fecha_fin:
                calendario_custom = {
                    'Marca': [selected_marca],
                    'Programa': ['Todos los programas'],
                    'Fecha inicio': [datetime.combine(fecha_inicio, datetime.min.time())],
                    'Fecha fin': [datetime.combine(fecha_fin, datetime.min.time())],
                    'Tipo': ['Convocatoria']
                }
                df_calendario = pd.DataFrame(calendario_custom)
            else:
                # Usar calendario manual para otras marcas
                df_calendario = df_calendario_manual
        
        progress_bar.progress(70)
        status_text.text("Calculando métricas...")
        
        # Filtrar por marca seleccionada
        marca_matriculados = df_matriculados[df_matriculados['Marca'] == selected_marca]
        marca_leads = df_leads[df_leads['Marca'] == selected_marca]
        marca_calendario = df_calendario[df_calendario['Marca'] == selected_marca]
        marca_inversion = df_inversion[df_inversion['Marca'] == selected_marca]
        
        # Sobrescribir la configuración del calendario si se proporcionó
        if selected_marca in ["GRADO", "UNISUD"] and fecha_inicio and fecha_fin:
            # Crear un DataFrame actualizado con las fechas proporcionadas
            calendario_custom = {
                'Marca': [selected_marca],
                'Programa': ['Todos los programas'],  # Ahora la convocatoria es para todos los programas
                'Fecha inicio': [datetime.combine(fecha_inicio, datetime.min.time())],
                'Fecha fin': [datetime.combine(fecha_fin, datetime.min.time())],
                'Tipo': ['Convocatoria']
            }
            # Reemplazar el calendario existente para esta marca
            marca_calendario = pd.DataFrame(calendario_custom)
        
        # Calcular métricas
        metrics = calculate_metrics(marca_matriculados, marca_leads, marca_calendario, marca_inversion, selected_marca, objetivo_matriculas)
        projections = project_results(metrics, marca_inversion, selected_marca)
        program_analysis = analyze_programs(marca_matriculados, marca_leads, marca_calendario)
        
        # Validar estructura de program_analysis
        for key in ['tabla_completa', 'top_matriculas', 'menor_conversion']:
            if key not in program_analysis:
                st.error(f"Error: Falta la clave '{key}' en el análisis de programas")
                progress_bar.empty()
                status_text.empty()
                st.stop()
        
        progress_bar.progress(90)
        status_text.text("Generando visualizaciones...")
        
        # Visualización del reporte
        st.header(f"Reporte Status Semanal - {selected_marca}")
        
        # 1. Estado actual
        st.subheader("Estado Actual")
        
        # Agregar barras de progreso para las métricas principales
        if selected_marca in ["GRADO", "UNISUD"] and metrics['tiempo_transcurrido'] is not None:
            st.markdown("##### Tiempo Transcurrido")
            st.progress(min(1.0, metrics['tiempo_transcurrido'] / 100), 
                        text=f"{metrics['tiempo_transcurrido']:.1f}%")
        
        # Mostrar progreso de matrículas respecto al objetivo con barra
        pct_objetivo = min(1.0, metrics['matriculas_acumuladas'] / max(1, metrics['objetivo_matriculas']))
        st.markdown("##### Matrículas vs Objetivo")
        st.progress(pct_objetivo, 
                   text=f"{metrics['matriculas_acumuladas']} de {metrics['objetivo_matriculas']} ({pct_objetivo*100:.1f}%)")
        
        # Mostrar otras métricas
        cols = st.columns(4)
        cols[0].metric("Leads Acumulados", f"{metrics['leads_acumulados']}")
        cols[1].metric("Tasa de Conversión", f"{metrics['tasa_conversion']:.2f}%")
        cols[2].metric("Programas Detectados", f"{metrics.get('programas_procesados', 0)}")
        
        # Proyección de cumplimiento
        pct_cumplimiento = projections['pct_cumplimiento_proyectado'] / 100
        cumplimiento_color = "normal"
        if pct_cumplimiento >= 1.0:
            cumplimiento_color = "off"  # Verde
        elif pct_cumplimiento >= 0.8:
            cumplimiento_color = "normal"  # Amarillo
        else:
            cumplimiento_color = "inverse"  # Rojo
            
        cols[3].metric(
            "Proyección de Cumplimiento", 
            f"{projections['pct_cumplimiento_proyectado']:.1f}%",
            delta_color=cumplimiento_color
        )
        
        # 2. Composición de resultados
        st.subheader("Composición de Resultados")
        
        # Mostrar barras para composición de matrículas
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("##### Matrículas por tipo de lead")
            st.progress(metrics['pct_matriculas_nuevos'] / 100, 
                       text=f"Leads Nuevos: {metrics['pct_matriculas_nuevos']:.1f}%")
            st.progress(metrics['pct_matriculas_remarketing'] / 100, 
                       text=f"Remarketing: {metrics['pct_matriculas_remarketing']:.1f}%")
        
        # 3. Estimación de cierre
        st.subheader("Estimación de Cierre (Monte Carlo)")
        cols = st.columns(3)
        cols[0].metric("Leads Proyectados", f"{projections['leads_proyectados']} ± {int(projections['leads_proyectados_std'])}")

        # Mostrar matrículas proyectadas
        matriculas_mean = projections['matriculas_proyectadas_mean']
        matriculas_std = int(projections['matriculas_proyectadas_std'])
        percentil_05 = projections['matriculas_proyectadas_min']
        percentil_95 = projections['matriculas_proyectadas_max']

        cols[1].metric("Matrículas Proyectadas (Media)", f"{matriculas_mean} ± {matriculas_std}")
        cols[2].metric("Intervalo 90% Confianza", f"{percentil_05} - {percentil_95}")

        # Añadir visualización de probabilidades con barras de progreso
        st.markdown("##### Probabilidades de Alcanzar Objetivo")
        
        umbrales = [80, 90, 100, 110, 120]
        for umbral in umbrales:
            prob_key = f'prob_meta_{umbral}'
            st.progress(min(1.0, projections[prob_key] / 100), 
                       text=f"{umbral}% del Objetivo: {projections[prob_key]:.1f}% de probabilidad")

        # Visualización de la distribución
        st.subheader("Distribución de Matrículas Proyectadas")

        # Crear histograma con matplotlib
        fig, ax = plt.subplots(figsize=(10, 5))
        simulacion = projections['simulacion_matriculas']
        ax.hist(simulacion, bins=30, alpha=0.7, color='blue')
        ax.axvline(x=matriculas_mean, color='red', linestyle='--', label=f'Media: {matriculas_mean}')
        ax.axvline(x=percentil_05, color='green', linestyle=':', label=f'P5: {percentil_05}')
        ax.axvline(x=percentil_95, color='green', linestyle=':', label=f'P95: {percentil_95}')
        
        # Línea para objetivo
        ax.axvline(x=metrics['objetivo_matriculas'], color='orange', linestyle='-', 
                   label=f'Objetivo: {metrics["objetivo_matriculas"]}')
        
        ax.set_xlabel('Matrículas Proyectadas')
        ax.set_ylabel('Frecuencia')
        ax.legend()

        st.pyplot(fig)
        
        # 4. Análisis por programa
        st.subheader("Análisis por Programa")
        
        # Tabla completa
        st.markdown("### Distribución de Resultados por Programa")
        st.dataframe(program_analysis['tabla_completa'])
        
        # Top 5 programas con más matrículas
        col1, col2 = st.columns(2)
        with col1:
            st.markdown("### Top 5 Programas con Más Matrículas")
            st.dataframe(program_analysis['top_matriculas'])
        
        # Top 5 programas con menor conversión
        with col2:
            st.markdown("### Top 5 Programas con Menor Conversión")
            st.dataframe(program_analysis['menor_conversion'])
        
        # Sección de comentarios
        st.subheader("Comentarios")
        comentarios = st.text_area("Añadir comentarios al reporte", height=150)
        
        # Opciones de exportación
        st.subheader("Exportar Reporte")
        col1, col2, col3 = st.columns(3)
        
        try:
            # Generar Excel
            excel_buffer = generate_excel(metrics, projections, program_analysis, comentarios, selected_marca)
            col1.download_button(
                label="Descargar Excel",
                data=excel_buffer,
                file_name=f"reporte_status_{selected_marca}_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.ms-excel"
            )
        except Exception as e:
            st.error(f"Error al generar Excel: {str(e)}")
            import traceback
            st.write(traceback.format_exc())
        
        try:
            # Generar PDF
            pdf_buffer = generate_pdf(metrics, projections, program_analysis, comentarios, selected_marca)
            col2.download_button(
                label="Descargar PDF",
                data=pdf_buffer,
                file_name=f"reporte_status_{selected_marca}_{datetime.now().strftime('%Y%m%d')}.pdf",
                mime="application/pdf"
            )
        except Exception as e:
            st.error(f"Error al generar PDF: {str(e)}")
            import traceback
            st.write(traceback.format_exc())
        
        try:
            # Generar PowerPoint
            pptx_buffer = generate_pptx(metrics, projections, program_analysis, comentarios, selected_marca)
            col3.download_button(
                label="Descargar PowerPoint",
                data=pptx_buffer,
                file_name=f"reporte_status_{selected_marca}_{datetime.now().strftime('%Y%m%d')}.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )
        except Exception as e:
            st.error(f"Error al generar PowerPoint: {str(e)}")
            import traceback
            st.write(traceback.format_exc())
        
        progress_bar.progress(100)
        status_text.text("¡Reporte generado con éxito!")
        
    except Exception as e:
        st.error(f"Error al generar el reporte: {str(e)}")
        import traceback
        st.write(traceback.format_exc())
        progress_bar.empty()
        status_text.empty()

# Información adicional
st.sidebar.title("Información")
st.sidebar.info("""
Este sistema genera reportes estratégicos semanales por marca y programa educativo, 
adaptados a los diferentes modelos como GRADO (convocatorias fijas) y 
POSGRADO (cohortes variables y continuas ADVANCE).
""")

st.sidebar.title("Marcas")
for marca in marcas:
    st.sidebar.markdown(f"- {marca}")

# Botón para generar datos de ejemplo
if st.sidebar.button("Generar Datos de Ejemplo"):
    try:
        from utils.data_generator import generate_sample_data
        generate_sample_data()
        st.sidebar.success("¡Datos de ejemplo generados correctamente! Verifica la carpeta 'sample_data'.")
    except Exception as e:
        st.sidebar.error(f"Error al generar datos de ejemplo: {str(e)}") 