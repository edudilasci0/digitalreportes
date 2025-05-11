import streamlit as st
import pandas as pd
import numpy as np
import os
import matplotlib.pyplot as plt
import io
from datetime import datetime
from utils.data_processor import load_data, process_matriculados, process_leads, process_planificacion
from utils.calculations import calculate_metrics, project_results, analyze_programs
from utils.report_generator import generate_excel, generate_pdf, generate_pptx

# Configuración de la página
st.set_page_config(
    page_title="Digital Reportes - Marketing Estratégico",
    page_icon="📊",
    layout="wide"
)

# Título y descripción
st.title("Reportador Estratégico de Marketing")
st.markdown("""
Esta aplicación genera reportes estratégicos por marca y programa educativo,
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
    3. **planificacion.xlsx**: Contiene la planificación mensual, inversión acumulada y calendario de convocatorias
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
    planificacion_file = st.file_uploader(f"Subir archivo de planificación - {selected_marca}", type=["xlsx"])

# Generación de reportes
if st.button("Generar Reporte") and matriculados_file and leads_file and planificacion_file:
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
        
        df_plan_mensual, df_inversion, df_calendario = process_planificacion(planificacion_file)
        progress_bar.progress(70)
        status_text.text("Calculando métricas...")
        
        # Filtrar por marca seleccionada
        marca_matriculados = df_matriculados[df_matriculados['Marca'] == selected_marca]
        marca_leads = df_leads[df_leads['Marca'] == selected_marca]
        marca_calendario = df_calendario[df_calendario['Marca'] == selected_marca]
        marca_inversion = df_inversion[df_inversion['Marca'] == selected_marca]
        
        # Calcular métricas
        metrics = calculate_metrics(marca_matriculados, marca_leads, marca_calendario, marca_inversion, selected_marca)
        projections = project_results(metrics, marca_inversion, selected_marca)
        program_analysis = analyze_programs(marca_matriculados, marca_leads, marca_calendario)
        
        progress_bar.progress(90)
        status_text.text("Generando visualizaciones...")
        
        # Visualización del reporte
        st.header(f"Reporte Estratégico - {selected_marca}")
        
        # 1. Estado actual
        st.subheader("Estado Actual")
        cols = st.columns(4)
        cols[0].metric("Tiempo Transcurrido", f"{metrics['tiempo_transcurrido']:.1f}%")
        cols[1].metric("Leads Acumulados", f"{metrics['leads_acumulados']}")
        cols[2].metric("Matrículas vs Meta", f"{metrics['matriculas_acumuladas']}/{metrics['meta_matriculas']}")
        cols[3].metric("Tasa de Conversión", f"{metrics['tasa_conversion']:.2f}%")
        
        # 2. Composición de resultados
        st.subheader("Composición de Resultados")
        cols = st.columns(2)
        cols[0].metric("% Matrículas Leads Nuevos", f"{metrics['pct_matriculas_nuevos']:.1f}%")
        cols[1].metric("% Matrículas Remarketing", f"{metrics['pct_matriculas_remarketing']:.1f}%")
        
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

        # Añadir visualización de probabilidades
        st.subheader("Probabilidades de Alcanzar Meta")
        prob_cols = st.columns(5)
        prob_cols[0].metric("80% de Meta", f"{projections['prob_meta_80']:.1f}%")
        prob_cols[1].metric("90% de Meta", f"{projections['prob_meta_90']:.1f}%")
        prob_cols[2].metric("100% de Meta", f"{projections['prob_meta_100']:.1f}%")
        prob_cols[3].metric("110% de Meta", f"{projections['prob_meta_110']:.1f}%")
        prob_cols[4].metric("120% de Meta", f"{projections['prob_meta_120']:.1f}%")

        # Visualización de la distribución
        st.subheader("Distribución de Matrículas Proyectadas")

        # Crear histograma con matplotlib
        fig, ax = plt.subplots(figsize=(10, 5))
        simulacion = projections['simulacion_matriculas']
        ax.hist(simulacion, bins=30, alpha=0.7, color='blue')
        ax.axvline(x=matriculas_mean, color='red', linestyle='--', label=f'Media: {matriculas_mean}')
        ax.axvline(x=percentil_05, color='green', linestyle=':', label=f'P5: {percentil_05}')
        ax.axvline(x=percentil_95, color='green', linestyle=':', label=f'P95: {percentil_95}')
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
        
        # Generar Excel
        excel_buffer = generate_excel(metrics, projections, program_analysis, comentarios, selected_marca)
        col1.download_button(
            label="Descargar Excel",
            data=excel_buffer,
            file_name=f"reporte_{selected_marca}_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.ms-excel"
        )
        
        # Generar PDF
        pdf_buffer = generate_pdf(metrics, projections, program_analysis, comentarios, selected_marca)
        col2.download_button(
            label="Descargar PDF",
            data=pdf_buffer,
            file_name=f"reporte_{selected_marca}_{datetime.now().strftime('%Y%m%d')}.pdf",
            mime="application/pdf"
        )
        
        # Generar PowerPoint
        pptx_buffer = generate_pptx(metrics, projections, program_analysis, comentarios, selected_marca)
        col3.download_button(
            label="Descargar PowerPoint",
            data=pptx_buffer,
            file_name=f"reporte_{selected_marca}_{datetime.now().strftime('%Y%m%d')}.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
        
        progress_bar.progress(100)
        status_text.text("¡Reporte generado con éxito!")
        
    except Exception as e:
        st.error(f"Error al generar el reporte: {str(e)}")
        progress_bar.empty()
        status_text.empty()

# Información adicional
st.sidebar.title("Información")
st.sidebar.info("""
Este sistema genera reportes estratégicos por marca y programa educativo, 
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