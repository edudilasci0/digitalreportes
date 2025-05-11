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
    page_title="Dashboard Marketing Educativo",
    page_icon="📊",
    layout="wide"
)

# Estilos CSS personalizados - Minimalistas y profesionales
st.markdown("""
<style>
    /* Estilo general */
    .main-container {
        padding: 1rem;
    }
    
    /* Tarjetas */
    .card {
        background-color: white;
        border-radius: 10px;
        padding: 20px;
        box-shadow: 0 2px 6px rgba(0, 0, 0, 0.05);
        margin-bottom: 20px;
        position: relative;
    }
    .card-title {
        font-size: 22px;
        font-weight: 600;
        margin-bottom: 15px;
        color: #333;
    }
    .card-metric {
        font-size: 34px;
        font-weight: 700;
        color: #1E88E5;
    }
    .card-metric-secondary {
        font-size: 18px;
        color: #666;
    }
    .card-value {
        font-size: 28px;
        font-weight: 700;
        margin-bottom: 5px;
    }
    .progress-bar-container {
        height: 20px;
        background-color: #f0f0f0;
        border-radius: 10px;
        margin: 5px 0 15px 0;
        overflow: hidden;
    }
    .progress-bar {
        height: 100%;
        background-color: #1E88E5;
    }
    
    /* Gráficos */
    .chart-container {
        margin-top: 20px;
    }
    
    /* Tablas */
    .styled-table {
        width: 100%;
        border-collapse: collapse;
        margin-bottom: 20px;
    }
    .styled-table th {
        background-color: #f5f7f9;
        padding: 10px;
        text-align: left;
        font-weight: 500;
        color: #333;
    }
    .styled-table td {
        padding: 10px;
        border-bottom: 1px solid #f0f0f0;
    }
    
    /* Observaciones */
    .observation-item {
        display: flex;
        margin-bottom: 10px;
    }
    .observation-bullet {
        margin-right: 10px;
        color: #1E88E5;
    }
    
    /* Ocultar elementos de Streamlit */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    .stDeployButton {display:none;}
    
    /* Adaptaciones específicas para parecerse al ejemplo */
    h1, h2, h3 {
        font-weight: 600 !important;
        color: #333 !important;
    }
    .block-container {
        padding-top: 1rem !important;
    }
    .main .block-container {
        max-width: 100% !important;
        padding-left: 1rem !important;
        padding-right: 1rem !important;
    }
</style>
""", unsafe_allow_html=True)

# Funciones para generar componentes UI personalizados
def create_dashboard_card(title, value, subtitle=None, percentage=None, color="#1E88E5"):
    """Crear una tarjeta de dashboard similar al ejemplo"""
    
    percentage_html = ""
    if percentage is not None:
        percentage_color = "green" if float(percentage.rstrip("%")) >= 50 else "red"
        percentage_html = f'<span style="color: {percentage_color}; font-size: 18px;">({percentage})</span>'
    
    subtitle_html = f'<div class="card-metric-secondary">{subtitle}</div>' if subtitle else ''
    
    html = f"""
    <div class="card">
        <div class="card-title">{title}</div>
        <div class="card-metric" style="color: {color};">{value} {percentage_html}</div>
        {subtitle_html}
    </div>
    """
    return st.markdown(html, unsafe_allow_html=True)

def create_progress_bar(current, total, label=None):
    """Crear barra de progreso personalizada"""
    percent = min(100, int((current / max(1, total)) * 100))
    
    label_html = f'<div style="margin-bottom: 5px; font-weight: 500;">{label}</div>' if label else ''
    
    html = f"""
    <div style="margin-bottom: 15px;">
        {label_html}
        <div class="progress-bar-container">
            <div class="progress-bar" style="width: {percent}%;"></div>
        </div>
        <div style="display: flex; justify-content: space-between;">
            <span style="font-size: 14px; color: #666;">0</span>
            <span style="font-size: 14px; font-weight: 500;">{current}/{total} ({percent}%)</span>
            <span style="font-size: 14px; color: #666;">{total}</span>
        </div>
    </div>
    """
    return st.markdown(html, unsafe_allow_html=True)

def create_observation_list(observations):
    """Crear lista de observaciones con viñetas"""
    html = '<div style="margin: 15px 0;">'
    
    for obs in observations:
        html += f"""
        <div class="observation-item">
            <div class="observation-bullet">•</div>
            <div>{obs}</div>
        </div>
        """
    
    html += '</div>'
    return st.markdown(html, unsafe_allow_html=True)

# Selector de marca
marcas = ["GRADO", "POSGRADO", "ADVANCE", "WIZARD", "AJA", "UNISUD"]
selected_marca = st.sidebar.selectbox("Seleccionar Marca", marcas)

# Cargar datos
usar_demo = st.sidebar.button("Usar Datos de Ejemplo")

if usar_demo:
    # Generar datos de demo
    from utils.data_generator import generate_demo_data
    df_matriculados, df_leads, df_plan_mensual, df_inversion, df_calendario = generate_demo_data(selected_marca)
    
    # Configurar objetivo de matrículas
    objetivo_matriculas = 120
    
    # Mensaje de datos de ejemplo
    st.sidebar.success("Usando datos de ejemplo")
else:
    # Espacio para cargar archivos
    st.sidebar.markdown("### Carga de Archivos")
    
    matriculados_file = st.sidebar.file_uploader(f"Matriculados - {selected_marca}", type=["xlsx"])
    leads_file = st.sidebar.file_uploader(f"Leads Activos - {selected_marca}", type=["xlsx"])
    planificacion_file = st.sidebar.file_uploader(f"Planificación (opcional)", type=["xlsx"])
    
    # Configurar objetivo de matrículas
    objetivo_matriculas = st.sidebar.number_input("Objetivo de Matrículas", min_value=1, value=120)
    
    # Salir si no hay archivos
    if not (matriculados_file and leads_file):
        st.info("Por favor, carga los archivos necesarios o usa los datos de ejemplo.")
        st.stop()
    
    # Procesar datos
    try:
        df_matriculados = process_matriculados(matriculados_file)
        df_leads = process_leads(leads_file)
        
        if planificacion_file:
            df_plan_mensual, df_inversion, df_calendario = process_planificacion(planificacion_file)
        else:
            # Crear DataFrames vacíos si no hay archivo de planificación
            df_plan_mensual = pd.DataFrame(columns=["Marca", "Canal", "Presupuesto total mes", "CPL estimado", "Leads estimados"])
            df_inversion = pd.DataFrame(columns=["Fecha", "Marca", "Canal", "Inversión acumulada", "CPL estimado"])
            df_calendario = pd.DataFrame(columns=["Marca", "Programa", "Fecha inicio", "Fecha fin", "Tipo"])
    except Exception as e:
        st.error(f"Error al procesar los archivos: {str(e)}")
        st.stop()

# Configuración para GRADO y UNISUD (convocatorias)
today = datetime.now()
if selected_marca in ["GRADO", "UNISUD"]:
    # Buscar fecha de convocatoria en el calendario o usar valores predeterminados
    if not df_calendario.empty and 'Todos los programas' in df_calendario['Programa'].values:
        convocatoria_row = df_calendario[(df_calendario['Marca'] == selected_marca) & 
                                         (df_calendario['Programa'] == 'Todos los programas')]
        if not convocatoria_row.empty:
            fecha_inicio = convocatoria_row['Fecha inicio'].iloc[0]
            fecha_fin = convocatoria_row['Fecha fin'].iloc[0]
        else:
            fecha_inicio = today - timedelta(days=30)
            fecha_fin = today + timedelta(days=60)
    else:
        fecha_inicio = today - timedelta(days=30)
        fecha_fin = today + timedelta(days=60)
        
    # Crear DataFrame de calendario si no existe
    if df_calendario.empty:
        df_calendario = pd.DataFrame([{
            'Marca': selected_marca,
            'Programa': 'Todos los programas',
            'Fecha inicio': fecha_inicio,
            'Fecha fin': fecha_fin,
            'Tipo': 'Convocatoria'
        }])

# Filtrar datos por marca
marca_matriculados = df_matriculados[df_matriculados['Marca'] == selected_marca]
marca_leads = df_leads[df_leads['Marca'] == selected_marca]
marca_calendario = df_calendario[df_calendario['Marca'] == selected_marca]
marca_inversion = df_inversion[df_inversion['Marca'] == selected_marca]

# Calcular métricas y proyecciones
metrics = calculate_metrics(marca_matriculados, marca_leads, marca_calendario, marca_inversion, selected_marca, objetivo_matriculas)
projections = project_results(metrics, marca_inversion, selected_marca)
program_analysis = analyze_programs(marca_matriculados, marca_leads, marca_calendario)

# Observaciones según resultados
observaciones = []
if metrics['matriculas_acumuladas'] / objetivo_matriculas >= 0.6:
    observaciones.append("Ratamos dentro del ritmo planificado.")
else:
    observaciones.append("Estamos por debajo del ritmo necesario.")

if projections['pct_cumplimiento_proyectado'] >= 90:
    observaciones.append("Priproyecta el logro de meta sobre base de remarketing.")
else:
    observaciones.append("Necesario intensificar acciones para cumplir la meta.")

observaciones.append("Refuerce contacto sobre leads activos.")

# ------- DASHBOARD PRINCIPAL -------
# Título del Dashboard
st.markdown(f"<h1 style='margin-bottom: 20px;'>Dashboard Principal - {selected_marca}</h1>", unsafe_allow_html=True)

# Primera sección: KPIs principales
col1, col2 = st.columns([2, 3])

with col1:
    # Tarjeta de matrículas
    matriculas_valor = f"{metrics['matriculas_acumuladas']} / {objetivo_matriculas}"
    matriculas_porcentaje = f"{(metrics['matriculas_acumuladas']/objetivo_matriculas*100):.0f}%"
    create_dashboard_card("Matrículas:", matriculas_valor, percentage=matriculas_porcentaje)
    
    # Tarjeta de leads
    create_dashboard_card("Leads:", f"{metrics['leads_acumulados']:,}", color="#4CAF50")
    
    # Tarjeta de conversión
    create_dashboard_card("Conversión:", f"{metrics['tasa_conversion']:.1f}%", color="#FF9800")
    
    # Observaciones
    st.markdown("<h3 style='margin-top: 20px;'>Observaciones</h3>", unsafe_allow_html=True)
    create_observation_list(observaciones)

with col2:
    # Gráfico de barras para leads y matrículas
    fig, ax = plt.subplots(figsize=(10, 6))
    
    # Datos simplificados como en el ejemplo
    categorias = ['LEADS\nESTIMADOS', 'LEADS\nACUMULADOS', 'MATRÍCULAS']
    
    # Valores estimados vs reales
    if 'leads_proyectados' in projections:
        leads_estimados = projections['leads_proyectados'] + metrics['leads_acumulados']
    else:
        leads_estimados = int(metrics['leads_acumulados'] * 1.5)  # Valor aproximado si no hay proyección
        
    valores = [leads_estimados, metrics['leads_acumulados'], metrics['matriculas_acumuladas']]
    
    # Colores como en el ejemplo (tonos de azul)
    colores = ['#90CAF9', '#2196F3', '#1565C0']
    
    # Crear barras
    bars = ax.bar(categorias, valores, color=colores, width=0.6)
    
    # Añadir porcentajes dentro de las barras (solo para matrículas)
    for i, bar in enumerate(bars):
        height = bar.get_height()
        ax.text(bar.get_x() + bar.get_width()/2., height/2,
                f"{(valores[i]/valores[0]*100):.0f}%" if i > 0 else "",
                ha='center', va='center', color='white', fontweight='bold')
    
    # Añadir valores arriba de las barras
    for i, bar in enumerate(bars):
        height = bar.get_height()
        ax.text(bar.get_x() + bar.get_width()/2., height + 0.1,
                f"{int(height):,}", ha='center', va='bottom', fontweight='bold')
    
    # Añadir línea horizontal para el objetivo
    ax.axhline(y=objetivo_matriculas, color='#FF5722', linestyle='--', linewidth=2)
    ax.text(len(categorias)-1, objetivo_matriculas + objetivo_matriculas*0.05, 
            f"Objetivo: {objetivo_matriculas}", ha='right', color='#FF5722', fontweight='bold')
    
    # Configuración del gráfico
    ax.set_ylim(0, max(valores + [objetivo_matriculas]) * 1.15)  # Espacio para etiquetas
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['left'].set_visible(False)
    ax.set_yticks([])  # Quitar escala vertical para un look más limpio
    
    # Mostrar el gráfico
    st.pyplot(fig)

# ------- SEGUNDA SECCIÓN - ESTIMACIÓN DE RESULTADOS -------
st.markdown("<h2 style='margin-top: 40px; margin-bottom: 20px;'>Estimación de Resultados</h2>", unsafe_allow_html=True)

col1, col2 = st.columns([3, 2])

with col1:
    # Gráfico de proyección simplificado
    fig, ax = plt.subplots(figsize=(10, 5))
    
    # Datos para proyección
    simulacion = np.array(projections['simulacion_matriculas'])
    matriculas_mean = projections['matriculas_proyectadas_mean']
    percentil_05 = projections['matriculas_proyectadas_min']
    percentil_95 = projections['matriculas_proyectadas_max']
    
    # Crear curva suavizada en lugar de histograma
    import scipy.stats as stats
    x = np.linspace(min(simulacion), max(simulacion), 100)
    kde = stats.gaussian_kde(simulacion)
    y = kde(x)
    
    # Área bajo la curva
    ax.fill_between(x, y, color='#E3F2FD', alpha=0.7)
    ax.plot(x, y, color='#2196F3', linewidth=2)
    
    # Línea para el valor esperado
    ax.axvline(x=matriculas_mean, color='#0D47A1', linestyle='-', linewidth=2)
    
    # Línea para el objetivo
    if objective_reached := (matriculas_mean >= objetivo_matriculas):
        color_objetivo = '#4CAF50'  # Verde si se alcanza
    else:
        color_objetivo = '#F44336'  # Rojo si no se alcanza
        
    ax.axvline(x=objetivo_matriculas, color=color_objetivo, linestyle='--', linewidth=2)
    
    # Eliminar ejes para un look más limpio
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['left'].set_visible(False)
    ax.set_yticks([])
    
    # Etiquetas en puntos clave
    ax.text(matriculas_mean, max(y)*1.05, f"{int(matriculas_mean)}", 
            ha='center', va='bottom', color='#0D47A1', fontweight='bold')
    ax.text(objetivo_matriculas, max(y)*0.8, f"Meta: {objetivo_matriculas}", 
            ha='center', va='bottom', color=color_objetivo, fontweight='bold')
    
    # Título del gráfico
    ax.set_title("ESPROYECTÓ LOS LOGRO DE META", fontsize=14, fontweight='bold', pad=20)
    
    # Mostrar gráfico
    st.pyplot(fig)
    
    # Valor grande para matrículas totales proyectadas
    st.markdown(f"""
    <div style="text-align: center; margin-top: 10px;">
        <div style="font-size: 48px; font-weight: 700; color: #1E88E5;">{int(matriculas_mean)}</div>
        <div style="font-size: 18px; color: #666;">MATRÍCULAS</div>
        <div style="font-size: 14px; color: #666;">(IC: {percentil_05}-{percentil_95})</div>
    </div>
    """, unsafe_allow_html=True)

with col2:
    # Contenido de la tarjeta de estimación
    st.markdown("""
    <div class="card">
        <div style="margin-bottom: 15px;">
            <div style="color: #1E88E5; font-weight: 500;">Estimación de Resursos</div>
        </div>
    """, unsafe_allow_html=True)
    
    # Recomendaciones según los resultados
    recomendaciones = [
        "Continuar con captacion efectiva.",
        "Se proyecta el logro de la meta.",
        "Continuar captración effective de meta."
    ]
    
    create_observation_list(recomendaciones)
    
    st.markdown("</div>", unsafe_allow_html=True)

# ------- TERCERA SECCIÓN - DISTRIBUCIÓN DE RESULTADOS -------
st.markdown(f"<h2 style='margin-top: 40px; margin-bottom: 20px;'>Distribución de Resultados - {selected_marca}</h2>", unsafe_allow_html=True)

# Tabla de Top Programas
if 'top_matriculas' in program_analysis and not program_analysis['top_matriculas'].empty:
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("<h3>Top Programas</h3>", unsafe_allow_html=True)
        
        # Datos de programas
        programas_data = program_analysis['top_matriculas'].head(5)
        
        # Crear tabla HTML para mayor control sobre el estilo
        html_table = """
        <table class="styled-table" style="width:100%">
            <tr>
                <th>Programa</th>
                <th style="text-align:center">volumen</th>
                <th style="text-align:center">$$</th>
            </tr>
        """
        
        for _, row in programas_data.iterrows():
            html_table += f"""
            <tr>
                <td>{row['Programa']}</td>
                <td style="text-align:center">{int(row['Leads'])}</td>
                <td style="text-align:center">{int(row['Matrículas'])}</td>
            </tr>
            """
        
        html_table += "</table>"
        st.markdown(html_table, unsafe_allow_html=True)
    
    with col2:
        st.markdown("<h3>Top Programas</h3>", unsafe_allow_html=True)
        
        # Ordenar por tasa de conversión
        programas_por_conversion = program_analysis['tabla_completa'].copy()
        programas_por_conversion = programas_por_conversion[programas_por_conversion['Leads'] >= 5].sort_values('Tasa Conversión (%)', ascending=False).head(5)
        
        # Crear tabla HTML para mayor control sobre el estilo
        html_table = """
        <table class="styled-table" style="width:100%">
            <tr>
                <th>Programa</th>
                <th style="text-align:center">$$</th>
                <th style="text-align:center">Conversion</th>
            </tr>
        """
        
        for _, row in programas_por_conversion.iterrows():
            html_table += f"""
            <tr>
                <td>{row['Programa']}</td>
                <td style="text-align:center">{int(row['Matrículas'])}</td>
                <td style="text-align:center">{row['Tasa Conversión (%)']:.1f}%</td>
            </tr>
            """
        
        html_table += "</table>"
        st.markdown(html_table, unsafe_allow_html=True)
else:
    st.info("No hay suficientes datos para mostrar la distribución por programas.")

# ------- CONFIGURACIÓN EN SIDEBAR -------
st.sidebar.markdown("### Herramientas")

# Observaciones personalizables
with st.sidebar.expander("Editar Observaciones", expanded=False):
    # Inicializar observaciones en session_state si no existen
    if 'observaciones_customizadas' not in st.session_state:
        st.session_state.observaciones_customizadas = observaciones.copy()
    
    # Campo para cada observación
    for i in range(len(st.session_state.observaciones_customizadas)):
        st.session_state.observaciones_customizadas[i] = st.text_input(
            f"Observación {i+1}", 
            value=st.session_state.observaciones_customizadas[i],
            key=f"obs_{i}"
        )
    
    # Botón para añadir nueva observación
    if st.button("Añadir Observación") and len(st.session_state.observaciones_customizadas) < 5:
        st.session_state.observaciones_customizadas.append("")
        st.experimental_rerun()

# Exportar reportes
with st.sidebar.expander("Exportar Reporte", expanded=False):
    # Texto para comentarios
    comentarios = st.text_area("Comentarios adicionales", height=100)
    
    # Botones de exportación
    col1, col2 = st.columns(2)
    
    try:
        excel_buffer = generate_excel(metrics, projections, program_analysis, comentarios, selected_marca)
        col1.download_button(
            "Excel", excel_buffer,
            file_name=f"reporte_{selected_marca}_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.ms-excel"
        )
    except:
        col1.warning("Error Excel")
    
    try:
        pptx_buffer = generate_pptx(metrics, projections, program_analysis, comentarios, selected_marca)
        col2.download_button(
            "PowerPoint", pptx_buffer,
            file_name=f"reporte_{selected_marca}_{datetime.now().strftime('%Y%m%d')}.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
    except:
        col2.warning("Error PPTX") 