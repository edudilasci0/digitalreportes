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
    page_title="Dashboard Educativo",
    page_icon="��",
    layout="wide",
    initial_sidebar_state="expanded"  # Asegurar que la sidebar esté expandida inicialmente
)

# Estilos CSS personalizados para un diseño limpio
st.markdown("""
<style>
    /* Estilo general */
    body {
        font-family: 'Helvetica Neue', Arial, sans-serif;
        color: #333;
        background-color: #f8f9fa;
    }
    
    /* Secciones/Slides */
    .slide {
        background-color: white;
        border-radius: 12px;
        padding: 30px;
        margin-bottom: 30px;
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.05);
    }
    
    /* Encabezados de sección */
    .slide-header {
        display: flex;
        align-items: center;
        margin-bottom: 25px;
    }
    
    .slide-indicator {
        width: 18px;
        height: 18px;
        border-radius: 50%;
        margin-right: 12px;
        display: inline-block;
    }
    
    .indicator-blue {
        background-color: #2196F3;
    }
    
    .indicator-purple {
        background-color: #9C27B0;
    }
    
    .indicator-yellow {
        background-color: #FFC107;
    }
    
    .slide-title {
        font-size: 22px;
        font-weight: 500;
        color: #333;
        margin: 0;
    }
    
    /* KPIs y métricas */
    .kpi-container {
        display: flex;
        justify-content: space-between;
        margin-bottom: 20px;
    }
    
    .kpi-card {
        background-color: #f8f9fa;
        border-radius: 8px;
        padding: 20px;
        text-align: center;
        width: 32%;
    }
    
    .kpi-title {
        font-size: 14px;
        font-weight: 500;
        color: #666;
        margin-bottom: 10px;
    }
    
    .kpi-value {
        font-size: 28px;
        font-weight: 700;
        color: #333;
        margin-bottom: 5px;
    }
    
    .kpi-meta {
        font-size: 14px;
        color: #666;
    }
    
    /* Barras de progreso */
    .progress-container {
        margin-bottom: 20px;
    }
    
    .progress-label {
        display: flex;
        justify-content: space-between;
        margin-bottom: 5px;
    }
    
    .progress-bar-bg {
        background-color: #f0f0f0;
        border-radius: 8px;
        height: 20px;
        overflow: hidden;
    }
    
    .progress-bar-fill {
        height: 100%;
        border-radius: 8px;
        transition: width 0.3s ease;
    }
    
    .progress-bar-fill-blue {
        background-color: #2196F3;
    }
    
    .progress-bar-fill-purple {
        background-color: #9C27B0;
    }
    
    .progress-bar-fill-yellow {
        background-color: #FFC107;
    }
    
    /* Observaciones y textos */
    .observation-box {
        background-color: #f8f9fa;
        border-radius: 8px;
        padding: 20px;
        margin-top: 20px;
    }
    
    .observation-title {
        font-weight: 600;
        margin-bottom: 10px;
    }
    
    /* Tablas */
    .data-table {
        width: 100%;
        border-collapse: collapse;
    }
    
    .data-table th {
        background-color: #f5f7f9;
        padding: 12px;
        text-align: left;
        font-weight: 500;
    }
    
    .data-table td {
        padding: 12px;
        border-bottom: 1px solid #f0f0f0;
    }
    
    /* Proyección central */
    .projection-central {
        text-align: center;
        padding: 30px;
    }
    
    .projection-value {
        font-size: 56px;
        font-weight: 700;
        color: #9C27B0;
        margin-bottom: 0;
    }
    
    .projection-label {
        font-size: 18px;
        color: #666;
        margin-top: 5px;
    }
    
    .projection-interval {
        font-size: 16px;
        color: #666;
        margin-top: 5px;
    }
    
    /* Estado visual */
    .status-indicator {
        display: inline-block;
        padding: 6px 12px;
        border-radius: 16px;
        font-weight: 500;
        font-size: 14px;
    }
    
    .status-on-track {
        background-color: #E8F5E9;
        color: #388E3C;
    }
    
    .status-behind {
        background-color: #FFEBEE;
        color: #D32F2F;
    }
    
    .status-just-in-time {
        background-color: #FFF8E1;
        color: #FFA000;
    }
    
    /* Ocultar elementos de Streamlit */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    .stDeployButton {display:none;}
    
    /* Ajustes responsivos */
    @media (max-width: 768px) {
        .kpi-container {
            flex-direction: column;
        }
        .kpi-card {
            width: 100%;
            margin-bottom: 10px;
        }
    }
    
    /* Eliminar márgenes de streamlit */
    .block-container {
        padding-top: 1rem !important;
        padding-bottom: 1rem !important;
    }
    .main .block-container {
        max-width: 100% !important;
        padding-left: 1rem !important;
        padding-right: 1rem !important;
    }
    
    /* Asegurar que la sidebar sea visible */
    .css-1d391kg, .css-1lcbmhc {
        visibility: visible !important;
        opacity: 1 !important;
        display: block !important;
    }
    
    /* Botón para mostrar panel de configuración */
    .show-sidebar-button {
        position: fixed;
        top: 10px;
        left: 10px;
        background-color: #2196F3;
        color: white;
        padding: 8px 12px;
        border-radius: 5px;
        z-index: 1000;
        cursor: pointer;
        font-weight: bold;
    }
</style>

<script>
// Script para controlar el botón de mostrar/ocultar sidebar
document.addEventListener('DOMContentLoaded', function() {
    // Crear botón si no existe
    if (!document.querySelector('.show-sidebar-button')) {
        const button = document.createElement('button');
        button.className = 'show-sidebar-button';
        button.innerHTML = 'Mostrar Configuración';
        button.onclick = function() {
            // Buscar el elemento de la sidebar por sus clases comunes
            const sidebar = document.querySelector('.css-1d391kg, .css-1lcbmhc');
            if (sidebar) {
                if (sidebar.style.display === 'none') {
                    sidebar.style.display = 'block';
                    this.innerHTML = 'Ocultar Configuración';
                } else {
                    sidebar.style.display = 'none';
                    this.innerHTML = 'Mostrar Configuración';
                }
            }
        };
        document.body.appendChild(button);
    }
});
</script>
""", unsafe_allow_html=True)

# Botón adicional para mostrar la sidebar (para dispositivos móviles o si está oculta)
if st.button("📋 Mostrar/Ocultar Panel de Configuración"):
    st.session_state['show_sidebar'] = not st.session_state.get('show_sidebar', True)
    if st.session_state['show_sidebar']:
        st.sidebar.markdown("### Panel de Configuración Visible")
    else:
        st.sidebar.markdown("### Panel de Configuración Oculto")
    st.rerun()

# Funciones para componentes UI
def create_slide_header(title, color):
    """Crear el encabezado de una sección/slide"""
    indicator_class = {
        "blue": "indicator-blue",
        "purple": "indicator-purple",
        "yellow": "indicator-yellow"
    }.get(color, "indicator-blue")
    
    html = f"""
    <div class="slide-header">
        <div class="slide-indicator {indicator_class}"></div>
        <h2 class="slide-title">{title}</h2>
    </div>
    """
    return st.markdown(html, unsafe_allow_html=True)

def create_kpi_card(title, value, meta=None):
    """Crear una tarjeta KPI"""
    meta_html = f'<div class="kpi-meta">{meta}</div>' if meta else ''
    
    html = f"""
    <div class="kpi-card">
        <div class="kpi-title">{title}</div>
        <div class="kpi-value">{value}</div>
        {meta_html}
    </div>
    """
    return html

def create_progress_bar(label, current, total, current_text="", color="blue"):
    """Crear una barra de progreso personalizada"""
    percent = min(100, int((current / max(1, total)) * 100))
    color_class = f"progress-bar-fill-{color}"
    
    html = f"""
    <div class="progress-container">
        <div class="progress-label">
            <span>{label}</span>
            <span>{current_text} ({percent}%)</span>
        </div>
        <div class="progress-bar-bg">
            <div class="progress-bar-fill {color_class}" style="width: {percent}%;"></div>
        </div>
    </div>
    """
    return html

def create_status_indicator(status):
    """Crear indicador de estado"""
    if status == "on_track":
        class_name = "status-on-track"
        text = "En ritmo"
    elif status == "behind":
        class_name = "status-behind"
        text = "Retrasado"
    else:  # just_in_time
        class_name = "status-just-in-time"
        text = "Justo"
    
    html = f'<span class="status-indicator {class_name}">{text}</span>'
    return html

def create_observation_box(title, text):
    """Crear caja de observación"""
    html = f"""
    <div class="observation-box">
        <div class="observation-title">{title}</div>
        <div>{text}</div>
    </div>
    """
    return html

def create_data_table(headers, rows):
    """Crear tabla de datos"""
    html = '<table class="data-table">\n<thead>\n<tr>'
    
    # Encabezados
    for header in headers:
        html += f'<th>{header}</th>'
    html += '</tr>\n</thead>\n<tbody>'
    
    # Filas
    for row in rows:
        html += '<tr>'
        for cell in row:
            html += f'<td>{cell}</td>'
        html += '</tr>'
    
    html += '</tbody>\n</table>'
    return html

# Cargar datos
# Sidebar para configuración
st.sidebar.title("Configuración")

# Selector de marca
marcas = ["GRADO", "POSGRADO", "ADVANCE", "WIZARD", "AJA", "UNISUD"]
selected_marca = st.sidebar.selectbox("Seleccionar Marca", marcas)

# Opción para usar datos de ejemplo
usar_demo = st.sidebar.checkbox("Usar datos de ejemplo", value=True)

# Configurar objetivo y fechas
objetivo_matriculas = st.sidebar.number_input("Objetivo de Matrículas", min_value=1, value=120)

# Cargar datos (de ejemplo o reales)
if usar_demo:
    # Generar datos de demo
    from utils.data_generator import generate_demo_data
    df_matriculados, df_leads, df_plan_mensual, df_inversion, df_calendario = generate_demo_data(selected_marca)
    
    # Mensaje de datos de ejemplo
    st.sidebar.success("Usando datos de ejemplo")
else:
    # Cargar archivos
    matriculados_file = st.sidebar.file_uploader(f"Matriculados - {selected_marca}", type=["xlsx"])
    leads_file = st.sidebar.file_uploader(f"Leads Activos - {selected_marca}", type=["xlsx"])
    planificacion_file = st.sidebar.file_uploader(f"Planificación (opcional)", type=["xlsx"])
    
    # Salir si no hay archivos necesarios
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

# Calcular valores adicionales para el dashboard
pct_objetivo = (metrics['matriculas_acumuladas'] / objetivo_matriculas) * 100
pct_tiempo = metrics['tiempo_transcurrido'] if metrics['tiempo_transcurrido'] is not None else 50
if 'leads_proyectados' in projections:
    leads_estimados = projections['leads_proyectados'] + metrics['leads_acumulados']
else:
    leads_estimados = int(metrics['leads_acumulados'] * 1.5)
pct_leads = (metrics['leads_acumulados'] / leads_estimados) * 100

# Determinar el estado basado en las métricas
if pct_objetivo >= pct_tiempo - 5:
    status = "on_track"  # En ritmo
elif pct_objetivo >= pct_tiempo - 15:
    status = "just_in_time"  # Justo
else:
    status = "behind"  # Retrasado

# CONTENIDO PRINCIPAL DEL DASHBOARD
st.title(f"Dashboard {selected_marca}")

# ---------------------------------------
# SECCIÓN 1: ESTADO ACTUAL / RITMO DE AVANCE
# ---------------------------------------
st.markdown('<div class="slide">', unsafe_allow_html=True)
create_slide_header("ESTADO ACTUAL / RITMO DE AVANCE", "blue")

# KPIs destacados
kpi_html = f"""
<div class="kpi-container">
    {create_kpi_card("Matrículas", f"{metrics['matriculas_acumuladas']} / {objetivo_matriculas}", f"{pct_objetivo:.1f}% de meta")}
    {create_kpi_card("Leads", f"{metrics['leads_acumulados']} / {leads_estimados}", f"{pct_leads:.1f}% de estimados")}
    {create_kpi_card("Tiempo Transcurrido", f"{pct_tiempo:.1f}%", "de la campaña")}
</div>
"""
st.markdown(kpi_html, unsafe_allow_html=True)

# Barras horizontales de progreso
progress_html = f"""
<div style="margin-top: 30px;">
    {create_progress_bar("Tiempo transcurrido", pct_tiempo, 100, f"{pct_tiempo:.1f}%", "blue")}
    {create_progress_bar("Leads acumulados", metrics['leads_acumulados'], leads_estimados, f"{metrics['leads_acumulados']} de {leads_estimados}", "blue")}
    {create_progress_bar("Matrículas confirmadas", metrics['matriculas_acumuladas'], objetivo_matriculas, f"{metrics['matriculas_acumuladas']} de {objetivo_matriculas}", "blue")}
</div>
"""
st.markdown(progress_html, unsafe_allow_html=True)

# Indicador visual de estado
status_html = f"""
<div style="margin-top: 30px;">
    <strong>Estado:</strong> {create_status_indicator(status)}
</div>
"""
st.markdown(status_html, unsafe_allow_html=True)

# Observación estratégica
observacion_texto = ""
if status == "on_track":
    observacion_texto = "El ritmo actual sostiene la meta. No se requieren ajustes."
elif status == "just_in_time":
    observacion_texto = "El ritmo está justo para alcanzar la meta. Se recomienda vigilar conversión."
else:
    observacion_texto = "El ritmo actual no alcanzará la meta. Se requieren ajustes en la estrategia."

observation_html = create_observation_box("Observación estratégica", observacion_texto)
st.markdown(observation_html, unsafe_allow_html=True)

st.markdown('</div>', unsafe_allow_html=True)

# ---------------------------------------
# SECCIÓN 2: PROYECCIÓN A CIERRE
# ---------------------------------------
st.markdown('<div class="slide">', unsafe_allow_html=True)
create_slide_header("PROYECCIÓN A CIERRE", "purple")

# Distribuir contenido en dos columnas
col1, col2 = st.columns([3, 2])

with col1:
    # Visual de proyección
    fig, ax = plt.subplots(figsize=(10, 5))
    
    # Datos para proyección
    simulacion = np.array(projections['simulacion_matriculas'])
    matriculas_mean = projections['matriculas_proyectadas_mean']
    percentil_05 = projections['matriculas_proyectadas_min']
    percentil_95 = projections['matriculas_proyectadas_max']
    
    # Crear visualización simplificada
    import scipy.stats as stats
    x = np.linspace(min(simulacion) * 0.9, max(simulacion) * 1.1, 100)
    
    try:
        # Verificar si hay suficiente variación en los datos
        if len(np.unique(simulacion)) <= 1:
            # Si todos los valores son iguales, crear una distribución artificial
            y = np.zeros_like(x)
            idx = np.abs(x - matriculas_mean).argmin()  # índice más cercano al valor medio
            y[idx] = 1.0  # Poner un pico en el valor medio
            
            # Suavizar ligeramente para visualización
            from scipy.ndimage import gaussian_filter1d
            y = gaussian_filter1d(y, sigma=2)
        else:
            # Intentar añadir pequeña variación si los datos son muy similares
            if np.std(simulacion) < 0.01:
                # Añadir pequeño ruido aleatorio
                noise = np.random.normal(0, 0.01, size=len(simulacion))
                simulacion_adj = simulacion + noise
                kde = stats.gaussian_kde(simulacion_adj)
            else:
                # Usar KDE normal si hay suficiente variación
                kde = stats.gaussian_kde(simulacion)
            
            y = kde(x)
    
    except np.linalg.LinAlgError:
        # Método alternativo sin KDE si falla
        y = np.zeros_like(x)
        hist, bin_edges = np.histogram(simulacion, bins=20, density=True)
        bin_centers = (bin_edges[:-1] + bin_edges[1:]) / 2
        
        # Interpolar histograma para crear curva suave
        from scipy.interpolate import interp1d
        if len(bin_centers) > 3:  # Necesitamos al menos algunos puntos para interpolar
            f = interp1d(bin_centers, hist, kind='quadratic', bounds_error=False, fill_value=0)
            y = f(x)
        else:
            # Fallback simple si no hay suficientes puntos
            idx = np.abs(x - matriculas_mean).argmin()
            y[max(0, idx-5):min(len(y), idx+5)] = 0.2
    
    # Área bajo la curva
    ax.fill_between(x, y, color='#E8EAF6', alpha=0.7)
    ax.plot(x, y, color='#9C27B0', linewidth=2)
    
    # Línea para el valor esperado
    ax.axvline(x=matriculas_mean, color='#7B1FA2', linestyle='-', linewidth=2)
    
    # Línea para el objetivo
    ax.axvline(x=objetivo_matriculas, color='#4CAF50' if matriculas_mean >= objetivo_matriculas else '#F44336', 
               linestyle='--', linewidth=2)
    
    # Configuración visual limpia
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['left'].set_visible(False)
    ax.set_yticks([])
    ax.set_xticks([objetivo_matriculas, int(matriculas_mean)])
    ax.set_xticklabels([f'Meta: {objetivo_matriculas}', f'Proyección: {int(matriculas_mean)}'])
    
    # Mostrar gráfico
    st.pyplot(fig)

with col2:
    # Proyección central
    st.markdown(f"""
    <div class="projection-central">
        <div class="projection-value">{int(matriculas_mean)}</div>
        <div class="projection-label">MATRÍCULAS</div>
        <div class="projection-interval">Intervalo de confianza: {percentil_05} – {percentil_95}</div>
    </div>
    """, unsafe_allow_html=True)
    
    # Texto lateral breve
    st.markdown(create_observation_box("Metodología de proyección", 
                "La proyección se basa en el rendimiento actual acumulado, aplicando simulación Monte Carlo sobre las tendencias de conversión históricas."), 
                unsafe_allow_html=True)

st.markdown('</div>', unsafe_allow_html=True)

# ---------------------------------------
# SECCIÓN 3: DISTRIBUCIÓN DE RESULTADOS POR PROGRAMA
# ---------------------------------------
st.markdown('<div class="slide">', unsafe_allow_html=True)
create_slide_header("DISTRIBUCIÓN DE RESULTADOS POR PROGRAMA", "yellow")

# Preparar datos para tablas
if 'top_matriculas' in program_analysis and not program_analysis['top_matriculas'].empty:
    # Top 5 programas con más matrículas
    top_matriculas = program_analysis['top_matriculas'].head(5)
    top_matriculas_rows = []
    for _, row in top_matriculas.iterrows():
        top_matriculas_rows.append([
            row['Programa'],
            int(row['Leads']),
            int(row['Matrículas']),
            f"{row['Tasa Conversión (%)']:.1f}%"
        ])
    
    # Top 5 programas con menor conversión
    programas_menor_conversion = program_analysis['menor_conversion'].head(5)
    menor_conversion_rows = []
    for _, row in programas_menor_conversion.iterrows():
        menor_conversion_rows.append([
            row['Programa'],
            int(row['Leads']),
            int(row['Matrículas']),
            f"{row['Tasa Conversión (%)']:.1f}%"
        ])
    
    # Mostrar tablas en dos columnas
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("Top 5 programas con más matrículas")
        top_table = create_data_table(
            ["Programa", "Leads", "Matrículas", "Conversión"],
            top_matriculas_rows
        )
        st.markdown(top_table, unsafe_allow_html=True)
    
    with col2:
        st.subheader("Top 5 programas con menor conversión")
        bottom_table = create_data_table(
            ["Programa", "Leads", "Matrículas", "Conversión"],
            menor_conversion_rows
        )
        st.markdown(bottom_table, unsafe_allow_html=True)
    
    # Opción para ver tabla completa
    with st.expander("Ver tabla completa de programas"):
        # Preparar datos para tabla completa
        tabla_completa = program_analysis['tabla_completa']
        tabla_completa_rows = []
        for _, row in tabla_completa.iterrows():
            # Determinar clasificación
            if row['Tasa Conversión (%)'] > 15:
                clasificacion = "Excelente conversión"
            elif row['Tasa Conversión (%)'] < 5 and row['Leads'] > 10:
                clasificacion = "Bajo rendimiento"
            else:
                clasificacion = "Normal"
            
            tabla_completa_rows.append([
                row['Programa'],
                int(row['Leads']),
                int(row['Matrículas']),
                f"{row['Tasa Conversión (%)']:.1f}%",
                clasificacion
            ])
        
        # Mostrar tabla completa
        complete_table = create_data_table(
            ["Programa", "Leads", "Matrículas", "Conversión", "Clasificación"],
            tabla_completa_rows
        )
        st.markdown(complete_table, unsafe_allow_html=True)
    
    # Insight al pie
    insights = []
    # Detectar programas con buena conversión pero pocos leads
    buena_conversion = tabla_completa[(tabla_completa['Tasa Conversión (%)'] > 15) & (tabla_completa['Leads'] < 50)]
    if not buena_conversion.empty:
        programas = ", ".join(buena_conversion['Programa'].head(3).tolist())
        insights.append(f"Programas con excelente conversión pero poca inversión ({programas}): oportunidad de escalar.")
    
    # Detectar programas con muchos leads pero baja conversión
    baja_conversion = tabla_completa[(tabla_completa['Tasa Conversión (%)'] < 5) & (tabla_completa['Leads'] > 100)]
    if not baja_conversion.empty:
        programas = ", ".join(baja_conversion['Programa'].head(3).tolist())
        insights.append(f"Programas con alta inversión pero baja conversión ({programas}): revisar propuesta de valor.")
    
    # Si no hay insights específicos, añadir uno general
    if not insights:
        insights.append("Algunos programas muestran oportunidades de optimización y balance en la inversión.")
    
    # Mostrar insights
    insight_html = create_observation_box("Insight estratégico", insights[0])
    st.markdown(insight_html, unsafe_allow_html=True)
else:
    st.info("No hay datos suficientes para mostrar la distribución por programas.")

st.markdown('</div>', unsafe_allow_html=True)

# ---------------------------------------
# HERRAMIENTAS DE EXPORTACIÓN
# ---------------------------------------
st.sidebar.markdown("### Exportar Reporte")

# Texto para comentarios
comentarios = st.sidebar.text_area("Comentarios adicionales", height=100)

# Botones de exportación
col1, col2 = st.sidebar.columns(2)

try:
    excel_buffer = generate_excel(metrics, projections, program_analysis, comentarios, selected_marca)
    col1.download_button(
        "Descargar Excel", excel_buffer,
        file_name=f"reporte_{selected_marca}_{datetime.now().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.ms-excel"
    )
except:
    col1.warning("Error Excel")

try:
    pdf_buffer = generate_pdf(metrics, projections, program_analysis, comentarios, selected_marca)
    col2.download_button(
        "Descargar PDF", pdf_buffer,
        file_name=f"reporte_{selected_marca}_{datetime.now().strftime('%Y%m%d')}.pdf",
        mime="application/pdf"
    )
except:
    col2.warning("Error PDF") 