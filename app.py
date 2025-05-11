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

# Función para crear tarjeta de métrica
def metric_card(title, value, description=None, color="#2196F3", icon=None, percentage=None):
    st.markdown(
        f"""
        <div style="
            background-color: white;
            border-radius: 10px;
            padding: 15px;
            border-left: 5px solid {color};
            margin-bottom: 10px;
            box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);">
            <h3 style="margin: 0; color: #333; font-size: 1rem;">{title}</h3>
            <h2 style="margin: 5px 0; color: {color}; font-size: 1.8rem; font-weight: bold;">{value}</h2>
            {f'<p style="margin: 0; color: #666; font-size: 0.8rem;">{description}</p>' if description else ''}
            {f'<div style="margin-top: 5px;"><span style="color: {"green" if float(percentage.replace("%","")) >= 50 else "red"}; font-weight: bold;">{percentage}</span></div>' if percentage else ''}
        </div>
        """, 
        unsafe_allow_html=True
    )

# Función para crear barra de progreso personalizada
def custom_progress_bar(current, total, label=None, color="#2196F3"):
    percent = min(100, int((current / max(1, total)) * 100))
    st.markdown(
        f"""
        <div style="margin-bottom: 5px;">
            <div style="display: flex; justify-content: space-between; margin-bottom: 5px;">
                <span style="font-size: 0.9rem; color: #666;">{label if label else ''}</span>
                <span style="font-size: 0.9rem; font-weight: bold; color: #333;">{current}/{total} ({percent}%)</span>
            </div>
            <div style="height: 20px; background-color: #f0f0f0; border-radius: 10px; overflow: hidden;">
                <div style="width: {percent}%; height: 100%; background-color: {color};"></div>
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

# Configuración de la página
st.set_page_config(
    page_title="Dashboard Marketing Educativo",
    page_icon="📊",
    layout="wide"
)

# Estilos CSS personalizados
st.markdown("""
<style>
    .main-title {
        font-size: 2.5rem !important;
        font-weight: 600 !important;
        margin-bottom: 1rem !important;
    }
    .section-title {
        font-size: 1.5rem !important;
        font-weight: 500 !important;
        margin-top: 1rem !important;
        margin-bottom: 1rem !important;
        padding-bottom: 0.5rem !important;
        border-bottom: 1px solid #f0f0f0 !important;
    }
    .observation-card {
        background-color: #f8f9fa;
        border-radius: 10px;
        padding: 15px;
        margin-bottom: 10px;
    }
    .stButton button {
        width: 100%;
    }
    .stDataFrame {
        border-radius: 10px !important;
        overflow: hidden !important;
    }
    .st-emotion-cache-1r6slb0 {
        max-width: 100% !important;
    }
</style>
""", unsafe_allow_html=True)

# Pestañas de la aplicación
tab1, tab2 = st.tabs(["📊 Dashboard", "⚙️ Configuración"])

with tab2:
    st.markdown("<h1 class='main-title'>Configuración del Reporte</h1>", unsafe_allow_html=True)
    
    # Carga de archivos
    st.markdown("<h2 class='section-title'>Carga de Archivos</h2>", unsafe_allow_html=True)
    
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
    st.markdown("<h2 class='section-title'>Parámetros de Reporte</h2>", unsafe_allow_html=True)
    
    # Configurar objetivo de matrículas
    objetivo_matriculas = st.number_input(
        "Objetivo de Matrículas", 
        min_value=1, 
        value=120, 
        help="Establece el objetivo de matrículas para esta marca y período"
    )
    
    # Opción para ingresar datos de planificación manualmente
    usar_planificacion_manual = not planificacion_file and st.checkbox("Ingresar datos de planificación manualmente", value=not planificacion_file)
    
    # Datos de planificación manual
    df_plan_mensual_manual = None
    df_inversion_manual = None 
    df_calendario_manual = None
    
    if usar_planificacion_manual:
        st.markdown("<h2 class='section-title'>Datos de Planificación Manual</h2>", unsafe_allow_html=True)
        
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
        st.markdown("<h2 class='section-title'>Calendario de Convocatoria</h2>", unsafe_allow_html=True)
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
                
                # Barra de progreso personalizada
                st.markdown("<h3>Tiempo transcurrido</h3>", unsafe_allow_html=True)
                custom_progress_bar(transcurrido, duracion_total, "Días de convocatoria")
                
                # Agregar contexto adicional
                st.write(f"Día {transcurrido} de {duracion_total} ({(fecha_fin - today.date()).days} días restantes)")
            else:
                st.error("La fecha de fin debe ser posterior a la fecha de inicio.")
    else:
        # Para marcas que no usan convocatorias
        fecha_inicio = None
        fecha_fin = None
        st.info(f"La marca {selected_marca} no se organiza por convocatorias con fechas fijas.")
    
    # Observaciones personalizables
    st.markdown("<h2 class='section-title'>Observaciones Personalizables</h2>", unsafe_allow_html=True)
    
    # Recuperar observaciones guardadas
    if 'observaciones' not in st.session_state:
        st.session_state.observaciones = [
            "Continuar con captación efectiva.",
            "Reforzar contacto sobre leads activos.",
            "Priorizar acciones de remarketing.",
            "Se proyecta el logro de la meta."
        ]
    
    observaciones = st.session_state.observaciones.copy()
    
    max_obs = 5
    num_obs = min(len(observaciones) + 1, max_obs)
    
    for i in range(num_obs):
        if i < len(observaciones):
            obs = st.text_input(f"Observación #{i+1}", value=observaciones[i], key=f"obs_{i}")
            observaciones[i] = obs
        else:
            nueva_obs = st.text_input(f"Observación #{i+1}", value="", key=f"obs_{i}")
            if nueva_obs:
                observaciones.append(nueva_obs)
    
    # Actualizar observaciones en session_state
    st.session_state.observaciones = [obs for obs in observaciones if obs]

with tab1:
    # Mostrar dashboard solo si hay datos cargados
    if not (matriculados_file and leads_file) and not st.button("Mostrar Demo (Datos de Ejemplo)"):
        st.info("Por favor, cargue los archivos requeridos en la pestaña de configuración o pulse 'Mostrar Demo' para ver un ejemplo.")
    else:
        # Si se pulsa el botón de demo, crear datos de ejemplo en memoria
        if not matriculados_file or not leads_file:
            from utils.data_generator import generate_demo_data
            df_matriculados, df_leads, df_plan_mensual, df_inversion, df_calendario = generate_demo_data(selected_marca)
        else:
            # Procesar datos
            try:
                df_matriculados = process_matriculados(matriculados_file)
                df_leads = process_leads(leads_file)
                
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
            except Exception as e:
                st.error(f"Error al procesar los archivos: {str(e)}")
                st.stop()
        
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
                'Programa': ['Todos los programas'],
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
        
        # DASHBOARD PRINCIPAL
        st.markdown(f"<h1 class='main-title'>Dashboard Principal - {selected_marca}</h1>", unsafe_allow_html=True)
        
        # Primera fila: KPIs principales
        col1, col2, col3 = st.columns(3)
        
        with col1:
            # Tarjeta de matrículas
            pct_objetivo = (metrics['matriculas_acumuladas'] / max(1, metrics['objetivo_matriculas'])) * 100
            metric_card(
                "Matrículas",
                f"{metrics['matriculas_acumuladas']} / {metrics['objetivo_matriculas']}",
                f"{pct_objetivo:.1f}% del objetivo",
                color="#2196F3",
                percentage=f"{pct_objetivo:.1f}%"
            )
        
        with col2:
            # Tarjeta de leads
            metric_card(
                "Leads Acumulados",
                f"{metrics['leads_acumulados']}",
                f"CPL: ${metrics['cpl_promedio']:.2f}",
                color="#4CAF50"
            )
        
        with col3:
            # Tarjeta de conversión
            metric_card(
                "Conversión",
                f"{metrics['tasa_conversion']:.1f}%",
                f"Programas: {metrics.get('programas_procesados', 0)}",
                color="#FF9800"
            )
        
        # Segunda fila: Gráficos y progreso
        col1, col2 = st.columns([3, 2])
        
        with col1:
            # Gráfico de progreso
            fig, ax = plt.subplots(figsize=(10, 5))
            
            # Datos para el gráfico
            if 'leads_proyectados' in projections:
                leads_estimados = projections['leads_proyectados'] + metrics['leads_acumulados']
            else:
                leads_estimados = metrics['leads_acumulados'] * 1.5
                
            leads_acumulados = metrics['leads_acumulados']
            matriculas = metrics['matriculas_acumuladas']
            objetivo = metrics['objetivo_matriculas']
            
            # Crear barras
            categorias = ['LEADS\nESTIMADOS', 'LEADS\nACUMULADOS', 'MATRÍCULAS']
            valores = [leads_estimados, leads_acumulados, matriculas]
            colores = ['#90CAF9', '#2196F3', '#1565C0']
            
            # Crear barras con anchos diferentes
            bar_width = 0.6
            bar_positions = np.arange(len(categorias))
            
            bars = ax.bar(bar_positions, valores, bar_width, color=colores)
            
            # Añadir valores encima de las barras
            for bar in bars:
                height = bar.get_height()
                ax.text(bar.get_x() + bar.get_width()/2., height + 0.1,
                        f'{int(height)}', ha='center', va='bottom', fontsize=11, fontweight='bold')
            
            # Configuración del gráfico
            ax.set_xticks(bar_positions)
            ax.set_xticklabels(categorias, fontsize=10, fontweight='bold')
            ax.spines['top'].set_visible(False)
            ax.spines['right'].set_visible(False)
            
            # Añadir línea de objetivo
            ax.axhline(y=objetivo, color='#FF5722', linestyle='--', linewidth=2, alpha=0.7)
            ax.text(len(categorias)-1, objetivo*1.05, f'Objetivo: {objetivo}', color='#FF5722', fontweight='bold')
            
            # Añadir porcentajes
            for i, v in enumerate(valores):
                if i == 2:  # Solo para matrículas
                    pct = (v / objetivo) * 100
                    ax.text(i, v/2, f'{pct:.1f}%', ha='center', color='white', fontweight='bold')
            
            ax.set_title('Estado de Campaña', fontsize=14, fontweight='bold', pad=20)
            
            st.pyplot(fig)
        
        with col2:
            st.markdown("<h2 class='section-title'>Progreso</h2>", unsafe_allow_html=True)
            
            # Tiempo transcurrido (solo para marcas con convocatorias)
            if selected_marca in ["GRADO", "UNISUD"] and metrics['tiempo_transcurrido'] is not None:
                custom_progress_bar(
                    current=f"{metrics['tiempo_transcurrido']:.1f}%", 
                    total="100%", 
                    label="Tiempo transcurrido",
                    color="#4CAF50"
                )
            
            # Progreso de matrículas
            custom_progress_bar(
                current=metrics['matriculas_acumuladas'],
                total=metrics['objetivo_matriculas'],
                label="Matrículas vs Objetivo",
                color="#2196F3"
            )
            
            # Composición de matrículas
            st.markdown("<h3>Composición de Matrículas</h3>", unsafe_allow_html=True)
            custom_progress_bar(
                current=f"{metrics['pct_matriculas_nuevos']:.1f}%",
                total="100%",
                label="Leads Nuevos",
                color="#1976D2"
            )
            custom_progress_bar(
                current=f"{metrics['pct_matriculas_remarketing']:.1f}%",
                total="100%",
                label="Remarketing",
                color="#FF9800"
            )
            
            # Observaciones
            st.markdown("<h3>Observaciones</h3>", unsafe_allow_html=True)
            st.markdown('<div class="observation-card">', unsafe_allow_html=True)
            for obs in st.session_state.observaciones:
                st.markdown(f"• {obs}", unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)
        
        # ESTIMACIÓN DE RESULTADOS
        st.markdown("<h1 class='section-title'>Estimación de Resultados</h1>", unsafe_allow_html=True)
        
        col1, col2 = st.columns([3, 2])
        
        with col1:
            # Gráfico de proyección
            fig, ax = plt.subplots(figsize=(10, 5))
            
            # Datos para la simulación
            simulacion = np.array(projections['simulacion_matriculas'])
            matriculas_mean = projections['matriculas_proyectadas_mean']
            percentil_05 = projections['matriculas_proyectadas_min']
            percentil_95 = projections['matriculas_proyectadas_max']
            
            # Crear el histograma
            n, bins, patches = ax.hist(simulacion, bins=30, alpha=0.4, color='#2196F3', density=True)
            
            # Añadir una curva de densidad
            import scipy.stats as stats
            mu, sigma = np.mean(simulacion), np.std(simulacion)
            x = np.linspace(mu - 3*sigma, mu + 3*sigma, 100)
            ax.plot(x, stats.norm.pdf(x, mu, sigma), color='#0D47A1', linewidth=2)
            
            # Añadir líneas verticales para la media y percentiles
            ax.axvline(x=matriculas_mean, color='#1976D2', linestyle='-', linewidth=2, 
                      label=f'Media: {matriculas_mean}')
            ax.axvline(x=percentil_05, color='#4CAF50', linestyle='--', linewidth=2,
                      label=f'P5: {percentil_05}')
            ax.axvline(x=percentil_95, color='#4CAF50', linestyle='--', linewidth=2,
                      label=f'P95: {percentil_95}')
            
            # Línea para objetivo
            ax.axvline(x=metrics['objetivo_matriculas'], color='#FF5722', linestyle=':', linewidth=2, 
                      label=f'Objetivo: {metrics["objetivo_matriculas"]}')
            
            # Áreas sombreadas para intervalos de confianza
            ax.fill_between(x, stats.norm.pdf(x, mu, sigma), 
                           where=(x >= percentil_05) & (x <= percentil_95),
                           color='#BBDEFB', alpha=0.5)
            
            ax.set_xlabel('Matrículas proyectadas')
            ax.set_ylabel('Densidad')
            ax.legend(loc='upper right')
            ax.set_title('Distribución de Matrículas Proyectadas', fontsize=14, fontweight='bold', pad=20)
            
            ax.spines['top'].set_visible(False)
            ax.spines['right'].set_visible(False)
            
            st.pyplot(fig)
            
        with col2:
            # Valor esperado prominente
            st.markdown(
                f"""
                <div style="
                    background-color: white;
                    border-radius: 10px;
                    padding: 20px;
                    text-align: center;
                    margin-bottom: 20px;
                    box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);">
                    <h1 style="font-size: 3rem; font-weight: bold; color: #2196F3; margin: 0;">
                        {metrics['matriculas_acumuladas'] + projections['matriculas_proyectadas_mean']}
                    </h1>
                    <p style="font-size: 1.5rem; margin: 5px 0;">MATRÍCULAS</p>
                    <p style="font-size: 1.2rem; color: #666;">(IC: {metrics['matriculas_acumuladas'] + percentil_05} - {metrics['matriculas_acumuladas'] + percentil_95})</p>
                </div>
                """, 
                unsafe_allow_html=True
            )
            
            # Probabilidades de logro
            st.markdown("<h3>Probabilidades de Logro</h3>", unsafe_allow_html=True)
            
            umbrales = [80, 90, 100, 110, 120]
            for umbral in umbrales:
                prob_key = f'prob_meta_{umbral}'
                
                # Determinar color según probabilidad
                if projections[prob_key] >= 75:
                    color = "#4CAF50"  # Verde
                elif projections[prob_key] >= 50:
                    color = "#FF9800"  # Naranja
                else:
                    color = "#F44336"  # Rojo
                
                custom_progress_bar(
                    current=f"{projections[prob_key]:.1f}%",
                    total="100%",
                    label=f"{umbral}% del objetivo",
                    color=color
                )
            
            # Recommendations summary
            st.markdown("<h3>Recomendaciones</h3>", unsafe_allow_html=True)
            
            # Determinar recomendaciones según proyecciones
            recomendaciones = []
            
            if projections['pct_cumplimiento_proyectado'] >= 100:
                recomendaciones.append("Se proyecta el logro de la meta.")
                recomendaciones.append("Continuar con captación efectiva.")
            elif projections['pct_cumplimiento_proyectado'] >= 80:
                recomendaciones.append("Se proyecta alcanzar entre el 80% y 100% de la meta.")
                recomendaciones.append("Intensificar acciones sobre leads activos.")
            else:
                recomendaciones.append("Proyección por debajo del 80% de la meta.")
                recomendaciones.append("Aumentar inversión y estrategias de remarketing.")
            
            # Mostrar recomendaciones
            st.markdown('<div class="observation-card">', unsafe_allow_html=True)
            for rec in recomendaciones:
                st.markdown(f"• {rec}", unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)

        # DISTRIBUCIÓN DE RESULTADOS POR PROGRAMA
        st.markdown(f"<h1 class='section-title'>Distribución de Resultados - {selected_marca}</h1>", unsafe_allow_html=True)
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("<h3>Top Programas (Volumen)</h3>", unsafe_allow_html=True)
            
            # Formatear datos para mostrar
            if 'top_matriculas' in program_analysis and not program_analysis['top_matriculas'].empty:
                top_programas = program_analysis['top_matriculas'].copy()
                # Límite a 5 programas
                top_programas = top_programas.head(5)
                
                # Estilizar DataFrame
                st.dataframe(
                    top_programas,
                    column_config={
                        "Programa": st.column_config.TextColumn("Programa"),
                        "Leads": st.column_config.NumberColumn("Leads", format="%d"),
                        "Matrículas": st.column_config.NumberColumn("Matrículas", format="%d"),
                        "Tasa Conversión (%)": st.column_config.NumberColumn("Conversión", format="%.1f%%"),
                    },
                    hide_index=True,
                    use_container_width=True
                )
            else:
                st.info("No hay datos suficientes para mostrar programas.")
        
        with col2:
            st.markdown("<h3>Top Programas (Conversión)</h3>", unsafe_allow_html=True)
            
            # Formatear datos para mostrar
            if 'tabla_completa' in program_analysis and not program_analysis['tabla_completa'].empty:
                top_conversion = program_analysis['tabla_completa'].copy()
                # Ordenar por tasa de conversión (descendente) y filtrar a los que tengan al menos 5 leads
                top_conversion = top_conversion[top_conversion['Leads'] >= 5].sort_values('Tasa Conversión (%)', ascending=False).head(5)
                
                # Estilizar DataFrame
                st.dataframe(
                    top_conversion,
                    column_config={
                        "Programa": st.column_config.TextColumn("Programa"),
                        "Leads": st.column_config.NumberColumn("Leads", format="%d"),
                        "Matrículas": st.column_config.NumberColumn("Matrículas", format="%d"),
                        "Tasa Conversión (%)": st.column_config.NumberColumn("Conversión", format="%.1f%%"),
                    },
                    hide_index=True,
                    use_container_width=True
                )
            else:
                st.info("No hay datos suficientes para mostrar programas.")
        
        # Tabla detallada y exportación
        with st.expander("Ver detalles por programa", expanded=False):
            if 'tabla_completa' in program_analysis and not program_analysis['tabla_completa'].empty:
                st.dataframe(
                    program_analysis['tabla_completa'],
                    column_config={
                        "Programa": st.column_config.TextColumn("Programa"),
                        "Leads": st.column_config.NumberColumn("Leads", format="%d"),
                        "Matrículas": st.column_config.NumberColumn("Matrículas", format="%d"),
                        "Tasa Conversión (%)": st.column_config.NumberColumn("Conversión", format="%.1f%%"),
                        "Clasificación": st.column_config.TextColumn("Clasificación"),
                    },
                    hide_index=True,
                    use_container_width=True
                )
            else:
                st.info("No hay datos detallados para mostrar.")

        # Exportar reportes
        st.markdown("<h2 class='section-title'>Exportar Reporte</h2>", unsafe_allow_html=True)
        
        # Sección de comentarios
        comentarios = st.text_area("Comentarios adicionales para el reporte", height=100)
        
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
            col1.error(f"Error al generar Excel")
        
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
            col2.error(f"Error al generar PDF")
        
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
            col3.error(f"Error al generar PowerPoint")

# Información adicional en la barra lateral
with st.sidebar:
    st.title("Información")
    st.info("""
    Este sistema genera reportes estratégicos semanales por marca y programa educativo, 
    adaptados a los diferentes modelos como GRADO (convocatorias fijas) y 
    POSGRADO (cohortes variables y continuas ADVANCE).
    """)
    
    st.title("Marcas")
    for marca in marcas:
        st.markdown(f"- {marca}")
    
    # Botón para generar datos de ejemplo
    if st.button("Generar Datos de Ejemplo"):
        try:
            from utils.data_generator import generate_sample_data
            generate_sample_data()
            st.success("¡Datos de ejemplo generados correctamente! Verifica la carpeta 'sample_data'.")
        except Exception as e:
            st.error(f"Error al generar datos de ejemplo: {str(e)}") 