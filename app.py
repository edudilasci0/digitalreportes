import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import io
from datetime import datetime
from fpdf import FPDF
import xlsxwriter

# Configuración de la página
st.set_page_config(
    page_title="Editor de Reportes Estratégicos",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Estilos CSS simplificados
st.markdown("""
<style>
    .main {
        background-color: #f8f9fa;
        font-family: 'Helvetica Neue', Arial, sans-serif;
    }
    
    .stButton button {
        background-color: #2196F3;
        color: white;
        border: none;
        border-radius: 4px;
        padding: 8px 16px;
    }
    
    h1, h2, h3 {
        color: #333;
    }
    
    .report-section {
        background-color: white;
        padding: 20px;
        border-radius: 8px;
        margin-bottom: 20px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.1);
    }
    
    .color-sample {
        display: inline-block;
        width: 20px;
        height: 20px;
        margin-right: 8px;
        border-radius: 3px;
    }
</style>
""", unsafe_allow_html=True)

# Inicializar estado si no existe
if 'kpi_data' not in st.session_state:
    st.session_state.kpi_data = {
        'matriculas': {'actual': 80, 'objetivo': 120},
        'leads': {'actual': 1200, 'objetivo': 1200},
        'tiempo': {'valor': 50}
    }

if 'proyeccion_data' not in st.session_state:
    st.session_state.proyeccion_data = {
        'valor': 110,
        'min': 95,
        'max': 125
    }

if 'programas_data' not in st.session_state:
    st.session_state.programas_data = [
        {'programa': 'Administración', 'leads': 300, 'matriculas': 25, 'conversion': 8.3},
        {'programa': 'Derecho', 'leads': 250, 'matriculas': 20, 'conversion': 8.0},
        {'programa': 'Marketing', 'leads': 200, 'matriculas': 15, 'conversion': 7.5},
        {'programa': 'Psicología', 'leads': 180, 'matriculas': 12, 'conversion': 6.7},
        {'programa': 'Economía', 'leads': 150, 'matriculas': 8, 'conversion': 5.3}
    ]

if 'colores_tema' not in st.session_state:
    st.session_state.colores_tema = {
        'estado_actual': '#2196F3',  # Azul
        'proyeccion': '#9C27B0',     # Morado
        'programas': '#FFC107'       # Amarillo
    }

if 'titulo_reporte' not in st.session_state:
    st.session_state.titulo_reporte = "GRADO - REPORTE ESTRATÉGICO"

if 'observaciones' not in st.session_state:
    st.session_state.observaciones = {
        'estado_actual': "El ritmo actual de matrículas está ligeramente por debajo de lo esperado.",
        'proyeccion': "Se proyecta alcanzar el objetivo con una probabilidad del 75%.",
        'programas': "Los programas de Administración y Derecho muestran el mejor rendimiento."
    }

# Título principal
st.title("Editor de Reportes Estratégicos")
st.write("Crea, personaliza y exporta reportes estratégicos de marketing educativo")

# SIDEBAR - Configuración general
st.sidebar.title("Configuración")

# Título del reporte
st.sidebar.subheader("Título del Reporte")
nuevo_titulo = st.sidebar.text_input("", st.session_state.titulo_reporte)
if nuevo_titulo != st.session_state.titulo_reporte:
    st.session_state.titulo_reporte = nuevo_titulo

# Colores del tema
st.sidebar.subheader("Colores del Tema")
col1, col2, col3 = st.sidebar.columns(3)
with col1:
    st.markdown(f"<div class='color-sample' style='background-color:{st.session_state.colores_tema['estado_actual']}'></div> Sección 1", unsafe_allow_html=True)
    nuevo_color1 = st.color_picker("", st.session_state.colores_tema['estado_actual'], key="color_seccion1")
    if nuevo_color1 != st.session_state.colores_tema['estado_actual']:
        st.session_state.colores_tema['estado_actual'] = nuevo_color1

with col2:
    st.markdown(f"<div class='color-sample' style='background-color:{st.session_state.colores_tema['proyeccion']}'></div> Sección 2", unsafe_allow_html=True)
    nuevo_color2 = st.color_picker("", st.session_state.colores_tema['proyeccion'], key="color_seccion2")
    if nuevo_color2 != st.session_state.colores_tema['proyeccion']:
        st.session_state.colores_tema['proyeccion'] = nuevo_color2

with col3:
    st.markdown(f"<div class='color-sample' style='background-color:{st.session_state.colores_tema['programas']}'></div> Sección 3", unsafe_allow_html=True)
    nuevo_color3 = st.color_picker("", st.session_state.colores_tema['programas'], key="color_seccion3")
    if nuevo_color3 != st.session_state.colores_tema['programas']:
        st.session_state.colores_tema['programas'] = nuevo_color3

# Opciones de exportación
st.sidebar.subheader("Exportar Reporte")
formato_exportacion = st.sidebar.selectbox("Formato", ["PDF", "Excel"])
if st.sidebar.button("Exportar Reporte"):
    st.sidebar.success(f"Reporte exportado en formato {formato_exportacion}")

# CONTENIDO PRINCIPAL - Tres pestañas para las secciones
tab1, tab2, tab3 = st.tabs(["ESTADO ACTUAL", "PROYECCIÓN", "PROGRAMAS"])

# TAB 1: ESTADO ACTUAL / RITMO DE AVANCE
with tab1:
    st.subheader("ESTADO ACTUAL / RITMO DE AVANCE")
    
    # KPIs editables
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.subheader("Matrículas")
        m_actual = st.number_input("Actual", min_value=0, value=st.session_state.kpi_data['matriculas']['actual'], key="mat_actual")
        m_objetivo = st.number_input("Objetivo", min_value=1, value=st.session_state.kpi_data['matriculas']['objetivo'], key="mat_objetivo")
        
        # Actualizar estado
        if m_actual != st.session_state.kpi_data['matriculas']['actual']:
            st.session_state.kpi_data['matriculas']['actual'] = m_actual
        if m_objetivo != st.session_state.kpi_data['matriculas']['objetivo']:
            st.session_state.kpi_data['matriculas']['objetivo'] = m_objetivo
            
        # Mostrar visualmente
        porcentaje = min(100, int((m_actual / max(1, m_objetivo)) * 100))
        st.write(f"**{m_actual} / {m_objetivo}** ({porcentaje}% de meta)")
        
    with col2:
        st.subheader("Leads")
        l_actual = st.number_input("Actual", min_value=0, value=st.session_state.kpi_data['leads']['actual'], key="leads_actual")
        l_objetivo = st.number_input("Objetivo", min_value=1, value=st.session_state.kpi_data['leads']['objetivo'], key="leads_objetivo")
        
        # Actualizar estado
        if l_actual != st.session_state.kpi_data['leads']['actual']:
            st.session_state.kpi_data['leads']['actual'] = l_actual
        if l_objetivo != st.session_state.kpi_data['leads']['objetivo']:
            st.session_state.kpi_data['leads']['objetivo'] = l_objetivo
            
        # Mostrar visualmente
        porcentaje = min(100, int((l_actual / max(1, l_objetivo)) * 100))
        st.write(f"**{l_actual} / {l_objetivo}** ({porcentaje}% de estimados)")
        
    with col3:
        st.subheader("Tiempo Transcurrido")
        t_valor = st.slider("Porcentaje", 0, 100, st.session_state.kpi_data['tiempo']['valor'])
        
        # Actualizar estado
        if t_valor != st.session_state.kpi_data['tiempo']['valor']:
            st.session_state.kpi_data['tiempo']['valor'] = t_valor
            
        # Mostrar visualmente
        st.write(f"**{t_valor}%** de la campaña")
    
    # Barras de progreso
    st.subheader("Progreso")
    
    # Barra de tiempo
    st.write("Tiempo transcurrido")
    st.progress(t_valor/100)
    
    # Barra de leads
    porcentaje_leads = min(100, int((l_actual / max(1, l_objetivo)) * 100))
    st.write(f"Leads acumulados: {l_actual} de {l_objetivo} ({porcentaje_leads}%)")
    st.progress(porcentaje_leads/100)
    
    # Barra de matrículas
    porcentaje_matriculas = min(100, int((m_actual / max(1, m_objetivo)) * 100))
    st.write(f"Matrículas confirmadas: {m_actual} de {m_objetivo} ({porcentaje_matriculas}%)")
    st.progress(porcentaje_matriculas/100)
    
    # Determinar estado
    if porcentaje_matriculas >= t_valor - 5:
        estado_actual = "En ritmo"
    elif porcentaje_matriculas >= t_valor - 15:
        estado_actual = "Justo"
    else:
        estado_actual = "Retrasado"
    
    # Mostrar estado
    st.write(f"**Estado:** {estado_actual}")
    
    # Observación estratégica (editable)
    st.subheader("Observación estratégica")
    nueva_observacion = st.text_area("", st.session_state.observaciones['estado_actual'], key="obs_estado")
    if nueva_observacion != st.session_state.observaciones['estado_actual']:
        st.session_state.observaciones['estado_actual'] = nueva_observacion

# TAB 2: PROYECCIÓN A CIERRE
with tab2:
    st.subheader("PROYECCIÓN A CIERRE")
    
    # Edición de valores de proyección
    col1, col2 = st.columns([3, 2])
    
    with col1:
        st.subheader("Valores de proyección")
        p_valor = st.number_input("Proyección", min_value=0, value=st.session_state.proyeccion_data['valor'], key="proy_valor")
        p_min = st.number_input("Mínimo", min_value=0, value=st.session_state.proyeccion_data['min'], key="proy_min")
        p_max = st.number_input("Máximo", min_value=0, value=st.session_state.proyeccion_data['max'], key="proy_max")
        
        # Actualizar estado
        if p_valor != st.session_state.proyeccion_data['valor']:
            st.session_state.proyeccion_data['valor'] = p_valor
        if p_min != st.session_state.proyeccion_data['min']:
            st.session_state.proyeccion_data['min'] = p_min
        if p_max != st.session_state.proyeccion_data['max']:
            st.session_state.proyeccion_data['max'] = p_max
            
        # Visualización simple
        fig, ax = plt.subplots(figsize=(10, 5))
        
        # Rango de valores para representar la curva
        rango = max(p_max - p_min, 10)  # Asegurar un rango mínimo
        x = np.linspace(p_min - rango*0.2, p_max + rango*0.2, 100)
        
        # Crear una curva de campana simple centrada en el valor proyectado
        media = p_valor
        desv_est = (p_max - p_min) / 4  # Una estimación razonable
        y = np.exp(-0.5 * ((x - media) / desv_est) ** 2) / (desv_est * np.sqrt(2 * np.pi))
        
        # Normalizar para mejor visualización
        y = y / max(y) * 0.8
        
        # Gráfico
        ax.fill_between(x, y, color=st.session_state.colores_tema['proyeccion'] + '40')  # Añadir transparencia
        ax.plot(x, y, color=st.session_state.colores_tema['proyeccion'], linewidth=2)
        
        # Línea para el valor esperado
        ax.axvline(x=p_valor, color=st.session_state.colores_tema['proyeccion'], linestyle='-', linewidth=2)
        
        # Línea para el objetivo
        objetivo = st.session_state.kpi_data['matriculas']['objetivo']
        ax.axvline(x=objetivo, color='#4CAF50' if p_valor >= objetivo else '#F44336', 
                linestyle='--', linewidth=2)
        
        # Configuración visual limpia
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)
        ax.spines['left'].set_visible(False)
        ax.set_yticks([])
        ax.set_xticks([objetivo, p_valor])
        ax.set_xticklabels([f'Meta: {objetivo}', f'Proyección: {p_valor}'])
        
        # Mostrar gráfico
        st.pyplot(fig)
    
    with col2:
        # Valor central destacado
        st.markdown(f"### Matrículas proyectadas")
        st.markdown(f"## {p_valor}")
        st.markdown(f"Intervalo de confianza: {p_min} – {p_max}")
        
        # Observación (editable)
        st.subheader("Observación")
        nueva_observacion = st.text_area("", st.session_state.observaciones['proyeccion'], key="obs_proyeccion")
        if nueva_observacion != st.session_state.observaciones['proyeccion']:
            st.session_state.observaciones['proyeccion'] = nueva_observacion

# TAB 3: DISTRIBUCIÓN DE RESULTADOS POR PROGRAMA
with tab3:
    st.subheader("DISTRIBUCIÓN DE RESULTADOS POR PROGRAMA")
    
    # Tabla editable con DataFrame
    st.subheader("Editar datos de programas")
    df_programas = pd.DataFrame(st.session_state.programas_data)
    
    # Convertir a editor
    edited_df = st.data_editor(
        df_programas,
        column_config={
            "programa": st.column_config.TextColumn("Programa"),
            "leads": st.column_config.NumberColumn("Leads"),
            "matriculas": st.column_config.NumberColumn("Matrículas"),
            "conversion": st.column_config.NumberColumn("Conversión (%)", format="%.1f")
        },
        num_rows="dynamic",
        use_container_width=True
    )
    
    # Actualizar datos si hay cambios
    if not edited_df.equals(df_programas):
        st.session_state.programas_data = edited_df.to_dict('records')
    
    # Mostrar visualización de la tabla
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("Top 5 programas con más matrículas")
        
        # Ordenar por matrículas descendente
        top_programas = edited_df.sort_values('matriculas', ascending=False).head(5)
        st.dataframe(top_programas)
    
    with col2:
        st.subheader("Top 5 programas con menor conversión")
        
        # Filtrar programas con al menos algunos leads
        programas_con_leads = edited_df[edited_df['leads'] > 5]
        
        # Ordenar por conversión ascendente
        menor_conversion = programas_con_leads.sort_values('conversion', ascending=True).head(5)
        st.dataframe(menor_conversion)
    
    # Observación estratégica (editable)
    st.subheader("Insight estratégico")
    nueva_observacion = st.text_area("", st.session_state.observaciones['programas'], key="obs_programas")
    if nueva_observacion != st.session_state.observaciones['programas']:
        st.session_state.observaciones['programas'] = nueva_observacion

# Vista previa del reporte
if st.button("Vista Previa del Reporte"):
    st.subheader("Vista Previa del Reporte")
    st.write(f"**Título:** {st.session_state.titulo_reporte}")
    
    st.write("### 1. ESTADO ACTUAL / RITMO DE AVANCE")
    st.write(f"Matrículas: {st.session_state.kpi_data['matriculas']['actual']} / {st.session_state.kpi_data['matriculas']['objetivo']}")
    st.write(f"Leads: {st.session_state.kpi_data['leads']['actual']} / {st.session_state.kpi_data['leads']['objetivo']}")
    st.write(f"Tiempo Transcurrido: {st.session_state.kpi_data['tiempo']['valor']}%")
    st.write(f"Observación: {st.session_state.observaciones['estado_actual']}")
    
    st.write("### 2. PROYECCIÓN A CIERRE")
    st.write(f"Matrículas Proyectadas: {st.session_state.proyeccion_data['valor']} ({st.session_state.proyeccion_data['min']} - {st.session_state.proyeccion_data['max']})")
    st.write(f"Observación: {st.session_state.observaciones['proyeccion']}")
    
    st.write("### 3. DISTRIBUCIÓN DE RESULTADOS POR PROGRAMA")
    st.dataframe(pd.DataFrame(st.session_state.programas_data))
    st.write(f"Insight: {st.session_state.observaciones['programas']}")
