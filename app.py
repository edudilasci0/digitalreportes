import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import io
from datetime import datetime
from fpdf import FPDF
import base64

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
        'matriculas': 110,
        'leads': 1500
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
        'proyeccion': "Se proyecta alcanzar el objetivo con un buen ritmo de conversión.",
        'programas': "Los programas de Administración y Derecho muestran el mejor rendimiento."
    }

# Funciones para exportación
def generate_excel():
    """Genera un reporte en formato Excel"""
    buffer = io.BytesIO()
    writer = pd.ExcelWriter(buffer, engine='xlsxwriter')
    
    # Hoja 1: Resumen General
    df_resumen = pd.DataFrame({
        'Métrica': [
            'Título del Reporte',
            'Matrículas Actuales',
            'Objetivo de Matrículas',
            'Leads Actuales',
            'Objetivo de Leads',
            'Tiempo Transcurrido',
            'Matrículas Proyectadas',
            'Leads Proyectados'
        ],
        'Valor': [
            st.session_state.titulo_reporte,
            st.session_state.kpi_data['matriculas']['actual'],
            st.session_state.kpi_data['matriculas']['objetivo'],
            st.session_state.kpi_data['leads']['actual'],
            st.session_state.kpi_data['leads']['objetivo'],
            f"{st.session_state.kpi_data['tiempo']['valor']}%",
            st.session_state.proyeccion_data['matriculas'],
            st.session_state.proyeccion_data['leads']
        ]
    })
    df_resumen.to_excel(writer, sheet_name='Resumen', index=False)
    
    # Hoja 2: Observaciones
    df_obs = pd.DataFrame({
        'Sección': [
            'Estado Actual',
            'Proyección', 
            'Programas'
        ],
        'Observación': [
            st.session_state.observaciones['estado_actual'],
            st.session_state.observaciones['proyeccion'],
            st.session_state.observaciones['programas']
        ]
    })
    df_obs.to_excel(writer, sheet_name='Observaciones', index=False)
    
    # Hoja 3: Datos de Programas
    df_programas = pd.DataFrame(st.session_state.programas_data)
    df_programas.to_excel(writer, sheet_name='Programas', index=False)
    
    # Dar formato
    workbook = writer.book
    
    # Formato para título
    header_format = workbook.add_format({
        'bold': True,
        'bg_color': '#2196F3',
        'font_color': 'white',
        'border': 1
    })
    
    # Aplicar formato a las hojas
    for sheet_name in writer.sheets:
        worksheet = writer.sheets[sheet_name]
        # Ajustar ancho de columnas
        worksheet.set_column('A:A', 25)
        worksheet.set_column('B:D', 18)
        
        # Formato para encabezados
        for col_num, value in enumerate(df_resumen.columns.values):
            worksheet.write(0, col_num, value, header_format)
    
    writer.close()
    buffer.seek(0)
    return buffer

def generate_pdf():
    """Genera un reporte en formato PDF"""
    class PDF(FPDF):
        def header(self):
            # Título del documento
            self.set_font('Arial', 'B', 15)
            self.cell(0, 10, st.session_state.titulo_reporte, 0, 1, 'C')
            self.ln(10)
        
        def footer(self):
            # Pie de página
            self.set_y(-15)
            self.set_font('Arial', 'I', 8)
            self.cell(0, 10, f'Página {self.page_no()}', 0, 0, 'C')
    
    pdf = PDF()
    pdf.add_page()
    
    # Sección 1: ESTADO ACTUAL / RITMO DE AVANCE
    pdf.set_font('Arial', 'B', 12)
    pdf.set_fill_color(33, 150, 243)  # Color azul
    pdf.cell(0, 10, 'ESTADO ACTUAL / RITMO DE AVANCE', 0, 1, 'L', 1)
    pdf.ln(5)
    
    pdf.set_font('Arial', '', 11)
    # Tabla de KPIs
    pdf.cell(50, 10, 'Métricas', 1)
    pdf.cell(35, 10, 'Actual', 1)
    pdf.cell(35, 10, 'Objetivo', 1)
    pdf.cell(35, 10, 'Porcentaje', 1)
    pdf.ln()
    
    # Datos de Matrículas
    m_actual = st.session_state.kpi_data['matriculas']['actual']
    m_objetivo = st.session_state.kpi_data['matriculas']['objetivo']
    m_porcentaje = min(100, int((m_actual / max(1, m_objetivo)) * 100))
    
    pdf.cell(50, 10, 'Matrículas', 1)
    pdf.cell(35, 10, str(m_actual), 1)
    pdf.cell(35, 10, str(m_objetivo), 1)
    pdf.cell(35, 10, f"{m_porcentaje}%", 1)
    pdf.ln()
    
    # Datos de Leads
    l_actual = st.session_state.kpi_data['leads']['actual']
    l_objetivo = st.session_state.kpi_data['leads']['objetivo']
    l_porcentaje = min(100, int((l_actual / max(1, l_objetivo)) * 100))
    
    pdf.cell(50, 10, 'Leads', 1)
    pdf.cell(35, 10, str(l_actual), 1)
    pdf.cell(35, 10, str(l_objetivo), 1)
    pdf.cell(35, 10, f"{l_porcentaje}%", 1)
    pdf.ln()
    
    # Tiempo Transcurrido
    pdf.cell(50, 10, 'Tiempo Transcurrido', 1)
    pdf.cell(35, 10, f"{st.session_state.kpi_data['tiempo']['valor']}%", 1)
    pdf.cell(35, 10, "100%", 1)
    pdf.cell(35, 10, f"{st.session_state.kpi_data['tiempo']['valor']}%", 1)
    pdf.ln(15)
    
    # Observación
    pdf.set_font('Arial', 'B', 11)
    pdf.cell(0, 10, 'Observación:', 0, 1)
    pdf.set_font('Arial', '', 11)
    pdf.multi_cell(0, 10, st.session_state.observaciones['estado_actual'])
    pdf.ln(5)
    
    # Sección 2: PROYECCIÓN A CIERRE
    pdf.add_page()
    pdf.set_font('Arial', 'B', 12)
    pdf.set_fill_color(156, 39, 176)  # Color morado
    pdf.cell(0, 10, 'PROYECCIÓN A CIERRE', 0, 1, 'L', 1)
    pdf.ln(5)
    
    pdf.set_font('Arial', '', 11)
    # Tabla de Proyecciones
    pdf.cell(100, 10, 'Métrica', 1)
    pdf.cell(55, 10, 'Valor Proyectado', 1)
    pdf.ln()
    
    pdf.cell(100, 10, 'Matrículas Proyectadas', 1)
    pdf.cell(55, 10, str(st.session_state.proyeccion_data['matriculas']), 1)
    pdf.ln()
    
    pdf.cell(100, 10, 'Leads Proyectados', 1)
    pdf.cell(55, 10, str(st.session_state.proyeccion_data['leads']), 1)
    pdf.ln(15)
    
    # Observación
    pdf.set_font('Arial', 'B', 11)
    pdf.cell(0, 10, 'Observación:', 0, 1)
    pdf.set_font('Arial', '', 11)
    pdf.multi_cell(0, 10, st.session_state.observaciones['proyeccion'])
    pdf.ln(5)
    
    # Sección 3: DISTRIBUCIÓN DE RESULTADOS POR PROGRAMA
    pdf.add_page()
    pdf.set_font('Arial', 'B', 12)
    pdf.set_fill_color(255, 193, 7)  # Color amarillo
    pdf.cell(0, 10, 'DISTRIBUCIÓN DE RESULTADOS POR PROGRAMA', 0, 1, 'L', 1)
    pdf.ln(5)
    
    pdf.set_font('Arial', '', 11)
    # Tabla de Programas
    df_programas = pd.DataFrame(st.session_state.programas_data)
    df_top = df_programas.sort_values('matriculas', ascending=False).head(10)
    
    # Encabezados
    pdf.cell(65, 10, 'Programa', 1)
    pdf.cell(30, 10, 'Leads', 1)
    pdf.cell(30, 10, 'Matrículas', 1)
    pdf.cell(30, 10, 'Conversión (%)', 1)
    pdf.ln()
    
    # Datos
    for _, row in df_top.iterrows():
        pdf.cell(65, 10, str(row['programa']), 1)
        pdf.cell(30, 10, str(int(row['leads'])), 1)
        pdf.cell(30, 10, str(int(row['matriculas'])), 1)
        pdf.cell(30, 10, f"{row['conversion']:.1f}%", 1)
        pdf.ln()
    
    pdf.ln(5)
    
    # Observación
    pdf.set_font('Arial', 'B', 11)
    pdf.cell(0, 10, 'Insight Estratégico:', 0, 1)
    pdf.set_font('Arial', '', 11)
    pdf.multi_cell(0, 10, st.session_state.observaciones['programas'])
    
    # Crear PDF como bytes
    buffer = io.BytesIO()
    buffer.write(pdf.output(dest='S').encode('latin1'))
    buffer.seek(0)
    return buffer

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
formato_exportacion = st.sidebar.selectbox("Formato", ["Excel", "PDF"])

if st.sidebar.button("Exportar Reporte"):
    try:
        if formato_exportacion == "Excel":
            buffer = generate_excel()
            # Crear link de descarga
            b64 = base64.b64encode(buffer.read()).decode()
            href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="reporte_{st.session_state.titulo_reporte.replace(" ", "_")}_{datetime.now().strftime("%Y%m%d")}.xlsx">Descargar Excel</a>'
            st.sidebar.markdown(href, unsafe_allow_html=True)
            st.sidebar.success("Excel generado correctamente")
        
        elif formato_exportacion == "PDF":
            buffer = generate_pdf()
            # Crear link de descarga
            b64 = base64.b64encode(buffer.read()).decode()
            href = f'<a href="data:application/pdf;base64,{b64}" download="reporte_{st.session_state.titulo_reporte.replace(" ", "_")}_{datetime.now().strftime("%Y%m%d")}.pdf">Descargar PDF</a>'
            st.sidebar.markdown(href, unsafe_allow_html=True)
            st.sidebar.success("PDF generado correctamente")

    except Exception as e:
        st.sidebar.error(f"Error al generar el reporte: {str(e)}")

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
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("Matrículas proyectadas")
        p_matriculas = st.number_input("Valor esperado", min_value=0, value=st.session_state.proyeccion_data['matriculas'], key="proy_matriculas")
        
        # Actualizar estado
        if p_matriculas != st.session_state.proyeccion_data['matriculas']:
            st.session_state.proyeccion_data['matriculas'] = p_matriculas
        
        # Visualizar con barra de progreso
        objetivo_matriculas = st.session_state.kpi_data['matriculas']['objetivo']
        pct_cumplimiento = min(100, int((p_matriculas / max(1, objetivo_matriculas)) * 100))
        
        st.write(f"Matrículas proyectadas vs Objetivo ({pct_cumplimiento}%)")
        st.progress(pct_cumplimiento/100)
        
        if pct_cumplimiento >= 100:
            st.success(f"Se espera CUMPLIR el objetivo con {p_matriculas} matrículas")
        elif pct_cumplimiento >= 90:
            st.info(f"Se proyecta alcanzar el {pct_cumplimiento}% del objetivo")
        else:
            st.warning(f"Se proyecta alcanzar el {pct_cumplimiento}% del objetivo")
    
    with col2:
        st.subheader("Leads proyectados")
        p_leads = st.number_input("Valor esperado", min_value=0, value=st.session_state.proyeccion_data['leads'], key="proy_leads")
        
        # Actualizar estado
        if p_leads != st.session_state.proyeccion_data['leads']:
            st.session_state.proyeccion_data['leads'] = p_leads
        
        # Visualizar con barra de progreso
        objetivo_leads = st.session_state.kpi_data['leads']['objetivo']
        pct_cumplimiento_leads = min(100, int((p_leads / max(1, objetivo_leads)) * 100))
        
        st.write(f"Leads proyectados vs Objetivo ({pct_cumplimiento_leads}%)")
        st.progress(pct_cumplimiento_leads/100)
        
        if pct_cumplimiento_leads >= 100:
            st.success(f"Se espera CUMPLIR el objetivo con {p_leads} leads")
        elif pct_cumplimiento_leads >= 90:
            st.info(f"Se proyecta alcanzar el {pct_cumplimiento_leads}% del objetivo")
        else:
            st.warning(f"Se proyecta alcanzar el {pct_cumplimiento_leads}% del objetivo")
    
    # Tabla de Proyección vs Objetivo
    st.subheader("Resumen de Proyección")
    
    data = {
        "Métrica": ["Matrículas", "Leads"],
        "Valor Actual": [st.session_state.kpi_data['matriculas']['actual'], st.session_state.kpi_data['leads']['actual']],
        "Valor Proyectado": [p_matriculas, p_leads],
        "Objetivo": [objetivo_matriculas, objetivo_leads],
        "% Proyectado del Objetivo": [f"{pct_cumplimiento}%", f"{pct_cumplimiento_leads}%"]
    }
    
    df_proyeccion = pd.DataFrame(data)
    st.table(df_proyeccion)
    
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
