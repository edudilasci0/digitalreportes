# utils/report_generator.py

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import io
from datetime import datetime
from fpdf import FPDF
from pptx import Presentation
from pptx.util import Inches, Pt
import collections

def generate_excel(metrics, projections, program_analysis, comentarios, marca):
    """Generar informe en formato Excel"""
    buffer = io.BytesIO()
    
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        # 1. Hoja de resumen
        df_resumen = pd.DataFrame({
            'Métrica': [
                'Tiempo Transcurrido (%)',
                'Leads Acumulados',
                'Matrículas Acumuladas',
                'Meta de Matrículas',
                'Tasa de Conversión (%)',
                '% Matrículas Leads Nuevos',
                '% Matrículas Remarketing',
                'Inversión Acumulada',
                'CPL Promedio'
            ],
            'Valor': [
                f"{metrics['tiempo_transcurrido']:.1f}%",
                metrics['leads_acumulados'],
                metrics['matriculas_acumuladas'],
                metrics['meta_matriculas'],
                f"{metrics['tasa_conversion']:.2f}%",
                f"{metrics['pct_matriculas_nuevos']:.1f}%",
                f"{metrics['pct_matriculas_remarketing']:.1f}%",
                f"${metrics['inversion_acumulada']:,.2f}",
                f"${metrics['cpl_promedio']:,.2f}"
            ]
        })
        
        df_resumen.to_excel(writer, sheet_name='Resumen', index=False)
        
        # 2. Hoja de proyecciones
        df_proyecciones = pd.DataFrame({
            'Métrica': [
                'Leads Proyectados',
                'Matrículas Proyectadas (Min)',
                'Matrículas Proyectadas (Max)',
                '% Cumplimiento Proyectado'
            ],
            'Valor': [
                projections['leads_proyectados'],
                projections['matriculas_proyectadas_min'],
                projections['matriculas_proyectadas_max'],
                f"{projections['pct_cumplimiento_proyectado']:.1f}%"
            ]
        })
        
        df_proyecciones.to_excel(writer, sheet_name='Proyecciones', index=False)
        
        # 3. Hoja de distribución de resultados
        # Asegurarnos que program_analysis['tabla_completa'] sea un DataFrame
        if not isinstance(program_analysis['tabla_completa'], pd.DataFrame):
            print("Convirtiendo 'tabla_completa' a DataFrame")
            df_tabla_completa = pd.DataFrame(program_analysis['tabla_completa'])
        else:
            df_tabla_completa = program_analysis['tabla_completa']
            
        df_tabla_completa.to_excel(writer, sheet_name='Distribución Resultados', index=False)
        
        # 4. Hoja de Top 5 programas
        # Asegurarnos que program_analysis['top_matriculas'] sea un DataFrame
        if not isinstance(program_analysis['top_matriculas'], pd.DataFrame):
            print("Convirtiendo 'top_matriculas' a DataFrame")
            df_top_matriculas = pd.DataFrame(program_analysis['top_matriculas'])
        else:
            df_top_matriculas = program_analysis['top_matriculas']
            
        df_top_matriculas.to_excel(writer, sheet_name='Top Programas', index=False)
        
        # 5. Hoja de programas con menor conversión
        # Asegurarnos que program_analysis['menor_conversion'] sea un DataFrame
        if not isinstance(program_analysis['menor_conversion'], pd.DataFrame):
            print("Convirtiendo 'menor_conversion' a DataFrame")
            df_menor_conversion = pd.DataFrame(program_analysis['menor_conversion'])
        else:
            df_menor_conversion = program_analysis['menor_conversion']
            
        df_menor_conversion.to_excel(writer, sheet_name='Menor Conversión', index=False)
        
        # 6. Hoja de comentarios
        df_comentarios = pd.DataFrame({
            'Fecha': [datetime.now().strftime('%Y-%m-%d %H:%M:%S')],
            'Comentarios': [comentarios]
        })
        
        df_comentarios.to_excel(writer, sheet_name='Comentarios', index=False)
        
        # Dar formato a las hojas
        workbook = writer.book
        
        # Formato para títulos
        header_format = workbook.add_format({
            'bold': True,
            'bg_color': '#4472C4',
            'font_color': 'white',
            'border': 1
        })
        
        # Aplicar formato a cada hoja
        for sheet_name in writer.sheets:
            worksheet = writer.sheets[sheet_name]
            worksheet.set_column('A:A', 30)
            worksheet.set_column('B:Z', 15)
            
            # Aplicar formato a los encabezados de forma segura
            try:
                for col_num, value in enumerate(writer.sheets[sheet_name].table.columns):
                    worksheet.write(0, col_num, value, header_format)
            except AttributeError as e:
                print(f"Error al formatear encabezados en {sheet_name}: {e}")
    
    buffer.seek(0)
    return buffer

def generate_pdf(metrics, projections, program_analysis, comentarios, marca):
    """Generar informe en formato PDF"""
    pdf = FPDF()
    pdf.add_page()
    
    # Configuración de estilo
    pdf.set_font('Arial', 'B', 16)
    pdf.cell(190, 10, f"Reporte Estratégico - {marca}", 0, 1, 'C')
    pdf.ln(10)
    
    # 1. Estado Actual
    pdf.set_font('Arial', 'B', 14)
    pdf.cell(190, 10, "Estado Actual", 0, 1, 'L')
    pdf.ln(5)
    
    pdf.set_font('Arial', '', 12)
    pdf.cell(95, 10, f"Tiempo Transcurrido: {metrics['tiempo_transcurrido']:.1f}%", 0, 0, 'L')
    pdf.cell(95, 10, f"Leads Acumulados: {metrics['leads_acumulados']}", 0, 1, 'L')
    pdf.cell(95, 10, f"Matrículas vs Meta: {metrics['matriculas_acumuladas']}/{metrics['meta_matriculas']}", 0, 0, 'L')
    pdf.cell(95, 10, f"Tasa de Conversión: {metrics['tasa_conversion']:.2f}%", 0, 1, 'L')
    pdf.ln(10)
    
    # 2. Composición de Resultados
    pdf.set_font('Arial', 'B', 14)
    pdf.cell(190, 10, "Composición de Resultados", 0, 1, 'L')
    pdf.ln(5)
    
    pdf.set_font('Arial', '', 12)
    pdf.cell(95, 10, f"% Matrículas Leads Nuevos: {metrics['pct_matriculas_nuevos']:.1f}%", 0, 0, 'L')
    pdf.cell(95, 10, f"% Matrículas Remarketing: {metrics['pct_matriculas_remarketing']:.1f}%", 0, 1, 'L')
    pdf.ln(10)
    
    # 3. Estimación de Cierre
    pdf.set_font('Arial', 'B', 14)
    pdf.cell(190, 10, "Estimación de Cierre", 0, 1, 'L')
    pdf.ln(5)
    
    pdf.set_font('Arial', '', 12)
    pdf.cell(95, 10, f"Leads Proyectados: {projections['leads_proyectados']}", 0, 0, 'L')
    pdf.cell(95, 10, f"Matrículas Proyectadas: {projections['matriculas_proyectadas_min']} - {projections['matriculas_proyectadas_max']}", 0, 1, 'L')
    pdf.cell(190, 10, f"% Cumplimiento Proyectado: {projections['pct_cumplimiento_proyectado']:.1f}%", 0, 1, 'L')
    pdf.ln(10)
    
    # 4. Top 5 Programas
    pdf.set_font('Arial', 'B', 14)
    pdf.cell(190, 10, "Top 5 Programas con Más Matrículas", 0, 1, 'L')
    pdf.ln(5)
    
    # Crear tabla de top 5 programas
    pdf.set_font('Arial', 'B', 10)
    pdf.cell(95, 10, "Programa", 1, 0, 'C')
    pdf.cell(30, 10, "Leads", 1, 0, 'C')
    pdf.cell(30, 10, "Matrículas", 1, 0, 'C')
    pdf.cell(35, 10, "Tasa Conv. (%)", 1, 1, 'C')
    
    # Asegurarnos que program_analysis['top_matriculas'] sea un DataFrame
    if not isinstance(program_analysis['top_matriculas'], pd.DataFrame):
        df_top_matriculas = pd.DataFrame(program_analysis['top_matriculas'])
    else:
        df_top_matriculas = program_analysis['top_matriculas']
    
    pdf.set_font('Arial', '', 10)
    for _, row in df_top_matriculas.iterrows():
        pdf.cell(95, 10, str(row['Programa']), 1, 0, 'L')
        pdf.cell(30, 10, str(row['Leads']), 1, 0, 'C')
        pdf.cell(30, 10, str(row['Matrículas']), 1, 0, 'C')
        pdf.cell(35, 10, str(row['Tasa Conversión (%)']), 1, 1, 'C')
    
    pdf.ln(10)
    
    # 5. Comentarios
    pdf.set_font('Arial', 'B', 14)
    pdf.cell(190, 10, "Comentarios", 0, 1, 'L')
    pdf.ln(5)
    
    pdf.set_font('Arial', '', 12)
    pdf.multi_cell(190, 10, comentarios)
    
    # Convertir a buffer
    buffer = io.BytesIO()
    buffer.write(pdf.output(dest='S').encode('latin1'))
    buffer.seek(0)
    
    return buffer

def generate_pptx(metrics, projections, program_analysis, comentarios, marca):
    """Generar presentación en formato PowerPoint"""
    prs = Presentation()
    
    # 1. Portada
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    title = slide.shapes.title
    title.text = f"Reporte Estratégico - {marca}"
    subtitle = slide.placeholders[1]
    subtitle.text = f"Fecha: {datetime.now().strftime('%Y-%m-%d')}"
    
    # 2. Estado Actual
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    title = slide.shapes.title
    title.text = "Estado Actual"
    
    content = slide.placeholders[1]
    tf = content.text_frame
    tf.text = f"Tiempo Transcurrido: {metrics['tiempo_transcurrido']:.1f}%\n"
    tf.text += f"Leads Acumulados: {metrics['leads_acumulados']}\n"
    tf.text += f"Matrículas vs Meta: {metrics['matriculas_acumuladas']}/{metrics['meta_matriculas']}\n"
    tf.text += f"Tasa de Conversión: {metrics['tasa_conversion']:.2f}%"
    
    # 3. Composición de Resultados
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    title = slide.shapes.title
    title.text = "Composición de Resultados"
    
    content = slide.placeholders[1]
    tf = content.text_frame
    tf.text = f"% Matrículas Leads Nuevos: {metrics['pct_matriculas_nuevos']:.1f}%\n"
    tf.text += f"% Matrículas Remarketing: {metrics['pct_matriculas_remarketing']:.1f}%"
    
    # 4. Estimación de Cierre
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    title = slide.shapes.title
    title.text = "Estimación de Cierre"
    
    content = slide.placeholders[1]
    tf = content.text_frame
    tf.text = f"Leads Proyectados: {projections['leads_proyectados']}\n"
    tf.text += f"Matrículas Proyectadas: {projections['matriculas_proyectadas_min']} - {projections['matriculas_proyectadas_max']}\n"
    tf.text += f"% Cumplimiento Proyectado: {projections['pct_cumplimiento_proyectado']:.1f}%"
    
    # 5. Top 5 Programas
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    title = slide.shapes.title
    title.text = "Top 5 Programas con Más Matrículas"
    
    # Asegurarnos que program_analysis['top_matriculas'] sea un DataFrame
    if not isinstance(program_analysis['top_matriculas'], pd.DataFrame):
        df_top_matriculas = pd.DataFrame(program_analysis['top_matriculas'])
    else:
        df_top_matriculas = program_analysis['top_matriculas']
    
    # Crear tabla
    if not df_top_matriculas.empty:
        rows = len(df_top_matriculas) + 1  # +1 para el encabezado
        cols = 4
        
        left = Inches(1)
        top = Inches(2)
        width = Inches(8)
        height = Inches(0.5 * rows)
        
        table = slide.shapes.add_table(rows, cols, left, top, width, height).table
        
        # Encabezados
        table.cell(0, 0).text = "Programa"
        table.cell(0, 1).text = "Leads"
        table.cell(0, 2).text = "Matrículas"
        table.cell(0, 3).text = "Tasa Conv. (%)"
        
        # Datos
        for i, (_, row) in enumerate(df_top_matriculas.iterrows(), 1):
            table.cell(i, 0).text = str(row['Programa'])
            table.cell(i, 1).text = str(row['Leads'])
            table.cell(i, 2).text = str(row['Matrículas'])
            table.cell(i, 3).text = str(row['Tasa Conversión (%)'])
    else:
        content = slide.placeholders[1]
        tf = content.text_frame
        tf.text = "No hay datos disponibles para mostrar"
    
    # 6. Comentarios
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    title = slide.shapes.title
    title.text = "Comentarios"
    
    content = slide.placeholders[1]
    tf = content.text_frame
    tf.text = comentarios
    
    # Guardar en buffer
    buffer = io.BytesIO()
    prs.save(buffer)
    buffer.seek(0)
    
    return buffer 