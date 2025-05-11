# Reportador Estratégico de Marketing

Aplicación para generar reportes estratégicos por marca y programa educativo, adaptados a diferentes modelos como GRADO (convocatorias fijas) y POSGRADO (cohortes variables y continuas ADVANCE).

## Características

- **Análisis por Marca**: Visualización de métricas clave para GRADO, POSGRADO, ADVANCE, WIZARD, AJA, UNISUD.
- **Estado Actual**: Tiempo transcurrido, leads acumulados, matrículas vs meta, tasa de conversión.
- **Composición de Resultados**: Análisis de matrículas por leads nuevos vs remarketing.
- **Estimación de Cierre**: Proyección de resultados mediante simulación Monte Carlo con intervalos de confianza.
- **Análisis por Programa**: Identificación de programas top, con baja conversión y oportunidades.
- **Exportación**: Generación de reportes en formatos Excel, PDF y PowerPoint.

## Instalación

1. Clonar el repositorio:
   ```
   git clone https://github.com/tu-usuario/digitalreportes.git
   cd digitalreportes
   ```

2. Instalar dependencias:
   ```
   pip install -r requirements.txt
   ```

## Uso

1. Ejecutar la aplicación:
   ```
   streamlit run app.py
   ```

2. Cargar archivos:
   - Seleccionar una marca en el menú desplegable
   - Cargar los archivos de matriculados, leads activos y planificación correspondientes a esa marca
   - Hacer clic en "Generar Reporte"

3. Datos de Ejemplo:
   - Para generar datos de muestra, hacer clic en el botón "Generar Datos de Ejemplo" en la barra lateral
   - Los datos se guardarán en la carpeta 'sample_data'

## Archivos de Entrada

1. **matriculados.xlsx** (Pestaña: matriculados)
   - Columnas: ID lead, Fecha ingreso, Fecha matrícula, Marca, Programa

2. **leads_activos.xlsx** (Pestaña: leads_activos)
   - Columnas: ID lead, Fecha ingreso, Estado actual, Marca, Programa

3. **planificacion.xlsx**
   - **plan_mensual**: Marca, Canal, Presupuesto total mes, CPL estimado, Leads estimados
   - **inversion_acumulada**: Fecha, Marca, Canal, Inversión acumulada, CPL estimado
   - **calendario_convocatoria**: Marca, Programa, Fecha inicio, Fecha fin, Tipo (Convocatoria/Cohorte)

## Estructura del Proyecto

```
digitalreportes/
├── app.py                     # Aplicación principal Streamlit
├── requirements.txt           # Dependencias del proyecto
├── README.md                  # Este archivo
├── utils/                     # Utilidades y módulos
│   ├── __init__.py
│   ├── data_processor.py      # Procesamiento de datos de entrada
│   ├── calculations.py        # Cálculos y análisis estadísticos
│   ├── report_generator.py    # Generación de reportes en diferentes formatos
│   └── data_generator.py      # Generación de datos de ejemplo
└── sample_data/               # Carpeta para datos de ejemplo
```

## Requisitos

- Python 3.8+
- Streamlit
- Pandas
- NumPy
- Matplotlib
- FPDF
- python-pptx
- XlsxWriter 