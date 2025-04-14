import streamlit as st

try:
    import openpyxl
    st.success("✅ openpyxl está instalado correctamente")
except ModuleNotFoundError as e:
    st.error(f"❌ Error: {e}")
    
import pandas as pd
from datetime import datetime
import io
import base64


def generate_excel(df):
    # Crear un buffer en memoria para el archivo Excel
    output = io.BytesIO()
    
    # Generar nombre de archivo con fecha y hora actual
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    excel_filename = f"analisis_datos_{timestamp}.xlsx"
    
    # Crear el archivo Excel en memoria
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Escribir los datos en una hoja
        df.to_excel(writer, sheet_name='Datos', index=False)
        
        # Crear una hoja para el resumen
        workbook = writer.book
        summary_sheet = workbook.create_sheet('Resumen')
        
        # Añadir información del resumen
        summary_sheet['A1'] = 'Resumen del Análisis'
        summary_sheet['A3'] = 'Número total de filas:'
        summary_sheet['B3'] = len(df)
        summary_sheet['A4'] = 'Número total de columnas:'
        summary_sheet['B4'] = len(df.columns)
        
        # Añadir lista de columnas
        summary_sheet['A6'] = 'Lista de Columnas:'
        for idx, col in enumerate(df.columns, start=7):
            summary_sheet[f'A{idx}'] = col
    
    return output, excel_filename

def get_download_link(buffer, filename):
    """Genera un link de descarga para el archivo Excel"""
    b64 = base64.b64encode(buffer.getvalue()).decode()
    return f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="{filename}">Descargar archivo Excel</a>'

# Configuración de la página
st.set_page_config(page_title="Conversor TXT a Excel", layout="wide")

# Título de la aplicación
st.title("Conversor de archivo TXT a Excel")

# Área de carga de archivo
uploaded_file = st.file_uploader("Selecciona un archivo TXT", type=['txt'])

if uploaded_file is not None:
    try:
        # Contenedor para los separadores
        with st.expander("Configuración de separador", expanded=True):
            separator = st.selectbox(
                "Selecciona el separador del archivo",
                options=[',', ';', '\t'],
                format_func=lambda x: "Coma (,)" if x == ',' else "Punto y coma (;)" if x == ';' else "Tabulación (\\t)"
            )

        # Botón para procesar el archivo
        if st.button("Procesar archivo"):
            # Leer el archivo
            df = pd.read_csv(uploaded_file, sep=separator)
            
            # Mostrar información del DataFrame
            st.subheader("Vista previa de los datos")
            st.dataframe(df.head())
            
            # Mostrar información de las columnas
            st.subheader("Información de columnas")
            col1, col2 = st.columns(2)
            with col1:
                st.write("Número de filas:", len(df))
                st.write("Número de columnas:", len(df.columns))
            
            with col2:
                st.write("Columnas:")
                for idx, columna in enumerate(df.columns, 1):
                    st.write(f"{idx}. {columna}")
            
            # Generar archivo Excel
            excel_buffer, excel_filename = generate_excel(df)
            
            # Mostrar botón de descarga
            st.markdown("### Descargar archivo Excel")
            st.download_button(
                label="Descargar Excel",
                data=excel_buffer.getvalue(),
                file_name=excel_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
            st.success("¡Archivo procesado exitosamente!")
            
    except Exception as e:
        st.error(f"Error al procesar el archivo: {str(e)}")
else:
    st.info("Por favor, sube un archivo TXT para comenzar.")

# Información adicional
st.sidebar.header("Información")
st.sidebar.write("""
Esta aplicación te permite:
- Subir archivos TXT
- Visualizar los datos
- Convertir a Excel con resumen
- Descargar el archivo Excel resultante
""")
