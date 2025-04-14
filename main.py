import streamlit as st
import pandas as pd
from datetime import datetime
import io
import openpyxl

# Configuración de la página
st.set_page_config(page_title="Conversor TXT a Excel", layout="wide")

def generate_excel(df):
    # Crear un buffer en memoria para el archivo Excel
    buffer = io.BytesIO()
    
    # Generar nombre de archivo con fecha y hora actual
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    excel_filename = f"analisis_datos_{timestamp}.xlsx"
    
    # Crear el archivo Excel usando XlsxWriter
with pd.ExcelWriter(buffer, engine='openpyxl') as writer        # Escribir los datos en una hoja
        df.to_excel(writer, sheet_name='Datos', index=False)
        
        # Obtener el objeto workbook y worksheet
        workbook = writer.book
        worksheet = workbook.add_worksheet('Resumen')
        
        # Añadir información del resumen
        worksheet.write('A1', 'Resumen del Análisis')
        worksheet.write('A3', 'Número total de filas:')
        worksheet.write('B3', len(df))
        worksheet.write('A4', 'Número total de columnas:')
        worksheet.write('B4', len(df.columns))
        
        # Añadir lista de columnas
        worksheet.write('A6', 'Lista de Columnas:')
        for idx, col in enumerate(df.columns):
            worksheet.write(idx + 6, 0, col)  # Corregido aquí
    
    return buffer, excel_filename

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
            st.download_button(
                label="📥 Descargar Excel",
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
