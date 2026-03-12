# app.py
import streamlit as st
import pandas as pd
import os
from procesamiento import procesar_tiempos
from exportacion import generar_excel_con_semaforo

st.set_page_config(page_title="Control RRHH Tortuguitas", page_icon="⏱️", layout="wide")

st.title("⏱️ Panel de Control de Tiempos Muertos - Tortuguitas")
st.markdown("Sube la bajada de **SPEC MANAGER** para analizar tiempos. La base de horarios se cruza automáticamente.")

# 1. VERIFICAMOS SI EXISTE LA BASE DE HORARIOS FIJA
if not os.path.exists("horarios.csv"):
    st.error("⚠️ No se encontró el archivo 'horarios.csv' en la carpeta. Por favor, guarda la base de horarios con ese nombre en la misma carpeta del programa.")
else:
    # Si existe, lo leemos de fondo (AGREGAMOS encoding='latin1' y sep=';' para leer Excel en español)
    try:
        df_horarios = pd.read_csv("horarios.csv", sep=';', encoding='latin1')
        st.success("✅ Base de horarios cargada exitosamente.")
    except Exception as e:
        st.error(f"Hubo un problema leyendo 'horarios.csv'. Detalle: {e}")
        st.stop() 

    # 2. INTERFAZ PARA SUBIR EL ARCHIVO DE SPEC
    st.subheader("📂 Archivo de Registros (TAL Y COMO SE BAJA DE SPEC MANAGER)")
    archivo_registros = st.file_uploader("Arrastra aquí el archivo crudo de entradas y salidas", type=["csv", "xlsx"])

    # 3. PROCESAMIENTO
    if archivo_registros is not None:
        try:
            # Leemos el archivo de SPEC (Atajamos si también viene con formato latino)
            if archivo_registros.name.endswith('.csv'):
                try:
                    df_registros = pd.read_csv(archivo_registros, encoding='utf-8')
                except UnicodeDecodeError:
                    archivo_registros.seek(0)
                    df_registros = pd.read_csv(archivo_registros, sep=';', encoding='latin1')
            else:
                df_registros = pd.read_excel(archivo_registros)
                
            st.info("¡Archivo cargado! Procesando tiempos y cruzando datos...")
            
            with st.spinner('Realizando cálculos matemáticos...'):
                df_resumen, df_detalle = procesar_tiempos(df_registros, df_horarios)
                
            st.subheader("📊 RESUMEN:")
            st.dataframe(df_resumen)
            
            archivo_excel = generar_excel_con_semaforo(df_resumen, df_detalle)
            
            st.download_button(
                label="📥 Descargar Reporte",
                data=archivo_excel,
                file_name="Reporte_Tiempos_Ociosos.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )
            
        except Exception as e:
            st.error(f"Ocurrió un error al procesar el archivo. Detalle del error: {e}")