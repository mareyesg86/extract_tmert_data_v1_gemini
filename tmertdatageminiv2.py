import streamlit as st
import pandas as pd
from io import BytesIO
import json
import os
from pathlib import Path
import configparser
from google.cloud import aiplatform
import time  # Para el indicador de carga

# Configuración de Gemini
def cargar_configuracion():
    config = configparser.ConfigParser()
    config_path = Path.home() / 'tmert_config.ini'

    if config_path.exists():
        config.read(config_path)
        if 'gemini' not in config:
            config['gemini'] = {}
        if 'model_id' not in config['gemini'] or config['gemini']['model_id'] not in ['gemini-2.0-flash-lite']:
            config['gemini']['model_id'] = os.getenv('GEMINI_MODEL_ID', 'gemini-2.0-flash-lite')
    else:
        config['gemini'] = {
            'api_key': os.getenv('GEMINI_API_KEY', ''),
            'project_id': os.getenv('GEMINI_PROJECT_ID', ''),
            'location': os.getenv('GEMINI_LOCATION', 'us-central1'),
            'model_id': os.getenv('GEMINI_MODEL_ID', 'gemini-2.0-flash-lite')
        }
        with open(config_path, 'w') as configfile:
            config.write(configfile)

    return config

config = cargar_configuracion()

# Función para eliminar columnas vacías
def eliminar_columnas_vacias(df):
    df = df.loc[:, (df != "").any(axis=0)]
    return df

def extraer_texto_desde_excel(file, hoja):
    try:
        df = pd.read_excel(file, sheet_name=hoja, header=None, dtype=str).fillna("")
        texto = df.astype(str).apply(lambda row: " ".join(row), axis=1).str.cat(sep="\n")
        return texto
    except Exception as e:
        st.error(f"Error al leer la hoja '{hoja}' del archivo Excel: {e}")
        return None

def consultar_gemini_datos_generales(texto):
    project_id = config.get('gemini', 'project_id')
    location = config.get('gemini', 'location')
    model_id = config.get('gemini', 'model_id', fallback='gemini-2.0-flash-lite')
    api_key = config.get('gemini', 'api_key')

    if not api_key or not project_id or not location or not model_id:
        st.error("La configuración de Gemini no está completa.")
        return None

    client_options = {"api_key": api_key}
    try:
        client = aiplatform.gapic.PredictionServiceClient(client_options=client_options)

        prompt_base = '''Actúa como un experto en la identificación y extracción de datos generales de empresas desde texto no estructurado.
        Analiza el siguiente texto, que proviene de una hoja de cálculo de una matriz TMERT, e identifica los siguientes campos clave.
        Devuelve un objeto JSON limpio con los siguientes campos:
        - empresa_razon_social
        - rut_empresa
        - actividad_economica
        - codigo_ciiu
        - direccion_matriz
        - comuna_matriz
        - representante_legal

        Sé flexible en la forma en que estos campos pueden aparecer en el texto. Busca patrones, palabras clave y el contexto para identificarlos correctamente.
        Si un campo no se encuentra, devuelve "" para ese campo.
        '''

        instance = aiplatform.gapic.Schema(
            prediction_format="json",
            instances=[{"content": prompt_base + "\n\nTexto de entrada:\n" + texto}]
        )

        endpoint = f"projects/{project_id}/locations/{location}/models/{model_id}"
        response = client.predict(
            endpoint=endpoint,
            instances=instance.instances,
        )
        return response.predictions[0]
    except Exception as e:
        st.error(f"Error al contactar a Gemini para datos generales: {e}")
        return None

st.title("Procesador de Matrices TMERT")
st.subheader("Carga tu archivo Excel para analizar")

uploaded_file = st.file_uploader("Cargar archivo Excel (.xlsx)", type="xlsx")

if uploaded_file is not None:
    st.success("Archivo cargado exitosamente.")

    hojas_excel = pd.ExcelFile(uploaded_file).sheet_names

    with st.expander("Configuración del Análisis", expanded=True):
        st.info("Selecciona las hojas que contienen los datos de la empresa y las tareas.")

        hoja_general_nombre = st.selectbox("Hoja con datos generales de la empresa", hojas_excel)
        hoja_tareas_nombre = st.selectbox("Hoja con datos de tareas", hojas_excel)

        if st.button("Procesar"):
            output = BytesIO()
            writer = pd.ExcelWriter(output, engine='xlsxwriter')
            procesamiento_exitoso = True

            with st.spinner(f"Extrayendo texto de la hoja '{hoja_general_nombre}'..."):
                texto_general = extraer_texto_desde_excel(uploaded_file, hoja_general_nombre)
                if texto_general:
                    with st.spinner("Analizando datos generales con Gemini..."):
                        json_resultado_general = consultar_gemini_datos_generales(texto_general)
                        if json_resultado_general:
                            try:
                                datos_dict_general = json.loads(json_resultado_general)
                                df_general = pd.DataFrame([datos_dict_general])
                                df_general = eliminar_columnas_vacias(df_general)
                                df_general.to_excel(writer, sheet_name='Antecedentes Empresa', index=False)
                            except json.JSONDecodeError as e:
                                st.error(f"Error al decodificar respuesta de Gemini (datos generales): {e}. Respuesta: {json_resultado_general}")
                                procesamiento_exitoso = False
                            except Exception as e:
                                st.error(f"Error al crear DataFrame de datos generales: {e}")
                                procesamiento_exitoso = False
                        else:
                            st.warning("No se encontraron datos generales de la empresa.")
                else:
                    procesamiento_exitoso = False

            with st.spinner(f"Extrayendo y procesando datos de tareas de la hoja '{hoja_tareas_nombre}'..."):
                try:
                    encabezados_raw = pd.read_excel(uploaded_file, sheet_name=hoja_tareas_nombre, header=None, skiprows=11, nrows=2, dtype=str).fillna("")
                    encabezados = []

                    for col in range(encabezados_raw.shape[1]):
                        fila1 = str(encabezados_raw.iat[0, col]).strip()
                        fila2 = str(encabezados_raw.iat[1, col]).strip()
                        if fila1 and fila2 and fila1 != fila2:
                            encabezados.append(f"{fila1} - {fila2}")
                        else:
                            encabezados.append(fila1 or fila2 or f"Col_{col}")

                    tareas = pd.read_excel(uploaded_file, sheet_name=hoja_tareas_nombre, header=None, skiprows=13, names=encabezados, dtype=str).fillna("")
                    tareas = tareas[tareas.astype(str).apply(lambda x: x.str.strip() != "").any(axis=1)]

                    riesgos = ["TRMS", "POSTURA", "MMC LDT", "MMC EA", "VIBRACIONES CC", "VIBRACIONES SMB"]
                    for r in riesgos:
                        if r not in tareas.columns:
                            tareas[r] = ""

                    tareas = eliminar_columnas_vacias(tareas)
                    tareas.to_excel(writer, sheet_name='Tareas', index=False)
                except Exception as e:
                    st.error(f"Error al procesar la hoja de tareas: {e}")
                    procesamiento_exitoso = False

            if procesamiento_exitoso:
                try:
                    writer.close()
                    output.seek(0)
                    st.download_button(
                        label="Descargar archivo Excel procesado",
                        data=output,
                        file_name="resultados_tmert.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    st.success("Procesamiento completado y archivo listo para descargar.")
                except Exception as e:
                    st.error(f"Error al generar el archivo Excel de descarga: {e}")
