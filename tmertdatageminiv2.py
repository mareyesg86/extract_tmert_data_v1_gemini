import streamlit as st
import pandas as pd
from io import BytesIO
import json
import os
from pathlib import Path
import configparser
import openai
import time  # Para el indicador de carga

# Configuración de OpenAI
def cargar_configuracion_openai():
    config = configparser.ConfigParser()
    config_path = Path.home() / 'tmert_config.ini'

    # Intentar cargar la API Key de OpenAI directamente desde el input de Streamlit
    openai_api_key_input = st.text_input("Ingresa tu clave de API de OpenAI (temporal)", type="password")
    if openai_api_key_input:
        config['openai'] = {
            'api_key': openai_api_key_input,
            'model_id': 'gpt-3.5-turbo'  # Modelo predeterminado
        }
        return config

    # Si no se ingresa la clave de API, intentar cargar desde secrets.toml
    if "openai_api_key" in st.secrets:
        config['openai'] = {
            'api_key': st.secrets["openai_api_key"],
            'model_id': st.secrets.get("openai_model_id", 'gpt-3.5-turbo')
        }
        return config

    # Si no se encuentra en secrets.toml, intentar cargar desde tmert_config.ini
    if config_path.exists():
        config.read(config_path)
        if 'openai' not in config:
            config['openai'] = {}
        config['openai']['api_key'] = config['openai'].get('api_key', os.getenv('OPENAI_API_KEY') or "")
        config['openai']['model_id'] = config['openai'].get('model_id', os.getenv('OPENAI_MODEL_ID', 'gpt-3.5-turbo') or 'gpt-3.5-turbo')
        return config
    else:
        # Si ninguno existe, crea tmert_config.ini con variables de entorno (respaldo)
        config['openai'] = {
            'api_key': os.getenv('OPENAI_API_KEY') or "",
            'model_id': os.getenv('OPENAI_MODEL_ID', 'gpt-3.5-turbo') or 'gpt-3.5-turbo'
        }
        with open(config_path, 'w') as configfile:
            config.write(configfile)
        return config

config_openai = cargar_configuracion_openai()
openai.api_key = config_openai.get('openai', 'api_key')
if config_openai.has_option('openai', 'model_id'):
    openai_model_id = config_openai.get('openai', 'model_id')
else:
    openai_model_id = 'gpt-3.5-turbo'

# Función para eliminar columnas vacías (sin cambios)
def eliminar_columnas_vacias(df):
    df = df.loc[:, (df != "").any(axis=0)]
    return df

# Función para extraer texto desde Excel (sin cambios)
def extraer_texto_desde_excel(file, hoja):
    try:
        df = pd.read_excel(file, sheet_name=hoja, header=None, dtype=str).fillna("")
        texto = df.astype(str).apply(lambda row: " ".join(row), axis=1).str.cat(sep="\n")
        return texto
    except Exception as e:
        st.error(f"Error al leer la hoja '{hoja}' del archivo Excel: {e}")
        return None

def consultar_openai_datos_generales(texto):
    if not openai.api_key:
        st.error("La clave de API de OpenAI no está configurada.")
        return None

    prompt_base = '''Actúa como un experto en la identificación y extracción de datos generales de empresas desde texto no estructurado.
    Analiza el siguiente texto, que proviene de una hoja de cálculo de una matriz TMERT, e identifica los siguientes campos clave.
    Devuelve un objeto JSON limpio con los siguientes campos:
    - empresa_razon_social: Nombre completo de la empresa.
    - rut_empresa: RUT de la empresa (sin puntos ni guion).
    - actividad_economica: Descripción de la actividad económica principal de la empresa.
    - codigo_ciiu: Código CIIU asociado a la actividad económica.
    - direccion_matriz: Dirección de la matriz de la empresa.
    - comuna_matriz: Comuna donde se encuentra la matriz de la empresa.
    - representante_legal: Nombre del representante legal de la empresa.
    - centro_trabajo: Nombre del centro de trabajo (si está disponible).
    - direccion_centro: Dirección del centro de trabajo (si está disponible).
    - comuna_centro: Comuna donde se encuentra el centro de trabajo (si está disponible).
    - trabajadores_hombres: Número de trabajadores hombres en la empresa.
    - trabajadores_mujeres: Número de trabajadoras mujeres en la empresa.
    - responsable_nombre: Nombre del responsable de la matriz TMERT.
    - responsable_cargo: Cargo del responsable de la matriz TMERT.
    - responsable_email: Correo electrónico del responsable de la matriz TMERT.
    - responsable_telefono: Teléfono del responsable de la matriz TMERT.

    Sé flexible en la forma en que estos campos pueden aparecer en el texto. Busca patrones, palabras clave y el contexto para identificarlos correctamente.
    Si un campo no se encuentra, devuelve "" para ese campo.
    '''

    try:
        response = openai.chat.completions.create(
            model=openai_model_id,
            messages=[
                {"role": "system", "content": prompt_base},
                {"role": "user", "content": texto}
            ],
            response_format={"type": "json_object"}
        )
        return response.choices[0].message.content
    except openai.error.InvalidRequestError as e:
        st.error(f"Error al contactar a OpenAI: {e.message}")
        return None
    except Exception as e:
        st.error(f"Error inesperado al contactar a OpenAI: {e}")
        return None

st.title("Procesador de Matrices TMERT")
st.subheader("Carga tu archivo Excel para analizar")

uploaded_file = st.file_uploader("Cargar archivo Excel (.xlsx)", type="xlsx")

if uploaded_file is not None:
    st.success("Archivo cargado exitosamente.")

    hojas_excel = pd.ExcelFile(uploaded_file).sheet_names

    with st.expander("Configuración del Análisis", expanded=True):
        st.info("Selecciona las hojas que contienen los datos de la empresa y las tareas.")

        hoja_general_nombre = st.selectbox("Hoja con datos generales de la empresa", hojas_excel, index=0 if "1" in hojas_excel else 0)
        hoja_tareas_nombre = st.selectbox("Hoja con datos de tareas", hojas_excel, index=1 if len(hojas_excel) > 1 and "2" in hojas_excel else 1)

        if st.button("Procesar"):
            output = BytesIO()
            writer = pd.ExcelWriter(output, engine='xlsxwriter')
            procesamiento_exitoso = True

            with st.spinner(f"Extrayendo texto de la hoja '{hoja_general_nombre}'..."):
                texto_general = extraer_texto_desde_excel(uploaded_file, hoja_general_nombre)
                if texto_general:
                    st.text_area("Texto Hoja Datos Generales (para depuración)", texto_general, height=200) # <- DEBUG

                    with st.spinner("Analizando datos generales con OpenAI..."):
                        json_resultado_general = consultar_openai_datos_generales(texto_general)
                        st.json(json_resultado_general) # <- DEBUG

                        if json_resultado_general:
                            try:
                                datos_dict_general = json.loads(json_resultado_general)
                                df_general = pd.DataFrame([datos_dict_general])
                                df_general = eliminar_columnas_vacias(df_general)
                                df_general.to_excel(writer, sheet_name='Antecedentes Empresa', index=False)
                            except json.JSONDecodeError as e:
                                st.error(f"Error al decodificar respuesta de OpenAI (datos generales): {e}. Respuesta: {json_resultado_general}")
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
