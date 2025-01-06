import os
import pandas as pd
from shareplum import Office365
from shareplum import Site
from shareplum.site import Version
from time import sleep
from castear import process_file
from descargar_archivos import obtener_rutas_so
import logging
from io import BytesIO
import time
import random
import urllib3

# Desactivar advertencias de SSL
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# Configuración de logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Función para obtener la conexión a SharePoint
def obtener_conexion_sharepoint():
    username = "jean.pier@grow-analytics.com.pe"
    password = "Grow123*$%"  
    url = "https://growanalyticscom.sharepoint.com"
    site_url = "https://growanalyticscom.sharepoint.com/sites/PROYECTOS/"

    try:
        # Autenticación y conexión a SharePoint
        authcookie = Office365(url, username=username, password=password).GetCookies()
        site = Site(site_url, version=Version.v2016, authcookie=authcookie)

        # Configuración del timeout para las solicitudes
        session = site._session  
        session.timeout = (60, 1800)  # Timeout de 30 segundos a 20 minutos
        session.verify = False
        logging.info("Conexión a SharePoint exitosa")
        return site
    except Exception as e:
        logging.error(f"Error al conectar a SharePoint: {e}")
        return None

# Función para obtener el nombre de la tabla y la ruta correcta de SharePoint (según la clave)
def obtener_nombre_tabla_y_ruta(clave):
    if 'STK' in clave:
        if 'Maestro' in clave:
            ruta_sharepoint = obtener_rutas_so()[2]  # Transaccional_STK/Maestro
            nombre_tabla = 'Stock'
        else:
            ruta_sharepoint = obtener_rutas_so()[3]  # Transaccional_STK/Sodimac
            nombre_tabla = 'Stock'
    elif 'SO' in clave:
        if 'Maestro' in clave:
            ruta_sharepoint = obtener_rutas_so()[0]  # Transaccional_SO/Maestro
            nombre_tabla = 'Ventas_SO'
        else:
            ruta_sharepoint = obtener_rutas_so()[1]  # Transaccional_SO/Sodimac
            nombre_tabla = 'Ventas_SO'
    else:
        logging.error(f"Ruta no encontrada para el archivo {clave}")
        return None, None  # Ruta y nombre de tabla no encontrados

    return nombre_tabla, ruta_sharepoint

# Función para manejar la subida con reintentos
def upload_with_retries(folder, file_data, nombre_subida, retries=5):
    for attempt in range(retries):
        try:
            logging.info(f"Intentando subir el archivo: {nombre_subida}, intento {attempt + 1}")
            folder.upload_file(file_data, nombre_subida)
            logging.info(f"Archivo casteado subido a SharePoint: {nombre_subida}")
            return
        except Exception as e:
            logging.warning(f"Error al subir el archivo: {e}. Intento {attempt + 1} de {retries}")
            # Retroceso exponencial con jitter
            sleep_time = min(2 ** attempt + random.random(), 60)  
            time.sleep(sleep_time)
    logging.error(f"Fallo al subir el archivo después de {retries} intentos")

# Función para procesar y subir archivos a SharePoint
def procesar_y_subir_archivos():
    # Directorio base
    script_directory = os.path.dirname(os.path.abspath(__file__))
    descargas_dir = os.path.join(script_directory, "Descargas")

    # Archivos a procesar
    archivos = {
        "SO_Maestro": os.path.join(descargas_dir, "script_Transaccional_de_Ventas_B2B_Maestro_2024.xlsx"),
        "SO_Sodimac": os.path.join(descargas_dir, "script_Transaccional_de_Ventas_B2B_Sodimac_2024.xlsx"),
        "STK_Maestro": os.path.join(descargas_dir, "script_Transaccional_de_STK_B2B_Maestro_2024.xlsx"),
        "STK_Sodimac": os.path.join(descargas_dir, "script_Transaccional_de_STK_B2B_Sodimac_2024.xlsx")
    }

    # Obtener conexión a SharePoint
    site = obtener_conexion_sharepoint()
    if not site:
        return

    # Procesar cada archivo
    for clave, ruta_script in archivos.items():
        nombre_archivo_descargado = os.path.basename(ruta_script).replace("script_", "")
        ruta_descargado = os.path.join(descargas_dir, f"001_{os.path.basename(nombre_archivo_descargado)}")
        
        logging.info(f'Ruta del script: {ruta_script}')
        logging.info(f'Ruta del archivo descargado: {ruta_descargado}')
        
        # Verificar si los archivos existen
        if not os.path.exists(ruta_script) or not os.path.exists(ruta_descargado):
            logging.error(f"Uno o ambos archivos no se encuentran: {ruta_script}, {ruta_descargado}")
            continue

        # Leer los dos archivos
        try:
            df_script = pd.read_excel(ruta_script)
            df_descargado = pd.read_excel(ruta_descargado)
        except Exception as e:
            logging.error(f"Error al leer los archivos Excel: {e}")
            continue
        
        logging.info(f"Datos del archivo script leído: {df_script.head()}")
        logging.info(f"Datos del archivo descargado leído: {df_descargado.head()}")

        # Concatenar los DataFrames
        df_concatenado = pd.concat([df_descargado, df_script], ignore_index=True)
        logging.info(f"Datos concatenados: {df_concatenado.head()}")

        # Casteo y formato en el DataFrame
        logging.info(f"Casteando archivo concatenado: {clave}")
        try:
            process_file(df_concatenado)  # No se pasa `output_path` ya que no guardamos el archivo
            logging.info(f"Archivo casteado correctamente para {clave}")
        except Exception as e:
            logging.error(f"Error al procesar el archivo casteado para {clave}: {e}")
            continue  # Pasar al siguiente archivo en caso de error

        # Obtener nombre de tabla y ruta de SharePoint
        nombre_tabla, ruta_sharepoint = obtener_nombre_tabla_y_ruta(clave)

        # Si no se encontró la ruta correcta, continuar con el siguiente archivo
        if ruta_sharepoint is None:
            continue

        folder = site.Folder(ruta_sharepoint)
        nombre_base = os.path.splitext(nombre_archivo_descargado)[0]
        nombre_subida = "Prueba_001_" + nombre_base + "_prueba.xlsx"
        
        chunk_size = 100000
        chunks = [df_concatenado[i:i + chunk_size] for i in range(0, len(df_concatenado), chunk_size)]
        logging.info(f"Archivo dividido en {len(chunks)} fragmentos para subir.")

        # Procesar y cargar cada fragmento por separado
        for idx, chunk in enumerate(chunks):
            logging.info(f"Procesando fragmento {idx + 1}/{len(chunks)}")
            # Guardar el archivo casteado a un buffer en memoria (sin guardar en disco)
            with BytesIO() as excel_buffer:
                with pd.ExcelWriter(excel_buffer, engine='xlsxwriter') as writer:
                    chunk.to_excel(writer, index=False, sheet_name='Sheet1')
                    workbook = writer.book
                    worksheet = writer.sheets['Sheet1']

                    # Formatos
                    date_format = workbook.add_format({'num_format': 'dd/mm/yyyy', 'align': 'right'})  # Formato de fecha
                    text_format = workbook.add_format({'num_format': '@'})  # Formato de texto

                    # Formato para la columna 'DATE_TRANSACTION' si existe
                    if 'DATE_TRANSACTION' in chunk.columns:
                        date_col_idx = chunk.columns.get_loc('DATE_TRANSACTION')
                        worksheet.set_column(date_col_idx, date_col_idx, 30, date_format)
                    
                    # Formato para las columnas de texto
                    worksheet.set_column('B:C', None, text_format)

                    # Crear tabla estructurada en Excel
                    worksheet.add_table(0, 0, len(chunk), len(chunk.columns) - 1,
                                        {'name': nombre_tabla, 'columns': [{'header': col} for col in chunk.columns]})

                # Finalizar y cargar el contenido en memoria
                excel_buffer.seek(0)
                file_data = excel_buffer.read()

                # Verificar que el archivo tiene contenido
                if not file_data:
                    logging.error(f"El archivo está vacío para {clave}. No se subirá a SharePoint.")
                    continue
                else:
                    upload_with_retries(folder, file_data, nombre_subida)
                    logging.info(f"Archivo casteado subido a SharePoint: {ruta_sharepoint}/{nombre_subida}")

# Ejecutar el proceso
procesar_y_subir_archivos()

