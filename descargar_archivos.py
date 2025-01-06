import os
import pandas as pd
from datetime import datetime
import pytz
from shareplum import Office365
from shareplum import Site
from shareplum.site import Version
import logging

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


# Función para obtener las rutas actualizadas
def obtener_rutas_so():
    return [
        'Documentos compartidos/01 PERU/001 Softys/Proyecto Thanos Moderno Institucional/05_Transaccional/Transaccional_SO/Maestro',
        'Documentos compartidos/01 PERU/001 Softys/Proyecto Thanos Moderno Institucional/05_Transaccional/Transaccional_SO/Sodimac',
        'Documentos compartidos/01 PERU/001 Softys/Proyecto Thanos Moderno Institucional/05_Transaccional/Transaccional_STK/Maestro',
        'Documentos compartidos/01 PERU/001 Softys/Proyecto Thanos Moderno Institucional/05_Transaccional/Transaccional_STK/Sodimac'
    ]

# Filtrar archivos que terminan en "2024" después del último guion bajo
def filtrar_archivos(site, ruta, patron):
    carpeta = site.Folder(ruta)
    archivos = carpeta.files
    archivos_filtrados = []
    
    for archivo in archivos:
        if isinstance(archivo, dict) and 'Name' in archivo:
            nombre_archivo = archivo['Name']
            if nombre_archivo.split('_')[-1].startswith(patron):
                archivos_filtrados.append(nombre_archivo)
        elif isinstance(archivo, str):
            if archivo.split('_')[-1].startswith(patron):
                archivos_filtrados.append(archivo)
    return archivos_filtrados

# Descargar el archivo
def descargar_archivo(site, ruta, archivo, destino_local):
    # Crear la carpeta si no existe
    carpeta_destino = os.path.dirname(destino_local)
    if not os.path.exists(carpeta_destino):
        os.makedirs(carpeta_destino)  # Crear todas las carpetas necesarias
        print(f"Carpeta creada: {carpeta_destino}")
    
    # Descargar el archivo
    carpeta = site.Folder(ruta)
    with open(destino_local, 'wb') as f:
        f.write(carpeta.get_file(archivo))
    print(f"Archivo descargado: {archivo}")


        #agregado
def obtener_y_descargar_archivos(site, folder_path):
    # Obtener rutas y descargar archivos
    rutas = obtener_rutas_so()
    
    for ruta in rutas:
        archivos = filtrar_archivos(site, ruta, patron="2024")
        if archivos:
            for archivo in archivos:
                destino = os.path.join(folder_path, archivo)
                descargar_archivo(site, ruta, archivo, destino)
        else:
            print(f"No se encontraron archivos en la ruta {ruta}.")

# Función principal para integrar el flujo
def main():
   
    site = obtener_conexion_sharepoint()
    folder_path_tmp = os.path.abspath(__file__)
    folder_path = os.path.join(folder_path_tmp, 'Descargas')
    obtener_y_descargar_archivos(site, folder_path)
    
if __name__ == "__main__":
    main()
