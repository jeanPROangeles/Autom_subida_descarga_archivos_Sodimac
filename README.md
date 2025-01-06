# Automización de Carga de Archivos

Este proceso automatiza el login en el sistema de **Sodimac**, la descarga de todas sus bases de datos, así como la concatenación de los archivos descargados desde **SharePoint** y su posterior carga automática.

## 1. Login Automático en Sodimac
Se realiza un inicio de sesión automático en el sistema de **Sodimac**, permitiendo la descarga de todas las bases de datos de manera eficiente.

## 2. Descarga de Bases de Datos
El sistema descarga automáticamente todas las bases de datos disponibles desde **Sodimac**, asegurando que todos los archivos requeridos se obtengan correctamente.

## 3. Concatenación de Archivos
Se realiza la concatenación de los archivos descargados desde **Sodimac** con los archivos obtenidos desde **SharePoint**, fusionándolos de manera automatizada.

## 4. Subida Automática
Una vez concatenados los archivos, se ejecuta el proceso de subida automática a la plataforma correspondiente, garantizando que todos los datos estén actualizados y listos para su uso.

---

# Documentación de `descargar_archivos.py`

Este script automatiza la descarga de archivos desde **SharePoint** basándose en rutas específicas, filtrando los archivos por un patrón determinado (en este caso, archivos que terminan en "2024" después del último guion bajo).

## Dependencias

Este script requiere las siguientes bibliotecas:

- `os`: Para la manipulación de rutas y directorios locales.
- `pandas`: Para manejar y procesar datos (aunque no se utiliza en el código proporcionado, se incluye en caso de futuras modificaciones).
- `datetime`: Para trabajar con fechas y horas.
- `pytz`: Para manejar zonas horarias.
- `shareplum`: Para interactuar con la API de SharePoint.

Para instalar las dependencias necesarias, puedes usar el siguiente comando:

```bash
pip install pandas pytz shareplum
