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
```
---

# Documentación de `subir_archivos.py`

Este script se encarga de procesar archivos Excel, realizar transformaciones y casteo en los datos, y luego cargar los resultados procesados en SharePoint. A continuación, se detalla cómo funciona el script, sus componentes principales y su flujo de ejecución.

---

## Dependencias

Este script requiere las siguientes bibliotecas para su correcto funcionamiento:

- `os`: Para manejar rutas de archivos y directorios.
- `pandas`: Para leer, manipular y escribir archivos Excel.
- `shareplum`: Para interactuar con SharePoint y cargar archivos.
- `time`: Para gestionar tiempos de espera (reintentos, retrocesos).
- `random`: Para agregar un componente aleatorio a los tiempos de espera.
- `urllib3`: Para gestionar advertencias de SSL.
- `logging`: Para registrar eventos e información sobre la ejecución.
- `io`: Para crear buffers en memoria para archivos Excel.
- `castear`: Para procesar y transformar los archivos antes de cargarlos en SharePoint.
- `descargar_archivos`: Para obtener las rutas correctas de los archivos de SharePoint.

### Instalación de dependencias

Para instalar las dependencias necesarias, utiliza `pip`:

```bash
pip install pandas shareplum urllib3 castear
```
## Función principal

El script realiza el siguiente flujo de procesamiento de archivos:

- `Conexión a SharePoint`: El script se conecta a SharePoint utilizando las credenciales configuradas en el código. Utiliza la librería shareplum para autenticar y establecer una conexión al sitio de SharePoint.

- `Obtención de archivos desde SharePoint`: Los archivos que coincidan con el patrón de búsqueda (por ejemplo, que terminen en "2024") se descargan de la carpeta específica en SharePoint utilizando las rutas configuradas.

- `Procesamiento de archivos Excel`: Una vez descargados los archivos, el script lee los archivos Excel y realiza una concatenación de los datos de múltiples archivos en un solo DataFrame. Esto permite combinar datos de diferentes fuentes o periodos.

- `Casteo y transformación de datos`: Los datos concatenados se procesan para asegurar que tengan el formato correcto. Esto puede incluir la conversión de tipos de datos, limpieza de valores nulos o ajustes específicos a las necesidades del proceso de carga.

- `Subida a SharePoint`: Después de realizar el procesamiento, los datos transformados se suben de nuevo a SharePoint. La subida se hace en fragmentos si el archivo es muy grande, utilizando la función de reintentos para manejar posibles fallos de red.

- `Manejo de errores y reintentos`: Si se produce un error durante el proceso de carga, el script intenta cargar el archivo nuevamente en función de un sistema de reintentos. Esto ayuda a manejar interrupciones temporales en la conexión o problemas con SharePoint.
