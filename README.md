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


Este script automatiza el proceso de conexión a **SharePoint**, filtrado de archivos según un patrón específico y descarga de esos archivos a un directorio local. El script está diseñado para trabajar con archivos que contienen datos relevantes para proyectos específicos y organiza esos archivos en una estructura de carpetas local.

---

## Funcionalidades principales

### 1. **Conexión a SharePoint**
   El script establece una conexión con **SharePoint** a través de la librería `shareplum`. Utiliza un nombre de usuario y una contraseña para autenticarse y obtener una cookie de autenticación. Con esta cookie, se accede a un sitio específico de SharePoint para realizar operaciones como la obtención de archivos.

### 2. **Filtrado de Archivos**
   El script filtra los archivos en SharePoint que terminan en "2024" después del último guion bajo (`_`). Esto se logra mediante la función `filtrar_archivos`, que recorre todos los archivos en una carpeta de SharePoint y selecciona aquellos que cumplen con este patrón.

### 3. **Creación de Carpeta Local**
   Antes de descargar los archivos, el script asegura que el directorio de destino local exista. Si no existe, el script lo crea automáticamente. De esta forma, garantiza que los archivos descargados se almacenen en un lugar organizado y accesible.

### 4. **Descarga de Archivos**
   Una vez que los archivos son filtrados, el script los descarga a la carpeta local especificada. Utiliza la función `descargar_archivo` para realizar esta operación, que guarda los archivos desde SharePoint al sistema local.

### 5. **Obtención y Descarga de Archivos desde Múltiples Rutas**
   La función `obtener_y_descargar_archivos` recorre varias rutas predefinidas dentro de SharePoint, filtra los archivos según el patrón de "2024", y los descarga a la carpeta local correspondiente.

### 6. **Estructura del Proyecto**
   El script está organizado de manera que se pueden agregar nuevas rutas o patrones de filtrado de archivos sin modificar las funciones clave. Esto facilita la extensión y personalización del flujo de trabajo.


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



# Flujo de ejecución
- `Conexión a SharePoint`: El script establece una conexión con SharePoint usando las credenciales y configuraciones de conexión definidas.

- `Procesamiento de archivos`: El script procesa los archivos descargados (como archivos de ventas o inventarios), transformándolos según las reglas especificadas.

- `Carga a SharePoint`: Los archivos procesados se cargan a la ruta especificada en SharePoint, asegurando que los datos estén listos para su uso en otros procesos o análisis.

- `Manejo de errores`: En caso de fallos, el script se encarga de registrar los errores y reintentar el proceso hasta un número determinado de veces, lo cual es útil para gestionar problemas de conectividad.


# Parámetros de configuración
- `Usuario y Contraseña de SharePoint`: El script utiliza credenciales codificadas de usuario y contraseña. Se recomienda utilizar un método de autenticación más seguro para entornos de producción, como la autenticación de Azure AD o mediante un servicio de autenticación.

- `Rutas de los archivos en SharePoint`: Las rutas de los archivos que deben ser procesados se configuran directamente en el script. Estas rutas deben ser precisas y tener los permisos adecuados para leer y escribir archivos en las ubicaciones de SharePoint.

- `Archivos a procesar`: Los archivos que se procesan deben estar ubicados en el directorio de descargas o en una carpeta específica que el script debe identificar para procesarlos.

# Consideraciones de seguridad
- `Gestión de credenciales`: Las credenciales de SharePoint están almacenadas de forma directa en el script. Para un entorno de producción, se recomienda usar una solución de almacenamiento seguro de credenciales, como un gestor de secretos o la autenticación basada en tokens.

- `Reintentos`: La estrategia de reintentos implementada con retrocesos exponenciales y jitter asegura que el script pueda manejar conexiones intermitentes sin fallar inmediatamente.

---

# Documentación de la función `process_file(df)`

La función `process_file` procesa un archivo cargado como un `DataFrame` de **pandas**, realizando transformaciones en las columnas del archivo de acuerdo a un formato específico. A continuación se describe el funcionamiento de la función, sus entradas y salidas, así como los detalles del procesamiento de datos.

---

## Descripción general

La función recibe un **DataFrame de pandas** y realiza las siguientes operaciones de procesamiento sobre las columnas del archivo:

1. **Validación de columnas**: Verifica que el archivo tenga al menos tres columnas.
2. **Conversión de formato de fecha**: Convierte la primera columna, que debe contener fechas, al formato `'dd/mm/yyyy'`.
3. **Conversión de columnas numéricas a texto**: Las segunda y tercera columna se convierten a texto, eliminando cualquier punto decimal en los valores flotantes (si existiera alguno).

---

## Entradas

### Parámetro

- `df` (DataFrame de pandas): 
  - Un **DataFrame** que contiene los datos del archivo que se va a procesar.
  - Las primeras tres columnas deben tener los siguientes tipos de datos esperados:
    - **Columna 1 (Fecha)**: Contiene datos de tipo fecha, que deben convertirse al formato `dd/mm/yyyy`.
    - **Columna 2 y 3 (Numéricas)**: Contienen datos numéricos o decimales, los cuales deben ser transformados a texto y sin decimales.

---

## Salidas

- **DataFrame procesado**: La función retorna el **DataFrame** con las modificaciones realizadas en las columnas, siguiendo las reglas especificadas.

---

## Detalles de procesamiento

1. **Verificación de columnas**:
    - La función verifica si el `DataFrame` tiene al menos tres columnas con `df.shape[1]`. Si el archivo tiene menos de tres columnas, se imprime un mensaje de advertencia y se retorna `None`.

2. **Conversión de la primera columna a fecha**:
    - La función convierte la primera columna (`DATE_TRANSACTION`) a un formato de fecha en el formato `'dd/mm/yyyy'`. Utiliza la función `pd.to_datetime` de pandas para convertir la columna, especificando que el día aparece antes del mes (con `dayfirst=True`).
    - Los valores de la columna se formatean como cadenas con el formato deseado utilizando el método `.dt.strftime('%d/%m/%Y')`.

3. **Conversión de las columnas segunda y tercera a texto sin decimales**:
    - Las columnas segunda y tercera se convierten a texto. Si la columna contiene valores flotantes, se eliminan los decimales convirtiendo el valor a entero.
    - Se usa `astype(str)` para convertir los valores numéricos a cadenas de texto.
    - La operación `lambda x: str(int(float(x))) if x else ""` asegura que los valores nulos sean reemplazados por una cadena vacía, y los valores flotantes son redondeados a enteros sin decimales.

---

## Ejemplo de uso

```python
import pandas as pd

# Ejemplo de DataFrame
data = {
    'DATE_TRANSACTION': ['01/02/2024', '15/03/2024'],
    'COLUMN_2': [1234.56, 7890.12],
    'COLUMN_3': [2345.67, 1234.89]
}

df = pd.DataFrame(data)

# Llamar a la función process_file
df_procesado = process_file(df)

print(df_procesado)
