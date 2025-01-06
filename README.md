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
```



# Documentación del Proyecto: `Limpieza de Sodimac y Maestro`

## Descripción
Este proyecto está diseñado para automatizar el proceso de descarga, organización y limpieza de archivos provenientes de las plataformas **Sodimac** y **Maestro**. Utiliza Selenium para la automatización web, pandas para manipulación de datos, y diversas técnicas de manejo de archivos en Python.

## Funcionalidades

### 1. **Descarga de Archivos**
La función `descargarArchivos` automatiza la descarga de archivos desde la página de **Sodimac**. El script accede al portal, introduce las credenciales, selecciona las tiendas y descarga los archivos correspondientes. Existen dos categorías: **Unidades** y **Soles**, y para cada una de ellas, los archivos se guardan en carpetas separadas.

### 2. **Creación de Carpetas**
El proyecto utiliza varias funciones para crear carpetas basadas en la fecha actual y en las categorías de los archivos descargados:
- `crearCarpetaFecha`: Crea una carpeta con la fecha actual y una subcarpeta con la hora para realizar copias de seguridad de los archivos.
- `crearCarpetasUnidadesSoles`: Crea carpetas para diferentes categorías como **unidades-maestro**, **unidades-sodimac**, **soles_maestro**, etc.

### 3. **Mover Archivos**
Una vez descargados los archivos, se mueven a sus respectivas carpetas utilizando la función `moverArchivos`. Esta función garantiza que no haya archivos duplicados mediante el uso de un contador.

### 4. **Limpieza y Concatenación de Archivos**
Las funciones `concatenarArchivos`, `limpiezaArchivosStock` y `limpiezaArchivosVentas` permiten limpiar y combinar los archivos descargados para obtener una estructura de datos consolidada.

- `concatenarArchivos`: Concatena los archivos descargados en una sola hoja de Excel.
- `limpiezaArchivosStock`: Filtra y organiza los datos relacionados con el stock de productos.
- `limpiezaArchivosVentas`: Filtra y organiza los datos relacionados con las ventas de productos.

### 5. **Generación de Datos**
Las funciones `stock` y `sell` generan dos conjuntos de datos:
- **Stock**: Calcula el stock de productos combinando los datos de unidades y soles.
- **Ventas**: Genera un reporte de ventas a partir de los datos de soles.

Ambas funciones aseguran que los datos se formateen correctamente para su posterior uso en otros procesos.

### 6. **Fecha**
La función `obtener_fecha_un_dia_antes` calcula la fecha del día anterior en un formato específico. Este dato es útil para la creación de registros históricos.

## Librerías Utilizadas

- **Selenium**: Automatización de navegadores para la descarga de archivos.
- **Pandas**: Manipulación y análisis de datos.
- **Tkinter**: Interfaz gráfica para la confirmación de acciones.
- **Openpyxl**: Manipulación de archivos Excel.
- **Shutil**: Manejo de archivos y directorios.

## Instrucciones de Ejecución

### 1. Instalación de Dependencias
Para ejecutar este script, es necesario instalar las siguientes librerías:

```bash
pip install selenium pandas openpyxl xlrd requests tkinter webdriver-manager
```

# Documentación del Código

Este código está diseñado para descargar, procesar y manipular archivos de datos relacionados con las unidades y ventas de las tiendas de Sodimac y Maestro. Utiliza Selenium para la automatización del navegador, Pandas para la manipulación de datos y varias librerías de Python para el manejo de archivos.

## Librerías Importadas

### Selenium y WebDriver:
- `selenium`: Para la automatización de navegación web.
- `selenium.webdriver`: Permite la interacción con los elementos de la página.
- `selenium.webdriver.support`: Proporciona condiciones de espera y otros métodos útiles.
- `selenium.webdriver.chrome.service` y `selenium.webdriver.edge.service`: Para iniciar el servicio del navegador Edge o Chrome.

### Manejo de Archivos y Datos:
- `pandas`: Librería principal para manipulación de datos.
- `os`, `shutil`, `glob`: Para interactuar con el sistema de archivos (crear carpetas, mover archivos, etc.).
- `xlrd`, `openpyxl`: Para leer y escribir archivos de Excel.
- `xlsxwriter`: Para generar archivos de Excel.
- `requests`: Para realizar solicitudes HTTP (no se usa activamente en el código).
- `tkinter`: Interfaz gráfica para mostrar mensajes de confirmación.

### Funciones Principales

#### `descargarArchivos`

Esta función descarga archivos desde el portal de Sodimac y Maestro, y mueve los archivos descargados a las carpetas correspondientes.

- **Parámetros**:
  - `downloads_path`: Ruta de la carpeta donde se almacenan los archivos descargados.
  - `uni_maestro`, `uni_sodimac`, `soles_maestro`, `soles_sodimac`, `clientes_sodimac`, `clientes_maestro`: Rutas de las carpetas donde se moverán los archivos descargados.
  
- **Flujo**:
  - Abre una página web de Sodimac y Maestro.
  - Inicia sesión con credenciales predefinidas.
  - Selecciona diferentes códigos de tienda y descarga archivos en formato `.xls`.
  - Los archivos descargados se mueven a las carpetas correspondientes según el tipo de archivo (Unidades o Soles, Maestro o Sodimac).

#### `crearCarpetasUnidadesSoles`

Crea una nueva carpeta si no existe.

- **Parámetros**:
  - `folder_path`: Ruta base donde se creará la nueva carpeta.
  - `nombre`: Nombre de la carpeta a crear.
  
- **Retorna**: La ruta completa de la nueva carpeta creada.

#### `fechaActual`

Obtiene la fecha actual en formato `dd-mm-yyyy`.

- **Retorna**: La fecha actual en formato `dd-mm-yyyy`.

#### `crearCarpeta`

Crea una carpeta llamada `Descargas` en el directorio donde se encuentra el script, si no existe.

- **Retorna**: Ruta de la carpeta creada o ya existente.

#### `extraerCarpetaDescargas`

Obtiene la ruta de la carpeta de descargas del sistema operativo.

- **Retorna**: Ruta de la carpeta de descargas.

#### `crearCarpetaFecha`

Crea una carpeta con la fecha actual y la hora actual como nombre.

- **Retorna**: La ruta completa de la carpeta creada.

#### `moverArchivos`

Mueve un archivo desde una ubicación original a una carpeta destino. Si el archivo ya existe en la carpeta destino, se renombra para evitar sobrescribirlo.

- **Parámetros**:
  - `original`: Ruta del archivo original.
  - `destino`: Ruta de la carpeta destino donde se moverá el archivo.

#### `concatenarArchivos`

Concatena los archivos de Excel ubicados en una carpeta dada y los guarda en un nuevo archivo combinado.

- **Parámetros**:
  - `uni_maestro`: Ruta de la carpeta con los archivos de unidades Maestro.
  - `maestra`: DataFrame de referencia para agregar columnas de códigos y clientes.
  
#### `poner_codigo_sodimac_maestro`

Agrega los códigos de clientes de Sodimac a un DataFrame que contiene información de unidades.

- **Parámetros**:
  - `df`: DataFrame que contiene los datos de unidades.
  - `maestra`: DataFrame con los códigos de clientes de referencia.
  
- **Retorna**: El DataFrame con los códigos de clientes agregados.

#### `limpiezaArchivosStock`

Limpia y organiza los archivos de stock de unidades y soles, seleccionando solo las columnas relevantes y concatenándolos en un único DataFrame.

- **Parámetros**:
  - `unidades`: Ruta de la carpeta de unidades.
  - `soles`: Ruta de la carpeta de soles.
  
- **Retorna**: DataFrame concatenado de las unidades y soles con columnas relevantes.

#### `limpiezaArchivosVentas`

Limpia y organiza los archivos de ventas, seleccionando solo las columnas relevantes.

- **Parámetros**:
  - `soles`: Ruta de la carpeta de soles.
  
- **Retorna**: DataFrame de ventas con las columnas relevantes.

#### `obtener_fecha_un_dia_antes`

Obtiene la fecha de ayer en formato `dd/mm/yyyy`.

- **Retorna**: La fecha de ayer en formato `dd/mm/yyyy`.

#### `stock`

Genera un DataFrame con información de stock, combinando los datos de unidades y soles y agregando columnas adicionales.

- **Parámetros**:
  - `soles_cadena`: DataFrame con los datos de soles.
  - `unidades_cadena`: DataFrame con los datos de unidades.
  - `codigo`: Código de cliente.

- **Retorna**: DataFrame con la información de stock.

#### `sell`

Genera un DataFrame con información de ventas, agregando columnas adicionales como el código de cliente y la fecha de la transacción.

- **Parámetros**:
  - `soles_cadena`: DataFrame con los datos de soles.
  - `codigo`: Código de cliente.

- **Retorna**: DataFrame con la información de ventas.

---

## Notas
- El código hace uso de Selenium para interactuar con páginas web de manera automatizada, especialmente para la descarga de archivos desde un portal de Sodimac y Maestro.
- La función principal, `descargarArchivos`, automatiza el proceso de selección de códigos de tienda y descarga de archivos, moviéndolos a las carpetas de destino correspondientes.
- Varias funciones están orientadas a la creación de carpetas, la manipulación de archivos Excel y la limpieza de datos, utilizando principalmente Pandas para el manejo de los datos.

---

Este código es útil para procesos de automatización de descargas y procesamiento de datos en un entorno de ventas y gestión de stock.

