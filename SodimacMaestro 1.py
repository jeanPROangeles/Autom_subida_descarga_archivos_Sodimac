# -*- coding: utf-8 -*-
"""
Creado el 15/05
Terminado el 2

@author: Joshep

Limpieza de sodimac y maestro

@Version: 4.0
"""
#Librerias
from selenium import webdriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
import time
from selenium.webdriver.edge.options import Options
from selenium.webdriver.support.ui import Select
from selenium import webdriver
from selenium import webdriver
from selenium.webdriver.edge.service import Service as EdgeService
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager

import pandas as pd
from datetime import datetime
import os
import requests
import tkinter as tk
from tkinter import messagebox
from selenium.common.exceptions import NoSuchElementException
import shutil
import os
import glob
import xlsxwriter
import xlrd
import html5lib
from pandas.errors import EmptyDataError 
from datetime import datetime
from datetime import datetime, timedelta
from descargar_archivos import obtener_conexion_sharepoint, obtener_rutas_so, filtrar_archivos, descargar_archivo, obtener_y_descargar_archivos
from subir_archivos import procesar_y_subir_archivos
from eliminar_archivo import eliminar_carpeta
import openpyxl as px
    

def descargarArchivos(downloads_path,uni_maestro,uni_sodimac,soles_maestro,soles_sodimac,clientes_sodimac,clientes_maestro):
    
    #Abrir pagina en incognito
    options = webdriver.EdgeOptions()
    options.add_argument("-inprivate")
    options.add_argument("--start-maximized")
    #Ejecutar driver de edge
    driver = webdriver.Edge(service=EdgeService(EdgeChromiumDriverManager().install()))
    #Abrir pagina de ketal
    driver.get("https://b2b.sodimac.com/b2bsoclpr/grafica/html/main_home.html")
    #Credenciales
    B2B = 'Sodimac Perú' # 7
    Empresa_key= '20266352337'
    Usuario_key = '77350925'
    Contrasenia = 'softys.2052'
    # AQUÍ VAN LOS CÓDIGOS DE LAS TIENDAS DE LA PÁGINA WEB
    listaCodigosMaestro = ['95','97','98','99','100','101','102','103','104',
                        '109','110','111','114','115','116','118','119',
                        '108','106','117','120','121','123','87','112']
    
    listaCodigosSodimac = ['29','32','30','41','36','40','37','39','16','43','44','56','60',
                        '62','61','70','22','42','107','113','175','25','28','65','64',
                        '54','67','166','46','68','124','339','105']
    
    Espera = WebDriverWait(driver, 10)

    empresa = driver.find_element(By.XPATH,"//*[@id='empresa']")
    empresa.send_keys(Empresa_key)
    usuario = driver.find_element(By.XPATH,"//*[@id='usuario']")
    usuario.send_keys(Usuario_key)
    clave = driver.find_element(By.XPATH,"//*[@id='clave']")
    clave.send_keys(Contrasenia)
    seleccionar = driver.find_element(By.XPATH,"//*[@id='CADENA']")
    select = Select(seleccionar)
    select.select_by_value("7")
    btnEntrar = driver.find_element(By.XPATH,"//*[@id='entrar2']")
    btnEntrar.click()
    
    # btnCerrar_anuncio = driver.find_element(By.XPATH,"//input[@value='Ok']")
    # btnCerrar_anuncio.click()
    # time.sleep(5)
    time.sleep(15)
    
    seleccionar = driver.find_element(By.XPATH,"//*[@id='awmobjeto5']/select")
    select = Select(seleccionar)
    select.select_by_value('1')
    
    check_box =driver.find_element(By.XPATH,"/html/body/table/tbody/tr[3]/td/table/tbody/tr[2]/td/table[4]/tbody/tr/td[2]/input")
    
    
    if not check_box.is_selected():
        check_box.click()
    
    root = tk.Tk()
    root.withdraw()
    
    respuesta = messagebox.askyesno("Ventana de confirmación", "Escoga la fecha, presiona 'Si' cuando ya este listo :)")
    root.destroy()
    
    # Verificar la respuesta del usuario
    
    if respuesta:
        for j in range(2):
            nombre_archivo_descargado = 'datos.xls'
            if j == 0:
                listaCodigos = listaCodigosMaestro
                carpeta = 'Maestro'
            else:
                listaCodigos = listaCodigosSodimac
                carpeta = 'Sodimac'
            
            #Maestro
            for i in range(2):
                for numero in listaCodigos:
                    seleccionar = driver.find_element(By.XPATH,"//*[@id='awmobjeto3']/select")
                    select = Select(seleccionar)
                    select.select_by_value(numero)
                    btnConsultar = driver.find_element(By.XPATH,"/html/body/table/tbody/tr[3]/td/table/tbody/tr[2]/td/table[3]/tbody/tr[3]/td/input")
                    btnConsultar.click()
                    try:
                        download = driver.find_element(By.XPATH,'/html/body/table/tbody/tr[3]/td/table/tbody/tr[2]/td/div/a[2]')
                        download.click()
                        time.sleep(1.8)
                    except NoSuchElementException:
                        print("El enlace no existe en la página")
                        time.sleep(1.5)
                        break
                                            
                    while True:
                        time.sleep(1.5)
                        try:
                            if i ==0: #Unidades Maestro y Sodimac
                                if j ==0: #Unidades Maestro
                                   moverArchivos(downloads_path+'\\'+nombre_archivo_descargado,uni_maestro)
                                   break
                                else: #Unidades Sodimac
                                   moverArchivos(downloads_path+'\\'+nombre_archivo_descargado,uni_sodimac)
                                   break
                            else: #Soles Maestro y Sodimac
                                if j ==0:#Soles Maestro
                                   moverArchivos(downloads_path+'\\'+nombre_archivo_descargado,soles_maestro)
                                   break
                                else:#Soles Sodimac
                                    moverArchivos(downloads_path+'\\'+nombre_archivo_descargado,soles_sodimac)
                                    break                                   
                        except FileNotFoundError:
                            time.sleep(2)
                    
                time.sleep(2)
                seleccionar = driver.find_element(By.XPATH,"//*[@id='awmobjeto6']/select")
                select = Select(seleccionar)
                select.select_by_value('S')
                time.sleep(2)
                
                    
                print("Listo :)")
            
            seleccionar = driver.find_element(By.XPATH,"//*[@id='awmobjeto6']/select")
            select = Select(seleccionar)
            select.select_by_value('UN')
            time.sleep(1)

    else:
        print("No seas habil porfa :)")

    
    time.sleep(3)


def crearCarpetasUnidadesSoles(folder_path,nombre):
    ruta_nueva_carpeta = folder_path+'//'+nombre
    if not os.path.exists(ruta_nueva_carpeta):
        # Crear la nueva carpeta
        os.mkdir(ruta_nueva_carpeta)
        print("La carpeta ha sido creada.")
    else:
        print("La carpeta ya existe.")
    
    return ruta_nueva_carpeta
    
def fechaActual():
    # Obtener la fecha y hora actuales
    now = datetime.now()
    # Obtener el año, mes y día actuales
    year = now.year
    month = now.month
    day = now.day
    
    if(month<10):
        month = '0'+str(month)
    if(day<10):
        day = '0'+ str(day)
        
    #Fecha
    fecha = str(day)+'-'+str(month)+'-'+str(year)
    
    return fecha

def crearCarpeta():
    
    folder_path_tmp = os.path.dirname(os.path.abspath(__file__))
    folder_path = os.path.join(folder_path_tmp, 'Descargas')
    if not os.path.exists(folder_path):
     os.makedirs(folder_path)
     print("Carpeta creada en: ", folder_path)
    else:
        # Si la carpeta ya existe, entonces muestra un mensaje en la consola
        print("La carpeta ya existe en: ", folder_path)
    
    return folder_path

def extraerCarpetaDescargas():
    if os.name == 'nt':
        # Si el sistema operativo es Windows
        downloads_path = os.path.join(os.environ['USERPROFILE'], 'Downloads')
    else:
        # Si el sistema operativo es Unix
        downloads_path = os.path.join(os.environ['HOME'], 'Downloads') 
    
    return downloads_path

def crearCarpetaFecha(folder_path):
    fecha = fechaActual()
    nombre_carpeta = fecha
    
    ruta_nueva_carpeta = os.path.join(folder_path, nombre_carpeta)

    
    hora_actual = datetime.now().strftime('%H.%M.%S')
    nombre_carpeta_backup = nombre_carpeta +'_'+hora_actual
    nombre_carpeta_completo = os.path.join(ruta_nueva_carpeta, nombre_carpeta_backup)
    os.makedirs(nombre_carpeta_completo, exist_ok=True)
    
    return nombre_carpeta_completo

def creacionUniMaestro(path_completo):
    unidades_maestro = 'unidades-maestro'
    return crearCarpetasUnidadesSoles(path_completo,unidades_maestro)

def creacionUniSodimac(path_completo):
    unidades_sodimac = 'unidades-sodimac'
    return crearCarpetasUnidadesSoles(path_completo,unidades_sodimac)
    
def creacionSolesMaestro(path_completo):
    soles_maestro = 'soles_maestro'
    return crearCarpetasUnidadesSoles(path_completo,soles_maestro)
    
def creacionSolesSodimac(path_completo):
    soles_sodimac = 'soles_sodimac'
    return crearCarpetasUnidadesSoles(path_completo,soles_sodimac)

def moverArchivos(original,destino):
    # Ruta completa del archivo a mover
    ruta_archivo = original

    # Ruta completa de la carpeta destino
    ruta_destino = destino
    
    
    nombre_archivo, extension = os.path.splitext(os.path.basename(ruta_archivo))
    
    ruta_archivo_destino = os.path.join(ruta_destino, nombre_archivo + extension)

    contador = 1
    
    while os.path.exists(ruta_archivo_destino):
        # Generar un nuevo nombre de archivo con un número al costado
        nuevo_nombre_archivo = f"{nombre_archivo}_{contador}{extension}"
        ruta_archivo_destino = os.path.join(ruta_destino, nuevo_nombre_archivo)
        contador += 1

    # Mover el archivo a la carpeta destino
    
    shutil.move(ruta_archivo, ruta_archivo_destino)

def concatenarArchivos(uni_maestro,maestra):
    folder_path = uni_maestro

    # Obtener la lista de archivos en la carpeta
    file_list = os.listdir(folder_path)
    
    # Lista para almacenar los DataFrames de cada archivo
    dataframes = []
    
    # Leer cada archivo y convertirlo en un DataFrame
    for file_name in file_list:
        # Comprobar si el archivo es de tipo Excel
        if file_name.endswith('.xls'):
            file_path = os.path.join(folder_path, file_name)
            try:
                df  = pd.read_table(file_path,encoding='latin-1')
                df1 = poner_codigo_sodimac_maestro(df,maestra)
                dataframes.append(df1)
            except EmptyDataError:
                print("El archivo "+ file_path+ " esta vacio, revise porfa")
                
    # Concatenar los DataFrames verticalmente
    combined_data = pd.concat(dataframes, axis=0, ignore_index=True)
    
    # Guardar el DataFrame combinado en un nuevo archivo
    combined_data.to_excel(uni_maestro+'//combinado.xlsx', index=False)

def poner_codigo_sodimac_maestro(df,maestra):
    filtro = df['Nro1.'] == 'Tienda'
    df_filtrado = df[filtro]
    print(df_filtrado)
    df_codigo = pd.merge(df_filtrado, maestra , left_on='SKU', right_on='CLIENT_SO')
    
    cliente_so = df_codigo.at[0,'CLIENT_SO']
    cod_cliente_so = df_codigo.at[0,'COD_CLIENT_SO']
    
    df['CLIENT_SO'] = cliente_so
    df['COD_CLIENT_SO'] = cod_cliente_so
    
    df2 = df
    
    return df2
    
def limpiezaArchivosStock(unidades,soles):
    archivo_unidades =unidades +'//' +'combinado.xlsx'
    archivo_soles = soles +'//' +'combinado.xlsx'
    
    archivo_unidades_limp = archivo_unidades[['Nro1.','SKU','Stock Contable/Fisico','CLIENT_SO','COD_CLIENT_SO']]
    archivo_soles_limp = archivo_soles[['Nro1.','SKU','Stock Contable/Fisico','CLIENT_SO','COD_CLIENT_SO']]
    
    concatenado = pd.concat([archivo_unidades_limp, archivo_soles_limp], axis=0)
    
    
    return concatenado

def limpiezaArchivosVentas(soles):
    archivo_soles =soles +'//' +'combinado.xlsx'
    archivo_soles_limp = archivo_soles[['Nro1.','SKU','Unidades Vendidas','Monto Venta','CLIENT_SO','COD_CLIENT_SO']]
    
    return archivo_soles_limp



def obtener_fecha_un_dia_antes():

    fecha_actual = datetime.now()
    fecha_anterior = fecha_actual - timedelta(days=1)
    day = fecha_anterior.strftime("%d")
    monthn =fecha_anterior.strftime("%m") #01, 02, ... , 11, 12
    year = fecha_anterior.strftime("%Y")
    print(day+" "+monthn+" "+year)
    print(fecha_actual)
    print(fecha_anterior)
    # fecha_formateada = pd.to_datetime(fecha_anterior.strftime("%d/%m/%Y"))
    fecha_formateada= day+"/"+monthn+"/"+year
    return fecha_formateada

def stock(soles_cadena,unidades_cadena,codigo):
    
    nuevas_stock_sodimac = unidades_cadena[['COD_CLIENT_SO','Nro1.','CLIENT_SO','Stock Contable/Fisico']]
    auxiliar = soles_cadena[['COD_CLIENT_SO','Nro1.','Stock Contable/Fisico']]

    nuevas_stock_sodimac["PK"] = nuevas_stock_sodimac["COD_CLIENT_SO"].map(str) + "" + nuevas_stock_sodimac["Nro1."]
    auxiliar["PK"] = auxiliar["COD_CLIENT_SO"].map(str) + "" + auxiliar["Nro1."]
    
    nuevas_stock_sodimac['TOTAL'] = nuevas_stock_sodimac['PK'].map(auxiliar.set_index('PK')['Stock Contable/Fisico'])
    
    nuevas_stock_sodimac['DATE_TRANSACTION'] = obtener_fecha_un_dia_antes()
    nuevas_stock_sodimac['COD_CLIENT_SI'] = str(codigo)
    nuevas_stock_sodimac['COD_UOM'] = 'UND'
    nuevas_stock_sodimac['UOM'] = 'UNIDAD'
    nuevas_stock_sodimac['UNIT_PRICE'] = ''
    nuevas_stock_sodimac.rename(columns={'Nro1.': 'COD_MATERIAL_SO'}, inplace=True)
    nuevas_stock_sodimac.rename(columns={'COD_CLIENT_SO': 'COD_WAREHOUSE'}, inplace=True)
    nuevas_stock_sodimac.rename(columns={'CLIENT_SO': 'WAREHOUSE'}, inplace=True)
    nuevas_stock_sodimac.rename(columns={'Stock Contable/Fisico': 'QUANTITY'}, inplace=True)
    
    
    nuevo_orden = ['DATE_TRANSACTION', 'COD_CLIENT_SI', 'COD_WAREHOUSE', 'WAREHOUSE','COD_MATERIAL_SO','COD_UOM','UOM','UNIT_PRICE','QUANTITY','TOTAL']
    nuevas_stock_sodimac = nuevas_stock_sodimac.reindex(columns=nuevo_orden)
    
    return nuevas_stock_sodimac
    
    
def sell(soles_cadena,codigo):
    nueva_ventas_maestro = soles_cadena[['COD_CLIENT_SO','Nro1.','Unidades Vendidas','Monto Venta']]
    nueva_ventas_maestro['DATE_TRANSACTION'] = obtener_fecha_un_dia_antes()
    nueva_ventas_maestro['COD_CLIENT_SI'] = str(codigo)
    nueva_ventas_maestro['COD_SELLER_SO'] = ''
    nueva_ventas_maestro['COD_UOM'] = 'UND'
    nueva_ventas_maestro['UOM'] = 'UNIDAD'
    nueva_ventas_maestro['UNIT_PRICE'] = ''
    nueva_ventas_maestro['TOTAL COSTO'] = ''
    nueva_ventas_maestro.rename(columns={'Nro1.': 'COD_MATERIAL_SO'}, inplace=True)
    nueva_ventas_maestro.rename(columns={'Unidades Vendidas': 'QUANTITY'}, inplace=True)
    nueva_ventas_maestro.rename(columns={'Monto Venta': 'TOTAL'}, inplace=True)
    
    nuevo_orden = ['DATE_TRANSACTION', 'COD_CLIENT_SI', 'COD_CLIENT_SO', 'COD_SELLER_SO','COD_MATERIAL_SO','COD_UOM','UOM','UNIT_PRICE','QUANTITY','TOTAL','TOTAL COSTO']
    nueva_ventas_maestro = nueva_ventas_maestro.reindex(columns=nuevo_orden)
    return nueva_ventas_maestro


def format_tbl(writer, sheet_name, df):
    outcols = df.columns
    if len(outcols) > 25:
        raise ValueError('table width out of range for current logic')
    tbl_hdr = [{'header':c} for c in outcols]
    bottom_num = len(df)+1
    right_letter = chr(65-1+len(outcols))
    tbl_corner = right_letter + str(bottom_num)

    worksheet = writer.sheets[sheet_name]
    worksheet.add_table('A1:' + tbl_corner,  {'columns':tbl_hdr})


def convertir(export,ruta_excel,hoja):
    fn_out = export
    df = pd.read_excel(ruta_excel)
    with pd.ExcelWriter(fn_out, mode='w', engine='xlsxwriter') as writer: 
        sheet_name=hoja
        df.to_excel(writer, sheet_name=sheet_name, index=False)
        format_tbl(writer, sheet_name, df)
        
def validarArchivos(archivo_excel,columnas):
    if os.path.isfile(archivo_excel):
        print(f"El archivo '{archivo_excel}' existe.")
        return 1
    else:
        df = pd.DataFrame(columnas)
        writer = pd.ExcelWriter(archivo_excel, engine='xlsxwriter')
        df.to_excel(writer, sheet_name='Hoja1', index=False)
        writer._save()
        print(f"El archivo '{archivo_excel}' ha sido creado")
        return 0

def crearArchivosBase(filepath):
    tran_ventas_maestro ={'DATE_TRANSACTION': [1],
        'COD_CLIENT_SI':[1], 
        'COD_CLIENT_SO': [1],
        'COD_SELLER_SO': [1],
        'COD_UOM':[1],
        'UOM': [1],
        'UNIT_PRICE':[1], 
        'QUANTITY':[1], 
        'TOTAL': [1],
        'TOTAL COSTO':[1] 
        }
    tran_stock_maestro = {'DATE_TRANSACTION': [1],
        'COD_CLIENT_SI':[1], 
        'COD_WAREHOUSE': [1],
        'WAREHOUSE': [1],
        'COD_MATERIAL_SO':[1],
        'COD_UOM': [1],
        'UOM':[1], 
        'UNIT_PRICE':[1], 
        'QUANTITY': [1],
        'TOTAL':[1] 
        }
    tran_ventas_sodimac = {'DATE_TRANSACTION': [1],
        'COD_CLIENT_SI':[1], 
        'COD_CLIENT_SO': [1],
        'COD_SELLER_SO': [1],
        'COD_UOM':[1],
        'UOM': [1],
        'UNIT_PRICE':[1], 
        'QUANTITY':[1], 
        'TOTAL': [1],
        'TOTAL COSTO':[1] 
        }
    tran_stock_sodimac = {'DATE_TRANSACTION': [1],
        'COD_CLIENT_SI':[1], 
        'COD_WAREHOUSE': [1],
        'WAREHOUSE': [1],
        'COD_MATERIAL_SO':[1],
        'COD_UOM': [1],
        'UOM':[1], 
        'UNIT_PRICE':[1], 
        'QUANTITY': [1],
        'TOTAL':[1] 
        }
    
    
    archivo_excel = '001_Transaccional_de_STK_B2B_Maestro_2023.xlsx'
    archivo_excel_1 = '001_Transaccional_de_STK_B2B_Sodimac_2023.xlsx'
    archivo_excel_2 = '001_Transaccional_de_Ventas_B2B_Maestro_2023.xlsx'
    archivo_excel_3 = '001_Transaccional_de_Ventas_B2B_Sodimac_2023.xlsx'
    
    validarArchivos(filepath+'//'+archivo_excel,tran_stock_maestro)
    validarArchivos(filepath+'//'+archivo_excel_1,tran_stock_sodimac)
    validarArchivos(filepath+'//'+archivo_excel_2,tran_ventas_maestro)
    validarArchivos(filepath+'//'+archivo_excel_3,tran_ventas_sodimac)
        
def crearMaestras(filepath):
    # Crear el DataFrame con los datos
    data = {
        'COD_CLIENT_SO': [35, 36, 37, 38, 39, 40, 41, 42, 43, 44, 46, 48, 49, 54, 58, 59, 62, 64, 65, 47, 45, 63, 66, 67, 68, 69, 34, 56],
        'CLIENT_SO': ['Maestro Chacarilla', 'Maestro Surquillo', 'Maestro Pueblo Libre', 'Maestro Chorrillos', 'Maestro Ate', 'Maestro Arequipa', 'Maestro Naranjal', 'Maestro Colonial', 'Maestro Callao', 'Maestro Independencia', 'Maestro Chiclayo', 'Maestro Huancayo', 'MAESTRO CUZCO', 'Maestro Ica', 'Maestro San Luis', 'Maestro Tacna', 'Maestro Barrios Alto', 'Maestro Cajamarca', 'Maestro Sullana', 'Maestro Trujillo', 'Maestro Piura', 'Maestro Comas', 'Maestro Chincha', 'Maestro Puente Piedra', 'Maestro San Juan de Miraflores', 'Maestro Huacho', 'Maestro Amazonia', 'Maestro Villa el Salvador']
    }
    
    data_2 = {
    'COD_CLIENT_SO': [10, 11, 12, 14, 15, 16, 17, 18, 2, 20, 21, 24, 25, 26, 27, 30, 4, 5, 57, 575, 6, 8, 28, 29, 23, 31, 76, 146, 22, 33, 320, 46, 75, 44],
    'CLIENT_SO': ['Sodimac Lima Centro', 'Sodimac Trujillo', 'Sodimac Chiclayo', 'Sodimac Jockey Plaza', 'Sodimac Canta Callao', 'Sodimac Ica Mall', 'Sodimac Trujillo Open P', 'Sodimac Bellavista', 'Sodimac San Miguel', 'Sodimac Piura', 'Sodimac Arequipa', 'Sodimac Cajamarca', 'Sodimac Ate', 'Sodimac S. J. Lurigancho', 'Sodimac Huacho', 'Sodimac Villa El Salvador', 'Sodimac Cono Norte', 'Sodimac Primavera', 'Sodimac Cerro Colorado', 'VENTA A DISTANCIA SODIMAC', 'Sodimac Atocongo', 'Sodimac La Victoria', 'Sodimac Pucallpa', 'Sodimac Canete', 'Sodimac Chimbote', 'Sodimac Sullana', 'Sodimac Huancayo', 'Sodimac -CXI', 'Sodimac Puruchuco', 'Sodimac Comas','SODIMAC IQUITOS','Sodimac Chiclayo III','Sodimac Ventanilla','Sodimac Plaza Norte']
    }

    df = pd.DataFrame(data)
    
    df2 = pd.DataFrame(data_2)
    
    # Exportar el DataFrame a un archivo de Excel
    if os.path.isfile(filepath+'//Maestra Maestro.xlsx'):
        print("Ya existe la maestra")
    else:
        df.to_excel(filepath + '//Maestra Maestro.xlsx', index=False) 
        print("Maestra creada")
        
    if os.path.isfile(filepath+'//Maestra Sodimac.xlsx'):
        print("Ya existe la maestra")
    else:
        df2.to_excel(filepath + '//Maestra Sodimac.xlsx', index=False) 
        print("Maestra creada")

def eliminar_archivo_excel(archivo_excel):
    if os.path.exists(archivo_excel):
        os.remove(archivo_excel)
        print(f"El archivo '{archivo_excel}' ha sido eliminado.")
    else:
        print(f"El archivo '{archivo_excel}' no existe.")




        
def main():

    #Crea Carpeta Principal
    folder_path = crearCarpeta()
    #Creamos transaccionales
    crearArchivosBase(folder_path)
    #Creamos maestras
    crearMaestras(folder_path)
    
    clientes_sodimac = pd.read_excel(folder_path+'//'+'Maestra Sodimac.xlsx')
    clientes_maestro = pd.read_excel(folder_path+'//'+'Maestra Maestro.xlsx')
    
    #Extrae ruta de descargas
    downloads_path = extraerCarpetaDescargas()
    
    #Creamos subcarpeta global fecha y hora
    path_completo = crearCarpetaFecha(folder_path)
    
    #Creacion de sub subcarpeta unidades y soles
    uni_maestro = creacionUniMaestro(path_completo)
    uni_sodimac = creacionUniSodimac(path_completo)
    soles_maestro = creacionSolesMaestro(path_completo)
    soles_sodimac = creacionSolesSodimac(path_completo)
    
    
    
    #Descargamos archivos
    descargarArchivos(downloads_path,uni_maestro,uni_sodimac,soles_maestro,soles_sodimac,clientes_sodimac,clientes_maestro)
    
    
    #Concatenar y Limpiar archivos
    concatenarArchivos(uni_maestro,clientes_maestro)
    concatenarArchivos(soles_maestro,clientes_maestro) 
    concatenarArchivos(uni_sodimac,clientes_sodimac)
    concatenarArchivos(soles_sodimac,clientes_sodimac)
    
    
    combinado = '//combinado.xlsx'
    combinado_soles_sodimac = pd.read_excel(soles_sodimac+combinado)
    combinado_unidades_sodimac = pd.read_excel(uni_sodimac+combinado)
    combinado_soles_maestro = pd.read_excel(soles_maestro+combinado)
    combinado_unidades_maestro = pd.read_excel(uni_maestro+combinado)
    
    
    
    
    soles_sodimac_limpio = combinado_soles_sodimac.dropna(subset=['Código Proveedor'])
    unidades_sodimac_limpio = combinado_unidades_sodimac.dropna(subset=['Código Proveedor'])
    soles_maestro_limpio = combinado_soles_maestro.dropna(subset=['Código Proveedor'])
    unidades_maestro_limpio = combinado_unidades_maestro.dropna(subset=['Código Proveedor'])

 
    
    #ventas_maestro = '//001_Transaccional_de_Ventas_B2B_Maestro_2023.xlsx'
    #stock_maestro = '//001_Transaccional_de_STK_B2B_Maestro_2023.xlsx'
    #ventas_sodimac ='//001_Transaccional_de_Ventas_B2B_Sodimac_2023.xlsx' 
    #stock_sodimac = '//001_Transaccional_de_STK_B2B_Sodimac_2023.xlsx'

    #transaccional_ventas_maestro = pd.read_excel(folder_path+ventas_maestro)
    #transaccional_stock_maestro = pd.read_excel(folder_path+stock_maestro)
    #transaccional_ventas_sodimac = pd.read_excel(folder_path+ventas_sodimac)
    #transaccional_stock_sodimac = pd.read_excel(folder_path+stock_sodimac)

    
    sodimac_export_stock = stock(soles_sodimac_limpio,unidades_sodimac_limpio,17814)
    maestro_export_stock = stock(soles_maestro_limpio,unidades_maestro_limpio,150814)
    sodimac_export_ventas = sell(soles_sodimac_limpio,17814)
    maestro_export_ventas = sell(soles_maestro_limpio,150814)
    
    sodimac_export_stock_concatenado = sodimac_export_stock
    maestro_export_stock_concatenado  = maestro_export_stock
    sodimac_export_ventas_concatenado = sodimac_export_ventas
    maestro_export_ventas_concatenado = maestro_export_ventas

    #sodimac_export_stock_concatenado = pd.concat([sodimac_export_stock,transaccional_stock_sodimac], axis=0)
    #maestro_export_stock_concatenado  = pd.concat([maestro_export_stock,transaccional_stock_maestro], axis=0)
    #sodimac_export_ventas_concatenado = pd.concat([sodimac_export_ventas,transaccional_ventas_sodimac], axis=0)
    #maestro_export_ventas_concatenado = pd.concat([maestro_export_ventas,transaccional_ventas_maestro], axis=0)



    #sodimac_export_stock_concatenado['FECHA_FORMATEADA'] = pd.to_datetime(sodimac_export_stock_concatenado['DATE_TRANSACTION']).dt.strftime('%d/%m/%Y')
    #maestro_export_stock_concatenado['FECHA_FORMATEADA'] = pd.to_datetime(maestro_export_stock_concatenado['DATE_TRANSACTION']).dt.strftime('%d/%m/%Y')
    #sodimac_export_ventas_concatenado['FECHA_FORMATEADA'] = pd.to_datetime(sodimac_export_ventas_concatenado['DATE_TRANSACTION']).dt.strftime('%d/%m/%Y')
    #maestro_export_ventas_concatenado['FECHA_FORMATEADA'] = pd.to_datetime(maestro_export_ventas_concatenado['DATE_TRANSACTION']).dt.strftime('%d/%m/%Y')



    
    sodimac_export_stock_concatenado['QUANTITY'] = sodimac_export_stock_concatenado['QUANTITY'].apply(lambda x: float(str(x).replace(',', '')))
    sodimac_export_stock_concatenado['TOTAL'] = sodimac_export_stock_concatenado['TOTAL'].apply(lambda x: float(str(x).replace(',', '')))
    
    maestro_export_stock_concatenado['QUANTITY'] = maestro_export_stock_concatenado['QUANTITY'].apply(lambda x: float(str(x).replace(',', '')))
    maestro_export_stock_concatenado['TOTAL'] = maestro_export_stock_concatenado['TOTAL'].apply(lambda x: float(str(x).replace(',', '')))
    
    sodimac_export_ventas_concatenado['QUANTITY'] = sodimac_export_ventas_concatenado['QUANTITY'].apply(lambda x: float(str(x).replace(',', '')))
    sodimac_export_ventas_concatenado['TOTAL'] = sodimac_export_ventas_concatenado['TOTAL'].apply(lambda x: float(str(x).replace(',', '')))
    
    maestro_export_ventas_concatenado['QUANTITY'] = maestro_export_ventas_concatenado['QUANTITY'].apply(lambda x: float(str(x).replace(',', '')))
    maestro_export_ventas_concatenado['TOTAL'] = maestro_export_ventas_concatenado['TOTAL'].apply(lambda x: float(str(x).replace(',', '')))
  
    
    sodimac_export_stock_concatenado.to_excel(folder_path+'//stock_sodimac.xlsx',index=False)
    maestro_export_stock_concatenado.to_excel(folder_path+'//stock_maestro.xlsx',index=False)
    sodimac_export_ventas_concatenado.to_excel(folder_path+'//ventas_sodimac.xlsx',index=False)
    maestro_export_ventas_concatenado.to_excel(folder_path+'//ventas_maestro.xlsx',index=False)
    
    hoja = 'Hoja1'
    
    convertir(folder_path+'//script_Transaccional_de_STK_B2B_Sodimac_2024.xlsx',folder_path+'//stock_sodimac.xlsx',hoja)
    convertir(folder_path+'//script_Transaccional_de_STK_B2B_Maestro_2024.xlsx',folder_path+'//stock_maestro.xlsx',hoja)
    convertir(folder_path+'//script_Transaccional_de_Ventas_B2B_Sodimac_2024.xlsx',folder_path+'//ventas_sodimac.xlsx',hoja)
    convertir(folder_path+'//script_Transaccional_de_Ventas_B2B_Maestro_2024.xlsx',folder_path+'//ventas_maestro.xlsx',hoja)
    
    
    eliminar_archivo_excel(folder_path+'//stock_sodimac.xlsx')
    eliminar_archivo_excel(folder_path+'//stock_maestro.xlsx')
    eliminar_archivo_excel(folder_path+'//ventas_sodimac.xlsx')
    eliminar_archivo_excel(folder_path+'//ventas_maestro.xlsx')
    
    
    print("Gracias ya termino :)")

#  descargar_archivos_sharepoint
    site = obtener_conexion_sharepoint()
    obtener_y_descargar_archivos(site, folder_path)
    procesar_y_subir_archivos()
    eliminar_carpeta('./Descargas')
if __name__ == "__main__":
    main()
    

 