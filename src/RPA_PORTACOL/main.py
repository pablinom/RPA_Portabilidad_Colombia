from datetime import datetime
from datetime import timedelta
import logging
import json
from time import sleep
from Procesamiento.ProcesamientoWEB import *
from Procesamiento.ProcesamientoBBDD import *
import os
import os.path as path
import shutil
import subprocess
import pyautogui
import psutil  # Para comprobar procesos activos

def cerrar_chrome():
    try:
        subprocess.call(["taskkill", "/F", "/IM", "chrome.exe"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        logging.info("Chrome cerrado correctamente.")
    except Exception as e:
        logging.warning(f"Error cerrando Chrome: {str(e)}")

def chrome_ya_esta_abierto():
    for proc in psutil.process_iter(['name']):
        if proc.info['name'] == 'chrome.exe':
            return True
    return False

def abrir_chrome():
    ruta_a_chrome = "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"
    subprocess.Popen([ruta_a_chrome])
    # ruta_a_chrome = "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"
    # if not chrome_ya_esta_abierto():
    #     subprocess.Popen([ruta_a_chrome])
    #     sleep(5)  # Darle tiempo a abrirse
    #     logging.info("Chrome abierto.")
    # else:
    #     logging.info("Chrome ya estaba abierto.")

def last_day_of_month(date): 
    if date.month == 12: 
        return date.replace(day=31) 
    return date.replace(month=date.month+1, day=1) - timedelta(days=1)

def iniciar_selenium():
    """Inicializa y devuelve una sesión de Chrome."""
    #driver = uc.Chrome(use_subprocess=True)
    driver = Driver(uc=True)
    return driver

def ArchivoExistente(data,rutaAbsoluta):

    try: 
        logging.info("Inicia LecturaArchivo")
        os.path.isfile(rutaAbsoluta) 
        if (path.exists(rutaAbsoluta)): 
            #leer archivo
            ExisteArchivo = True
        else:       
            ExisteArchivo = False
        return ExisteArchivo
    except Exception as e:
        ExisteArchivo = False
        logging.error("Error buscar el archivo : " + str(e))
        return ExisteArchivo
'''
def main():
    while True:
        try:
            # Abrir archivo json de parámetros 
            with open('C:\\RPA_PORTACOL_INTERACTIVO\\src\\config.json', encoding='utf-8') as json_file:
                data = json.load(json_file)
            ruta_carpeta_logs = data["ruta_carpeta_logs"]
            if not os.path.exists(ruta_carpeta_logs):
                os.makedirs(ruta_carpeta_logs)
                print(f"Carpeta creada: {ruta_carpeta_logs}")
            ruta_carpeta_RPA_logs = data["ruta_carpeta_RPA_logs"]
            if not os.path.exists(ruta_carpeta_RPA_logs):
                os.makedirs(ruta_carpeta_RPA_logs)
                print(f"Carpeta creada: {ruta_carpeta_RPA_logs}")    

            current_date = datetime.today().strftime('%d-%m-%Y')
            logging.basicConfig(filename=data["ruta_carpeta_RPA_logs"]+current_date+'.txt', filemode='a', format='%(asctime)s %(levelname)-8s %(message)s',level=logging.INFO,datefmt='%Y-%m-%d %H:%M:%S')

            logging.info("Inicia ejecución RPA PORTA COL")
            ruta_origen = data["ruta_archivo_modelo"]
            ruta_destino = data["ruta_Archivo"]
            # Obtener la lista de archivos en la carpeta origen
            archivos_origen = os.listdir(ruta_origen)

            # Verificar si cada archivo existe en la ruta destino, si no, copiarlo
            for archivo in archivos_origen:
                ruta_archivo_origen = os.path.join(ruta_origen, archivo)
                ruta_archivo_destino = os.path.join(ruta_destino, archivo)

                if not os.path.exists(ruta_archivo_destino):  # Si el archivo no está en la ruta destino
                    shutil.copy2(ruta_archivo_origen, ruta_archivo_destino)  # Copia con metadatos
                    print(f"📂 Copiado: {archivo} → {ruta_destino}")
                else:
                    print(f"✅ Ya existe: {archivo}")

            rutaAbsoluta = os.path.join(ruta_destino, "RPA_PORTACOL.csv")
            rutaCSV = os.path.join(ruta_destino, "dfpendientesCSV.csv")

            if ArchivoExistente(data, rutaAbsoluta):
                if not ArchivoExistente(data, rutaCSV):
                    df = ProcesarExcel(data, rutaAbsoluta)
                    registros_pendientes, ruta_salida = obtener_registros_pendientes(df)
                    df_filtrado = ProcesarCSVFiltrado(data, ruta_salida)
                    
                    if registros_pendientes.empty:
                        print("No hay registros pendientes.")
                        sleep(60)
                        continue

                    abrir_chrome()
                    procesar_registros(data, df_filtrado, ruta_salida)  # usando pyautogui

                else:
                    df_filtrado = ProcesarCSVFiltrado(data, rutaCSV)
                    if df_filtrado.empty:
                        print("No hay registros pendientes.")
                        sleep(60)
                        continue

                    abrir_chrome()
                    procesar_registros(data, df_filtrado, rutaCSV)

                logging.info("Finaliza ejecución RPA PORTA COL")
                break

            else:
                print(f"No hay archivo llamado: {rutaAbsoluta}")
                sleep(60)

        except Exception as e:
            logging.error(f"Error en la ejecución: {str(e)}")
            cerrar_chrome()
            sleep(60)

        finally:
            # Asegúrate de que Chrome no quede abierto si algo sale mal
            cerrar_chrome()

if __name__ == '__main__':
    main()
'''
def main():
    while True:
        try:
            # Abrir archivo json de parámetros 
            with open('C:\\RPA_PORTACOL_INTERACTIVO\\src\\config.json', encoding='utf-8') as json_file:
                data = json.load(json_file)
            
            ruta_carpeta_logs = data["ruta_carpeta_logs"]
            if not os.path.exists(ruta_carpeta_logs):
                os.makedirs(ruta_carpeta_logs)
                print(f"Carpeta creada: {ruta_carpeta_logs}")
            
            ruta_carpeta_RPA_logs = data["ruta_carpeta_RPA_logs"]
            if not os.path.exists(ruta_carpeta_RPA_logs):
                os.makedirs(ruta_carpeta_RPA_logs)
                print(f"Carpeta creada: {ruta_carpeta_RPA_logs}")
                
            # Configurar el logging solo si no está ya configurado
            if not logging.getLogger().handlers:
                current_date = datetime.today().strftime('%d-%m-%Y')
                logging.basicConfig(
                    filename=data["ruta_carpeta_RPA_logs"]+current_date+'.txt', 
                    filemode='a', 
                    format='%(asctime)s %(levelname)-8s %(message)s',
                    level=logging.INFO,
                    datefmt='%Y-%m-%d %H:%M:%S'
                )
            
            logging.info("Inicia ejecución RPA PORTA COL")
            
            ruta_origen = data["ruta_archivo_modelo"]
            ruta_destino = data["ruta_Archivo"]
            
            # Obtener la lista de archivos en la carpeta origen
            archivos_origen = os.listdir(ruta_origen)
            
            # Verificar si cada archivo existe en la ruta destino, si no, copiarlo
            for archivo in archivos_origen:
                ruta_archivo_origen = os.path.join(ruta_origen, archivo)
                ruta_archivo_destino = os.path.join(ruta_destino, archivo)
                
                if not os.path.exists(ruta_archivo_destino):  # Si el archivo no está en la ruta destino
                    shutil.copy2(ruta_archivo_origen, ruta_archivo_destino)  # Copia con metadatos
                    print(f"📂 Copiado: {archivo} → {ruta_destino}")
                else:
                    print(f"✅ Ya existe: {archivo}")
            
            rutaAbsoluta = os.path.join(ruta_destino, "RPA_PORTACOL.csv")
            rutaCSV = os.path.join(ruta_destino, "dfpendientesCSV.csv")
            
            if ArchivoExistente(data, rutaAbsoluta):
                # Primero preparamos los datos
                if not ArchivoExistente(data, rutaCSV):
                    df = ProcesarExcel(data, rutaAbsoluta)
                    registros_pendientes, ruta_salida = obtener_registros_pendientes(df)
                    df_filtrado = ProcesarCSVFiltrado(data, ruta_salida)
                    
                    if registros_pendientes.empty:
                        print("No hay registros pendientes.")
                        logging.info("No hay registros pendientes. Esperando 60 segundos...")
                        sleep(60)
                        continue  # Vuelve al inicio del loop principal 
                    
                else:
                    df_filtrado = LeerCSVFiltrado(rutaCSV)
                    ruta_salida = rutaCSV
                    
                    if df_filtrado.empty:
                        print("No hay registros pendientes.")
                        logging.info("No hay registros pendientes. Esperando 60 segundos...")
                        sleep(60)
                        continue  # Vuelve al inicio del loop principal
                
                # Loop infinito para reintentar el procesamiento
                while True:
                    try:
                        print("Iniciando proceso de registros...")
                        logging.info("Iniciando procesamiento de registros")
                        
                        # Asegurarnos de que Chrome esté cerrado antes de abrirlo
                        cerrar_chrome()
                        sleep(2)
                        
                        # Abrimos Chrome y procesamos los registros
                        abrir_chrome()
                        
                        # Llamamos a procesar_registros con manejo de excepciones
                        procesar_exito = False
                        try:
                            procesar_registros(data, df_filtrado, ruta_salida)
                            procesar_exito = True
                        except Exception as e:
                            logging.error(f"Error en procesar_registros: {str(e)}")
                            print(f"⚠️ Error en procesar_registros: {str(e)}")
                            sleep(1)  # Pequeña pausa antes de cerrar Chrome
                            
                        # Si el procesamiento fue exitoso, salimos del bucle
                        if procesar_exito:
                            logging.info("Finaliza ejecución RPA PORTA COL exitosamente")
                            print("✅ Proceso completado con éxito")
                            return  # Salimos de la función si todo salió bien
                        
                        # Si llegamos aquí, es porque hubo un error en procesar_registros
                        print("Cerrando Chrome y reintentando en 10 segundos...")
                        cerrar_chrome()
                        sleep(10)  # Espera antes de reintentar
                        
                    except Exception as e:
                        # Error general durante el proceso
                        logging.error(f"Error general durante el procesamiento: {str(e)}")
                        print(f"Error general: {str(e)}")
                        cerrar_chrome()
                        sleep(10)  # Espera antes de reintentar
            else:
                print(f"No hay archivo llamado: {rutaAbsoluta}")
                logging.warning(f"No se encontró el archivo: {rutaAbsoluta}")
                sleep(60)
                
        except Exception as e:
            # Error en la inicialización general
            logging.error(f"Error en la inicialización: {str(e)}")
            cerrar_chrome()
            print(f"Error en la inicialización. Reintentando en 60 segundos...")
            sleep(60)
        finally:
            # Asegúrate de que Chrome no quede abierto si algo sale mal
            cerrar_chrome()

if __name__ == '__main__':
    main()    