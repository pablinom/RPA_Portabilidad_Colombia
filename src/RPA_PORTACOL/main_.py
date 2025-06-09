from time import sleep
from datetime import datetime, timedelta
import logging
import json
import shutil
import os
import pandas as pd
from Procesamiento.ProcesamientoWEB import *
from Procesamiento.ProcesamientoBBDD import *
import os.path as path

def last_day_of_month(date): 
    if date.month == 12: 
        return date.replace(day=31) 
    return date.replace(month=date.month+1, day=1) - timedelta(days=1)

def iniciar_selenium():
    driver = Driver(uc=True)
    return driver

def ArchivoExistente(rutaAbsoluta):
    try:
        logging.info("Inicia LecturaArchivo")
        return os.path.isfile(rutaAbsoluta)
    except Exception as e:
        logging.error("Error al buscar el archivo: " + str(e))
        return False

def main():
    driver = None
    while True:
        try:
            with open('C:\\Users\\USER\\Desktop\\RPA_PORTACOL_INTERACTIVO\\src\\config.json', encoding='utf-8') as json_file:
                data = json.load(json_file)

            # Crear carpetas si no existen
            ruta_carpeta_logs = data["ruta_carpeta_logs"]
            os.makedirs(ruta_carpeta_logs, exist_ok=True)

            ruta_carpeta_RPA_logs = data["ruta_carpeta_RPA_logs"]
            os.makedirs(ruta_carpeta_RPA_logs, exist_ok=True)

            # Obtener fecha actual
            current_date = datetime.today().strftime('%d-%m-%Y')

            # Definir nombre de archivo de logs
            nombre_archivo_log = f"Logs_Rpa_Portacol_{current_date}.txt"
            ruta_completa_log = os.path.join(ruta_carpeta_RPA_logs, nombre_archivo_log)

            # Configurar logging
            logging.basicConfig(
                filename=ruta_completa_log,
                filemode='a',
                format='%(asctime)s %(levelname)-8s %(message)s',
                level=logging.INFO,
                datefmt='%Y-%m-%d %H:%M:%S'
            )

            logging.info("Inicia ejecución RPA PORTA COL")


            ruta_origen = data["ruta_archivo_modelo"]
            ruta_destino = data["ruta_Archivo"]
            archivos_origen = os.listdir(ruta_origen)

            for archivo in archivos_origen:
                ruta_archivo_origen = os.path.join(ruta_origen, archivo)
                ruta_archivo_destino = os.path.join(ruta_destino, archivo)

                if not os.path.exists(ruta_archivo_destino):
                    shutil.copy2(ruta_archivo_origen, ruta_archivo_destino)
                    print(f"📂 Copiado: {archivo} → {ruta_destino}")
                else:
                    print(f"✅ Ya existe: {archivo}")

            ruta = data["ruta_Archivo"]
            rutaAbsoluta = os.path.join(ruta, "RPA_PORTA_COL.xlsx")  
            ruta_csv = os.path.join(ruta, "RPA_PORTA_COL.csv")  # Ruta del nuevo CSV

            ExisteArchivo = ArchivoExistente(rutaAbsoluta)

            if ExisteArchivo:
                # Convertir XLSX a CSV si no existe el CSV
                if not os.path.exists(ruta_csv):
                    print("Leyendo excel espere para procesarlo!!")
                    df_convert = pd.read_excel(rutaAbsoluta)
                    df_convert.to_csv(ruta_csv, index=False)
                    print(f"✅ Archivo convertido a CSV: {ruta_csv}")
                else:
                    print(f"✅ El archivo CSV ya existe: {ruta_csv}")

                ruta_filtrado = os.path.join(data["ruta_Archivo"], "RPA_PORTA_COL_FILTRADO.xlsx")

                if os.path.exists(ruta_filtrado):
                    print(f"✅ El archivo filtrado ya existe: {ruta_filtrado}")
                    df_filtrado = ProcesarExcel(data, ruta_filtrado)

                    columnas_necesarias = ["Celular1", "Email", "Operador", "Ventana_cambio", "Gestionado"]
                    df_filtrado = df_filtrado[columnas_necesarias]

                    print("🔄 Procesando el archivo filtrado existente...")
                    procesar_registros(data, driver, df_filtrado)

                else:
                    df = ProcesarExcel(data, ruta_csv)  # Ahora se procesa el CSV
                    registros_pendientes = obtener_registros_pendientes(df)

                    if registros_pendientes.empty:
                        print("No hay registros pendientes. No se abrirá Selenium.")
                        logging.info("No hay registros pendientes. Finaliza ejecución.")
                        if driver:
                            driver.quit()
                        break
                    else:
                        registros_pendientes.to_excel(ruta_filtrado, index=False)
                        print(f"📄 Archivo filtrado generado: {ruta_filtrado}")

                        if driver is None:
                            driver = iniciar_selenium()

                        df_filtrado = ProcesarExcel(data, ruta_filtrado)

                        columnas_necesarias = ["Celular1", "Email", "Operador", "Ventana_cambio", "Gestionado"]
                        df_filtrado = df_filtrado[columnas_necesarias]

                        procesar_registros(data, driver, df_filtrado)

                        print("✅ Proceso finalizado exitosamente.")
                        logging.info("Finaliza ejecución RPA PORTA COL exitosamente.")

                        if driver:
                            driver.quit()
                        break

            else:
                print(f"No hay archivo llamado: {rutaAbsoluta} por favor deje el archivo en la ruta indicada")
                logging.warning(f"No se encontró el archivo {rutaAbsoluta}")
                sleep(30)

        except Exception as e:
            print(f"⚠️ Error detectado: {e}")
            logging.error(f"Error en el main, reintentando: {e}")
            if driver:
                try:
                    driver.quit()
                except:
                    pass
            driver = None
            print("♻️ Reiniciando proceso en 15 segundos...")
            sleep(15)

if __name__ == '__main__':
    main()
