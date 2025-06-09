from seleniumbase import Driver
from time import sleep
import logging
import pandas as pd
import pyautogui
import subprocess
import pytesseract
from PIL import Image
import os
import re
from datetime import datetime
           
'''
def procesar_registros(data, df, ruta_excel):
    """
    Procesa los registros del DataFrame, consultando en la web de portabilidad
    y actualizando el estado en el archivo CSV.
    
    Si ocurre un error, la excepción será capturada por el bucle en main()
    para reiniciar el proceso.
    """
    pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
    ruta = data["ruta_Archivo"]
    df["Gestionado"] = df["Gestionado"].fillna("").astype(str).str.strip()

    try:
        # Esperamos a que cargue el navegador
        sleep(8)
        logging.info("Abriendo la URL de consulta")
        
        # Acceder a la página
        pyautogui.click("C:\\RPA_PORTACOL_INTERACTIVO\\resources\\img\\refresh.PNG")
        sleep(4)
        pyautogui.write("https://www.portabilidadcolombia.com.co/?handler=Check")
        sleep(1)
        pyautogui.press("Enter")
        sleep(6)
        
        # Contador para seguimiento
        total_registros = len(df)
        registros_procesados = 0
        
        for index, row in df.iterrows(): 
            gestionado = row["Gestionado"]
            if gestionado == "":  
                numero_linea = row["TeleNum"]
                numero_linea = str(numero_linea)
                logging.info(f"Procesando línea: {numero_linea} ({registros_procesados+1}/{total_registros})")
                print(f"📱 Procesando línea: {numero_linea} ({registros_procesados+1}/{total_registros})")
                correo = row["Email"]
                
                try:
                    # Buscamos el número en la página
                    pyautogui.click("C:\\RPA_PORTACOL_INTERACTIVO\\resources\\img\\check.PNG")
                    sleep(4)
                    
                    # Ingresar el número
                    pyautogui.hotkey('shift', 'tab')
                    sleep(1)
                    pyautogui.press('delete')
                    sleep(2)
                    pyautogui.write(numero_linea)
                    sleep(2)
                    
                    # Hacer clic en buscar (mediante tabulaciones)
                    for _ in range(4):
                        pyautogui.press('tab')
                        sleep(0.5)
                    pyautogui.press("Enter")
                    sleep(8)
                    
                    try:
                        # Capturar y analizar la pantalla
                        img = pyautogui.screenshot()
                        text = pytesseract.image_to_string(img)
                        logging.debug(f"Texto extraído para línea {numero_linea}")
                        print(f"text: {text}")
                        # Procesar el resultado
                        if "no existen registros" in text.lower():
                            logging.info(f"Línea {numero_linea}: No existen registros")
                            print(f"✅ Línea {numero_linea}: No existen registros")
                            
                            df.at[index, "Operador"] = "NO ENCONTRADO"
                            df.at[index, "Ventana_cambio"] = "-"
                            df.at[index, "Gestionado"] = "KO"
                            
                            # Guardar cambios y continuar
                            df.to_csv(ruta_excel, index=False)
                            print(f"📄 Archivo actualizado: {ruta_excel}")
                            insertar_numero_sin_duplicado(numero_linea, correo)
                            registros_procesados += 1
                            continue
                        
                        # Intentar extraer la fecha
                        fecha_match = re.search(r'\d{2}/\d{2}/\d{4}', text)
                        fecha_ventana = fecha_match.group(0) if fecha_match else None
                        
                        # Buscar operador
                        operador = None
                        if fecha_ventana:
                            match = re.search(r'(\d{10})\s+\d+\s+([A-Z\s\.]+?)\s+[A-Z\s\.]+\s+' + re.escape(fecha_ventana), text)
                            if match:
                                operador = match.group(2).strip()
                        
                        # Actualizar resultados
                        logging.info(f"Línea {numero_linea}: Operador={operador}, Fecha={fecha_ventana}")
                        print(f"📊 Línea {numero_linea}: Operador={operador}, Fecha={fecha_ventana}")
                        
                                               
                        # Procesar según operador
                        if operador in "COMCEL": 
                            df.at[index, "Operador"] = "COMCEL S.A"
                            df.at[index, "Ventana_cambio"] = fecha_ventana
                            df.at[index, "Gestionado"] = "OK"
                            df.to_csv(ruta_excel, index=False)
                            pass  # No se hace nada especial para COMCEL
                        elif operador in "TELEFONICA":  # Movistar TELEFÓNICA MÓVILES COLOMBIA S.A
                            df.at[index, "Operador"] = "TELEFONICA MÓVILES COLOMBIA S.A"
                            df.at[index, "Ventana_cambio"] = fecha_ventana
                            df.at[index, "Gestionado"] = "OK"
                            df.to_csv(ruta_excel, index=False)
                            print("🔄 Alimentando otro excel (Telefonica)")
                            insertar_numero_sin_duplicado_Telefonica(numero_linea)
                        elif operador in "COLOMBIA":  # TIGO COLOMBIA MÓVIL S.A.
                            df.at[index, "Operador"] = "TIGO COLOMBIA MÓVIL S.A."
                            df.at[index, "Ventana_cambio"] = fecha_ventana
                            df.at[index, "Gestionado"] = "OK"
                            df.to_csv(ruta_excel, index=False)
                            print("🔄 Alimentando otro excel (Tigo)")
                            insertar_numero_sin_duplicado_Tigo(numero_linea,correo)
                        elif operador in "PARTNERS":  # WOM PARTNERS TELECOM COLOMBIA S.A.S WOM
                            df.at[index, "Operador"] = "WOM PARTNERS TELECOM COLOMBIA S.A.S WOM"
                            df.at[index, "Ventana_cambio"] = fecha_ventana
                            df.at[index, "Gestionado"] = "OK"
                            df.to_csv(ruta_excel, index=False)
                            print("🔄 Alimentando otro excel (Wom)")
                            insertar_numero_sin_duplicado_Wom(numero_linea)
                        elif operador in "VIRGIN":  # WOM PARTNERS TELECOM COLOMBIA S.A.S WOM
                            df.at[index, "Operador"] = "VIRGIN MOBILE"
                            df.at[index, "Ventana_cambio"] = fecha_ventana
                            df.at[index, "Gestionado"] = "OK"
                            df.to_csv(ruta_excel, index=False)
                            print("🔄 Alimentando otro excel (Wom)")
                            insertar_numero_sin_duplicado_Wom(numero_linea)
                        else:
                            print("⚠️ No se pudo extraer el operador y la fecha.")
                            df.at[index, "Operador"] = "Error extraccion"
                            df.at[index, "Ventana_cambio"] = "-"
                            df.at[index, "Gestionado"] = "KO"
                            df.to_csv(ruta_excel, index=False)
                        
                        registros_procesados += 1
                    
                    except Exception as e:
                        # Error al procesar la captura de pantalla
                        error_msg = f"Error al procesar captura para línea {numero_linea}: {e}"
                        logging.error(error_msg)
                        print(f"❌ {error_msg}")
                        
                        # No marcamos como gestionado para que sea reprocesado
                        df.to_csv(ruta_excel, index=False)
                        # Lanzamos la excepción para que el main la capture y reinicie
                        raise Exception(f"Error al procesar captura: {e}")
                
                except Exception as e:
                    # Error al interactuar con la página
                    error_msg = f"Error al interactuar con la página para línea {numero_linea}: {e}"
                    logging.error(error_msg)
                    print(f"❌ {error_msg}")
                    
                    # Lanzamos la excepción para que el main la capture y reinicie
                    raise Exception(f"Error en la interacción: {e}")
            else:
                # Este registro ya fue gestionado
                registros_procesados += 1
        
        # Si completamos todos los registros, marcamos como exitoso
        logging.info(f"Proceso completado exitosamente. {registros_procesados} registros procesados.")
        print(f"✅ Proceso completado exitosamente. {registros_procesados} registros procesados.")
        return True
            
    except Exception as e:
        # Propagamos la excepción para que sea capturada en el main
        logging.error(f"Error en procesar_registros: {e}")
        print(f"❌ Error general en procesar_registros: {e}")
        raise
'''    
def procesar_registros(data, df, ruta_excel):
    """
    Procesa los registros del DataFrame, consultando en la web de portabilidad
    y actualizando el estado en el archivo CSV.
    
    Si ocurre un error, la excepción será capturada por el bucle en main()
    para reiniciar el proceso.
    """
    pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
    ruta = data["ruta_Archivo"]
    df["Gestionado"] = df["Gestionado"].fillna("").astype(str).str.strip()

    try:
        # Esperamos a que cargue el navegador
        sleep(8)
        logging.info("Abriendo la URL de consulta")
        
        # Acceder a la página
        pyautogui.click("C:\\RPA_PORTACOL_INTERACTIVO\\resources\\img\\refresh.PNG")
        sleep(4)
        pyautogui.write("https://www.portabilidadcolombia.com.co/?handler=Check")
        sleep(1)
        pyautogui.press("Enter")
        sleep(8)
        
        # Contador para seguimiento
        total_registros = len(df)
        registros_procesados = 0
        
        for index, row in df.iterrows(): 
            gestionado = row["Gestionado"]
            if gestionado == "":  
                numero_linea = row["TeleNum"]
                numero_linea = str(numero_linea)
                logging.info(f"Procesando línea: {numero_linea} ({registros_procesados+1}/{total_registros})")
                print(f"📱 Procesando línea: {numero_linea} ({registros_procesados+1}/{total_registros})")
                correo = row["Email"]
                
                try:
                    # Buscamos el número en la página
                    pyautogui.click("C:\\RPA_PORTACOL_INTERACTIVO\\resources\\img\\check.PNG")
                    sleep(4)
                    
                    # Ingresar el número
                    pyautogui.hotkey('shift', 'tab')
                    sleep(1)
                    pyautogui.press('delete')
                    sleep(2)
                    pyautogui.write(numero_linea)
                    sleep(2)
                    
                    # Hacer clic en buscar (mediante tabulaciones)
                    for _ in range(4):
                        pyautogui.press('tab')
                        sleep(0.5)
                    pyautogui.press("Enter")
                    sleep(8)
                    
                    try:
                        # Capturar y analizar la pantalla
                        img = pyautogui.screenshot()
                        text = pytesseract.image_to_string(img)
                        logging.debug(f"Texto extraído para línea {numero_linea}")
                        print(f"text: {text}")
                        
                        # Procesar el resultado
                        if "no existen registros" in text.lower():
                            logging.info(f"Línea {numero_linea}: No existen registros")
                            print(f"Línea {numero_linea}: No existen registros")
                            
                            df.at[index, "Operador"] = "NO ENCONTRADO"
                            df.at[index, "Ventana_cambio"] = "-"
                            df.at[index, "Gestionado"] = "KO"
                            
                            # Guardar cambios y continuar
                            df.to_csv(ruta_excel, index=False)
                            print(f"Archivo actualizado: {ruta_excel}")
                            insertar_numero_sin_duplicado(numero_linea, correo)
                            registros_procesados += 1
                            continue
                        
                        # MÉTODO MEJORADO DE EXTRACCIÓN DE DATOS
                        # Buscar directamente las operadoras en el texto
                        operador = None
                        fecha_ventana = None
                        
                        # Primero buscamos una fecha en formato DD/MM/YYYY
                        fecha_match = re.search(r'\d{2}/\d{2}/\d{4}', text)
                        if fecha_match:
                            fecha_ventana = fecha_match.group(0)
                        operador  =extraer_prestador_receptor(text)
                        if operador == "No encontrado":
                            operador = extraer_prestador_receptor_v2(text)
                        print(f"operador: {operador}, fecha: {fecha_ventana}")
                            
                        # Buscar operadoras conocidas en el texto
                        operadoras = [
                            ("TELEFONICA", "TELEFONICA MÓVILES COLOMBIA S.A"),
                            ("COMCEL", "COMCEL S.A"),
                            ("COLOMBIA", "TIGO COLOMBIA MÓVIL S.A."),
                            ("PARTNERS", "WOM PARTNERS TELECOM COLOMBIA S.A.S WOM"),
                            ("VIRGIN", "VIRGIN MOBILE"),
                            ("ETB", "ETB - EMPRESA DE TELECOMUNICACIONES DE BOGOTÁ S.A."),
                        ]
                        
                        # Buscar por coincidencia de texto
                        # for clave, nombre_completo in operadoras:
                        #     if clave in text:
                        #         operador = clave
                        #         break
                       
                        
                        # # Si no se encontró de esta manera, probar con pattern matching más específico
                        # if not operador and "Estado de la numeración" in text:
                        #     # Intenta encontrar la línea que contiene el número de teléfono
                        #     lineas = text.split('\n')
                        #     for linea in lineas:
                        #         if numero_linea in linea:
                        #             # Una vez encontrada la línea, buscar qué operadora aparece
                        #             for clave, _ in operadoras:
                        #                 if clave in linea:
                        #                     operador = clave
                        #                     break
                        
                        # logging.info(f"Línea {numero_linea}: Operador={operador}, Fecha={fecha_ventana}")
                        # print(f"📊 Línea {numero_linea}: Operador={operador}, Fecha={fecha_ventana}")
                        
                        # Procesar según operador
                        if not operador:
                            # Si todavía no se encuentra el operador, usar un enfoque más agresivo
                            if "TELEFONICA MOVILES" in text or "TELEFONICA MÓVILES" in text:
                                operador = "TELEFONICA"
                            elif "COMCEL" in text:
                                operador = "COMCEL"
                            elif "COLOMBIA MÓVIL" in text or "TIGO" in text:
                                operador = "COLOMBIA"
                            elif "WOM" in text or "PARTNERS" in text:
                                operador = "PARTNERS"
                            elif "VIRGIN" in text:
                                operador = "VIRGIN"
                            elif "ETB" in text:
                                operador = "ETB"

                        
                        # Actualizar basado en el operador encontrado
                        if operador.__contains__("COMCEL"): 
                            df.at[index, "Operador"] = "COMCEL S.A"
                            df.at[index, "Ventana_cambio"] = fecha_ventana
                            df.at[index, "Gestionado"] = "OK"
                            df.to_csv(ruta_excel, index=False)
                            pass  # No se hace nada especial para COMCEL
                        elif operador.__contains__("GRUPO"): 
                            df.at[index, "Operador"] = "GRUPO EXITO"
                            df.at[index, "Ventana_cambio"] = fecha_ventana
                            df.at[index, "Gestionado"] = "OK"
                            df.to_csv(ruta_excel, index=False)
                            pass  # No se hace nada especial para COMCEL
                        elif operador == "TELEFONICA":  # Movistar TELEFÓNICA MÓVILES COLOMBIA S.A
                            df.at[index, "Operador"] = "TELEFONICA MÓVILES COLOMBIA S.A"
                            df.at[index, "Ventana_cambio"] = fecha_ventana
                            df.at[index, "Gestionado"] = "OK"
                            df.to_csv(ruta_excel, index=False)
                            print("Alimentando otro excel (Telefonica)")
                            insertar_numero_sin_duplicado_Telefonica(numero_linea)
                            pass
                        elif operador == "COLOMBIA":  # TIGO COLOMBIA MÓVIL S.A.
                            df.at[index, "Operador"] = "TIGO COLOMBIA MÓVIL S.A."
                            df.at[index, "Ventana_cambio"] = fecha_ventana
                            df.at[index, "Gestionado"] = "OK"
                            df.to_csv(ruta_excel, index=False)
                            print("Alimentando otro excel (Tigo)")
                            insertar_numero_sin_duplicado_Tigo(numero_linea,correo)
                            pass
                        elif operador == "PARTNERS":  # WOM PARTNERS TELECOM COLOMBIA S.A.S WOM
                            df.at[index, "Operador"] = "WOM PARTNERS TELECOM COLOMBIA S.A.S WOM"
                            df.at[index, "Ventana_cambio"] = fecha_ventana
                            df.at[index, "Gestionado"] = "OK"
                            df.to_csv(ruta_excel, index=False)
                            print("Alimentando otro excel (Wom)")
                            insertar_numero_sin_duplicado_Wom(numero_linea)
                            pass
                        elif operador == "VIRGIN":  # VIRGIN MOBILE
                            df.at[index, "Operador"] = "VIRGIN MOBILE"
                            df.at[index, "Ventana_cambio"] = fecha_ventana
                            df.at[index, "Gestionado"] = "OK"
                            df.to_csv(ruta_excel, index=False)
                            pass  # No se hace nada especial para VIRGIN
                        elif operador.__contains__("ETB"):
                            df.at[index, "Operador"] = "ETB - EMPRESA DE TELECOMUNICACIONES DE BOGOTÁ S.A."
                            df.at[index, "Ventana_cambio"] = fecha_ventana
                            df.at[index, "Gestionado"] = "OK"
                            df.to_csv(ruta_excel, index=False)
                            print("🔄 Alimentando otro excel (ETB)")
                            pass
                        else:
                            print("⚠️ No se pudo extraer el operador y la fecha.")
                            logging.warning(f"Texto extraído: {text}")
                            df.at[index, "Operador"] = "Error extraccion"
                            df.at[index, "Ventana_cambio"] = "-"
                            df.at[index, "Gestionado"] = "KO"
                            df.to_csv(ruta_excel, index=False)
                            now = datetime.now().strftime("%Y%m%d_%H%M%S")
                            nombre_archivo = f"Error_{now}_{numero_linea}.png"
                            ruta_carpeta = "C:\\RPA_PORTA_COL_BASE\\error"
                            ruta_completa = os.path.join(ruta_carpeta, nombre_archivo)
                            screenshot = pyautogui.screenshot()
                            # Guardar la captura
                            screenshot.save(ruta_completa)
                            print(f"Captura guardada en: {ruta_completa}")
                            pass

            
                        registros_procesados += 1
                    
                    except Exception as e:
                        # Error al procesar la captura de pantalla
                        error_msg = f"Error al procesar captura para línea {numero_linea}: {e}"
                        logging.error(error_msg)
                        print(f"{error_msg}")
                        
                        # No marcamos como gestionado para que sea reprocesado
                        df.to_csv(ruta_excel, index=False)
                        # Lanzamos la excepción para que el main la capture y reinicie
                        raise Exception(f"Error al procesar captura: {e}")
                
                except Exception as e:
                    # Error al interactuar con la página
                    error_msg = f"Error al interactuar con la página para línea {numero_linea}: {e}"
                    logging.error(error_msg)
                    print(f"❌ {error_msg}")
                    
                    # Lanzamos la excepción para que el main la capture y reinicie
                    raise Exception(f"Error en la interacción: {e}")
            else:
                # Este registro ya fue gestionado
                registros_procesados += 1
        
        # Si completamos todos los registros, marcamos como exitoso
        logging.info(f"Proceso completado exitosamente. {registros_procesados} registros procesados.")
        print(f"✅ Proceso completado exitosamente. {registros_procesados} registros procesados.")
        return True
            
    except Exception as e:
        # Propagamos la excepción para que sea capturada en el main
        logging.error(f"Error en procesar_registros: {e}")
        print(f"❌ Error general en procesar_registros: {e}")
        raise
def insertar_numero_sin_duplicado(numero_linea, correo):
    """
    Inserta un número de línea y correo en un archivo CSV evitando duplicados.
    
    Args:
        numero_linea (str): Número de línea a insertar
        correo (str): Correo electrónico asociado
    """
    ruta_csv = r"C:\RPA_PORTA_COL_BASE\sin_resultados.csv"

    # Si no existe el archivo, creamos uno nuevo con las columnas necesarias
    if not os.path.exists(ruta_csv):
        try:
            df_nuevo = pd.DataFrame({"NUMERO_LINEA": [numero_linea], "CORREO": [correo]})
            df_nuevo.to_csv(ruta_csv, sep=";", index=False)
            print(f"✅ Archivo creado y número insertado: {numero_linea} con correo: {correo}")
            return
        except PermissionError as e:
            print(f"❌ Error: No tienes permiso para crear el archivo. {e}")
            return
        except Exception as e:
            print(f"❌ Error al crear el archivo: {e}")
            return

    # Detectamos automáticamente el separador
    separador = detectar_separador(ruta_csv)

    # Leemos el CSV con el separador correcto
    try:
        df_existente = pd.read_csv(ruta_csv, sep=separador, dtype=str, engine='python', on_bad_lines='skip')
    except Exception as e:
        print(f"❌ Error al leer el archivo CSV: {e}")
        return

    # Asegurarnos de que existan las columnas necesarias
    if "NUMERO_LINEA" not in df_existente.columns:
        print("⚠️ Columna 'NUMERO_LINEA' no encontrada en el archivo.")
        return

    if "CORREO" not in df_existente.columns:
        df_existente["CORREO"] = ""
        print("📧 Columna CORREO agregada al archivo")

    # Validamos si ya existe el número
    if numero_linea not in df_existente["NUMERO_LINEA"].values:
        df_nuevo = pd.DataFrame({"NUMERO_LINEA": [numero_linea], "CORREO": [correo]})
        df_existente = pd.concat([df_existente, df_nuevo], ignore_index=True)
        try:
            df_existente.to_csv(ruta_csv, sep=separador, index=False)
            print(f"✅ Número agregado a sin_resultados.csv: {numero_linea} con correo: {correo}")
        except PermissionError as e:
            print(f"❌ Error: No se puede escribir en el archivo. Está abierto o falta permiso. {e}")
        except Exception as e:
            print(f"❌ Error al guardar el archivo: {e}")
    else:
        # Si ya existe, verificar si el correo es diferente y actualizarlo si es necesario
        indice = df_existente.index[df_existente["NUMERO_LINEA"] == numero_linea].tolist()[0]
        correo_actual = df_existente.at[indice, "CORREO"]

        if pd.isna(correo_actual) or correo_actual.strip() == "":
            # Actualizar si el correo está vacío
            df_existente.at[indice, "CORREO"] = correo
            try:
                df_existente.to_csv(ruta_csv, sep=separador, index=False)
                print(f"📧 Correo agregado para el número {numero_linea}: {correo}")
            except PermissionError as e:
                print(f"❌ Error: No se puede escribir en el archivo. Está abierto o falta permiso. {e}")
        elif correo != correo_actual and correo.strip() != "":
            # Actualizar si es diferente y no está vacío
            df_existente.at[indice, "CORREO"] = correo
            try:
                df_existente.to_csv(ruta_csv, sep=separador, index=False)
                print(f"🔄 Correo actualizado para el número {numero_linea}: {correo}")
            except PermissionError as e:
                print(f"❌ Error: No se puede escribir en el archivo. Está abierto o falta permiso. {e}")
        else:
            print(f"ℹ El número {numero_linea} ya está en sin_resultados.csv con el mismo correo.")
            
def detectar_separador(ruta_archivo):
    """Detecta automáticamente el separador del CSV (coma o punto y coma)"""
    with open(ruta_archivo, "r", encoding="utf-8") as f:
        primera_linea = f.readline()
    if ";" in primera_linea:
        return ";"
    else:
        return ","
    
def insertar_numero_sin_duplicado_Telefonica(numero_linea):
    ruta_csv = r"C:\RPA_PORTA_COL_BASE\Telefonica.csv"

    # Si no existe el archivo, lo creamos con el número directamente
    if not os.path.exists(ruta_csv):
        df_nuevo = pd.DataFrame({"NUMERO_LINEA": [numero_linea]})
        df_nuevo.to_csv(ruta_csv, sep=";", index=False)
        print(f"Archivo creado y número insertado: {numero_linea}")
        return

    # Detectamos el separador automáticamente
    separador = detectar_separador(ruta_csv)

    # Leemos el CSV con el separador correcto
    try:
        df_existente = pd.read_csv(ruta_csv, sep=separador, dtype=str, engine='python', on_bad_lines='skip')
    except Exception as e:
        print(f"❌ Error al leer el archivo CSV: {e}")
        return

    # Verificar si la columna existe
    if "NUMERO_LINEA" not in df_existente.columns:
        print("⚠️ Columna 'NUMERO_LINEA' no encontrada en el archivo.")
        return

    # Validamos si ya existe el número
    if numero_linea not in df_existente["NUMERO_LINEA"].values:
        df_nuevo = pd.DataFrame({"NUMERO_LINEA": [numero_linea]})
        df_existente = pd.concat([df_existente, df_nuevo], ignore_index=True)
        df_existente.to_csv(ruta_csv, sep=separador, index=False)
        print(f"✅ Número agregado a Telefonica.csv: {numero_linea}")
    else:
        print(f"ℹ El número {numero_linea} ya está en Telefonica.csv")

def insertar_numero_sin_duplicado_Tigo(numero_linea, correo):
    ruta_csv = r"C:\RPA_PORTA_COL_BASE\tigo.csv"

    # Si no existe el archivo, creamos uno nuevo con las columnas necesarias
    if not os.path.exists(ruta_csv):
        df_nuevo = pd.DataFrame({"NUMERO_LINEA": [numero_linea], "CORREO": [correo]})
        df_nuevo.to_csv(ruta_csv, sep=";", index=False)
        print(f"Archivo creado y número insertado: {numero_linea} con correo: {correo}")
        return

    # Detectamos automáticamente el separador
    separador = detectar_separador(ruta_csv)

    # Leemos el CSV con el separador correcto
    try:
        df_existente = pd.read_csv(ruta_csv, sep=separador, dtype=str, engine='python', on_bad_lines='skip')
    except Exception as e:
        print(f"❌ Error al leer el archivo CSV: {e}")
        return

    # Asegurarnos de que exista la columna CORREO
    if "CORREO" not in df_existente.columns:
        df_existente["CORREO"] = ""

    # Validamos si ya existe el número
    if numero_linea not in df_existente["NUMERO_LINEA"].values:
        # Agregar nueva fila
        df_nuevo = pd.DataFrame({"NUMERO_LINEA": [numero_linea], "CORREO": [correo]})
        df_existente = pd.concat([df_existente, df_nuevo], ignore_index=True)
        df_existente.to_csv(ruta_csv, sep=separador, index=False)
        print(f"✅ Número agregado a tigo.csv: {numero_linea} con correo: {correo}")
    else:
        # Si ya existe, verificar si el correo es diferente
        indice = df_existente.index[df_existente["NUMERO_LINEA"] == numero_linea].tolist()[0]
        correo_actual = df_existente.at[indice, "CORREO"]

        if pd.isna(correo_actual) or correo_actual.strip() == "":
            # Actualizar si el correo está vacío
            df_existente.at[indice, "CORREO"] = correo
            df_existente.to_csv(ruta_csv, sep=separador, index=False)
            print(f"📧 Correo agregado para el número {numero_linea}: {correo}")
        elif correo != correo_actual:
            # Opcional: actualizar si es diferente
            df_existente.at[indice, "CORREO"] = correo
            df_existente.to_csv(ruta_csv, sep=separador, index=False)
            print(f"🔄 Correo actualizado para el número {numero_linea}: {correo}")
        else:
            print(f"ℹ El número {numero_linea} ya está en tigo.csv con el mismo correo.")

def insertar_numero_sin_duplicado_Wom(numero_linea):
    ruta_csv = r"C:\RPA_PORTA_COL_BASE\wom.csv"

    # Si no existe el archivo, lo creamos con el número directamente
    if not os.path.exists(ruta_csv):
        df_nuevo = pd.DataFrame({"NUMERO_LINEA": [numero_linea]})
        df_nuevo.to_csv(ruta_csv, sep=";", index=False)
        print(f"Archivo creado y número insertado: {numero_linea}")
        return

    # Detectamos el separador automáticamente
    separador = detectar_separador(ruta_csv)

    # Leemos el CSV con el separador correcto
    try:
        df_existente = pd.read_csv(ruta_csv, sep=separador, dtype=str, engine='python', on_bad_lines='skip')
    except Exception as e:
        print(f"❌ Error al leer el archivo CSV: {e}")
        return

    # Verificar si la columna existe
    if "NUMERO_LINEA" not in df_existente.columns:
        print("⚠️ Columna 'NUMERO_LINEA' no encontrada en el archivo.")
        return

    # Validamos si ya existe el número
    if numero_linea not in df_existente["NUMERO_LINEA"].values:
        df_nuevo = pd.DataFrame({"NUMERO_LINEA": [numero_linea]})
        df_existente = pd.concat([df_existente, df_nuevo], ignore_index=True)
        df_existente.to_csv(ruta_csv, sep=separador, index=False)
        print(f"✅ Número agregado a Telefonica.csv: {numero_linea}")
    else:
        print(f"ℹ El número {numero_linea} ya está en Telefonica.csv")


def extraer_prestador_receptor(texto):
    lineas = texto.splitlines()
    for linea in lineas:
        # Eliminar espacios múltiples y normalizar
        linea = re.sub(r'\s+', ' ', linea).strip()

        # Buscar patrón flexible: Número seguido de espacio(s) o coma, luego nombre
        match = re.search(r"\d+[,\.]?\s+\d+\s+([A-Z\s\.]+?)(?:\s+[A-Z]{2,})+", linea)
        if match:
            prestador_receptor = match.group(1).strip()
            return prestador_receptor
    return "No encontrado"

def extraer_prestador_receptor_v2(texto):
    texto = texto.upper()
    texto = re.sub(r'\s+', ' ', texto)  # Normaliza espacios

    # Palabras clave comunes de operadores móviles en Colombia
    operadores = {
        "TELEFONICA": r"TELEFONICA|TELEFONICA MOVILES COLOMBIA",
        "COLOMBIA": r"TIGO|TIGO UNE",
        "CLARO": r"CLARO|AMERICAS MOVIL",
        "WOM": r"WOM",
        "ETB": r"ETB|EMPRESA DE TELECOMUNICACIONES DE BOGOTA",
        "UNE": r"UNE|EPM|EMPRESAS PUBLICAS DE MEDELLIN"
    }

    # Buscar líneas que empiecen por número celular seguido de código
    lineas = texto.split(" ")

    for i, palabra in enumerate(lineas):
        # Buscar si la palabra parece un número celular (10 dígitos)
        if re.match(r"\d{10}", palabra):
            # Verificar las siguientes palabras (hasta 15) como posible nombre
            fragmento = " ".join(lineas[i+1:i+15])
            fragmento = re.sub(r"[^\w\s]", "", fragmento)  # Limpiar signos raros

            # Buscar coincidencia con cada operador usando su expresión regular
            for operador, patron in operadores.items():
                if re.search(patron, fragmento):
                    return operador

    return "No encontrado"