from seleniumbase import Driver
from time import sleep
import logging
import pandas as pd
from datetime import datetime
import os


def procesar_registros(data,driver, df):
    try:    
        """Abre la URL y procesa cada registro con Selenium."""
        # ruta = data["ruta_Archivo"]
        # ruta_excel = ruta+"\\"+"RPA_PORTA_COL.xlsx"
        ruta_excel = os.path.join(data["ruta_Archivo"], "RPA_PORTA_COL_FILTRADO.xlsx")
        df["Gestionado"] = df["Gestionado"].fillna("").astype(str).str.strip()
        
        for index, row in df.iterrows(): 
            gestionado = row["Gestionado"]
            celular = str(row["Celular1"]).strip()
            if gestionado == "" and pd.notna(row["Celular1"]) and celular != "":
                # Aquí puedes agregar más acciones en la página si es necesario
               
                numero_linea = row["Celular1"]
                correo = row["Email"]
                operador = row["Operador"]
                Ventana_cambio = row["Ventana_cambio"]
                gestionado = row["Gestionado"]  # Campo a validar
                url = data["url_Publica"]
                driver.uc_open_with_reconnect(url, 4)
                sleep(1)
                driver.uc_gui_click_captcha()
                #sleep(4)
                sleep(2)
                print(f"Procesando línea: {numero_linea}")
                sleep(1)
                # Ingresar el número de teléfono en el campo
                driver.type("#hf-number", numero_linea)  # Usando el ID del input
                sleep(1)
                # Hacer clic en el botón de envío usando XPath
                driver.click("//button[@type='submit']")
                # Esperar para ver el resultado
                sleep(5)
                #validar si no existe tabla
                # Validar si el mensaje "No existen registros" está presente
                if driver.is_text_visible("No existen registros"):
                    print(" No existen registros en la tabla.")
                    df.at[index, "Operador"] = "NO ENCONTRADO"
                    df.at[index, "Ventana_cambio"] = "-"
                    df.at[index, "Gestionado"] = "KO"
                    # Guardar los cambios en el archivo Excel
                    print(ruta_excel)
                    df.to_excel(ruta_excel, index=False)
                    print(f"📄 Archivo actualizado: {ruta_excel}")
                    # Alimentar el archivo de "No encontrados"
                    ruta_sinResultado = data["ruta_Archivo_NoEncontrado"]
                    insertar_en_no_encontrado(ruta_sinResultado, numero_linea)
                    continue

                else:
                    print("✅ Hay registros disponibles.")
                    # Extraer datos de la tabla
                    try:
                        operador = driver.get_text("//tbody/tr[1]/td[3]")  # Prestador Receptor
                        Ventana_cambio = driver.get_text("//tbody/tr[1]/td[5]")  # Fecha Ventana

                        df.at[index, "Operador"] = operador
                        df.at[index, "Ventana_cambio"] = Ventana_cambio
                        df.at[index, "Gestionado"] = "OK"
                        # Guardar los cambios en el archivo Excel
                        print(ruta_excel)
                        df.to_excel(ruta_excel, index=False)
                        print(f"📄 Archivo actualizado: {ruta_excel}")
                        #continue

                    except Exception as e:
                        print(f"⚠️ Error extrayendo datos: {e}")
                        df.at[index, "Operador"] = "NO ENCONTRADO"
                        df.at[index, "Ventana_cambio"] = "-"
                        df.at[index, "Gestionado"] = "KO"
                        # Guardar los cambios en el archivo Excel
                        print(ruta_excel)
                        df.to_excel(ruta_excel, index=False)
                        print(f"📄 Archivo actualizado: {ruta_excel}")
                        # Alimentar el archivo de "No encontrados"
                        ruta_sinResultado = data["ruta_Archivo_NoEncontrado"]
                        insertar_en_no_encontrado(ruta_sinResultado, numero_linea)
                        continue
                        


                        #continue
                if operador == "TELEFÓNICA MÓVILES COLOMBIA S.A": #Movistar
                    print("alimentar otro excel")
                    ruta_telefonica = data["ruta_Archivo_telefonica"]
                    # Datos del registro actual
                    numero_linea = df.at[index, "Numero_linea"]
                    nuevo_registro = {
                        "NUMERO_LINEA": numero_linea,
                        "VENTANA_CAMBIO": Ventana_cambio
                        }
                        # Verificar si el archivo existe
                    
                    df_telefonica = pd.read_excel(ruta_telefonica, engine="openpyxl", dtype=str)
                   
                    # Asegurar que los datos sean tipo texto para evitar errores en comparación
                    df_telefonica["NUMERO_LINEA"] = df_telefonica["NUMERO_LINEA"].astype(str)
                    numero_linea = str(numero_linea)

                    # Verificar si el número ya está en el archivo
                    if numero_linea not in df_telefonica["NUMERO_LINEA"].values:
                        df_telefonica = pd.concat([df_telefonica, pd.DataFrame([nuevo_registro])], ignore_index=True)
                        df_telefonica.to_excel(ruta_telefonica, index=False)
                        print(f"📥 Registro agregado a Telefonica.xlsx: {numero_linea}")
                    else:
                        print(f"✅ El número ya existe en Telefonica.xlsx: {numero_linea}")
                # "COLOMBIA MÓVIL S.A." es tigo
                if operador == "COLOMBIA MÓVIL S.A.": #TIGO
                    print("alimentar otro excel")
                    ruta = data["ruta_Archivo_tigo"]
                    # Datos del registro actual
                    numero_linea = df.at[index, "Numero_linea"]
                    nuevo_registro = {
                        "NUMERO_LINEA": numero_linea,
                        "VENTANA_CAMBIO": Ventana_cambio,
                        "CORREO": correo,
                        }
                        # Verificar si el archivo existe
                    
                    df_telefonica = pd.read_excel(ruta, engine="openpyxl", dtype=str)
                   
                    # Asegurar que los datos sean tipo texto para evitar errores en comparación
                    df_telefonica["NUMERO_LINEA"] = df_telefonica["NUMERO_LINEA"].astype(str)
                    numero_linea = str(numero_linea)

                    # Verificar si el número ya está en el archivo
                    if numero_linea not in df_telefonica["NUMERO_LINEA"].values:
                        df_telefonica = pd.concat([df_telefonica, pd.DataFrame([nuevo_registro])], ignore_index=True)
                        df_telefonica.to_excel(ruta, index=False)
                        print(f"📥 Registro agregado a Telefonica.xlsx: {numero_linea}")
                    else:
                        print(f"✅ El número ya existe en Telefonica.xlsx: {numero_linea}")  
                # "PARTNERS TELECOM COLOMBIA S.A.S WOM" es WOM
                if operador == "PARTNERS TELECOM COLOMBIA S.A.S WOM": #WOM
                    print("alimentar otro excel")
                    ruta = data["ruta_Archivo_wom"]
                    # Datos del registro actual
                    numero_linea = df.at[index, "NUMERO_LINEA"]
                    nuevo_registro = {
                        "NUMERO_LINEA": numero_linea,
                        "VENTANA_CAMBIO": Ventana_cambio
                        
                        }
                        # Verificar si el archivo existe
                    
                    df_telefonica = pd.read_excel(ruta, engine="openpyxl", dtype=str)
                   
                    # Asegurar que los datos sean tipo texto para evitar errores en comparación
                    df_telefonica["NUMERO_LINEA"] = df_telefonica["NUMERO_LINEA"].astype(str)
                    numero_linea = str(numero_linea)

                    # Verificar si el número ya está en el archivo
                    if numero_linea not in df_telefonica["NUMERO_LINEA"].values:
                        df_telefonica = pd.concat([df_telefonica, pd.DataFrame([nuevo_registro])], ignore_index=True)
                        df_telefonica.to_excel(ruta, index=False)
                        print(f"📥 Registro agregado a Telefonica.xlsx: {numero_linea}")
                    else:
                        print(f"✅ El número ya existe en Telefonica.xlsx: {numero_linea}")
            
    except Exception as e:
        logging.error(f"Error en procesar registros: {e}")            

def insertar_en_no_encontrado(ruta_excel_sinresultado, numero_linea):
    """
    Inserta un registro en el archivo sinResultado si no existe.

    :param ruta_excel_sinresultado: Ruta del archivo sinResultado (.xlsx)
    :param numero_linea: Número de línea a insertar
    """
    try:
        # Leer archivo existente
        df_sinResultado = pd.read_excel(ruta_excel_sinresultado, engine="openpyxl", dtype=str)
    except FileNotFoundError:
        # Si el archivo no existe, crear DataFrame vacío con las columnas esperadas
        df_sinResultado = pd.DataFrame(columns=["NUMERO_LINEA", "DEUDA", "GESTIONADO"])

    # Asegurar que las columnas estén como texto
    df_sinResultado["NUMERO_LINEA"] = df_sinResultado["NUMERO_LINEA"].astype(str)
    numero_linea = str(numero_linea)

    # Verificar si ya existe
    if numero_linea not in df_sinResultado["NUMERO_LINEA"].values:
        nuevo_registro = {
            "NUMERO_LINEA": numero_linea,
            "OPERADOR": "",
            "DEUDA": "",
            "GESTIONADO": ""
        }

        df_sinResultado = pd.concat([df_sinResultado, pd.DataFrame([nuevo_registro])], ignore_index=True)
        df_sinResultado.to_excel(ruta_excel_sinresultado, index=False)
        print(f"📥 Registro agregado a sinResultado: {numero_linea}")
    else:
        print(f"✅ El número ya existe en sinResultado: {numero_linea}")
 