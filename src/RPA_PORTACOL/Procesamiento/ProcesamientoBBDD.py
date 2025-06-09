from time import sleep
from datetime import datetime
from datetime import timedelta
import os
import logging
import locale
import pandas as pd

# Leer Excel
def ProcesarExcel(data, rutaExcel):
    try: 
        logging.info("Lectura de Excel")
       
        # Leer el archivo, asegurando que todas las columnas sean tratadas como texto
        # Cargar el archivo usando openpyxl (más confiable que el motor por defecto)
        print("Leyendo Excel espere porfavor ....")
        #df = pd.read_excel(rutaExcel, dtype=str)  # Leer todo como texto para evitar errores de tipo de dato
        df = pd.read_csv(rutaExcel, dtype=str, encoding="latin1", sep=";")
        #print(df)

        return df
    
    except Exception as e:
        logging.error(f"Error al procesar el Excel: {e}")
        return None

def ProcesarCSVFiltrado(data, rutaExcel):
    """
    Lee un archivo CSV, aplica filtros específicos y verifica/crea columnas necesarias.
    
    Filtros:
    - NombreCampaña: Solo valores específicos de PORTABILIDAD
    - TeleNum: Solo números que empiecen por 3 y no estén vacíos
    - Verifica/crea campos: Operador, Ventana_cambio, Gestionado
    
    Returns:
        pd.DataFrame: DataFrame filtrado y procesado
    """
    logging.info("Lectura de CSV")
    print("Leyendo CSV, espere por favor...")
    
    # Lista de encodings y separadores a probar
    encodings = ["latin1", "utf-8", "cp1252", "ISO-8859-1"]
    separadores = [";", ","]
    
    # Campañas permitidas
    campanas_permitidas = [
        "PORTABILIDAD (POS)",
        "PORTABILIDAD (POS) MODELO",
        "PORTABILIDAD (POS) OPE2",
        "PORTABILIDAD (POS) OPE3",
        "PORTABILIDAD (POS) OPE4"
    ]
    
    for encoding in encodings:
        for sep in separadores:
            try:
                # Intentar leer el archivo
                df = pd.read_csv(rutaExcel, dtype=str, encoding=encoding, sep=sep)
                
                # Verificar si la columna "NombreCampaña" existe
                if "NombreCampaña" not in df.columns:
                    nombre_columnas = [col for col in df.columns if 'campaña' in col.lower() or 'campana' in col.lower()]
                    if nombre_columnas:
                        print(f"Renombrando columna {nombre_columnas[0]} a 'NombreCampaña'")
                        df.rename(columns={nombre_columnas[0]: "NombreCampaña"}, inplace=True)
                    else:
                        print("No se encontró la columna 'NombreCampaña'")
                        return pd.DataFrame()  # Retornar DataFrame vacío en lugar de None
                
                # Verificar si la columna "TeleNum" existe
                if "TeleNum" not in df.columns:
                    telefono_columnas = [col for col in df.columns if 'tele' in col.lower() or 'phone' in col.lower() or 'movil' in col.lower() or 'celular' in col.lower()]
                    if telefono_columnas:
                        print(f"Renombrando columna {telefono_columnas[0]} a 'TeleNum'")
                        df.rename(columns={telefono_columnas[0]: "TeleNum"}, inplace=True)
                    else:
                        print("No se encontró la columna 'TeleNum'")
                        return pd.DataFrame()  # Retornar DataFrame vacío en lugar de None
                
                # 1. Filtrar por NombreCampaña (solo las campañas permitidas)
                df = df[df["NombreCampaña"].isin(campanas_permitidas)]
                
                # 2. Filtrar TeleNum: números que empiecen por 3 y que no estén vacíos
                df = df[df["TeleNum"].notna()]  # No nulos
                df = df[df["TeleNum"] != ""]    # No vacíos
                df = df[df["TeleNum"].str.startswith("3")]  # Que empiecen por 3
                
                # 3. Verificar/crear el campo Operador
                if "Operador" not in df.columns:
                    print("Creando columna 'Operador'")
                    df["Operador"] = ""
                
                # 4. Verificar/crear el campo Ventana_cambio
                if "Ventana_cambio" not in df.columns:
                    print("Creando columna 'Ventana_cambio'")
                    df["Ventana_cambio"] = ""
                
                # 5. Verificar/crear el campo Gestionado
                if "Gestionado" not in df.columns:
                    print("La columna 'Gestionado' no se encontró.")
                    
                    gestionado_columns = [col for col in df.columns if 'gestionado' in col.lower()]
                    if gestionado_columns:
                        print(f"Columnas similares encontradas: {gestionado_columns}")
                        df.rename(columns={gestionado_columns[0]: "Gestionado"}, inplace=True)
                        print(f"Renombrando columna {gestionado_columns[0]} a 'Gestionado'")
                    else:
                        print("Creando columna 'Gestionado'")
                        df["Gestionado"] = ""
                
                print(f"CSV leído y filtrado correctamente con encoding {encoding} y separador {sep}")
                print(f"Registros después del filtrado: {len(df)}")
                return df
                
            except Exception as e:
                logging.warning(f"No se pudo leer con encoding {encoding} y separador {sep}: {e}")
    
    # Si llegamos aquí, no se pudo leer el archivo
    logging.error("No se pudo leer el archivo CSV con ninguna configuración")
    return pd.DataFrame()  # Retornar DataFrame vacío en lugar de None

def obtener_registros_pendientes(df):
    """
    Filtra los registros donde la columna 'Gestionado' está vacía, aplica filtros de campaña
    permitida y crea un archivo CSV con ellos.
    Solo usa el campo TeleNum para validar.
    
    Args:
        df (pd.DataFrame): DataFrame con los datos a filtrar
        
    Returns:
        tuple: (df_pendientes, ruta_salida) - DataFrame filtrado y ruta del archivo guardado
    """
    # Campañas permitidas
    campanas_permitidas = [
        "PORTABILIDAD (POS)",
        "PORTABILIDAD (POS) MODELO",
        "PORTABILIDAD (POS) OPE2",
        "PORTABILIDAD (POS) OPE3",
        "PORTABILIDAD (POS) OPE4"
    ]
    
    # Verificar si la columna "NombreCampaña" existe
    if "NombreCampaña" not in df.columns:
        print("La columna 'NombreCampaña' no existe en el CSV. Buscando columnas similares...")
        nombre_columnas = [col for col in df.columns if 'campaña' in col.lower() or 'campana' in col.lower()]
        if nombre_columnas:
            print(f"Renombrando columna {nombre_columnas[0]} a 'NombreCampaña'")
            df.rename(columns={nombre_columnas[0]: "NombreCampaña"}, inplace=True)
        else:
            print("Error: No se encontró la columna 'NombreCampaña'")
            return pd.DataFrame(), ""
    
    # Verificar si la columna "Gestionado" existe
    if "Gestionado" not in df.columns:
        print("La columna 'Gestionado' no existe en el CSV. Buscando columnas similares...")
        
        # Buscar columnas que puedan contener "Gestionado" en parte del nombre
        gestionado_columns = [col for col in df.columns if 'gestionado' in col.lower()]
        if gestionado_columns:
            print(f"Columnas similares encontradas: {gestionado_columns}")
            # Usar la primera columna similar encontrada
            df.rename(columns={gestionado_columns[0]: "Gestionado"}, inplace=True)
            print(f"Renombrando columna {gestionado_columns[0]} a 'Gestionado'")
        else:
            # Si no hay columnas similares, crear la columna
            print("Creando columna 'Gestionado'")
            df["Gestionado"] = ""
    
    # Verificar si la columna "TeleNum" existe
    if "TeleNum" not in df.columns:
        print("Error: No se encontró la columna 'TeleNum'. Verificando columnas disponibles...")
        print(f"Columnas disponibles: {', '.join(df.columns)}")
        telefono_columnas = [col for col in df.columns if 'tele' in col.lower() or 'phone' in col.lower() or 'movil' in col.lower() or 'celular' in col.lower()]
        if telefono_columnas:
            print(f"Renombrando columna {telefono_columnas[0]} a 'TeleNum'")
            df.rename(columns={telefono_columnas[0]: "TeleNum"}, inplace=True)
        else:
            print("No se pudo encontrar una columna adecuada para TeleNum")
            return pd.DataFrame(), ""
    
    print("Filtrando registros pendientes...")
    
    # 1. Filtrar por NombreCampaña (solo las campañas permitidas)
    df_filtrado = df[df["NombreCampaña"].isin(campanas_permitidas)]
    print(f"Registros después de filtrar campañas: {len(df_filtrado)}")
    
    # 2. Filtrar TeleNum: solo números que empiecen por 3 y no estén vacíos
    df_filtrado = df_filtrado[df_filtrado["TeleNum"].notna()]
    df_filtrado = df_filtrado[df_filtrado["TeleNum"] != ""]
    df_filtrado = df_filtrado[df_filtrado["TeleNum"].str.startswith("3")]
    print(f"Registros después de filtrar números: {len(df_filtrado)}")
    
    # 3. Filtrar registros pendientes (donde Gestionado está vacío)
    df_pendientes = df_filtrado[
        (df_filtrado["Gestionado"].isna()) | 
        (df_filtrado["Gestionado"].str.strip() == "")
    ]
    print(f"Registros pendientes finales: {len(df_pendientes)}")
    
    # 4. Verificar/crear campos adicionales necesarios
    if "Operador" not in df_pendientes.columns:
        print("Creando columna 'Operador'")
        df_pendientes["Operador"] = ""
    
    if "Ventana_cambio" not in df_pendientes.columns:
        print("Creando columna 'Ventana_cambio'")
        df_pendientes["Ventana_cambio"] = ""

    # Usar el nombre de archivo fijo requerido
    nombre_archivo = "dfpendientesCSV.csv"
    ruta_salida = os.path.join(r"C:\RPA_PORTA_COL_BASE", nombre_archivo)

    # Asegurar que el directorio existe
    os.makedirs(os.path.dirname(ruta_salida), exist_ok=True)

    # Guardar CSV asegurando compatibilidad de encoding
    try:
        df_pendientes.to_csv(ruta_salida, index=False, encoding="latin1", sep=";")
        print(f"Archivo guardado: {ruta_salida}")
    except Exception as e:
        print(f"Error al guardar con encoding latin1: {e}")
        try:
            df_pendientes.to_csv(ruta_salida, index=False, encoding="utf-8", sep=";")
            print(f"Archivo guardado con encoding utf-8: {ruta_salida}")
        except Exception as e2:
            print(f"Error al guardar también con utf-8: {e2}")
            print("Intentando guardado simplificado...")
            # Intentar una última opción con menos configuración
            try:
                df_pendientes.to_csv(ruta_salida, index=False)
                print(f"Archivo guardado con configuración por defecto: {ruta_salida}")
            except Exception as e3:
                print(f"Error al guardar con configuración por defecto: {e3}")

    return df_pendientes, ruta_salida

def LeerCSVFiltrado(ruta_csv):
    """
    Lee un archivo CSV filtrado existente.
    Esta función es más liviana que ProcesarCSVFiltrado ya que solo lee el CSV
    sin realizar ningún procesamiento adicional.
    
    Args:
        ruta_csv (str): Ruta al archivo CSV filtrado
        
    Returns:
        pandas.DataFrame: DataFrame con los datos del CSV filtrado
    
    Raises:
        Exception: Si hay problemas al leer el archivo
    """
    try:
        print(f"Leyendo CSV filtrado: {ruta_csv}")
        # Leer el CSV directamente
        df = pd.read_csv(ruta_csv)
        
        # Asegurar que Gestionado sea string para evitar problemas
        df["Gestionado"] = df["Gestionado"].fillna("").astype(str).str.strip()
        
        # Contar registros pendientes
        pendientes = len(df[df["Gestionado"] == ""])
        print(f"ℹRegistros pendientes: {pendientes}")
        
        return df
    except Exception as e:
        print(f"Error al leer CSV filtrado: {e}")
        raise Exception(f"Error al leer CSV filtrado: {e}")
# Leer Excel
# def ProcesarCSVFiltrado(data, rutaExcel):
#     try: 
#         logging.info("Lectura de Excel")
       
#         # Leer el archivo, asegurando que todas las columnas sean tratadas como texto
#         # Cargar el archivo usando openpyxl (más confiable que el motor por defecto)
#         print("Leyendo Excel espere porfavor ....")
#         #df = pd.read_excel(rutaExcel, dtype=str)  # Leer todo como texto para evitar errores de tipo de dato
#         try:
#             df = pd.read_csv(rutaExcel, dtype=str, encoding="latin1", sep=";")
#             df.columns = df.columns.str.strip()
#             return df
#         except Exception as e:
#             logging.error(f"Error al procesar el Excel: {e}")
          
#         try:
#             df = pd.read_csv(rutaExcel, dtype=str, encoding="utf-8", sep=";")
#             df.columns = df.columns.str.strip()
#             return df
#         except Exception as e:
#             logging.error(f"Error al procesar el Excel: {e}")
        
#         #df = pd.read_csv(rutaExcel, sep=",", encoding="latin1", dtype=str)

#         #print(df)

#         return df
    
#     except Exception as e:
#         logging.error(f"Error al procesar el Excel: {e}")
#         return None
# def ProcesarCSVFiltrado(data, rutaExcel):
#     logging.info("Lectura de CSV")
#     print("Leyendo CSV, espere por favor...")
    
#     # Lista de encodings y separadores a probar
#     encodings = ["latin1", "utf-8", "cp1252", "ISO-8859-1"]
#     separadores = [";", ","]
    
#     for encoding in encodings:
#         for sep in separadores:
#             try:
#                 # Intentar leer el archivo
#                 df = pd.read_csv(rutaExcel, dtype=str, encoding=encoding, sep=sep)
                
#                 # Verificar si la columna "Gestionado" existe
#                 if "Gestionado" not in df.columns:
#                     print(f"La columna 'Gestionado' no se encontró. Columnas disponibles: {', '.join(df.columns)}")
                    
#                     # Buscar columnas que puedan contener "Gestionado" en parte del nombre
#                     gestionado_columns = [col for col in df.columns if 'gestionado' in col.lower()]
#                     if gestionado_columns:
#                         print(f"Columnas similares encontradas: {gestionado_columns}")
#                         # Usar la primera columna similar encontrada
#                         df.rename(columns={gestionado_columns[0]: "Gestionado"}, inplace=True)
#                         print(f"Renombrando columna {gestionado_columns[0]} a 'Gestionado'")
#                     else:
#                         # Si no hay columnas similares, crear la columna
#                         print("Creando columna 'Gestionado'")
#                         df["Gestionado"] = ""
                
#                 print(f"CSV leído correctamente con encoding {encoding} y separador {sep}")
#                 return df
#             except Exception as e:
#                 logging.warning(f"No se pudo leer con encoding {encoding} y separador {sep}: {e}")
    
#     # Si llegamos aquí, no se pudo leer el archivo
#     logging.error("No se pudo leer el archivo CSV con ninguna configuración")
#     return None

# def obtener_registros_pendientes(df):
#     """
#     Filtra los registros donde la columna 'Gestionado' está vacía.
#     """
#     # Verificar si la columna "Gestionado" existe
#     if "Gestionado" not in df.columns:
#         print("La columna 'Gestionado' no existe en el CSV. Buscando columnas similares...")
        
#         # Buscar columnas que puedan contener "Gestionado" en parte del nombre
#         gestionado_columns = [col for col in df.columns if 'gestionado' in col.lower()]
#         if gestionado_columns:
#             print(f"Columnas similares encontradas: {gestionado_columns}")
#             # Usar la primera columna similar encontrada
#             df.rename(columns={gestionado_columns[0]: "Gestionado"}, inplace=True)
#             print(f"Renombrando columna {gestionado_columns[0]} a 'Gestionado'")
#         else:
#             # Si no hay columnas similares, crear la columna
#             print("Creando columna 'Gestionado'")
#             df["Gestionado"] = ""
    
#     print("Filtrando registros pendientes...")
#     df_pendientes = df[
#         (df["Gestionado"].isna() | (df["Gestionado"].str.strip() == "")) &
#         (~df["Celular1"].isna()) & (df["Celular1"].str.strip() != "")
#     ]

#     # Obtener fecha actual
#     nombre_archivo = f"dfpendientesCSV.csv"
#     ruta_salida = os.path.join(r"C:\RPA_PORTA_COL_BASE", nombre_archivo)

#     # Guardar CSV asegurando compatibilidad de encoding
#     try:
#         df_pendientes.to_csv(ruta_salida, index=False, encoding="latin1", sep=";")
#         print(f"Archivo guardado: {ruta_salida}")
#     except Exception as e:
#         print(f"Error al guardar con encoding latin1: {e}")
#         try:
#             df_pendientes.to_csv(ruta_salida, index=False, encoding="utf-8", sep=";")
#             print(f"Archivo guardado con encoding utf-8: {ruta_salida}")
#         except Exception as e2:
#             print(f"Error al guardar también con utf-8: {e2}")

#     return df_pendientes, ruta_salida

# def obtener_registros_pendientes(df):
#     """
#     Filtra los registros donde la columna 'Gestionado' está vacía.
#     """
#     df_pendientes = df[
#     (df["Gestionado"].isna() | (df["Gestionado"].str.strip() == "")) &
#     (~df["Celular1"].isna()) & (df["Celular1"].str.strip() != "")
#     ]

#     # Obtener fecha actual
#     nombre_archivo = f"dfpendientesCSV.csv"
#     ruta_salida = os.path.join(r"C:\RPA_PORTA_COL_BASE", nombre_archivo)

#     # Guardar CSV
#     df_pendientes.to_csv(ruta_salida, index=False, encoding="latin1", sep=";")
#     print(f"Archivo guardado: {ruta_salida}")

#     return df_pendientes,ruta_salida
