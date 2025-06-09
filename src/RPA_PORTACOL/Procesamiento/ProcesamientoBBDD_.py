import pandas as pd
import logging
import os

def ProcesarExcel(data, ruta):
    try:
        logging.info("Lectura de archivo")

        print("Leyendo archivo por favor espere...")

        # Verificar si es CSV o Excel
        extension = os.path.splitext(ruta)[1]

        if extension.lower() == ".csv":
            df = pd.read_csv(ruta, dtype=str)
        else:
            df = pd.read_excel(ruta, dtype=str)

        return df

    except Exception as e:
        logging.error(f"Error al procesar el archivo: {e}")
        return None

def obtener_registros_pendientes(df):
    try:
        df_pendientes = df[
            (df["Gestionado"].isna() | (df["Gestionado"] == "")) &
            (df["Celular1"].notna() & (df["Celular1"] != "")) &
            (df["Celular1"].str.len() == 10) &
            (df["Celular1"].str.startswith("3"))
        ]
        
        print(f"Se filtraron {len(df_pendientes)} registros válidos con 'Celular1' correcto")
        return df_pendientes

    except Exception as e:
        logging.error(f"Error al filtrar el archivo: {e}")
        return None

def crear_excel_filtrado(df_pendientes, ruta_filtrado):
    try:
        df_pendientes.to_excel(ruta_filtrado, index=False)
        print(f"Archivo filtrado creado: {ruta_filtrado}")
    except Exception as e:
        logging.error(f"Error al crear el Excel filtrado: {e}")
