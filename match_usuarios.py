import pandas as pd
import os

def procesar_match_usuarios():
    carpeta = "ARCHIVOS"
    
    # Definir las rutas exactas de entrada y salida
    ruta_archivo1 = os.path.join(carpeta, "User_Download_File.xlsx")
    ruta_archivo2 = os.path.join(carpeta, "users_logs_File.xlsx")
    ruta_archivo3 = os.path.join(carpeta, "Gemini_User_Reports_File.xlsx")
    ruta_salida = os.path.join(carpeta, "match_de_usuarios.xlsx")

    try:
        print("Cargando y filtrando columnas de los archivos...")

        # --- ARCHIVO 1 ---
        cols_df1 = [
            'First Name [Required]', 'Last Name [Required]', 
            'Email Address [Required]', 'Org Unit Path [Required]', 
            'Last Sign In [READ ONLY]'
        ]
        # usecols nos permite cargar solamente las columnas que necesitamos
        df1 = pd.read_excel(ruta_archivo1, usecols=cols_df1)

        # --- ARCHIVO 2 ---
        cols_df2 = [
            'Usuario', 'Classroom: Fecha del último uso', 
            'Gmail (Web): Fecha del último uso',
            'Correos electrónicos enviados [2026-04-03 GMT+0]', 
            'Correos electrónicos recibidos [2026-04-03 GMT+0]', 
            'Drive: fecha de la última actividad'
        ]
        df2 = pd.read_excel(ruta_archivo2, usecols=cols_df2)
        # Renombramos 'Usuario' para que coincida exactamente con el Archivo 1
        df2 = df2.rename(columns={'Usuario': 'Email Address [Required]'})

        # --- ARCHIVO 3 ---
        cols_df3 = [
            'Correo electrónico', 'Uso general', 'Días activos'
        ]
        df3 = pd.read_excel(ruta_archivo3, usecols=cols_df3)
        # Renombramos 'Correo electrónico' para que coincida exactamente con el Archivo 1
        df3 = df3.rename(columns={'Correo electrónico': 'Email Address [Required]'})

        print("Realizando el cruce de datos (Match)...")

        # Unimos Archivo 1 y Archivo 2
        # how='outer' asegura que si un usuario existe en un archivo pero no en el otro, no se pierda.
        resultado = pd.merge(df1, df2, on='Email Address [Required]', how='outer')

        # Unimos el resultado anterior con el Archivo 3
        resultado_final = pd.merge(resultado, df3, on='Email Address [Required]', how='outer')

        print("Guardando el archivo consolidado...")
        resultado_final.to_excel(ruta_salida, index=False)
        
        print("-" * 40)
        print("¡Proceso exitoso!")
        print(f"Total de registros en el archivo final: {len(resultado_final)}")
        print(f"Archivo creado en: {ruta_salida}")
        print("-" * 40)

    except FileNotFoundError as e:
        print(f"Error: No se encontró uno de los archivos en la carpeta '{carpeta}'.")
        print(f"Detalle técnico: {e}")
    except ValueError as e:
        print("Error: Revisa los nombres de las columnas. Uno de los Excel no tiene la columna exacta que pediste.")
        print(f"Detalle técnico: {e}")
    except Exception as e:
        print(f"Ocurrió un error inesperado: {e}")

def main():
    # Validamos primero si la carpeta ARCHIVOS existe
    if not os.path.exists("ARCHIVOS"):
        print("Error: No se encontró la carpeta 'ARCHIVOS'. Asegúrate de crearla y poner los Excel dentro.")
        return
        
    procesar_match_usuarios()

if __name__ == "__main__":
    main()