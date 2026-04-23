import pandas as pd
import os

def evaluar_actividad(valor, texto_exclusion):
    """
    Evalúa si el valor indica actividad.
    """
    if pd.isna(valor) or str(valor).strip().lower() == 'nan':
        return 'No'
        
    valor_str = str(valor).strip()
    
    if valor_str.lower() == texto_exclusion.lower():
        return 'No'
        
    return 'Sí'

def evaluar_dias_activos(valor):
    """
    Evalúa específicamente la columna original de 'Días activos'.
    (El criterio se mantiene: distinto de 0).
    """
    if pd.isna(valor):
        return 'No'
    
    valor_str = str(valor).strip().split('.')[0] 
    
    if valor_str == '0': 
        return 'No'
        
    return 'Sí'

def generar_reporte_uso():
    carpeta = "ARCHIVOS"
    ruta_entrada = os.path.join(carpeta, "usuarios_separados.xlsx")
    ruta_salida = os.path.join(carpeta, "uso_workspace.xlsx")
    
    hojas_esperadas = ['Docentes', 'Estudiantes', 'Asistentes Educación', 'Otros']

    try:
        if not os.path.exists(ruta_entrada):
            print(f"Error: No se encuentra el archivo '{ruta_entrada}'.")
            return

        print("Cargando el archivo usuarios_separados.xlsx...")
        dict_dfs = pd.read_excel(ruta_entrada, sheet_name=None)
        
        print("Analizando uso de aplicaciones y generando resúmenes...")
        
        with pd.ExcelWriter(ruta_salida, engine='openpyxl') as writer:
            
            fila_inicio_resumen = 0 
            
            for hoja in hojas_esperadas:
                if hoja in dict_dfs:
                    df = dict_dfs[hoja].copy()
                    total_usuarios = len(df)
                    
                    # Inicializamos columnas
                    df['Uso Cuenta (Sí/No)'] = 'No'
                    df['Uso Classroom (Sí/No)'] = 'No'
                    df['Uso Gmail (Sí/No)'] = 'No'
                    df['Uso Drive (Sí/No)'] = 'No'
                    df['Uso de Gemini (Sí/No)'] = 'No' # Aquí le damos el nuevo nombre
                    
                    # Evaluaciones
                    if 'Last Sign In [READ ONLY]' in df.columns:
                        df['Uso Cuenta (Sí/No)'] = df['Last Sign In [READ ONLY]'].apply(lambda x: evaluar_actividad(x, 'Never logged in'))
                    if 'Classroom: Fecha del último uso' in df.columns:
                        df['Uso Classroom (Sí/No)'] = df['Classroom: Fecha del último uso'].apply(lambda x: evaluar_actividad(x, 'Nunca'))
                    if 'Gmail (Web): Fecha del último uso' in df.columns:
                        df['Uso Gmail (Sí/No)'] = df['Gmail (Web): Fecha del último uso'].apply(lambda x: evaluar_actividad(x, 'No en los últimos 30 días'))
                    if 'Drive: fecha de la última actividad' in df.columns:
                        df['Uso Drive (Sí/No)'] = df['Drive: fecha de la última actividad'].apply(lambda x: evaluar_actividad(x, 'Nunca'))
                        
                    # Validación de uso de Gemini leyendo la columna "Días activos" original
                    if 'Días activos' in df.columns:
                        df['Uso de Gemini (Sí/No)'] = df['Días activos'].apply(evaluar_dias_activos)

                    # Guardamos los datos detallados de la hoja
                    df.to_excel(writer, sheet_name=hoja, index=False)
                    print(f" -> Datos de '{hoja}' procesados.")
                    
                    # --- CREACIÓN DEL CUADRO RESUMEN PARA ESTA HOJA ESPECÍFICA ---
                    if total_usuarios > 0:
                        con_acceso = len(df[df['Uso Cuenta (Sí/No)'] == 'Sí'])
                        sin_acceso = len(df[df['Uso Cuenta (Sí/No)'] == 'No'])
                        uso_classroom = len(df[df['Uso Classroom (Sí/No)'] == 'Sí'])
                        uso_gmail = len(df[df['Uso Gmail (Sí/No)'] == 'Sí'])
                        uso_drive = len(df[df['Uso Drive (Sí/No)'] == 'Sí'])
                        uso_gemini = len(df[df['Uso de Gemini (Sí/No)'] == 'Sí']) 
                        
                        def calc_pct(valor):
                            return f"{(valor / total_usuarios) * 100:.2f}%"
                        
                        # Armamos la estructura visual del cuadro con el nuevo texto explicativo
                        datos_resumen = [
                            [f"--- RESUMEN: {hoja.upper()} (Total usuarios: {total_usuarios}) ---", "", ""],
                            ["Con acceso (Distinto a Never logged in)", con_acceso, calc_pct(con_acceso)],
                            ["Sin acceso (Celda vacía / Never logged in)", sin_acceso, calc_pct(sin_acceso)],
                            ["Uso Classroom (Distinto a Nunca)", uso_classroom, calc_pct(uso_classroom)],
                            ["Uso Gmail (Distinto a No en los últimos 30 días)", uso_gmail, calc_pct(uso_gmail)],
                            ["Uso Drive (Distinto a Nunca)", uso_drive, calc_pct(uso_drive)],
                            ["Uso de Gemini (Días Activos distinto a 0)", uso_gemini, calc_pct(uso_gemini)], 
                            ["", "", ""] # Fila en blanco para separar el siguiente cuadro
                        ]
                        
                        df_resumen = pd.DataFrame(datos_resumen)
                        
                        df_resumen.to_excel(writer, sheet_name='Resúmenes por Tipo', startrow=fila_inicio_resumen, index=False, header=False)
                        
                        fila_inicio_resumen += len(datos_resumen)
                        
                else:
                    print(f" -> Advertencia: La hoja '{hoja}' no se encontró.")

        print("-" * 50)
        print("¡Proceso de auditoría exitoso!")
        print(f"El reporte final se guardó en: {ruta_salida}")
        print("-" * 50)

    except Exception as e:
        print(f"Ocurrió un error inesperado: {e}")

def main():
    generar_reporte_uso()

if __name__ == "__main__":
    main()