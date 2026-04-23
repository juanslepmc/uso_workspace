import pandas as pd
import os
import re

def extraer_datos_ruta(ruta):
    """
    Toma la ruta de la unidad organizativa y extrae: Comuna, RBD, Establecimiento, Tipo
    """
    if pd.isna(ruta):
        return pd.Series([None, None, None, None])
    
    ruta_str = str(ruta)
    partes = ruta_str.strip('/').split('/')
    
    comuna, rbd, estab, tipo = None, None, None, None
    
    # Extracción de Comuna y Establecimiento/RBD
    if len(partes) >= 3:
        comuna = partes[1]
        
        tercera_parte = partes[2].strip()
        match = re.match(r'^(\d+)\s+(.*)$', tercera_parte)
        if match:
            rbd = int(match.group(1))
            estab = match.group(2)
        else:
            estab = tercera_parte
            
    # Extracción del Tipo
    if len(partes) >= 4:
        tipo = partes[3].strip() 
        
        # --- REGLAS DE NEGOCIO PARA "TIPO" ---
        # 1. Corrección de error de tipeo
        if tipo == 'Esudiantes':
            tipo = 'Estudiantes'
            
        # 2. Caso especial PIE
        elif tipo == '299 Programa de Integración (PIE) opción 4':
            tipo = 'Estudiantes'
        
    return pd.Series([comuna, rbd, estab, tipo])

def clasificar_usuarios_y_egresados(row):
    """
    Revisa si el tipo es Egresados y los mueve a Estudiantes o Docentes
    basándose en su correo electrónico.
    """
    tipo = row['Tipo']
    
    # Solo entramos a clasificar si el tipo es exactamente "Egresados"
    if tipo != 'Egresados':
        return tipo
        
    if pd.isna(row['Email Address [Required]']):
        return 'Otros'
        
    email = str(row['Email Address [Required]']).strip()
    
    # --- REDIRECCIÓN DE EGRESADOS ---
    
    # 1. Si tiene 1 letra/número antes del punto -> Va a la hoja Estudiantes
    if re.match(r'^[A-Za-z0-9]\.', email):
        return 'Estudiantes'
        
    # 2. Si tiene 2 o más letras/números antes del punto -> Va a la hoja Docentes
    elif re.match(r'^[A-Za-z0-9]{2,}\.', email):
        return 'Docentes'
        
    # 3. Si es Egresado pero no cumple el formato de correo, se va a Otros
    else:
        return 'Otros'

def separar_usuarios_por_tipo():
    carpeta = "ARCHIVOS"
    ruta_entrada = os.path.join(carpeta, "match_de_usuarios.xlsx")
    ruta_salida = os.path.join(carpeta, "usuarios_separados.xlsx")
    
    # Mantenemos solo las 3 hojas principales
    tipos_esperados = ['Docentes', 'Estudiantes', 'Asistentes Educación']

    try:
        if not os.path.exists(ruta_entrada):
            print(f"Error: No se encuentra el archivo '{ruta_entrada}'.")
            return

        print("Cargando el archivo match_de_usuarios.xlsx...")
        df = pd.read_excel(ruta_entrada)

        # Validaciones de columnas
        for col in ['Org Unit Path [Required]', 'Email Address [Required]']:
            if col not in df.columns:
                print(f"Error: La columna '{col}' no existe en el archivo.")
                return

        print("Procesando rutas y aplicando reglas de negocio...")
        
        # 1. Extraemos los datos de la ruta
        df[['Comuna', 'RBD', 'Establecimiento', 'Tipo']] = df['Org Unit Path [Required]'].apply(extraer_datos_ruta)

        print("Redireccionando Egresados a sus hojas correspondientes...")
        # 2. Aplicamos la lógica de correos para mover Egresados a Estudiantes o Docentes
        df['Tipo'] = df.apply(clasificar_usuarios_y_egresados, axis=1)

        print("Generando archivo Excel con pestañas...")
        with pd.ExcelWriter(ruta_salida, engine='openpyxl') as writer:
            
            # Hojas principales (Docentes, Estudiantes, Asistentes Ed.)
            for tipo in tipos_esperados:
                df_filtrado = df[df['Tipo'] == tipo]
                df_filtrado.to_excel(writer, sheet_name=tipo, index=False)
                print(f" -> Pestaña '{tipo}' finalizada con {len(df_filtrado)} registros.")
            
            # Hoja Otros
            df_otros = df[~df['Tipo'].isin(tipos_esperados) | df['Tipo'].isna()]
            df_otros.to_excel(writer, sheet_name='Otros', index=False)
            print(f" -> Pestaña 'Otros' finalizada con {len(df_otros)} registros.")

        print("-" * 50)
        print("¡Proceso exitoso!")
        print(f"Archivo guardado en: {ruta_salida}")
        print("-" * 50)

    except Exception as e:
        print(f"Ocurrió un error inesperado: {e}")

def main():
    separar_usuarios_por_tipo()

if __name__ == "__main__":
    main()