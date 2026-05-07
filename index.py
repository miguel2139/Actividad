import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog

# 1. Configuración de la ventana oculta de Tkinter
root = tk.Tk()
root.withdraw()  # Oculta la ventana principal de Tkinter para que solo salga la de elegir archivos

try:
    print("Selecciona el PRIMER archivo (2026.xlsm)...")
    # Abrir ventana para el primer archivo
    ruta_archivo1 = filedialog.askopenfilename(
        title="Selecciona el PRIMER archivo (ej. 2026.xlsm)",
        filetypes=[("Archivos de Excel", "*.xlsx *.xlsm")]
    )

    print("Selecciona el SEGUNDO archivo...")
    # Abrir ventana para el segundo archivo
    ruta_archivo2 = filedialog.askopenfilename(
        title="Selecciona el SEGUNDO archivo",
        filetypes=[("Archivos de Excel", "*.xlsx *.xlsm")]
    )

    # Validar si se seleccionaron ambos archivos
    if ruta_archivo1 and ruta_archivo2:
        print("Cargando archivos...")
        df1 = pd.read_excel(ruta_archivo1, engine='openpyxl')
        df2 = pd.read_excel(ruta_archivo2, engine='openpyxl')

        # Normalizar nombres de columnas
        df1.columns = [c.upper().strip() for c in df1.columns]
        df2.columns = [c.upper().strip() for c in df2.columns]

        # Unificar
        df_unificado = pd.concat([df1, df2], ignore_index=True)

        # Lógica de autocompletado
        for col in ['NOMBRE', 'APELLIDO']:
            if col in df_unificado.columns:
                mapping = df_unificado.dropna(subset=[col])
                mapping = mapping[mapping[col] != 'Sin información']
                mapping_dict = mapping.drop_duplicates('CCOPERADOR').set_index('CCOPERADOR')[col].to_dict()
                df_unificado[col] = df_unificado[col].fillna(df_unificado['CCOPERADOR'].map(mapping_dict))
                df_unificado[col] = df_unificado[col].fillna('Sin información')

        # Guardar resultado
        nombre_salida = 'resultado_unificado.xlsx'
        df_unificado.to_excel(nombre_salida, index=False)
        
        print("-" * 30)
        print(f"¡Éxito! Archivo guardado como: {os.path.abspath(nombre_salida)}")
        print("-" * 30)
        print(df_unificado.head())
    else:
        print("Operación cancelada: No seleccionaste los archivos.")

except Exception as e:
    print(f"Error: {e}")

finally:
    root.destroy() # Cierra el proceso de la ventana)