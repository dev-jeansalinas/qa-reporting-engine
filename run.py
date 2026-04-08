import os
import json
import datetime
from src.generator import generar_excel_snva
from src.utils import cargar_casos_prueba, generar_word_snva

# GESTIÓN DE RUTAS
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_PATH = os.path.join(BASE_DIR, "data")
OUTPUT_PATH = os.path.join(BASE_DIR, "output")
LOGO_FILE = os.path.join(BASE_DIR, "assets", "logo.png")
PLANTILLA_WORD = os.path.join(BASE_DIR, "assets", "plantilla_sinova.docx")

def main():
    # Crear la carpeta de salida automáticamente si no existe
    if not os.path.exists(OUTPUT_PATH):
        os.makedirs(OUTPUT_PATH)    
        print(f"[+] Carpeta creada: {OUTPUT_PATH}")

    # CARGA DE CONFIGURACIÓN
    ruta_config = os.path.join(DATA_PATH, "config_proyecto.json")
    try:
        with open(ruta_config, 'r', encoding='utf-8') as f:
            config_proyecto = json.load(f)
    except FileNotFoundError:
        print(f"[-] Error: No se encontró {ruta_config}")
        return

    # CARGA DE CASOS DE PRUEBA
    ruta_casos = os.path.join(DATA_PATH, "casos.json")
    casos = cargar_casos_prueba(ruta_casos)

    if not casos:
        print("[-] Abortando: No hay casos de prueba para procesar.")
        return

    # DEFINICIÓN DEL ARCHIVO DE SALIDA
    nombre_archivo = f"SNVA_Test_Cases_{config_proyecto['modulo'].replace(' ', '_')}.xlsx"
    ruta_final = os.path.join(OUTPUT_PATH, nombre_archivo)

    # EJECUCIÓN DEL GENERADOR EXCEL
    print(f"[*] Generando Excel para: {config_proyecto['cliente']}...")
    
    # Pasamos los datos usando ** (unpacking) y las rutas de archivos
    generar_excel_snva(
        **config_proyecto, 
        casos_prueba=casos, 
        logo_path=LOGO_FILE, 
        ruta_final=ruta_final
    )

    # EJECUCIÓN DEL GENERADOR WORD
    nombre_archivo_word = f"SNVA_Evidencias_{config_proyecto['modulo'].replace(' ', '_')}.docx"
    ruta_word = os.path.join(OUTPUT_PATH, nombre_archivo_word)

    print(f"[*] Generando Word para: {config_proyecto['cliente']}...")
    
    generar_word_snva(
        datos_hu=casos, 
        info_proyecto=config_proyecto, 
        ruta_plantilla=PLANTILLA_WORD, 
        ruta_salida=ruta_word
    )

    print("-" * 30)
    print(f"[SUCCESS] Excel generado con éxito!")
    print(f"[EXCEL] {ruta_final}")
    print(f"[WORD]  {ruta_word}")
    print("-" * 30)

if __name__ == "__main__":
    main()