from dotenv import load_dotenv
import os

# Carga las variables de entorno
load_dotenv()

try: 
    workpath = os.path.dirname(os.path.abspath(__file__))
    #workpath2 = os.path.dirname(script_dir)
    print(f'Consecutivo {workpath}')

    directorio = os.path.join(workpath, 'Contabilidad FacturacionE')

    consecutivo = os.path.join(workpath, 'Facturacion Electronica IMPORTANTE', 'ConsecutivoFacturacionElectronica.txt')

    # Leer el número inicial desde el archivo de texto
    with open(consecutivo, 'r') as file:
        numero_inicial = int(file.read().strip())

    # Listar todos los archivos Excel en el directorio
    archivos_excel = [f for f in os.listdir(directorio) if f.endswith('.xlsx') or f.endswith('.xls')]

    # Variable para rastrear si se ha renombrado algún archivo
    renombrado = False
    ultimo_numero = numero_inicial

    # Renombrar los archivos Excel con números consecutivos a partir del número inicial
    for archivo in sorted(archivos_excel):
        # Obtener el nombre y la extensión del archivo
        nombre, extension = os.path.splitext(archivo)
        
        # Verificar si el nombre ya es un número (archivo ya procesado)
        if nombre.isdigit():
            continue 

        # Crear el nuevo nombre con el número consecutivo
        nuevo_nombre = f"{ultimo_numero}{extension}"
        
        # Obtener la ruta completa de origen y destino
        ruta_origen = os.path.join(directorio, archivo)
        ruta_destino = os.path.join(directorio, nuevo_nombre)
        
        # Renombrar el archivo
        os.rename(ruta_origen, ruta_destino)
        
        # Marcar que al menos un archivo ha sido renombrado
        renombrado = True
        
        # Actualizar el último número utilizado
        ultimo_numero += 1

    # Actualizar el archivo de texto con el último número utilizado solo si se renombró algún archivo
    if renombrado:
        with open(consecutivo, 'w') as file:
            file.write(str(ultimo_numero))
            print(f"Se han renombrado {ultimo_numero - numero_inicial} archivos con éxito.")

except Exception as e:
    print(f"Ocurrió un error: {e}")