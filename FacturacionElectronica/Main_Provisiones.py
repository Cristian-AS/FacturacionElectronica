import os
import time
import datetime
import sys
from dotenv import load_dotenv

load_dotenv()
#log_path = os.getenv('Raiz_Proyecto')

log_path = os.path.dirname(os.path.abspath(__file__))
print(f'Main Provisiones {log_path}')

def CrearConsoleLog(Path):
    # Definir la ruta de la carpeta de log
    log_path = os.path.join(Path, 'Console_log')

    # Crear la carpeta de log si no existe
    if not os.path.exists(log_path):
        os.mkdir(log_path)
        print(f"Se ha creado la carpeta de log en {log_path}")

    return log_path

original_stdout = sys.stdout

# Obtener la fecha y hora actual
now = datetime.datetime.now()

# Formatear la fecha y la hora como una cadena
timestamp = now.strftime("%Y-%m-%d_%H-%M-%S")

# Definir la ruta del archivo de console log
log_path = CrearConsoleLog(log_path)


with open(os.path.join(log_path, f'console_log_PROVISIONES-{timestamp}.txt'), 'a') as f:
    sys.stdout = f # Cambiar la salida estándar al archivo que acabamos de abrir
    
    # Cambiar el directorio de trabajo al directorio del script
    os.chdir(os.path.join(os.path.dirname(os.path.abspath(__file__)), 'BOTS'))

    exec(open('PlanoProvisiones.py', encoding='utf-8').read())

    time.sleep(15)
    
    sys.stdout = original_stdout # Restaurar la salida estándar original