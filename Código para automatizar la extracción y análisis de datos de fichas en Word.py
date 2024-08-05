import os
import pandas as pd
from docx import Document

# Función para extraer datos del documento de Word
def extraer_datos_word(file_path):
    # Cargar el documento de Word
    doc = Document(file_path)
    
    # Nombre del archivo sin la extensión
    captacions = os.path.splitext(os.path.basename(file_path))[0]
    
    # Inicializamos el diccionario con todas las columnas y sus valores vacíos
    data = { 
        "Captacions": captacions,  # Asignamos el nombre del archivo a la columna "Captacions"
        "Tipus": "", "Proveïdor": "", "Potència": "",
        "Consum": "", "MCA": "", "Cabal": "", "Impulsió": "", 
        "Longitud": "", "Variador": "", "Instal·lació": ""
    }
    
    # Mapeo de términos de búsqueda a claves del diccionario
    mapeo = {
        "TIPO": "Tipus",
        "PROVEEDOR": "Proveïdor",
        "POTENCIA": "Potència",
        "CONSUM": "Consum",
        "MCA": "MCA",
        "CAUDAL": "Cabal",
        "IMPULSIÓN": "Impulsió",
        "LONGITUD": "Longitud",
        "VARIADOR": "Variador",
        "INSTALACIÓN": "Instal·lació"
    }
    
    # Leer el contenido del documento y mapear los valores
    for para in doc.paragraphs:
        text = para.text.strip()
        for key, column in mapeo.items():
            if key in text:
                data[column] = text.split("\t")[-1]
    
    return data

# Ruta a la carpeta con los archivos de Word
folder_path = r'C:\Users\DGutierrez\Desktop\Bombes'

# Lista para almacenar los datos extraídos
datos_pozo_list = []

# Iterar sobre todos los archivos .docx en la carpeta
for filename in os.listdir(folder_path):
    if filename.endswith(".docx"):
        file_path = os.path.join(folder_path, filename)
        try:
            datos_pozo = extraer_datos_word(file_path)
            datos_pozo_list.append(datos_pozo)
        except Exception as e:
            print(f"Error al procesar el archivo {file_path}: {e}")

# Crear un DataFrame con los datos extraídos
df = pd.DataFrame(datos_pozo_list)

# Ruta para guardar el archivo Excel
file_path_excel = r'rutadelarchivo.xlsx'

# Guardar el DataFrame en un archivo Excel
try:
    df.to_excel(file_path_excel, index=False)
    print(f"El archivo Excel se ha guardado en: {file_path_excel}")
except Exception as e:
    print(f"Error al guardar el archivo Excel: {e}")
