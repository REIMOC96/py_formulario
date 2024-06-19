import os
import numpy as np
from pdf2image import convert_from_path
import cv2
from openpyxl import Workbook

# Función para convertir un PDF a imágenes
def convertir_pdf_a_imagenes(pdf_path):
    try:
        return convert_from_path(pdf_path)
    except Exception as e:
        print(f"Error al convertir PDF a imágenes: {e}")
        return []

# Función para cargar la plantilla limpia desde un PDF
def cargar_plantilla(plantilla_path):
    paginas = convertir_pdf_a_imagenes(plantilla_path)
    if not paginas:
        raise FileNotFoundError(f"No se pudo cargar la plantilla desde {plantilla_path}")
    plantilla = np.array(paginas[0])  # Asumimos que la plantilla es la primera página
    plantilla_gray = cv2.cvtColor(plantilla, cv2.COLOR_BGR2GRAY)
    return plantilla_gray

# Función para seleccionar y guardar el área de búsqueda de respuestas en la plantilla
def seleccionar_area_respuestas(plantilla, r):
    # Extraer el área seleccionada
    area_respuestas = plantilla[int(r[1]):int(r[1]+r[3]), int(r[0]):int(r[0]+r[2])]
    return area_respuestas

# Función para comparar el área de respuestas con una imagen escaneada y detectar las respuestas marcadas
def comparar_area_respuestas(area_respuestas, imagen_escaneada, r):
    # Convertir la imagen escaneada a escala de grises
    gray = cv2.cvtColor(imagen_escaneada, cv2.COLOR_BGR2GRAY)
    
    # Recortar el área de la imagen escaneada donde se esperan las respuestas
    y, h, x, w = int(r[1]), int(r[3]), int(r[0]), int(r[2])
    area_escaneada = gray[y:y+h, x:x+w]
    
    # Aplicar umbralización adaptativa para obtener una imagen binaria
    _, binary_escaneada = cv2.threshold(area_escaneada, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)
    
    # Aquí deberías implementar la lógica para comparar el área de respuestas con la imagen escaneada
    # y detectar las respuestas marcadas. Por simplicidad, dejo un ejemplo simulado.
    respuestas_marcadas = [(100, 50), (200, 80)]  # Ejemplo simulado de respuestas marcadas
    
    return respuestas_marcadas

# Función para escribir los resultados en un archivo Excel
def escribir_resultados(datos_leidos):
    # Crear un nuevo libro de Excel
    wb = Workbook()
    
    # Crear una nueva hoja en el libro
    ws = wb.active
    ws.title = "Resultados"
    
    # Escribir encabezados
    encabezados = ["Pregunta"]
    for i in range(1, 9):
        encabezados.append(f"Respuesta {i}")
    ws.append(encabezados)
    
    # Escribir los datos de cada pregunta y su conteo de respuestas
    for pregunta_num in range(1, 34):  # Suponiendo que hay 33 preguntas
        pregunta = f"Pregunta {pregunta_num}"
        fila = [pregunta]
        
        if pregunta in datos_leidos:
            respuestas = datos_leidos[pregunta]
            for i in range(1, 9):  # 8 opciones de respuesta
                coordenada_respuesta = f"({i}, 0)"  # Se debe ajustar la coordenada adecuadamente
                if coordenada_respuesta in respuestas:
                    fila.append(respuestas[coordenada_respuesta])
                else:
                    fila.append(0)  # Si no hay respuestas, se pone 0
        
        ws.append(fila)
    
    # Guardar el archivo de Excel
    resultado_path = "resultados.xlsx"
    wb.save(resultado_path)
    print(f"Resultados guardados en {resultado_path}")

# Función principal
def main():
    # Obtener el directorio actual del script
    script_dir = os.path.dirname(os.path.abspath(__file__))
    
    # Definir el path de la plantilla limpia (en el mismo directorio que este script)
    plantilla_path = os.path.join(script_dir, 'plantilla_limpia.pdf')
    
    # Cargar la plantilla limpia
    try:
        plantilla = cargar_plantilla(plantilla_path)
    except Exception as e:
        print(f"Error al cargar la plantilla: {e}")
        return
    
    # Definir las coordenadas del área de respuestas
    r = (100, 50, 200, 150)  # Debes proporcionar las coordenadas correctas
    
    # Seleccionar el área de respuestas en la plantilla
    area_respuestas = seleccionar_area_respuestas(plantilla, r)
    
    # Directorio donde están los documentos escaneados
    scans_dir = os.path.join(script_dir, 'scans')
    
    # Inicializar diccionario para almacenar las respuestas por pregunta
    datos_leidos = {}
    
    # Iterar sobre los archivos en el directorio de scans
    for filename in os.listdir(scans_dir):
        if filename.endswith(".pdf"):  # Asegurarse de procesar solo archivos PDF
            # Construir el path completo de la imagen escaneada
            imagen_escaneada_path = os.path.join(scans_dir, filename)
            
            # Convertir el PDF escaneado a imagen
            imagenes = convertir_pdf_a_imagenes(imagen_escaneada_path)
            if not imagenes:
                print(f"No se pudo convertir {filename} a imágenes.")
                continue
            
            # Tomar solo la primera página como imagen escaneada
            imagen_escaneada = np.array(imagenes[0])
            
            # Comparar el área de respuestas con la imagen escaneada y detectar respuestas marcadas
            respuestas_marcadas = comparar_area_respuestas(area_respuestas, imagen_escaneada, r)
            
            # Si no se encontraron respuestas marcadas, continuar con el siguiente archivo
            if not respuestas_marcadas:
                print(f"No se encontraron respuestas marcadas en {filename}.")
                continue
            
            # Aquí deberías implementar la lógica para leer las respuestas marcadas y contarlas
            # Supongamos que las respuestas se identifican por coordenadas, ejemplo:
            # respuestas_marcadas = [(100, 50), (200, 80), (150, 60), (250, 90)]
            
            # En este ejemplo, se simula el conteo de respuestas para cada pregunta
            preguntas = {'Pregunta1': {"(1, 0)": 2, "(2, 0)": 5, "(3, 0)": 3},
                         'Pregunta2': {"(1, 0)": 4, "(2, 0)": 1, "(4, 0)": 2}}
            
            # Actualizar el diccionario datos_leidos con las respuestas marcadas
            for pregunta, respuestas in preguntas.items():
                if pregunta not in datos_leidos:
                    datos_leidos[pregunta] = {}
                for respuesta, conteo in respuestas.items():
                    datos_leidos[pregunta][respuesta] = conteo
    
    # Reorganizar datos_leidos para mostrar conteo individual de respuestas
    datos_con_conteo_individual = {}
    for pregunta, respuestas in datos_leidos.items():
        datos_con_conteo_individual[pregunta] = {}
        for respuesta, conteo in respuestas.items():
            coordenadas_respuesta = respuesta
            conteo_individual = conteo
            datos_con_conteo_individual[pregunta][coordenadas_respuesta] = conteo_individual
    
    # Escribir los resultados en un archivo Excel
    escribir_resultados(datos_con_conteo_individual)

if __name__ == "__main__":
    main()
