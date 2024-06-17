import os
import cv2
import numpy as np
import pytesseract
from pdf2image import convert_from_path
from openpyxl import Workbook

# Función para convertir un PDF a imágenes
def convertir_pdf_a_imagenes(pdf_path):
    return convert_from_path(pdf_path)

# Función para cargar la plantilla limpia desde un PDF
def cargar_plantilla(plantilla_path):
    paginas = convertir_pdf_a_imagenes(plantilla_path)
    plantilla = np.array(paginas[0])  # Asumimos que la plantilla es la primera página
    plantilla_gray = cv2.cvtColor(plantilla, cv2.COLOR_BGR2GRAY)
    return plantilla_gray

# Función para comparar la plantilla con una imagen escaneada y detectar respuestas marcadas
def comparar_con_plantilla(plantilla, imagen_escaneada):
    # Convertir la imagen escaneada a escala de grises
    gray = cv2.cvtColor(imagen_escaneada, cv2.COLOR_BGR2GRAY)
    
    # Aplicar umbralización adaptativa para obtener una imagen binaria
    _, binary = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)
    
    # Aplicar detección de coincidencia de plantilla con la imagen escaneada
    resultado = cv2.matchTemplate(binary, plantilla, cv2.TM_CCOEFF_NORMED)
    
    # Definir umbral de confianza para la detección de coincidencia
    threshold = 0.8
    loc = np.where(resultado >= threshold)
    
    # Obtener las coordenadas de las coincidencias encontradas
    respuestas_marcadas = []
    for pt in zip(*loc[::-1]):
        respuestas_marcadas.append(pt)
    
    return respuestas_marcadas

# Función para promediar la distancia vertical entre preguntas
def obtener_distancia_media_entre_preguntas(preguntas_coords):
    distancias = [preguntas_coords[i+1][1] - preguntas_coords[i][1] for i in range(len(preguntas_coords) - 1)]
    return sum(distancias) / len(distancias)

# Función principal
def main():
    # Obtener el directorio actual del script
    script_dir = os.path.dirname(__file__)
    
    # Definir el path de la plantilla limpia (en el mismo directorio que este script)
    plantilla_path = os.path.join(script_dir, 'plantilla_limpia.pdf')
    
    # Cargar la plantilla limpia
    plantilla = cargar_plantilla(plantilla_path)
    
    # Directorio donde están los documentos escaneados
    scans_dir = os.path.join(script_dir, 'scans')
    
    # Inicializar diccionario para contar respuestas
    respuestas_totales = {}

    # Iterar sobre los archivos en el directorio de scans
    for filename in os.listdir(scans_dir):
        if filename.endswith(".pdf"):  # Asegurarse de procesar solo archivos PDF
            # Construir el path completo de la imagen escaneada
            imagen_escaneada_path = os.path.join(scans_dir, filename)
            
            # Convertir el PDF escaneado a imágenes
            imagenes_escaneadas = convertir_pdf_a_imagenes(imagen_escaneada_path)
            
            # Procesar cada página del PDF
            for pagina in imagenes_escaneadas:
                imagen_escaneada = np.array(pagina)
                
                # Comparar la plantilla con la imagen escaneada y detectar respuestas marcadas
                respuestas_marcadas = comparar_con_plantilla(plantilla, imagen_escaneada)
                
                # Aquí puedes procesar `respuestas_marcadas` para contar respuestas y agregarlas a `respuestas_totales`
                # ...

    # Crear un libro de Excel y agregar los resultados
    wb = Workbook()
    ws = wb.active
    ws.title = "Resultados"
    
    # Escribir los encabezados
    encabezados = ["Pregunta", "1", "2", "3", "4", "5", "6", "7", "8"]
    ws.append(encabezados)
    
    # Escribir los resultados en el archivo de Excel
    for pregunta, respuestas in respuestas_totales.items():
        fila = [pregunta] + [respuestas.get(str(i), 0) for i in range(1, 9)]
        ws.append(fila)
    
    # Guardar el archivo de Excel
    wb.save("resultados.xlsx")

if __name__ == "__main__":
    main()
