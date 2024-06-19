import os
import numpy as np
from pdf2image import convert_from_path
import cv2

# Función para convertir un PDF a imágenes
def convertir_pdf_a_imagenes(pdf_path):
    try:
        return convert_from_path(pdf_path)
    except Exception as e:
        print(f"Error al convertir PDF a imágenes: {e}")
        return []

# Función para cargar el área de respuestas desde una imagen
def cargar_area_respuestas(area_respuestas_path):
    area_respuestas = cv2.imread(area_respuestas_path, cv2.IMREAD_GRAYSCALE)
    if area_respuestas is None:
        raise FileNotFoundError(f"No se pudo cargar el área de respuestas desde {area_respuestas_path}")
    return area_respuestas

# Función para comparar el área de respuestas con una imagen escaneada y obtener coordenadas de coincidencias
def encontrar_coincidencias(area_respuestas, imagen_escaneada):
    # Convertir la imagen escaneada a escala de grises
    gray = cv2.cvtColor(imagen_escaneada, cv2.COLOR_BGR2GRAY)
    
    # Aplicar umbralización adaptativa para obtener una imagen binaria
    _, binary_escaneada = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)
    
    # Aplicar detección de coincidencia de plantilla con el área de respuestas de la imagen escaneada
    resultado = cv2.matchTemplate(binary_escaneada, area_respuestas, cv2.TM_CCOEFF_NORMED)
    
    # Definir umbral de confianza para la detección de coincidencia
    threshold = 0.8
    loc = np.where(resultado >= threshold)
    
    # Obtener las coordenadas de las coincidencias encontradas
    coincidencias = []
    for pt in zip(*loc[::-1]):
        coincidencias.append(pt)
    
    return coincidencias

# Función para guardar las coordenadas en un archivo de texto
def guardar_coordenadas(coordenadas, output_file):
    with open(output_file, 'w') as file:
        for coord in coordenadas:
            file.write(f"{coord[0]}, {coord[1]}\n")
    print(f"Coordenadas guardadas en {output_file}")

# Función principal
def main():
    # Obtener el directorio actual del script
    script_dir = os.path.dirname(os.path.abspath(__file__))
    
    # Definir el path del área de respuestas (imagen en formato JPG)
    area_respuestas_path = os.path.join(script_dir, 'area_respuestas.jpg')
    
    # Cargar el área de respuestas
    try:
        area_respuestas = cargar_area_respuestas(area_respuestas_path)
    except Exception as e:
        print(f"Error al cargar el área de respuestas: {e}")
        return
    
    # Directorio donde están los documentos escaneados
    scans_dir = os.path.join(script_dir, 'scans')
    
    # Crear una lista para almacenar todas las coordenadas encontradas
    todas_coordenadas = []
    
    # Iterar sobre los archivos en el directorio de scans
    for filename in os.listdir(scans_dir):
        if filename.endswith(".jpg"):  # Asegurarse de procesar solo archivos JPG
            # Construir el path completo de la imagen escaneada
            imagen_escaneada_path = os.path.join(scans_dir, filename)
            
            # Cargar la imagen escaneada
            imagen_escaneada = cv2.imread(imagen_escaneada_path)
            if imagen_escaneada is None:
                print(f"No se pudo cargar {filename}.")
                continue
            
            # Encontrar coincidencias del área de respuestas en la imagen escaneada
            coincidencias = encontrar_coincidencias(area_respuestas, imagen_escaneada)
            
            # Guardar todas las coordenadas encontradas
            todas_coordenadas.extend(coincidencias)
    
    # Guardar las coordenadas en un archivo de texto
    output_file = os.path.join(script_dir, 'areas.txt')
    guardar_coordenadas(todas_coordenadas, output_file)

if __name__ == "__main__":
    main()
