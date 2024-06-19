import os
import cv2
import numpy as np
from pdf2image import convert_from_path

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
def seleccionar_area_respuestas(plantilla):
    # Mostrar la plantilla y permitir al usuario dibujar un rectángulo
    cv2.namedWindow("Seleccione el área de respuestas", cv2.WINDOW_NORMAL)
    
    # Ajustar el tamaño de la ventana para que la imagen se ajuste a la pantalla
    cv2.resizeWindow("Seleccione el área de respuestas", 800, 600)
    
    r = cv2.selectROI("Seleccione el área de respuestas", plantilla)
    cv2.destroyWindow("Seleccione el área de respuestas")
    
    # Extraer el área seleccionada
    area_respuestas = plantilla[int(r[1]):int(r[1]+r[3]), int(r[0]):int(r[0]+r[2])]
    return area_respuestas, r

# Función principal para seleccionar y guardar el área de respuestas
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
    
    # Seleccionar el área de respuestas en la plantilla
    area_respuestas, r = seleccionar_area_respuestas(plantilla)
    
    # Guardar el área seleccionada como una imagen para referencia
    cv2.imwrite("area_respuestas.jpg", area_respuestas)

if __name__ == "__main__":
    main()
