import os
import numpy as np
from pdf2image import convert_from_path
import cv2
from openpyxl import Workbook
import matplotlib.pyplot as plt

# Función para convertir un PDF a imágenes
def convertir_pdf_a_imagenes(pdf_path):
    return convert_from_path(pdf_path)

# Función para cargar la plantilla limpia
def cargar_plantilla(plantilla_path):
    paginas = convertir_pdf_a_imagenes(plantilla_path)
    plantilla = np.array(paginas[0])  # Tomamos solo la primera página como plantilla
    plantilla_gray = cv2.cvtColor(plantilla, cv2.COLOR_BGR2GRAY)
    return plantilla_gray

# Función para detectar las coordenadas de las respuestas en la plantilla
def detectar_coordenadas_respuestas(plantilla, recorte_respuestas_path):
    recorte_respuestas = cv2.imread(recorte_respuestas_path, cv2.IMREAD_GRAYSCALE)
    _, binary = cv2.threshold(recorte_respuestas, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)
    
    contornos, _ = cv2.findContours(binary, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    
    coordenadas_respuestas = []
    for contorno in contornos:
        x, y, w, h = cv2.boundingRect(contorno)
        coordenadas_respuestas.append((x, y, w, h))
    
    coordenadas_respuestas = sorted(coordenadas_respuestas, key=lambda k: (k[1], k[0]))
    
    # Detectar las coordenadas de las respuestas en la plantilla completa
    resultado = cv2.matchTemplate(plantilla, recorte_respuestas, cv2.TM_CCOEFF_NORMED)
    threshold = 0.9  # Ajustar el umbral para mejorar la precisión
    loc = np.where(resultado >= threshold)
    
    coordenadas_plantilla = []
    for pt in zip(*loc[::-1]):
        for coord in coordenadas_respuestas:
            x, y, w, h = coord
            coordenadas_plantilla.append((pt[0] + x, pt[1] + y, w, h))
    
    return coordenadas_plantilla

# Función para detectar la respuesta marcada en un área específica
def detectar_respuesta_marcada(area):
    gray = cv2.cvtColor(area, cv2.COLOR_BGR2GRAY)
    _, binary = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)
    
    contornos, _ = cv2.findContours(binary, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    
    if contornos:
        return True
    return False

# Función para contar las respuestas en una página dada
def contar_respuestas_en_pagina(imagen, coordenadas_respuestas):
    conteos = {f'Pregunta {i+1}': {str(j): 0 for j in range(1, 9)} for i in range(33)}
    
    for pregunta in range(33):
        for idx, (x, y, w, h) in enumerate(coordenadas_respuestas[pregunta*8:(pregunta+1)*8]):
            area = imagen[y:y+h, x:x+w]
            
            # Mostrar el área que se está evaluando usando Matplotlib
            plt.imshow(cv2.cvtColor(imagen, cv2.COLOR_BGR2RGB))
            plt.gca().add_patch(plt.Rectangle((x, y), w, h, edgecolor='red', facecolor='none'))
            plt.title(f'Pregunta {pregunta+1}, Opción {idx+1}')
            plt.show()
            
            if detectar_respuesta_marcada(area):
                conteos[f'Pregunta {pregunta+1}'][str(idx+1)] += 1
    
    return conteos

# Función para procesar todas las encuestas en un directorio
def procesar_encuestas(pdfs_dir, coordenadas_respuestas):
    conteo_total = {f'Pregunta {i+1}': {str(j): 0 for j in range(1, 9)} for i in range(33)}
    
    for filename in os.listdir(pdfs_dir):
        if filename.endswith(".pdf"):
            pdf_path = os.path.join(pdfs_dir, filename)
            imagenes = convertir_pdf_a_imagenes(pdf_path)
            
            for imagen in imagenes:
                imagen_np = np.array(imagen)
                conteo_pagina = contar_respuestas_en_pagina(imagen_np, coordenadas_respuestas)
                
                for pregunta, conteos in conteo_pagina.items():
                    for respuesta, cantidad in conteos.items():
                        conteo_total[pregunta][respuesta] += cantidad
    
    return conteo_total

# Función para escribir los resultados en un archivo Excel
def escribir_resultados_en_excel(conteo_total, output_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Resultados"
    
    header = ["Pregunta"] + [str(i) for i in range(1, 9)]
    ws.append(header)
    
    for pregunta, conteos in conteo_total.items():
        fila = [pregunta] + [conteos[str(i)] for i in range(1, 9)]
        ws.append(fila)
    
    wb.save(output_path)

# Función principal
def main():
    plantilla_path = 'plantilla_limpia.pdf'
    recorte_respuestas_path = 'recorte_respuestas.jpg'
    pdfs_dir = 'scans'
    output_path = 'resultados.xlsx'
    
    plantilla = cargar_plantilla(plantilla_path)
    coordenadas_respuestas = detectar_coordenadas_respuestas(plantilla, recorte_respuestas_path)
    
    conteo_total = procesar_encuestas(pdfs_dir, coordenadas_respuestas)
    escribir_resultados_en_excel(conteo_total, output_path)
    print(f'Resultados escritos en {output_path}')

if __name__ == "__main__":
    main()
