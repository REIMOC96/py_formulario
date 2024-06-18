import os
import cv2
import numpy as np
from pdf2image import convert_from_path
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

# Función para encontrar coordenadas de respuestas marcadas
def encontrar_coordenadas_respuestas(plantilla, imagen_escaneada):
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
    coordenadas_respuestas = []
    for pt in zip(*loc[::-1]):
        x, y = pt
        coordenadas_respuestas.append((x, y))
    
    return coordenadas_respuestas

# Función para determinar la respuesta marcada en las coordenadas dadas
def determinar_respuesta_marcada(imagen_escaneada, x, y):
    # Aquí debes definir la lógica para determinar qué respuesta fue marcada
    # Por ejemplo, puedes usar los valores de píxel en las coordenadas (x, y)
    # y compararlos con valores de referencia para determinar la respuesta.
    # Esto depende de cómo se vean marcadas las respuestas en tus formularios.
    # Por simplicidad, aquí se devuelve directamente el valor de píxel en (x, y).
    return imagen_escaneada[y, x]

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
    
    # Directorio donde están los documentos escaneados
    scans_dir = os.path.join(script_dir, 'scans')
    
    # Inicializar diccionario para contar respuestas
    respuestas_totales = {}
    
    # Definir los números posibles de respuesta (del 1 al 8)
    numeros_respuesta = list(range(1, 9))

    # Lista para almacenar los nombres de archivo y preguntas
    datos_leidos = []

    # Iterar sobre los archivos en el directorio de scans
    for filename in os.listdir(scans_dir):
        if filename.endswith(".pdf"):  # Asegurarse de procesar solo archivos PDF
            # Construir el path completo de la imagen escaneada
            imagen_escaneada_path = os.path.join(scans_dir, filename)
            
            # Convertir el PDF escaneado a imagen
            imagen_escaneada = convertir_pdf_a_imagenes(imagen_escaneada_path)[0]  # Tomamos solo la primera página
            
            if imagen_escaneada is None:
                continue
            
            imagen_escaneada = np.array(imagen_escaneada)
            
            # Encontrar coordenadas de respuestas marcadas
            coordenadas_respuestas = encontrar_coordenadas_respuestas(plantilla, imagen_escaneada)
            
            # Identificar la pregunta utilizando el nombre del archivo
            pregunta = filename  # Puedes ajustar esto según cómo identifiques las preguntas
            
            # Contar las respuestas marcadas del 1 al 8
            conteo_respuestas = {str(num): 0 for num in numeros_respuesta}
            for (x, y) in coordenadas_respuestas:
                respuesta_marcada = determinar_respuesta_marcada(imagen_escaneada, x, y)
                if respuesta_marcada and str(respuesta_marcada) in conteo_respuestas:
                    conteo_respuestas[str(respuesta_marcada)] += 1
            
            # Guardar el conteo de respuestas para la pregunta actual
            respuestas_totales[pregunta] = conteo_respuestas
            
            # Guardar los datos leídos (nombre de archivo y pregunta)
            datos_leidos.append((filename, pregunta))
    
    # Crear un libro de Excel y agregar los resultados
    wb = Workbook()
    ws = wb.active
    ws.title = "Resultados"
    
    # Escribir los encabezados
    encabezados = ["Archivo", "Pregunta"] + [str(num) for num in numeros_respuesta]
    ws.append(encabezados)
    
    # Escribir los resultados en el archivo de Excel
    for datos in datos_leidos:
        filename, pregunta = datos
        conteo_respuestas = respuestas_totales[pregunta]
        fila = [filename, pregunta] + [conteo_respuestas[str(num)] for num in numeros_respuesta]
        ws.append(fila)
    
    # Guardar el archivo de Excel
    resultado_path = os.path.join(script_dir, "resultados.xlsx")
    wb.save(resultado_path)
    print(f"Resultados guardados en {resultado_path}")

if __name__ == "__main__":
    main()
