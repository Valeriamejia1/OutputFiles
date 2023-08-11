import os
import time
import shutil
import pyautogui
import glob

shortcut_path = r"C:\dist\pdfreader\pdfreader.exe"
pdf_folder = r"C:\Python tools\PDF reader\OutputFiles\QA\Default"
output_folder = r"C:\Python tools\PDF reader\OutputFiles\OUTPUT Default"

pdf_files = [
    #"06-11 minutes.pdf",
    #"1690203601050_1994364726 minutes.pdf",
    #"1690808400472_1671940182 minutes.pdf",
    #"Combined File minutes.pdf",
    #"Kronos Timecards TC 07-30-22 minutes.pdf",
    #"Scripps Approved Kronos we 6.25.22 minutes.pdf",
    "time weston minutes.pdf",
    # Agrega los nombres de los archivos PDF que desees procesar
]

for pdf_file in pdf_files:
    # Abre el tool en una nueva ventana
    # process = subprocess.Popen([shortcut_path])

    time.sleep(3)

    formato_pos = (136, 48)  # Reemplaza con las coordenadas reales
    pyautogui.click(formato_pos)
    pyautogui.click(formato_pos)

    # Espera unos segundos para que el tool se abra completamente
    time.sleep(5)

    # Selecciona el formato en "Select Report Type" y "default"
    formato_pos = (320, 162)  # Reemplaza con las coordenadas reales
    pyautogui.click(formato_pos)

    time.sleep(1)

    formato_pos = (300, 187)  # Reemplaza con las coordenadas reales
    pyautogui.click(formato_pos)

    time.sleep(1)

    # Haz click en el botón "Open File" y selecciona el archivo PDF
    formato_pos = (313, 242)  # Reemplaza con las coordenadas reales
    pyautogui.click(formato_pos)
    time.sleep(1)

    # Ingresa la ruta completa del archivo PDF en el diálogo de selección de archivo
    pyautogui.typewrite(os.path.join(pdf_folder, pdf_file))
    pyautogui.press("enter")

    time.sleep(1)

    # Haz click en el botón "yes minutes" y selecciona el archivo PDF
    formato_pos = (983, 614)  # Reemplaza con las coordenadas reales
    pyautogui.click(formato_pos)
    
    time.sleep(2)
    # Haz click en el botón "Proccess File" y selecciona el archivo PDF
    formato_pos = (326, 272)  # Reemplaza con las coordenadas reales
    pyautogui.click(formato_pos)
    pyautogui.click(formato_pos)

    # Espera un tiempo prudencial para que se procese el archivo PDF
    time.sleep(3)

    # Seleccionar "OK"
    formato_pos = (1073, 610)  # Reemplaza con las coordenadas reales
    pyautogui.click(formato_pos)

    time.sleep(1)

    # Cierra el tool
    formato_pos = (535, 78)  # Reemplaza con las coordenadas reales
    pyautogui.click(formato_pos)

    # Obtiene el nombre original del archivo PDF sin la extensión
    pdf_name = os.path.splitext(pdf_file)[0]

    time.sleep(3)

    # Busca el archivo de Excel generado en la carpeta de salida
    excel_files = glob.glob(r"C:\dist\output\*.xlsx")
    if len(excel_files) > 0:
        # Obtiene la ruta del primer archivo de Excel encontrado
        excel_path = excel_files[0]

        # Genera el nuevo nombre de archivo basado en el nombre del PDF
        new_excel_name = f"{pdf_name}.xlsx"

        # Mueve el archivo de Excel generado al directorio de salida y renómbralo
        new_excel_path = os.path.join(output_folder, new_excel_name)
        shutil.move(excel_path, new_excel_path)

