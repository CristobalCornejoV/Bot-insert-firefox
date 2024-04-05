from selenium import webdriver
from openpyxl import load_workbook
import time

# Rutas genéricas - Cambia estas rutas por las rutas específicas en tu sistema
ruta_firefox = r"RUTA_AL_EJECUTABLE_DE_FIREFOX"  # Ruta al ejecutable de Firefox
ruta_archivo_excel = r"RUTA_AL_ARCHIVO_EXCEL"  # Ruta al archivo Excel

# Cargar archivo Excel
wb = load_workbook(ruta_archivo_excel)
hojas = wb.sheetnames
hoja_datos = wb['Hoja1']
wb.close()

# Configuración del navegador
opciones = webdriver.FirefoxOptions()
opciones.binary_location = ruta_firefox
driver = webdriver.Firefox(options=opciones)

# Iniciar sesión en el sitio web - Cambia estas URL y credenciales por las de tu sitio
driver.get("URL_DEL_SITIO_WEB")  # URL del sitio web
time.sleep(2)
driver.find_element("name", "LOGIN").send_keys("USUARIO")  # Usuario
driver.find_element("name", "PASSWD").send_keys("CONTRASEÑA")  # Contraseña
driver.find_element("name", "Valid_CNX").click()
time.sleep(2)

# Procesar datos del archivo Excel y asignar etiquetas a los equipos
for i in range(2, 6):
    nombre, apellido = hoja_datos[f'A{i}:B{i}'][0]
    print(nombre.value)
    print(apellido.value)

    # Acceder a la URL del equipo específico - Cambia esta URL por la correspondiente a tu sitio
    driver.get("URL_DEL_SITIO_WEB_PARA_CADA_EQUIPO" + str(nombre.value))  # URL del equipo específico
    time.sleep(2)
    driver.find_element("name", "TAG").clear()
    time.sleep(2)
    driver.find_element("name", "TAG").send_keys(str(apellido.value))
    time.sleep(2)
    driver.find_element("name", "Valid_modif").click()
    time.sleep(2)

# Cerrar el navegador
driver.quit()