from selenium import webdriver
from openpyxl import load_workbook
import time

# Variables genéricas
RUTA_FIREFOX = r"RUTA_AL_EJECUTABLE_DE_FIREFOX"  # Cambia esta ruta por la ruta al ejecutable de Firefox en tu sistema
RUTA_ARCHIVO_EXCEL = r"RUTA_AL_ARCHIVO_EXCEL"  # Cambia esta ruta por la ruta al archivo Excel en tu sistema

# Función para iniciar sesión en el sitio web - Cambia estas URL y credenciales por las de tu sitio
def iniciar_sesion(driver):
    driver.get("URL_DEL_SITIO_WEB")  # URL del sitio web
    time.sleep(2)
    driver.find_element("name", "LOGIN").send_keys("USUARIO")  # Usuario
    driver.find_element("name", "PASSWD").send_keys("CONTRASEÑA")  # Contraseña
    driver.find_element("name", "Valid_CNX").click()
    time.sleep(2)

# Función para procesar los datos del archivo Excel y asignar etiquetas a los equipos
def procesar_excel_y_asignar_etiquetas(driver, archivo_excel):
    wb = load_workbook(archivo_excel)
    hoja_datos = wb['Hoja1']
    for i in range(2, 6):
        nombre, apellido = hoja_datos[f'A{i}:B{i}'][0]
        print(nombre.value)
        print(apellido.value)

        # Acceder a la URL del equipo específico - Cambia esta URL por la correspondiente a tu sitio
        driver.get("URL_DEL_SITIO_WEB_PARA_EQUIPO" + str(nombre.value))  # URL del equipo específico
        time.sleep(2)
        driver.find_element("name", "TAG").clear()
        time.sleep(2)
        driver.find_element("name", "TAG").send_keys(str(apellido.value))
        time.sleep(2)
        driver.find_element("name", "Valid_modif").click()
        time.sleep(2)

    wb.close()

# Función principal
def main():
    opciones = webdriver.FirefoxOptions()
    opciones.binary_location = RUTA_FIREFOX
    driver = webdriver.Firefox(options=opciones)
    iniciar_sesion(driver)
    procesar_excel_y_asignar_etiquetas(driver, RUTA_ARCHIVO_EXCEL)
    driver.quit()

if __name__ == "__main__":
    main()