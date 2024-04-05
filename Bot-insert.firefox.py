from selenium import webdriver
from openpyxl import load_workbook
import time

# Rutas y archivos genéricos - Cambia estas rutas por las rutas específicas en tu sistema
RUTA_FIREFOX = r"RUTA_AL_EJECUTABLE_DE_FIREFOX"
RUTA_ARCHIVO_EXCEL = r"RUTA_AL_ARCHIVO_EXCEL"

# Función para iniciar sesión en el sitio web - Cambia estas URL y credenciales por las de tu sitio
def iniciar_sesion(driver):
    # Abrir la página de inicio de sesión
    driver.get("URL_DEL_SITIO_WEB")
    time.sleep(2)
    # Introducir credenciales y hacer clic en el botón de inicio de sesión
    driver.find_element("name", "LOGIN").send_keys("USUARIO")
    driver.find_element("name", "PASSWD").send_keys("CONTRASEÑA")
    driver.find_element("name", "Valid_CNX").click()
    time.sleep(2)

# Función para procesar los datos del archivo Excel y asignar etiquetas a los equipos
def procesar_excel_y_asignar_etiquetas(driver, archivo_excel):
    wb = load_workbook(archivo_excel)
    hoja_nombres = wb['Hoja1']
    for i in range(2, 6):
        nombre, apellido = hoja_nombres[f'A{i}:B{i}'][0]
        print(nombre.value)
        print(apellido.value)

        # Abrir la página de cada equipo y asignar etiqueta - Cambia esta URL por la correspondiente a tu sitio
        driver.get("URL_DEL_SITIO_WEB_PARA_EQUIPO" + str(nombre.value))
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
    iniciar_sesion(driver)  # Iniciar sesión en el sitio web
    procesar_excel_y_asignar_etiquetas(driver, RUTA_ARCHIVO_EXCEL)  # Procesar Excel y asignar etiquetas
    driver.quit()  # Cerrar el navegador al finalizar

if __name__ == "__main__":
    main()