from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
import datetime
import time
import pandas as pd

# Paso 0: Rutas del sistema
chromedriver_path = r"C:\Users\La Pecera\Documents\CODEO\install\chromedriver-win64\chromedriver-win64\chromedriver.exe"
chrome_binary_path = r"C:\Program Files\Google\Chrome\Application\chrome.exe"

# Paso 1: Configurar navegador
options = webdriver.ChromeOptions()
options.binary_location = chrome_binary_path
options.add_argument("--start-maximized")

from selenium import webdriver
from selenium.webdriver.chrome.options import Options

options = Options()
options.binary_location = r"C:\Program Files\Google\Chrome\Application\chrome.exe"

# options.add_argument("--headless=new")  # si querés headless

driver = webdriver.Chrome(options=options)  # ← SIN Service(), Selenium Manager resuelve el driver

try:
    # Paso 3: Acceder a la página
    driver.get("https://portaltramites.inpi.gob.ar/marcasconsultas/busqueda/?Cod_Funcion=NQA0ADEA")
    wait = WebDriverWait(driver, 20)

    # Paso 4: Abrir acordeón
    wait.until(EC.element_to_be_clickable((By.XPATH, "//a[contains(text(),'BUSCADOR DE MARCAS')]"))).click()
    print("✅ Acordeón abierto")
    time.sleep(1)

    # Paso 5: Seleccionar fecha desde (miércoles anterior)
    fecha_input = wait.until(EC.element_to_be_clickable((By.ID, "IngresoDesde")))
    fecha_input.click()
    print("🗓️ Calendario desplegado")
    wait.until(EC.visibility_of_element_located((By.ID, "ui-datepicker-div")))

    # Calcular siempre el miércoles de la semana pasada
    today = datetime.date.today()
    dias_desde_miercoles = (today.weekday() - 2) % 7
    miercoles_esta_semana = today - datetime.timedelta(days=dias_desde_miercoles)
    miercoles_anterior = miercoles_esta_semana - datetime.timedelta(days=7)
    print(f"➡ Buscando miércoles anterior: {miercoles_anterior.strftime('%d/%m/%Y')}")

    # Paso 5.1: Función para obtener mes y año visibles en el calendario
    def obtener_mes_anyo_actual():
        titulo = wait.until(EC.presence_of_element_located((By.CLASS_NAME, "ui-datepicker-title")))
        mes = titulo.find_element(By.CLASS_NAME, "ui-datepicker-month").text.lower()
        anio = titulo.find_element(By.CLASS_NAME, "ui-datepicker-year").text
        return mes, int(anio)

    meses = {
        'enero': 1, 'febrero': 2, 'marzo': 3, 'abril': 4, 'mayo': 5, 'junio': 6,
        'julio': 7, 'agosto': 8, 'septiembre': 9, 'octubre': 10, 'noviembre': 11, 'diciembre': 12
    }

    # Navegar hasta mes/año del miércoles anterior
    while True:
        mes_visible, anio_visible = obtener_mes_anyo_actual()
        if meses[mes_visible] == miercoles_anterior.month and anio_visible == miercoles_anterior.year:
            break
        wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "ui-datepicker-prev"))).click()
        time.sleep(0.5)

    # Paso 5.3: Seleccionar día correcto
    xpath_dia = f"//div[@id='ui-datepicker-div']//a[text()='{miercoles_anterior.day}']"
    wait.until(EC.element_to_be_clickable((By.XPATH, xpath_dia))).click()
    print("✅ Fecha seleccionada")

    # Paso 5.4: Seleccionar tipo de búsqueda "CONTIENE"
    select_denominacion = Select(wait.until(EC.element_to_be_clickable((By.ID, "TxtDenominacionTipoBusqueda"))))
    select_denominacion.select_by_value("1")
    print("✅ Seleccionado tipo de búsqueda 'CONTIENE'")

    # Paso 5.5: Ingresar fecha hasta (hoy)
    hoy = datetime.date.today()
    print(f"➡ Ingresando fecha actual en campo 'IngresoHasta': {hoy.strftime('%d/%m/%Y')}")
    ingreso_hasta = wait.until(EC.element_to_be_clickable((By.ID, "IngresoHasta")))
    ingreso_hasta.click()
    wait.until(EC.visibility_of_element_located((By.ID, "ui-datepicker-div")))

    while True:
        mes_visible, anio_visible = obtener_mes_anyo_actual()
        if meses[mes_visible] == hoy.month and anio_visible == hoy.year:
            break
        wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "ui-datepicker-next"))).click()
        time.sleep(0.5)

    xpath_hoy = f"//div[@id='ui-datepicker-div']//a[text()='{hoy.day}']"
    wait.until(EC.element_to_be_clickable((By.XPATH, xpath_hoy))).click()
    print("✅ Fecha actual seleccionada en 'IngresoHasta'")

    # Paso 6: Buscar
    wait.until(EC.element_to_be_clickable((By.ID, "BtnBuscarAvanzada"))).click()
    print("🔘 Botón 'Buscar Avanzada' presionado")

    # Paso 7: Cambiar cantidad de registros por página a 100
    WebDriverWait(driver, 90).until(EC.presence_of_element_located((By.ID, "tblGrillaMarcas")))
    WebDriverWait(driver, 90).until(
        lambda d: len(d.find_elements(By.CSS_SELECTOR, "#tblGrillaMarcas tbody tr")) > 0
    )

    btn_dropdown = WebDriverWait(driver, 90).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "button.btn.btn-default.dropdown-toggle"))
    )
    driver.execute_script("arguments[0].scrollIntoView(true);", btn_dropdown)
    btn_dropdown = WebDriverWait(driver, 30).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, "button.btn.btn-default.dropdown-toggle"))
    )
    btn_dropdown.click()
    opciones = WebDriverWait(driver, 60).until(
        EC.presence_of_all_elements_located((By.CSS_SELECTOR, "ul.dropdown-menu > li > a"))
    )
    for opcion in opciones:
        if opcion.text.strip() == "100":
            opcion.click()
            print("✅ Seleccionado '100 registros por página'")
            break
 # 👇👇👇 NUEVO: Esperar a que la tabla efectivamente cargue 100 filas
    print("⏳ Esperando a que se carguen los 100 resultados de la primera página...")
    WebDriverWait(driver, 90).until(
        lambda d: len(d.find_elements(By.CSS_SELECTOR, "#tblGrillaMarcas tbody tr")) >= 100
    )
    print("✅ Primera página lista con 100 resultados")
    # Paso 8: Extraer datos de tabla paginada
    wait.until(EC.presence_of_element_located((By.ID, "tblGrillaMarcas")))
    datos_completos = []

    while True:
        # Extraer datos de la página actual
        tabla = wait.until(EC.presence_of_element_located((By.ID, "tblGrillaMarcas")))
        encabezados = [th.text.strip() for th in tabla.find_elements(By.CSS_SELECTOR, "thead th")]
        filas = tabla.find_elements(By.CSS_SELECTOR, "tbody > tr")

        for fila in filas:
            try:
                celdas = fila.find_elements(By.TAG_NAME, "td")
                fila_datos = [celda.text.strip() for celda in celdas]
                if fila_datos:
                    datos_completos.append(dict(zip(encabezados, fila_datos)))
            except Exception as e:
                print(f"⚠️ Fila omitida por error: {e}")
                continue

        # ✅ Comprobar número de página actual y total para evitar volver al inicio
        try:
            paginas = driver.find_elements(By.CSS_SELECTOR, "ul.pagination li.page-number")
            pagina_actual = driver.find_element(By.CSS_SELECTOR, "ul.pagination li.page-number.active").text.strip()
            total_paginas = paginas[-1].text.strip() if paginas else pagina_actual
            print(f"📄 Página actual: {pagina_actual} / {total_paginas}")

            if pagina_actual == total_paginas:
                print("➡ Última página alcanzada. Fin de la extracción.")
                break

            # Avanzar a la siguiente página
            btn_siguiente = driver.find_element(By.CSS_SELECTOR, "li.page-next > a")
            driver.execute_script("arguments[0].scrollIntoView(true);", btn_siguiente)
            time.sleep(0.5)
            driver.execute_script("arguments[0].click();", btn_siguiente)
            print("➡ Navegando a la siguiente página...")
            time.sleep(2)

        except Exception as e:
            print(f"⚠️ Error al procesar paginación: {e}")
            break

    # Paso 9: Guardar en Excel
    if datos_completos:
        df = pd.DataFrame(datos_completos)
        df.to_excel("marcas_grilla_completa.xlsx", index=False)
        print("📊 Archivo Excel generado con todos los datos: marcas_grilla_completa.xlsx")
    else:
        print("⚠️ No se encontraron datos para exportar.")

    input("🧪 Presione ENTER para cerrar...")

finally:
    # Paso 10: Cerrar navegador
    driver.quit()
