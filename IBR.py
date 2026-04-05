import time
import os
import win32com.client # Comunicación con Excel en ejecución
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from datetime import datetime

def extraer_datos_ibr():
    options = Options()
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")
    options.add_argument("--headless") 
    options.add_argument("--window-size=1920,1080")
    
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    url = "https://suameca.banrep.gov.co/estadisticas-economicas/informacionSerie/241/tasas_interes_indicador_bancario_referencia_ibr"
    
    try:
        print("Conectando con Banrep para extraer IBR...")
        driver.get(url)
        wait = WebDriverWait(driver, 40)
        
        wait.until(EC.presence_of_element_located((By.CLASS_NAME, "highcharts-label")))
        time.sleep(3) 

        elementos = driver.find_elements(By.CLASS_NAME, "highcharts-label")
        
        datos_finales = []
        for e in elementos:
            texto = e.text.strip().replace(',', '.') 
            if texto and any(c.isdigit() for c in texto):
                mitad = len(texto) // 2
                valor_str = texto[:mitad] if texto[:mitad] == texto[mitad:] else texto
                try:
                    datos_finales.append(float(valor_str))
                except:
                    continue

        if len(datos_finales) < 10:
            print(f"Error: Solo se detectaron {len(datos_finales)} valores.")
            return None

        fecha_actual = datetime.now().strftime('%d/%m/%Y')
        
        # Estructura: Fecha, Tipo, Nominal/Efectiva, overnight, 1m, 3m, 12m
        tabla = [
            [fecha_actual, "Nominal", datos_finales[0], datos_finales[2], datos_finales[4], datos_finales[8]],
            [fecha_actual, "Efectiva", datos_finales[1], datos_finales[3], datos_finales[5], datos_finales[9]]
        ]
        return tabla

    except Exception as e:
        print(f"ERROR en scraping: {str(e)}")
        return None
    finally:
        driver.quit()

def actualizar_excel_ibr_abierto(nuevos_datos):
    nombre_libro = "Renta fija.xlsm"
    try:
        # Conectar con la aplicación Excel abierta
        excel = win32com.client.GetActiveObject("Excel.Application")
        
        # Buscar el libro 
        wb = None
        for book in excel.Workbooks:
            if book.Name == nombre_libro:
                wb = book
                break
        
        if wb is None:
            print(f"Error: El archivo {nombre_libro} no está abierto.")
            return

        ws = wb.Worksheets("IBR")
        
        # Encontrar última fila 
        ultima_fila = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
        
        # Validar duplicados por fecha 
        fecha_ultima = str(ws.Cells(ultima_fila, 1).Value)
        fecha_hoy = nuevos_datos[0][0]

        # Limpieza simple de formato de fecha que devuelve Excel
        if "00:00:00" in fecha_ultima:
            fecha_ultima = datetime.strptime(fecha_ultima.split()[0], '%Y-%m-%d').strftime('%d/%m/%Y')

        if fecha_ultima == fecha_hoy:
            print(f"Los datos de IBR para {fecha_hoy} ya existen en el Excel.")
        else:
            proxima_fila = ultima_fila + 1
            for i, fila_datos in enumerate(nuevos_datos):
                fila_actual = proxima_fila + i
                # Pegar columna por columna (A a F)
                for j, valor in enumerate(fila_datos):
                    ws.Cells(fila_actual, j + 1).Value = valor
                
                # Aplicar formato de 3 decimales a las columnas numéricas 
                for col_idx in range(3, 7):
                    ws.Cells(fila_actual, col_idx).NumberFormat = "#,##0.000"
            
            print(f"¡Éxito! Datos de IBR pegados en Excel abierto para la fecha {fecha_hoy}.")

    except Exception as e:
        print(f"Error al conectar con la instancia de Excel: {e}")

if __name__ == "__main__":
    resultado = extraer_datos_ibr()
    if resultado:
        actualizar_excel_ibr_abierto(resultado)