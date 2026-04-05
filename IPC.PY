import time
import os
import win32com.client 
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

def extraer_ipc_selenium():
    options = Options()
    options.add_argument("--headless") # Ejecución en segundo plano
    options.add_argument("--disable-blink-features=AutomationControlled")
    
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    url = "https://www.dane.gov.co/index.php/estadisticas-por-tema/precios-y-costos/indice-de-precios-al-consumidor-ipc/ipc-informacion-tecnica"
    
    try:
        print("Accediendo al DANE...")
        driver.get(url)
        wait = WebDriverWait(driver, 30)
        time.sleep(5) 
        
        # Entrar al iframe
        driver.switch_to.frame(0) 

        # Localizar dato 
        xpath_dato = "//td[contains(normalize-space(), '5,29')]"
        elemento_ipc = wait.until(EC.presence_of_element_located((By.XPATH, xpath_dato)))
        
        valor_texto = elemento_ipc.text.strip().replace(',', '.')
        ipc_valor = float(valor_texto)
        
        print(f"IPC extraído: {ipc_valor}")
        return {"mes": "Febrero", "año": 2026, "valor": ipc_valor}

    except Exception as e:
        print(f"Error en scraping: {e}")
        return None
    finally:
        driver.quit()

def actualizar_excel_en_vivo(datos):
    nombre_libro = "Renta fija.xlsm"
    try:
        # Capturar la instancia de Excel abierta
        excel = win32com.client.GetActiveObject("Excel.Application")
        
        # Buscar el libro específico entre los abiertos
        wb = None
        for book in excel.Workbooks:
            if book.Name == nombre_libro:
                wb = book
                break
        
        if wb is None:
            print(f"Error: El archivo {nombre_libro} no está abierto en Excel.")
            return

        ws = wb.Worksheets("IPC") 
        # Encontrar última fila 
        ultima_fila = ws.Cells(ws.Rows.Count, 1).End(-4162).Row 
        # Comprobar duplicados
        mes_existente = ws.Cells(ultima_fila, 1).Value
        anio_existente = ws.Cells(ultima_fila, 2).Value
        
        if str(mes_existente).strip().lower() == datos['mes'].lower() and int(anio_existente) == datos['año']:
            print("El dato ya existe en la hoja.")
        else:
            proxima_fila = ultima_fila + 1
            # Insertar datos directamente en las celdas
            ws.Cells(proxima_fila, 1).Value = datos['mes']
            ws.Cells(proxima_fila, 2).Value = datos['año']
            ws.Cells(proxima_fila, 3).Value = datos['valor']
            
            # Formatear la celda del valor como número
            ws.Cells(proxima_fila, 3).NumberFormat = "0,00"
            
            print(f"¡Éxito! Dato pegado en la fila {proxima_fila}")
            
    except Exception as e:
        print(f"Error al conectar con la instancia de Excel: {e}")

if __name__ == "__main__":
    resultado = extraer_ipc_selenium()
    if resultado:
        actualizar_excel_en_vivo(resultado)