import os
import glob
import csv
import win32com.client  # Para interactuar con Excel abierto
from datetime import datetime

def limpiar_y_pegar_datos():
    # Detectar la ruta de script
    ruta_actual = os.path.dirname(os.path.abspath(__file__))
    nombre_maestro = "Renta fija.xlsm"
    
    # Buscar el archivo CSV descargado por el scraping
    archivos_descargados = glob.glob(os.path.join(ruta_actual, "DeudaCorporativa*.csv"))
    
    if not archivos_descargados:
        print("Error: No se encontró el archivo CSV en la carpeta.")
        return

    # Tomar el archivo más reciente
    archivo_reciente = max(archivos_descargados, key=os.path.getctime)
    print(f"Procesando archivo: {os.path.basename(archivo_reciente)}")

    try:
        #  Conexión con la instancia activa de Excel
        try:
            excel = win32com.client.GetActiveObject("Excel.Application")
        except Exception:
            print("Error: Excel no está detectado como abierto.")
            return

        # Localizar el libro 
        wb = None
        for libro in excel.Workbooks:
            if libro.Name == nombre_maestro:
                wb = libro
                break
        
        if wb is None:
            print(f"Error: El libro '{nombre_maestro}' debe estar abierto.")
            return

        ws = wb.Worksheets("BASE")
        
        # Identificar la última fila con datos 
        ultima_fila = ws.Cells(ws.Rows.Count, 1).End(-4162).Row # -4162 = xlUp
        proxima_fila = ultima_fila + 1
        
        fecha_ejecucion = datetime.now().strftime("%d/%m/%Y")

        #  Procesar el CSV e insertar datos
        with open(archivo_reciente, mode='r', encoding='utf-8-sig') as f:
            lector = csv.reader(f, delimiter=';')
            next(lector) # Saltar la fila de títulos
            
            filas_insertadas = 0
            for fila in lector:
                if len(fila) < 10: continue

                # --- CORRECCIÓN DE FORMATO DE FECHA
                dato_fecha_venc = fila[3]
                try:
                    
                    partes = dato_fecha_venc.split()
                    # Convertimos a objeto fecha real para que Excel lo entienda
                    fecha_limpia = datetime.strptime(f"{partes[1]} {partes[2]} {partes[3]}", "%b %d %Y")
                except:
                    fecha_limpia = dato_fecha_venc

                # Función interna para limpiar números (comas por puntos)
                def clean_num(valor):
                    try: return float(str(valor).replace(',', '.'))
                    except: return valor

                # Pegado de datos Columna por Columna (A a J)
                for i in range(10):
                    val_original = fila[i]
                    col_idx = i + 1
                    
                    celda = ws.Cells(proxima_fila, col_idx)
                    
                    if i == 3: # Columna D: Fecha de vencimiento
                        celda.Value = fecha_limpia
                        celda.NumberFormat = "dd/mm/yyyy" # formato fecha corta
                    
                    elif i == 5: # Columna F: Tasa Facial
                        celda.Value = clean_num(val_original)
                        celda.NumberFormat = "0.00"
                    
                    elif i in [7, 8]: # Columnas H e I: Cantidad y Volúmenes
                        celda.Value = clean_num(val_original)
                        celda.NumberFormat = "#,##0"
                    
                    else: # Resto de columnas (Texto)
                        celda.Value = val_original

                # Columna K: 
                ws.Cells(proxima_fila, 11).Value = fecha_ejecucion
                ws.Cells(proxima_fila, 11).NumberFormat = "dd/mm/yyyy"
                
                proxima_fila += 1
                filas_insertadas += 1

        print(f"¡ÉXITO! Se agregaron {filas_insertadas} filas a la base.")
        
        # 5. Cerrar archivo CSV y eliminarlo
        f.close()
        os.remove(archivo_reciente)

    except Exception as e:
        print(f"Se produjo un error durante el proceso: {e}")

if __name__ == "__main__":
    limpiar_y_pegar_datos()