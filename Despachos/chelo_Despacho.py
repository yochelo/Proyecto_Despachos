import json
from datetime import datetime
import sys
import os
from colorama import Fore, Style, init
import time  
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
import pandas as pd
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive
import re

# Inicializar colorama para habilitar los colores en la consola de Windows
init(autoreset=True)

# Redirigir errores al log pero también mostrarlos en consola
log_file = "error_log.txt"
sys.stderr = open(log_file, "w", buffering=1)

def log_error(message):
    print(f"{Fore.RED}{message}{Style.RESET_ALL}")
    with open(log_file, "a") as log:
        log.write(message + "\n")

# Configuración básica
try:
    # Ignorar errores de certificado
    chrome_options = Options()
    chrome_options.add_argument('--ignore-certificate-errors')
    chrome_options.add_argument('--ignore-ssl-errors')
    chrome_options.add_argument('--start-maximized')  # Opcional: iniciar maximizado

    # Conexión al navegador Chrome en modo depuración
    chrome_options.debugger_address = "localhost:9223"

    # Ruta al chromedriver
    driver_path = r"C:\Users\lgv\Desktop\Chelo\chromedriver_win32\chromedriver.exe"
    service = Service(driver_path)

    # Intentar conectar al navegador
    print(f"{Fore.CYAN}Intentando conectar con Chrome en modo depuración...{Style.RESET_ALL}")
    driver = webdriver.Chrome(service=service, options=chrome_options)
    print(f"{Fore.CYAN}¡Conexión exitosa con Chrome!{Style.RESET_ALL}")
except Exception as e:
    log_error(f"Error al iniciar WebDriver: {str(e)}")
    sys.exit()

# Configuración de espera dinámica
wait = WebDriverWait(driver, 15)  # Tiempo máximo para las esperas dinámicas

def capturar_stock(driver):
    try:
        contenedor = driver.find_element(By.ID, "contentwrapper")
        texto = contenedor.text
        match = re.search(r"Stock:\s*([\d\.]+)", texto)
        if match:
            stock_raw = match.group(1).replace(".", "")  # Quitamos el punto de miles
            return int(stock_raw)
        else:
            return None
    except Exception as e:
        log_error(f"No se pudo capturar el stock: {str(e)}")
        return None


# Verificar ventanas abiertas
try:
    print(f"{Fore.CYAN}Verificando ventanas abiertas...{Style.RESET_ALL}")
    windows = driver.window_handles
    if not windows:
        log_error("No se detectaron ventanas abiertas en Chrome.")
        sys.exit()
    print(f"{Fore.CYAN}Ventanas disponibles: {windows}{Style.RESET_ALL}")
except Exception as e:
    log_error(f"Error al obtener ventanas abiertas: {str(e)}")
    sys.exit()

# Verificar el título de la página
try:
    print(f"{Fore.CYAN}Verificando título de la página...{Style.RESET_ALL}")
    page_title = driver.title
    print(f"{Fore.CYAN}Título detectado: {page_title}{Style.RESET_ALL}")
except Exception as e:
    log_error(f"No se pudo obtener el título de la página: {str(e)}")

print(f"{Fore.CYAN}Todo listo para interactuar con la página.{Style.RESET_ALL}")

# Leer el archivo Excel
ruta_excel = "planilla.xlsx"
data = pd.read_excel(ruta_excel)


# Lista para el reporte
reporte = []

# Bandera para controlar si venimos de una demanda insatisfecha
saltar_paso_6 = False

# Bandera para controlar si venimos de "Seleccionar Otro Insumo"
seleccionar_otro_insumo = False

# Suponemos que el usuario ya está en el paso 2
print(f"{Fore.GREEN}INICIANDO OK{Style.RESET_ALL}")

# Iterar sobre los ítems del Excel
for index, row in data.iterrows():
    nombre_item = row["item"]
    solicitadas = row["solicitada"]
    otorgadas = row["otorgada"]

    # Saltar ítems sin cantidades válidas
    if pd.isna(solicitadas) or pd.isna(otorgadas):
        continue

    try:
        # Paso 2: Hacer clic en "Agregar Insumo" solo si no venimos de "Seleccionar Otro Insumo"
        if not seleccionar_otro_insumo:
            print("Haciendo clic en 'Agregar Insumo'...")
            boton_agregar_insumo = wait.until(
                EC.element_to_be_clickable((By.LINK_TEXT, "Agregar Insumo"))
            )
            time.sleep(2)  # Espera antes de interactuar con la pantalla principal (Agregar Insumo)
            boton_agregar_insumo.click()
        else:
            # Reiniciar la bandera para el próximo ítem
            seleccionar_otro_insumo = False

        # Paso 3: Buscar el NNE en el cuadro de búsqueda
        print(f"Buscando NNE para el ítem: {nombre_item}...")

        # Leer el archivo Excel del diccionario
        ruta_diccionario = ruta_diccionario = "diccionario.xlsx"

        diccionario = pd.read_excel(ruta_diccionario)

        # Filtrar el NNE correspondiente al ítem
        nne_row = diccionario[diccionario['item'] == nombre_item]

        # Verificar si se encontró el NNE
        if not nne_row.empty:
            nne = nne_row.iloc[0]['NNE']
            print(f"NNE encontrado: {nne}")
        else:
            raise Exception(f"NNE no encontrado para el ítem: {nombre_item}")

        # Interactuar con la página usando el NNE
        cuadro_busqueda = wait.until(EC.presence_of_element_located((By.NAME, "__buscar")))
        cuadro_busqueda.clear()  # Limpiar el cuadro antes de escribir
        cuadro_busqueda.send_keys(f" {nne}")  # Nota: espacio inicial requerido
        boton_buscar = driver.find_element(By.XPATH, "//input[@value='Buscar por sistema']")
        boton_buscar.click()



        # Paso 4: Seleccionar el ítem de la ventana emergente
        try:
            print(f"Esperando el ítem relacionado al NNE: {nne} en la ventana emergente.")

            # Detectar si hay un cambio de ventana
            main_window = driver.current_window_handle  # Ventana principal
            wait.until(EC.number_of_windows_to_be(2))  # Esperar a que haya 2 ventanas abiertas
            new_window = [window for window in driver.window_handles if window != main_window][0]
            
            # Cambiar al nuevo contexto (ventana emergente)
            driver.switch_to.window(new_window)
            print("Ventana emergente detectada, interactuando con el contenido.")

            # Buscar el enlace que contiene el NNE como texto
            xpath_nne = f"//a[contains(text(), '{nne}')]"
            elemento_nne = wait.until(EC.element_to_be_clickable((By.XPATH, xpath_nne)))
            print(f"Elemento encontrado: {elemento_nne.text}. Intentando hacer clic.")
            elemento_nne.click()
            print(f"Ítem relacionado al NNE {nne} seleccionado correctamente.")

            # Regresar al contexto principal
            driver.switch_to.window(main_window)  # Regresa a la ventana principal

        except Exception as e:
            print(f"{Fore.RED}ERROR: No se pudo seleccionar el ítem en la ventana emergente. Detalle: {str(e)}{Style.RESET_ALL}")
            break

        # Paso 5: Completar cantidades pedidas y despachadas
        try:
            print("Rellenando los campos de cantidad pedida y cantidad despachada...")
            saltar_paso_6 = False  # Definir la variable por defecto

            # Seleccionar y llenar el campo de "Cantidad pedida"
            campo_pedida = wait.until(EC.presence_of_element_located((By.NAME, "cant_ped")))
            campo_pedida.clear()
            campo_pedida.send_keys(str(int(solicitadas)))
            print(f"Cantidad pedida: {solicitadas}")

            # Seleccionar y llenar el campo de "Cantidad despachada"
            campo_despachada = wait.until(EC.presence_of_element_located((By.NAME, "cant")))
            campo_despachada.clear()
            campo_despachada.send_keys(str(int(otorgadas)))
            print(f"Cantidad despachada: {otorgadas}")

            # Capturar el stock antes de enviar
            stock_en_el_momento = capturar_stock(driver)
            print(f"Stock disponible al momento: {stock_en_el_momento}")

            # Verificar si es demanda insatisfecha (otorgadas == 0)
            if otorgadas == 0:
                print(f"{Fore.YELLOW}{Style.BRIGHT}Demanda Insatisfecha detectada para el ítem {nombre_item}. Procesando flujo correspondiente...{Style.RESET_ALL}")
                reporte.append(f"{Fore.YELLOW}{Style.BRIGHT}{nne} | {nombre_item} | pedidos: {solicitadas} - despachados: 0 | DEMANDA INSATISFECHA | stock: {stock_en_el_momento}{Style.RESET_ALL}")
 
                saltar_paso_6 = True  # Marcar para saltar el Paso 6
   

            # Intentar enviar el formulario
            boton_enviar = wait.until(EC.element_to_be_clickable(
                (By.XPATH, "//input[@type='submit' and @name='enviar' and @value='Enviar']")
            ))
            driver.execute_script("arguments[0].scrollIntoView();", boton_enviar)
            driver.execute_script("arguments[0].click();", boton_enviar)
            print("Formulario enviado correctamente.")
            time.sleep(2)  # Espera para que SIGHEOS termine de procesar el envío


            # Manejar alertas inesperadas (caso "Sin Stock")
            time.sleep(0.5)  # Dar tiempo para que aparezca la alerta
            try:
                alerta = driver.switch_to.alert
                mensaje_alerta = alerta.text
                print(f"{Fore.RED}{Style.BRIGHT}Stock Insuficiente: {mensaje_alerta}{Style.RESET_ALL}")
                alerta.accept()  # Hacer clic en "Aceptar" en la alerta
                print("Alerta cerrada correctamente.")

                if "Debe ingresar una cantidad menor o igual al stock del insumo" in mensaje_alerta:
                    # Reportar el ítem como "Stock Insuficiente"
                    reporte.append(f"{Fore.RED}{Style.BRIGHT}{nne} | {nombre_item} | pedidos: {solicitadas} - despachados: {otorgadas} | STOCK INSUFICIENTE | stock: {stock_en_el_momento}{Style.RESET_ALL}")

                    # Hacer clic en "Seleccionar Otro Insumo"
                    boton_otro_insumo = wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Seleccionar Otro Insumo")))
                    boton_otro_insumo.click()
                    print("Clic realizado en 'Seleccionar Otro Insumo'. Preparando flujo para el siguiente ítem...")

                    seleccionar_otro_insumo = True  # Activar la bandera para evitar repetir Paso 2
                    continue  # Saltar al siguiente ítem en el bucle
            except:
                print("No se detectó alerta después de enviar el formulario. Continuando flujo normal.")

            # Registro exitoso en caso de que no haya errores
            if not saltar_paso_6:
                print(f"{Fore.GREEN}Item {nombre_item} correctamente cargado{Style.RESET_ALL}")
                reporte.append(f"{Fore.GREEN}{Style.BRIGHT}{nne} | {nombre_item} | pedidos: {solicitadas} - despachados: {otorgadas} | OK | stock: {stock_en_el_momento}{Style.RESET_ALL}")


        except Exception as e:
            print(f"{Fore.RED}ERROR: Falló el envío del formulario. Detalle: {str(e)}{Style.RESET_ALL}")
            reporte.append(f"{Fore.RED}{nne} | {nombre_item} | ERROR al completar el formulario{Style.RESET_ALL}")
            continue

        # Paso 6: Verificar y hacer clic en "click aquí" si corresponde
        if not saltar_paso_6:  # Ejecutar el Paso 6 solo si no se detecta "Demanda Insatisfecha"
            try:
                print("Verificando si es necesario hacer clic en 'click aquí' para continuar...")
                enlace_click_aqui = wait.until(
                    EC.element_to_be_clickable((By.LINK_TEXT, "click aqui"))
                )
                enlace_click_aqui.click()
                print("Clic realizado en 'click aquí'. Continuando con el proceso...")
            except Exception as e:
                print(f"{Fore.RED}ERROR: No se pudo hacer clic en 'click aquí'. Detalle: {str(e)}{Style.RESET_ALL}")
                reporte.append(f"{Fore.RED}{nne} | {nombre_item} | ERROR al avanzar al siguiente paso{Style.RESET_ALL}")
                break
        else:
            print("Saltando Paso 6 debido a 'Demanda Insatisfecha'.")

    except Exception as e:
        print(
            f"{Fore.RED}ERROR: El bot se detuvo al procesar el ítem {nombre_item} con NNE {nne}. Detalle: {str(e)}{Style.RESET_ALL}"
        )
        reporte.append(f"{Fore.RED}{nne} | {nombre_item} | ERROR general{Style.RESET_ALL}")

# Crear listas separadas para cada color
verde = [linea for linea in reporte if Fore.GREEN in linea]
amarillo = [linea for linea in reporte if Fore.YELLOW in linea]
rojo = [linea for linea in reporte if Fore.RED in linea]

# Combinar las listas reordenadas
reporte_ordenado = verde + amarillo + rojo


# Imprimir el reporte ordenado y limpio
print("\n" + Fore.WHITE + "=" * 100 + Style.RESET_ALL)  # Línea blanca superior
print(f"{Fore.CYAN}{Style.BRIGHT}                                    R E P O R T E    F I N A L{Style.RESET_ALL}")  # Título en celeste
print(Fore.WHITE + "=" * 100 + Style.RESET_ALL)  # Línea blanca inferior
for linea in reporte_ordenado:
    # Limpieza de caracteres no deseados
    texto_limpio = linea.replace("←[0m", "").replace("←[", "").replace("[0m", "")
    print(texto_limpio)

#REPORTES GUARDADOS EN "/REPORTES" POR PLANILLA 

# Asegurarse de que la carpeta Reportes exista
os.makedirs("Reportes", exist_ok=True)

# Pedir metadatos al final, sin condicionar por errores
servicio = input("Ingrese el código del servicio (ej: u15): ").strip().lower()

# Pedir fecha en formato corto (DD/MM)
fecha_input = input("Ingrese la fecha del despacho (DD/MM): ").strip()

try:
    # Agregamos el año actual y convertimos a formato estándar
    dia, mes = map(int, fecha_input.split("/"))
    anio_actual = datetime.now().year
    fecha_obj = datetime(anio_actual, mes, dia)
    fecha = fecha_obj.strftime("%Y-%m-%d")  # Esto te da el formato final
except Exception as e:
    print(f"{Fore.RED}Formato de fecha inválido. Usá DD/MM. Error: {e}{Style.RESET_ALL}")
    sys.exit()

# Timestamp para nombre del archivo
ahora = datetime.now()
hora_str = ahora.strftime("%H-%M-%S")
nombre_archivo = f"Reportes/{fecha}_{hora_str}_{servicio}.json"

# Construir estructura del log completo
log_completo = {
    "servicio": servicio,
    "fecha": fecha,
    "resultados": []
}

for linea in reporte_ordenado:
    # Limpiar colores
    linea_limpia = linea.replace(Fore.RED, "").replace(Fore.GREEN, "").replace(Fore.YELLOW, "").replace(Style.BRIGHT, "").replace(Style.RESET_ALL, "")
    partes = linea_limpia.split("|")
    if len(partes) >= 4:
        log_completo["resultados"].append({
            "nne": partes[0].strip(),
            "item": partes[1].strip(),
            "pedidos": partes[2].split(":")[1].split("-")[0].strip(),
            "despachados": partes[2].split("-")[1].split(":")[1].strip(),
            "estado": partes[3].strip(),
            "stock": partes[4].split(":")[1].strip() if len(partes) > 4 else None
        })

# Guardar el archivo
with open(nombre_archivo, "w", encoding="utf-8") as f:
    json.dump(log_completo, f, indent=4, ensure_ascii=False)

print(f"{Fore.CYAN}Reporte JSON guardado en: {nombre_archivo}{Style.RESET_ALL}")

def registrar_remanentes_global(log_completo, archivo="datos/remanentes.json"):
    
    try:
        if not os.path.exists("remanentes"):
            os.makedirs("remanentes")

        # Cargar remanentes anteriores
        if os.path.exists(archivo):
            with open(archivo, "r", encoding="utf-8") as f:
                data = json.load(f)
        else:
            data = {}

        for item in log_completo["resultados"]:
            if "STOCK INSUFICIENTE" in item["estado"]:
                try:
                    nne = item["nne"]
                    nombre = item["item"]

                    try:
                        despachados = int(float(item["despachados"]))
                        stock = int(float(item["stock"]))
                    except (ValueError, TypeError):
                        continue

                    remanente = despachados
                    if remanente > 0:
                        if nne in data:
                            data[nne]["remanente"] += remanente
                        else:
                            data[nne] = {"item": nombre, "remanente": remanente}

                except Exception as e:
                    print(f"  ❌ ERROR en procesamiento: {e}")
                    continue

        with open(archivo, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=4, ensure_ascii=False)

        print(f"{Fore.MAGENTA}Remanentes actualizados en {archivo}{Style.RESET_ALL}")

    except Exception as e:
        log_error(f"Error al registrar remanentes globales: {e}")
        
cantidad_remanentes = sum(1 for r in log_completo["resultados"] if r["estado"] == "STOCK INSUFICIENTE")
registrar_remanentes_global(log_completo)


print(f"\n{Fore.BLUE}Reporte JSON guardado en: {nombre_archivo}{Style.RESET_ALL}")
print(f"{Fore.CYAN}Total ítems procesados: {len(log_completo['resultados'])}{Style.RESET_ALL}")


print("Autenticando...")
gauth = GoogleAuth()
print("Cargando config del cliente...")
gauth.LoadClientConfigFile("client_secret.json")

print("Cargando credenciales...")
gauth.LoadCredentialsFile("mycreds.txt")

if gauth.credentials is None:
    print("No hay credenciales. Lanzando auth por navegador...")
    gauth.LocalWebserverAuth()
elif gauth.access_token_expired:
    print("Token expirado. Refrescando...")
    gauth.Refresh()
else:
    print("Token válido. Autorizando...")
    gauth.Authorize()

gauth.SaveCredentialsFile("mycreds.txt")
print("Autenticación completada.")

drive = GoogleDrive(gauth)
print("Instancia de GoogleDrive creada.")

# ID de la carpeta "Reportes"
folder_id = "1Sk1cjwqzQF4Fqb1JOndPtplT8OK9bl6-"

print("Subiendo archivo...")
archivo_drive = drive.CreateFile({
    'title': os.path.basename(nombre_archivo),
    'parents': [{'id': folder_id}]
})
archivo_drive.SetContentFile(nombre_archivo)
archivo_drive.Upload()
print(f"{Fore.CYAN}✅ Subido correctamente: {archivo_drive['title']}{Style.RESET_ALL}")

input("Presioná Enter para cerrar...")
