import json
from datetime import datetime
import os
from colorama import Fore, Style, init
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
import time
import re

init(autoreset=True)

# Configurar Chrome en modo depuración
chrome_options = Options()
chrome_options.add_argument('--ignore-certificate-errors')
chrome_options.add_argument('--ignore-ssl-errors')
chrome_options.add_argument('--start-maximized')
chrome_options.debugger_address = "localhost:9229"
driver_path = r"C:\Users\lgv\Desktop\Chelo\chromedriver_win32\chromedriver.exe"
service = Service(driver_path)
driver = webdriver.Chrome(service=service, options=chrome_options)
wait = WebDriverWait(driver, 15)

def capturar_stock(driver):
    try:
        contenedor = driver.find_element(By.ID, "contentwrapper")
        texto = contenedor.get_attribute("innerText")
        match = re.search(r"Stock:\s*([\d\.]+)", texto)
        if match:
            stock_raw = match.group(1).replace(".", "")
            return int(stock_raw)
        else:
            return None
    except Exception as e:
        print(f"{Fore.RED}Error al capturar stock: {e}{Style.RESET_ALL}")
        return None

# Cargar el JSON de remanentes
with open("datos/remanentes.json", "r", encoding="utf-8") as f:
    remanentes = json.load(f)

reporte = []
# Bandera para controlar si venimos de "Seleccionar Otro Insumo"
seleccionar_otro_insumo = False
remanentes_validos = {k: v for k, v in remanentes.items() if v["remanente"] > 0}


for nne, info in remanentes_validos.items():
    item = info["item"]
    cantidad = info["remanente"]

    print(f"\n🔍 Procesando {item} ({nne}) | Remanente: {cantidad}")

    # Paso 1: Hacer clic en "Agregar Insumo" solo si no venimos de "Seleccionar Otro Insumo"
    if not seleccionar_otro_insumo:
        print(f"Haciendo clic en 'Agregar Insumo'...")
        boton_agregar = wait.until(
            EC.element_to_be_clickable((By.LINK_TEXT, "Agregar Insumo"))
        )
        boton_agregar.click()
    else:
        # Reiniciar la bandera para el próximo ítem
        seleccionar_otro_insumo = False

    
    cuadro_busqueda = wait.until(EC.presence_of_element_located((By.NAME, "__buscar")))
    cuadro_busqueda.clear()
    cuadro_busqueda.send_keys(f" {nne}")
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

    try:
        # 🕒 Esperar a que se actualice el área de stock
        
        wait.until(lambda d: "Stock:" in d.find_element(By.ID, "contentwrapper").text)

        stock_actual = capturar_stock(driver)
        if stock_actual is None:
            print(f"{Fore.RED}❌ No pude leer stock de {item} ({nne}){Style.RESET_ALL}")

        # 1) Stock = 0 -> pendiente
        if stock_actual == 0:
            print(f"{Fore.RED}🔴 Pendiente – Stock: {stock_actual} | Remanente: {cantidad}{Style.RESET_ALL}")
            reporte.append({
                "estado": "PENDIENTE",
                "nne": nne,
                "item": item,
                "stock": stock_actual,
                "intento": cantidad,
                "descontado": 0,
                "remanente": cantidad  # el valor se mantiene
            })

            # clic en “Seleccionar Otro Insumo”
            try:
                btn = wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Seleccionar Otro Insumo")))
                btn.click()
                seleccionar_otro_insumo = True
            except:
                print(f"{Fore.RED}No encontré 'Seleccionar Otro Insumo'{Style.RESET_ALL}")
            continue

        # 2) Stock >= remanente -> descuento total
        if stock_actual >= cantidad:
            print(f"{Fore.GREEN}🟢 Descuento Total – Stock: {stock_actual} | Descontado: {cantidad}{Style.RESET_ALL}")
            info["remanente"] = 0
            reporte.append({
                "estado": "TOTAL",
                "nne": nne,
                "item": item,
                "stock": stock_actual,
                "intento": cantidad,
                "descontado": cantidad,
                "remanente": 0  # importante!
            })

            try:
                print("✍️ Completando cantidades...")

                campo_pedida = wait.until(EC.presence_of_element_located((By.NAME, "cant_ped")))
                campo_pedida.clear()
                campo_pedida.send_keys(str(cantidad))
                print(f"→ Pedida: {cantidad}")

                campo_despachada = wait.until(EC.presence_of_element_located((By.NAME, "cant")))
                campo_despachada.clear()
                campo_despachada.send_keys(str(cantidad))
                print(f"→ Despachada: {cantidad}")

                # Enviar formulario
                boton_enviar = wait.until(EC.element_to_be_clickable(
                    (By.XPATH, "//input[@type='submit' and @name='enviar' and @value='Enviar']")
                ))
                driver.execute_script("arguments[0].scrollIntoView();", boton_enviar)
                driver.execute_script("arguments[0].click();", boton_enviar)
                print("🚀 Formulario enviado correctamente.")

                # Saltar pantalla intermedia con “click aquí”
                try:
                    enlace_click_aqui = wait.until(
                        EC.element_to_be_clickable((By.LINK_TEXT, "click aqui"))
                    )
                    enlace_click_aqui.click()
                    print("🔁 Clic en 'click aquí' realizado.")
                except Exception as e:
                    print(f"{Fore.RED}ERROR: No se pudo hacer clic en 'click aquí': {e}{Style.RESET_ALL}")
                    continue

            except Exception as e:
                print(f"{Fore.RED}❌ Error al completar o enviar formulario: {e}{Style.RESET_ALL}")
                continue

        # 3) 0 < Stock < remanente -> descuento parcial
        else:
            intento = cantidad
            descontado = stock_actual
            nuevo_remanente = cantidad - stock_actual
            print(f"{Fore.YELLOW}🟡 Descuento Parcial – Stock: {stock_actual} | Intentado: {intento} | Descontado: {descontado} | Nuevo remanente: {nuevo_remanente}{Style.RESET_ALL}")

            try:
                print("✍️ Completando cantidades para descuento parcial...")

                campo_pedida = wait.until(EC.presence_of_element_located((By.NAME, "cant_ped")))
                campo_pedida.clear()
                campo_pedida.send_keys(str(descontado))

                campo_despachada = wait.until(EC.presence_of_element_located((By.NAME, "cant")))
                campo_despachada.clear()
                campo_despachada.send_keys(str(descontado))

                # Enviar formulario
                boton_enviar = wait.until(EC.element_to_be_clickable(
                    (By.XPATH, "//input[@type='submit' and @name='enviar' and @value='Enviar']")
                ))
                driver.execute_script("arguments[0].scrollIntoView();", boton_enviar)
                driver.execute_script("arguments[0].click();", boton_enviar)
                print("🚀 Formulario enviado correctamente (descuento parcial).")

                # Actualizar remanente
                info["remanente"] = nuevo_remanente

                reporte.append({
                    "estado": "PARCIAL",
                    "nne": nne,
                    "item": item,
                    "stock": stock_actual,
                    "intento": cantidad,
                    "descontado": descontado,
                     "remanente": nuevo_remanente  # <- corregido acá
                })

                # Salir con click aquí
                try:
                    enlace_click_aqui = wait.until(
                        EC.element_to_be_clickable((By.LINK_TEXT, "click aqui"))
                    )
                    enlace_click_aqui.click()
                    print("🔁 Clic en 'click aquí' realizado.")
                except Exception as e:
                    print(f"{Fore.RED}ERROR: No se pudo hacer clic en 'click aquí': {e}{Style.RESET_ALL}")
                    continue

            except Exception as e:
                print(f"{Fore.RED}❌ Error procesando descuento parcial: {e}{Style.RESET_ALL}")
                continue


    except Exception as e:
        print(f"{Fore.RED}❌ Error procesando {item} ({nne}): {e}{Style.RESET_ALL}")

# Ordenar los reportes: TOTAL → PARCIAL → PENDIENTE
reporte_ordenado = (
    [r for r in reporte if r["estado"] == "TOTAL"] +
    [r for r in reporte if r["estado"] == "PARCIAL"] +
    [r for r in reporte if r["estado"] == "PENDIENTE"]
)

print("\n" + Fore.WHITE + "=" * 100 + Style.RESET_ALL)
print(f"{Fore.CYAN}{Style.BRIGHT}                            R E S U M E N   D E   R E M A N E N T E S{Style.RESET_ALL}")
print(Fore.WHITE + "=" * 100 + Style.RESET_ALL)

for r in reporte_ordenado:
    if r["estado"] == "TOTAL":
        print(f"{Fore.GREEN}🟢 {r['item']} ({r['nne']}) → Stock: {r['stock']} | Descontado: {r['descontado']}{Style.RESET_ALL}")
    elif r["estado"] == "PARCIAL":
        print(f"{Fore.YELLOW}🟡 {r['item']} ({r['nne']}) → Stock: {r['stock']} | Intentado: {r['descontado'] + r['remanente']} | Descontado: {r['descontado']} | Remanente: {r['remanente']}{Style.RESET_ALL}")
    elif r["estado"] == "PENDIENTE":
        print(f"{Fore.RED}🔴 {r['item']} ({r['nne']}) → Stock: {r['stock']} | Remanente: {r['remanente']}{Style.RESET_ALL}")

# Aplicar los descuentos al JSON original
for r in reporte:
    if r["estado"] in ("TOTAL", "PARCIAL"):
        nne = r["nne"]
        remanentes[nne]["remanente"] = r["remanente"]

# Guardar el nuevo JSON actualizado solo si todo salió bien
with open("datos/remanentes.json", "w", encoding="utf-8") as f:
    json.dump(remanentes, f, indent=4, ensure_ascii=False)

print(f"{Fore.MAGENTA}📂 Remanentes actualizados en datos/remanentes.json{Style.RESET_ALL}")



# Guardar cambios en el archivo original
with open("datos/remanentes.json", "w", encoding="utf-8") as f:
    json.dump(remanentes, f, indent=4, ensure_ascii=False)

# Guardar reporte
timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
os.makedirs("Reportes/Remanentes", exist_ok=True)
with open(f"Reportes/Remanentes/{timestamp}_remanentes.json", "w", encoding="utf-8") as f:
    json.dump(reporte, f, indent=4, ensure_ascii=False)

print("\n✅ Proceso de remanentes completado.")