import time
import pyautogui
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import WebDriverException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By

# --- CONFIGURACIÓN ---
URL_QLIK = "https://qlik.copservir.com/"
USUARIO = "Qlikzona29"
PASSWORD = "pF2A3f2x*"

def verificar_inicio_sesion(driver):
    """Verifica si la URL cambió o si estamos en el Hub de Qlik."""
    try:
        # Esperamos a que la URL cambie o contenga 'hub' o 'sense'
        time.sleep(5)
        if "hub" in driver.current_url.lower() or "sense" in driver.current_url.lower():
            return True
        return False
    except:
        return False


def esperar_carga_hub(driver, timeout=20):
    """Espera explícita a que la URL cambie indicando que el Hub o Sense está cargado.

    Usa WebDriverWait con una condición lambda que comprueba la URL.
    Devuelve True si detecta el cambio dentro del timeout, False si expira.
    """
    try:
        WebDriverWait(driver, timeout).until(
            lambda d: "hub" in d.current_url.lower() or "sense" in d.current_url.lower()
        )
        return True
    except Exception:
        return False

def login_con_action_chains(driver):
    """Intento 1: Usando la API de Selenium (Más seguro)"""
    print("Intentando acceso con Selenium ActionChains...")
    try:
        actions = ActionChains(driver)
        # Escribir usuario (asumiendo que el cursor ya está ahí)
        actions.send_keys(USUARIO)
        actions.pause(1)
        actions.send_keys(Keys.TAB)
        actions.pause(1)
        # Escribir contraseña
        actions.send_keys(PASSWORD)
        actions.pause(1)
        actions.send_keys(Keys.ENTER)
        actions.perform()
        return verificar_inicio_sesion(driver)
    except Exception as e:
        print(f"Error en ActionChains: {e}")
        return False

def login_con_pyautogui(driver):
    """Intento 2: Usando PyAutoGUI (Control total del teclado del PC)"""
    print("Fallo el primer intento. Recurriendo a PyAutoGUI (Sistema Operativo)...")
    try:
        # Aseguramos que la ventana del navegador tenga el foco real del sistema
        driver.maximize_window()
        time.sleep(2)
        
        # Simulamos pulsaciones físicas
        pyautogui.write(USUARIO, interval=0.1)
        pyautogui.press('tab')
        time.sleep(1)
        pyautogui.write(PASSWORD, interval=0.1)
        pyautogui.press('enter')
        
        return verificar_inicio_sesion(driver)
    except Exception as e:
        print(f"Error crítico en PyAutoGUI: {e}")
        return False

# --- FLUJO PRINCIPAL ---
def ejecutar_automatizacion():
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    driver = webdriver.Chrome(options=options)

    try:
        driver.get(URL_QLIK)
        # Tiempo de espera generoso para carga de Qlik
        print("Cargando página de Qlik...")
        time.sleep(8) 

        # INTENTO 1
        exito = login_con_action_chains(driver)

        # INTENTO 2 (Si el 1 falló)
        if not exito:
            # Refrescamos para limpiar campos si es necesario
            driver.refresh()
            time.sleep(5)
            exito = login_con_pyautogui(driver)

        if exito:
            print("¡Acceso exitoso!")
            # Espera explícita para asegurarnos de que el Hub/Sense esté completamente cargado
            print("Esperando que el Hub se cargue completamente...")
            cargado = esperar_carga_hub(driver, timeout=20)
            if cargado:
                print("Página cargada. Continuando...")
            else:
                print("Aviso: tiempo de espera agotado. La URL no cambió a 'hub' o 'sense'.")
            # Aquí continúa tu lógica de extracción de datos...
        else:
            print("No se pudo acceder después de ambos métodos. Verifique credenciales o conexión.")

    except WebDriverException as e:
        print(f"Error de conexión/navegador: {e}")
    finally:
        # Mantener abierto un momento para ver el resultado antes de cerrar
        time.sleep(5)
        # Descomenta la siguiente línea para cerrar el navegador al terminar
        # driver.quit()

if __name__ == "__main__":
    ejecutar_automatizacion()