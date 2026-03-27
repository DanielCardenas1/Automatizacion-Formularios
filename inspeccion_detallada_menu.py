"""
Script para inspeccionar exactamente los elementos del menú
"""

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import time


def inspeccionar_menu():
    """Inspecciona el menú después del login"""
    
    print("[*] Inicializando WebDriver...")
    options = webdriver.ChromeOptions()
    
    driver = webdriver.Chrome(
        service=Service(ChromeDriverManager().install()),
        options=options
    )
    
    wait = WebDriverWait(driver, 20)
    
    try:
        # LOGIN
        print("[*] Accediendo y haciendo login...")
        driver.get("https://rubonline.icbf.gov.co/DefaultF.aspx")
        time.sleep(3)
        
        campo_usuario = wait.until(EC.presence_of_element_located((By.ID, "UserName")))
        campo_usuario.send_keys("Usuario")
        
        campo_password = driver.find_element(By.ID, "Password")
        campo_password.send_keys("Contraseña")
        
        boton_login = driver.find_element(By.ID, "LoginButton")
        boton_login.click()
        
        print("[+] Esperando carga de página...")
        time.sleep(8)
        
        # INSPECCIONAR MENÚ
        print("\n" + "="*70)
        print("INSPECCIÓN DEL MENÚ")
        print("="*70)
        
        # Buscar el ul del menú
        ul_menu = driver.find_element(By.ID, "ulMenuPrincipal")
        print(f"\n✓ Menú principal encontrado")
        
        # Obtener todos los <a> del menú que contengan "Beneficiario"
        print("\n[*] Buscando todos los elementos con 'Beneficiario':")
        
        # Primera búsqueda general
        elementos = driver.find_elements(By.XPATH, 
            "//ul[@id='ulMenuPrincipal']//a[contains(text(), 'Beneficiario')]")
        
        print(f"\nEncontrados {len(elementos)} elementos con 'Beneficiario':")
        for idx, elem in enumerate(elementos):
            texto = elem.text
            href = elem.get_attribute("href")
            clase = elem.get_attribute("class")
            parent = elem.find_element(By.XPATH, "..").tag_name
            
            print(f"\n  [{idx}] Texto: '{texto}'")
            print(f"      Href: {href}")
            print(f"      Class: {clase}")
            print(f"      Parent tag: {parent}")
        
        # Buscar específicamente el que tiene href con BENEFICIARIO
        print("\n" + "="*70)
        print("BÚSQUEDA ESPECÍFICA DEL ENLACE CLICABLE")
        print("="*70)
        
        enlace_beneficiario = driver.find_elements(By.XPATH, 
            "//a[@class='EstiloMenuUla' and contains(@href, 'BENEFICIARIO')]")
        
        print(f"\nEncontrados {len(enlace_beneficiario)} enlaces con href BENEFICIARIO:")
        for idx, elem in enumerate(enlace_beneficiario):
            print(f"\n  [{idx}] Texto: '{elem.text}'")
            print(f"      Href: {elem.get_attribute('href')}")  
            print(f"      Class: {elem.get_attribute('class')}")
            print(f"      ID: {elem.get_attribute('id')}")
            
            # Intentar hacer click
            if idx == 0:  # Click en el primero
                print(f"\n[*] Intentando click en el primer elemento...")
                try:
                    elem.click()
                    print("[+] Click ejecutado")
                    time.sleep(4)
                    
                    # Ver si se cargó contenido
                    print(f"\n[*] URL después del click: {driver.current_url}")
                    
                    # Inspeccionar el iframe
                    print("\n[*] Contenido del iframe:")
                    try:
                        iframe = driver.find_element(By.ID, "frameContent")
                        driver.switch_to.frame(iframe)
                        
                        # Ver qué hay dentro del iframe
                        body_text = driver.find_element(By.TAG_NAME, "body").text[:200]
                        print(f"    {body_text}")
                        
                        # Buscar campos en el iframe
                        inputs = driver.find_elements(By.TAG_NAME, "input")
                        selects = driver.find_elements(By.TAG_NAME, "select")
                        buttons = driver.find_elements(By.TAG_NAME, "button")
                        
                        print(f"\n    Inputs: {len(inputs)}")
                        print(f"    Selects: {len(selects)}")
                        print(f"    Buttons: {len(buttons)}")
                        
                        # Mostrar detalles de selects
                        for sel in selects[:5]:
                            print(f"      - {sel.get_attribute('name')} (id: {sel.get_attribute('id')})")
                        
                        driver.switch_to.default_content()
                        
                    except Exception as e:
                        print(f"    Error inspeccionar iframe: {e}")
                        driver.switch_to.default_content()
                        
                except Exception as e:
                    print(f"[-] Error en click: {e}")
        
        print("\n" + "="*70)
        print("[+] Inspección completada")
        
        time.sleep(10)
        
    except Exception as e:
        print(f"\n[-] Error general: {str(e)}")
        import traceback
        traceback.print_exc()
    
    finally:
        driver.quit()


if __name__ == "__main__":
    inspeccionar_menu()
