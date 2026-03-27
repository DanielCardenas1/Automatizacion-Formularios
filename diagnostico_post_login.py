"""
Script de Diagnóstico - Inspecciona la página DESPUÉS del login
"""

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import time


def diagnosticar_pagina_post_login():
    """Accede a la página, hace login y luego inspecciona la estructura"""
    
    print("[*] Inicializando WebDriver...")
    options = webdriver.ChromeOptions()
    
    driver = webdriver.Chrome(
        service=Service(ChromeDriverManager().install()),
        options=options
    )
    
    wait = WebDriverWait(driver, 20)
    
    try:
        url = "https://rubonline.icbf.gov.co/DefaultF.aspx"
        print(f"\n[*] Accediendo a {url}...")
        driver.get(url)
        time.sleep(3)
        
        # LOGIN
        print("\n[*] Realizando login...")
        
        campo_usuario = wait.until(EC.presence_of_element_located((By.ID, "UserName")))
        campo_usuario.send_keys("angie.cardenas")
        
        campo_password = driver.find_element(By.ID, "Password")
        campo_password.send_keys("Celeste1020*")
        
        boton_login = driver.find_element(By.ID, "LoginButton")
        boton_login.click()
        
        print("[+] Esperando carga de página post-login...")
        time.sleep(8)
        
        print(f"\n[+] POST-LOGIN - Título: {driver.title}")
        print(f"[+] URL actual: {driver.current_url}")
        
        # DIAGNÓSTICO POST-LOGIN
        print("\n" + "="*60)
        print("ESTRUCTURA DE LA PÁGINA POST-LOGIN")
        print("="*60)
        
        # Buscar inputs y sus atributos
        print("\nCAMPOS DE BÚSQUEDA/ENTRADA:")
        inputs = driver.find_elements(By.TAG_NAME, "input")
        for idx, inp in enumerate(inputs[:20]):  # Primeros 20
            id_attr = inp.get_attribute("id")
            name = inp.get_attribute("name")
            tipo = inp.get_attribute("type")
            value = inp.get_attribute("value")
            placeholder = inp.get_attribute("placeholder")
            
            if tipo not in ["hidden"]:
                print(f"\n  Input #{idx}")
                print(f"    ID: {id_attr}")
                print(f"    Name: {name}")
                print(f"    Type: {tipo}")
                print(f"    Placeholder: {placeholder}")
        
        # Buscar dropdowns/select
        print("\n\nDROPDOWNS/SELECT:")
        selects = driver.find_elements(By.TAG_NAME, "select")
        for idx, select in enumerate(selects):
            id_attr = select.get_attribute("id")
            name = select.get_attribute("name")
            print(f"\n  Select #{idx}")
            print(f"    ID: {id_attr}")
            print(f"    Name: {name}")
        
        # Buscar botones
        print("\n\nBOTONES:")
        botones = driver.find_elements(By.TAG_NAME, "button")
        for idx, boton in enumerate(botones):
            id_attr = boton.get_attribute("id")
            texto = boton.text
            print(f"  • Button {idx}: {id_attr} - {texto}")
        
        # Buscar tablas/grillas
        print("\n\nTABLAS/GRILLAS:")
        tablas = driver.find_elements(By.TAG_NAME, "table")
        for idx, tabla in enumerate(tablas):
            id_attr = tabla.get_attribute("id")
            filas = tabla.find_elements(By.TAG_NAME, "tr")
            print(f"\n  Table #{idx}")
            print(f"    ID: {id_attr}")
            print(f"    Filas: {len(filas)}")
            
            # Mostrar encabezados
            if filas:
                encabezados = filas[0].find_elements(By.TAG_NAME, "th")
                if encabezados:
                    print(f"    Encabezados: {[th.text for th in encabezados[:5]]}")
        
        # Buscar divs con content o data
        print("\n\nPRINCIPALES DIVs:")
        divs = driver.find_elements(By.TAG_NAME, "div")
        for div in divs[:10]:
            id_attr = div.get_attribute("id")
            clase = div.get_attribute("class")
            if id_attr or "content" in clase or "data" in clase or "list" in clase:
                print(f"  • ID: {id_attr} | Class: {clase}")
        
        # Guardar HTML completo
        print("\n\n[*] Guardando HTML post-login...")
        with open("pagina_debug_post_login.html", "w", encoding="utf-8") as f:
            f.write(driver.page_source)
        
        print("[+] HTML guardado en: pagina_debug_post_login.html")
        
        print("\n" + "="*60)
        print("✅ Diagnóstico completado")
        print("="*60)
        
        # Mantener el navegador abierto por 30 segundos para inspeccionar manualmente
        print("\n[*] Navegador abierto - puedes inspeccionar manualmente")
        print("[*] Se cerrará en 30 segundos...")
        time.sleep(30)
        
    except Exception as e:
        print(f"\n[-] Error: {str(e)}")
        import traceback
        traceback.print_exc()
    
    finally:
        print("\n[*] Cerrando navegador...")
        driver.quit()
        print("[+] WebDriver cerrado")


if __name__ == "__main__":
    diagnosticar_pagina_post_login()
