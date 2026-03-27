"""
Script para inspeccionar la página DESPUÉS del login
Navegará por los menús y mostrará enlaces encontrados
"""

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import time


def explorar_menu():
    """Navega a través del menú post-login y muestra toda la estructura"""
    
    print("[*] Inicializando WebDriver...")
    options = webdriver.ChromeOptions()
    
    driver = webdriver.Chrome(
        service=Service(ChromeDriverManager().install()),
        options=options
    )
    
    wait = WebDriverWait(driver, 20)
    
    try:
        # LOGIN
        print("[*] Accediendo a RUB Online...")
        driver.get("https://rubonline.icbf.gov.co/DefaultF.aspx")
        time.sleep(3)
        
        print("[*] Realizando login...")
        campo_usuario = wait.until(EC.presence_of_element_located((By.ID, "UserName")))
        campo_usuario.send_keys("angie.cardenas")
        
        campo_password = driver.find_element(By.ID, "Password")
        campo_password.send_keys("Celeste1020*")
        
        boton_login = driver.find_element(By.ID, "LoginButton")
        boton_login.click()
        
        print("[+] Login completado, esperando carga de página...")
        time.sleep(8)
        
        print(f"\n✓ URL actual: {driver.current_url}")
        print(f"✓ Título: {driver.title}")
        
        # EXPLORACIÓN DE MENÚ
        print("\n" + "="*70)
        print("BÚSQUEDA DE ELEMENTOS DEL MENÚ")
        print("="*70)
        
        # Buscar todos los enlaces/elementos con texto
        print("\n[*] Todos los enlaces visibles:")
        enlaces = driver.find_elements(By.TAG_NAME, "a")
        for idx, enlace in enumerate(enlaces[:30]):  # Primeros 30
            texto = enlace.text.strip()
            href = enlace.get_attribute("href")
            onclick = enlace.get_attribute("onclick")
            id_attr = enlace.get_attribute("id")
            
            if texto:
                print(f"\n  [{idx}] {texto}")
                print(f"      ID: {id_attr}")
                if href:
                    print(f"      HREF: {href}")
                if onclick:
                    print(f"      ONCLICK: {onclick[:80]}...")
        
        # Buscar por texto específico
        print("\n" + "="*70)
        print("BÚSQUEDA ESPECÍFICA DE PALABRAS CLAVE")
        print("="*70)
        
        palabras_clave = ["Beneficiario", "Opción", "Nuevo", "Agregar", "Plus", "More"]
        
        for palabra in palabras_clave:
            elementos = driver.find_elements(By.XPATH, f"//*[contains(text(), '{palabra}')]")
            if elementos:
                print(f"\n✓ Encontrados {len(elementos)} elementos con '{palabra}':")
                for elem in elementos[:5]:
                    print(f"    - {elem.text.strip()[:50]}")
                    print(f"      Tag: {elem.tag_name}")
                    print(f"      ID: {elem.get_attribute('id')}")
            else:
                print(f"\n✗ No encontrado: '{palabra}'")
        
        # Buscar botones e inputs visibles
        print("\n" + "="*70)
        print("BOTONES E INPUTS VISIBLES")
        print("="*70)
        
        inputs = driver.find_elements(By.TAG_NAME, "input")
        print(f"\nInputs encontrados: {len(inputs)}")
        for inp in inputs[:20]:
            tipo = inp.get_attribute("type")
            id_attr = inp.get_attribute("id")
            name = inp.get_attribute("name")
            placeholder = inp.get_attribute("placeholder")
            
            if tipo not in ["hidden"]:
                print(f"  - {id_attr or name} ({tipo}) placeholder: {placeholder}")
        
        botones = driver.find_elements(By.TAG_NAME, "button")
        print(f"\nBotones encontrados: {len(botones)}")
        for btn in botones[:10]:
            print(f"  - {btn.get_attribute('id') or btn.text}")
        
        # Buscar imágenes con "+" o similar
        print("\n" + "="*70)
        print("IMÁGENES Y ICONOS")
        print("="*70)
        
        imgs = driver.find_elements(By.TAG_NAME, "img")
        print(f"\nImágenes encontradas: {len(imgs)}")
        for img in imgs:
            src = img.get_attribute("src")
            alt = img.get_attribute("alt")
            title = img.get_attribute("title")
            
            if "plus" in src.lower() or "add" in src.lower() or "+" in alt or "+" in title:
                print(f"  ✓ {alt or title or src}")
        
        # Guardar HTML
        print("\n[*] Guardando HTML post-login...")
        with open("pagina_post_login.html", "w", encoding="utf-8") as f:
            f.write(driver.page_source)
        print("[+] HTML guardado en: pagina_post_login.html")
        
        print("\n" + "="*70)
        print("✅ Exploración completada")
        print("="*70)
        print("\n[*] El navegador permanecerá abierto por 60 segundos para inspección manual")
        print("[*] Puedes usar F12 para abrir DevTools inspector")
        
        time.sleep(60)
        
    except Exception as e:
        print(f"\n[-] Error: {str(e)}")
        import traceback
        traceback.print_exc()
    
    finally:
        driver.quit()
        print("\n[+] WebDriver cerrado")


if __name__ == "__main__":
    explorar_menu()
