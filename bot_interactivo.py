"""
Bot INTERACTIVO - Navega manualmente y muestra dónde hacer click
"""

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import time


def bot_interactivo():
    """Bot que permite navegación interactiva para capturar ubicaciones de clicks"""
    
    print("[*] Inicializando WebDriver en modo VISIBLE...")
    options = webdriver.ChromeOptions()
    # NO usar headless - queremos ver la ventana
    
    driver = webdriver.Chrome(
        service=Service(ChromeDriverManager().install()),
        options=options
    )
    
    wait = WebDriverWait(driver, 20)
    
    try:
        # LOGIN AUTOMÁTICO
        print("\n" + "="*70)
        print("FASE 1: LOGIN AUTOMÁTICO")
        print("="*70)
        
        print("[*] Accediendo a RUB Online...")
        driver.get("https://rubonline.icbf.gov.co/DefaultF.aspx")
        time.sleep(3)
        
        print("[*] Ingresando credenciales...")
        campo_usuario = wait.until(EC.presence_of_element_located((By.ID, "UserName")))
        campo_usuario.send_keys("angie.cardenas")
        
        campo_password = driver.find_element(By.ID, "Password")
        campo_password.send_keys("Celeste1020*")
        
        boton_login = driver.find_element(By.ID, "LoginButton")
        boton_login.click()
        
        print("[+] Login completado\n")
        time.sleep(8)
        
        # NAVEGACIÓN INTERACTIVA
        print("="*70)
        print("FASE 2: NAVEGACIÓN INTERACTIVA")
        print("="*70)
        print("\n✓ La ventana del navegador está ABIERTA y visible")
        print("✓ Puedes navegar manualmente dentro")
        print("\nINSTRUCCIONES:")
        print("  1. Abre la consola Python con F12 en Chrome")
        print("  2. Navega a: RUB ONLINE > Rub online > Beneficiario > Beneficiario")
        print("  3. Luego presiona ENTER aquí para continuar")
        print("  4. El bot te mostrará los elementos clickeables\n")
        
        # PAUSA interactiva
        input("▶ Presiona ENTER cuando hayas navegado a Beneficiario...")
        
        print("\n[*] Continuando con la inspección...")
        time.sleep(2)
        
        # ANÁLISIS DEL ESTADO ACTUAL
        print("\n" + "="*70)
        print("FASE 3: ANÁLISIS DE LA PÁGINA ACTUAL")
        print("="*70)
        
        url_actual = driver.current_url
        titulo = driver.title
        
        print(f"\n✓ URL actual: {url_actual}")
        print(f"✓ Título: {titulo}")
        
        # Buscar el iframe
        try:
            iframe = driver.find_element(By.ID, "frameContent")
            print("✓ Iframe 'frameContent' encontrado")
            
            driver.switch_to.frame(iframe)
            print("✓ Estoy dentro del iframe")
            
            # Buscar campos específicos
            print("\n[*] Buscando campos en la página Beneficiario:")
            
            # Campos de entrada
            inputs = driver.find_elements(By.TAG_NAME, "input")
            print(f"\n  Inputs encontrados: {len(inputs)}")
            for idx, inp in enumerate(inputs):
                id_attr = inp.get_attribute("id")
                name = inp.get_attribute("name")
                tipo = inp.get_attribute("type")
                placeholder = inp.get_attribute("placeholder")
                
                if tipo not in ["hidden"]:
                    print(f"    [{idx}] ID: {id_attr} | Name: {name} | Type: {tipo} | Placeholder: {placeholder}")
            
            # Dropdowns
            selects = driver.find_elements(By.TAG_NAME, "select")
            print(f"\n  Selects/Dropdowns encontrados: {len(selects)}")
            for idx, sel in enumerate(selects):
                id_attr = sel.get_attribute("id")
                name = sel.get_attribute("name")
                print(f"    [{idx}] ID: {id_attr} | Name: {name}")
            
            # Botones
            botones = driver.find_elements(By.TAG_NAME, "button")
            print(f"\n  Botones encontrados: {len(botones)}")
            for idx, btn in enumerate(botones):
                texto = btn.text
                id_attr = btn.get_attribute("id")
                title = btn.get_attribute("title")
                print(f"    [{idx}] ID: {id_attr} | Texto: '{texto}' | Title: '{title}'")
            
            # Imágenes (incluyendo posibles botones de +)
            imgs = driver.find_elements(By.TAG_NAME, "img")
            print(f"\n  Imágenes encontradas: {len(imgs)}")
            for idx, img in enumerate(imgs[:10]):
                src = img.get_attribute("src")
                alt = img.get_attribute("alt")
                title = img.get_attribute("title")
                
                if alt or title:
                    print(f"    [{idx}] Src: {src[-30:]} | Alt: '{alt}' | Title: '{title}'")
            
            driver.switch_to.default_content()
            
        except Exception as e:
            print(f"[-] Error con iframe: {e}")
            driver.switch_to.default_content()
        
        # SIGUIENTE PASO
        print("\n" + "="*70)
        print("FASE 4: ¿QUÉ DESEAS HACER?")
        print("="*70)
        print("\nOpciones:")
        print("  1. Guardar inspector HTML de esta página")
        print("  2. Mostrar todos los enlaces de la página")
        print("  3. Buscar campo específico por nombre")
        print("  4. Hacer click en elemento por índice")
        print("  5. Salir\n")
        
        opcion = input("Elige una opción (1-5): ").strip()
        
        if opcion == "1":
            with open("pagina_beneficiario_debug.html", "w", encoding="utf-8") as f:
                f.write(driver.page_source)
            print("[+] HTML guardado en: pagina_beneficiario_debug.html")
        
        elif opcion == "2":
            print("\n[*] Buscando todos los enlaces visibles:\n")
            
            try:
                iframe = driver.find_element(By.ID, "frameContent")
                driver.switch_to.frame(iframe)
                
                enlaces = driver.find_elements(By.TAG_NAME, "a")
                for idx, enlace in enumerate(enlaces):
                    texto = enlace.text.strip()
                    href = enlace.get_attribute("href")
                    if texto or href:
                        print(f"  [{idx}] Texto: '{texto[:40]}' | HREF: {href[:50]}")
                
                driver.switch_to.default_content()
            except Exception as e:
                print(f"[-] Error: {e}")
        
        elif opcion == "3":
            nombre_campo = input("\n¿Nombre del campo a buscar? ").strip()
            print(f"\n[*] Buscando campo '{nombre_campo}'...\n")
            
            try:
                iframe = driver.find_element(By.ID, "frameContent")
                driver.switch_to.frame(iframe)
                
                # Buscar por name, id, placeholder
                elementos = driver.find_elements(By.XPATH, 
                    f"//*[contains(@name, '{nombre_campo}') or contains(@id, '{nombre_campo}') or contains(@placeholder, '{nombre_campo}')]")
                
                print(f"[+] Encontrados {len(elementos)} elemento(s):")
                for idx, elem in enumerate(elementos):
                    print(f"\n  [{idx}]")
                    print(f"    Tag: {elem.tag_name}")
                    print(f"    ID: {elem.get_attribute('id')}")
                    print(f"    Name: {elem.get_attribute('name')}")
                    print(f"    Tipo: {elem.get_attribute('type')}")
                    print(f"    Texto: {elem.text[:50]}")
                
                driver.switch_to.default_content()
            except Exception as e:
                print(f"[-] Error: {e}")
        
        elif opcion == "4":
            nombreque_elemento = input("\n¿Tipo de elemento? (input/select/button/a): ").strip().lower()
            indice = input("¿Índice a clickear?: ").strip()
            
            try:
                indice = int(indice)
                iframe = driver.find_element(By.ID, "frameContent")
                driver.switch_to.frame(iframe)
                
                elementos = driver.find_elements(By.TAG_NAME, nombreque_elemento)
                if indice < len(elementos):
                    elemento = elementos[indice]
                    print(f"\n[*] Clickeando elemento #{indice}...")
                    elemento.click()
                    print("[+] Click ejecutado")
                    time.sleep(2)
                else:
                    print(f"[-] Índice {indice} fuera de rango")
                
                driver.switch_to.default_content()
            except Exception as e:
                print(f"[-] Error: {e}")
        
        print("\n[*] Navegador abierto por 120 segundos más para inspección manual")
        print("[*] Usa F12 para abrir DevTools si necesitas")
        time.sleep(120)
        
    except Exception as e:
        print(f"\n[-] Error: {str(e)}")
        import traceback
        traceback.print_exc()
    
    finally:
        print("\n[+] Cerrando sesión...")
        driver.quit()
        print("[+] WebDriver cerrado")


if __name__ == "__main__":
    bot_interactivo()
