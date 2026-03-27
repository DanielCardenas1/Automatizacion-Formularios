#!/usr/bin/env python3
"""
BOT INTERACTIVO SIMPLIFICADO
Pasos claros en la terminal
"""

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import time
import os

def esperar_enterfor(mensaje):
    """Espera que presiones ENTER"""
    input(f"\n⏸ {mensaje}\n")

def bot_simple():
    print("\n" + "="*80)
    print("BOT RUB ONLINE - MODO INTERACTIVO")
    print("="*80)
    
    print("\n[1/3] Inicializando Chrome...")
    options = webdriver.ChromeOptions()
    driver = webdriver.Chrome(
        service=Service(ChromeDriverManager().install()),
        options=options
    )
    
    wait = WebDriverWait(driver, 20)
    
    try:
        # PASO 1: LOGIN
        print("\n" + "="*80)
        print("PASO 1: ACCESO Y LOGIN")
        print("="*80)
        
        print("\n→ Abriendo https://rubonline.icbf.gov.co/DefaultF.aspx")
        driver.get("https://rubonline.icbf.gov.co/DefaultF.aspx")
        time.sleep(2)
        
        print("→ Rellenando usuario: Usuario")
        campo_usuario = wait.until(EC.presence_of_element_located((By.ID, "UserName")))
        campo_usuario.send_keys("Usuario")
        time.sleep(1)
        
        print("→ Rellenando contraseña")
        campo_password = driver.find_element(By.ID, "Password")
        campo_password.send_keys("Contraseña")
        time.sleep(1)
        
        print("→ Presionando botón LOGIN")
        boton_login = driver.find_element(By.ID, "LoginButton")
        boton_login.click()
        time.sleep(5)
        
        print("✓ Login completado\n")
        
        # PASO 2: NAVEGACIÓN MANUAL
        print("="*80)
        print("PASO 2: NAVEGACIÓN MANUAL")
        print("="*80)
        
        print("""
✓ Chrome está ABIERTO y visible
✓ Deberías ver la página de RUB Online

INSTRUCCIONES EN CHROME:
  1. En el menú izquierdo, busca "Beneficiario"
  2. Haz click en: RUB ONLINE → Rub online → Beneficiario → Beneficiario
  3. Cuando veas la página de Beneficiario cargada...
  
""")
        esperar_enterfor("Presiona ENTER cuando hayas llegado a la página de Beneficiario")
        
        # PASO 3: ANÁLISIS
        print("\n" + "="*80)
        print("PASO 3: ANÁLISIS DE LA PÁGINA")
        print("="*80)
        
        print(f"\nURL actual: {driver.current_url}")
        print(f"Título: {driver.title}")
        
        # Entrar al iframe
        try:
            iframe = driver.find_element(By.ID, "frameContent")
            driver.switch_to.frame(iframe)
            
            print("\n✓ Estoy dentro del iframe de Beneficiario")
            
            # ANÁLISIS: Dropdowns
            selects = driver.find_elements(By.TAG_NAME, "select")
            print(f"\n📋 DROPDOWNS ENCONTRADOS: {len(selects)}")
            for idx, sel in enumerate(selects):
                name = sel.get_attribute("name")
                id_attr = sel.get_attribute("id")
                print(f"   [{idx}] Name: {name} | ID: {id_attr}")
            
            # ANÁLISIS: Inputs
            inputs = driver.find_elements(By.TAG_NAME, "input")
            input_count = 0
            for inp in inputs:
                if inp.get_attribute("type") not in ["hidden"]:
                    input_count += 1
            
            print(f"\n📝 INPUTS VISIBLES: {input_count}")
            for idx, inp in enumerate(driver.find_elements(By.TAG_NAME, "input")):
                if inp.get_attribute("type") not in ["hidden"]:
                    name = inp.get_attribute("name")
                    id_attr = inp.get_attribute("id")
                    tipo = inp.get_attribute("type")
                    print(f"   [{idx}] Name: {name} | ID: {id_attr} | Type: {tipo}")
            
            # ANÁLISIS: Botones
            botones = driver.find_elements(By.TAG_NAME, "button")
            print(f"\n🔘 BOTONES: {len(botones)}")
            for idx, btn in enumerate(botones):
                texto = btn.text.strip()
                id_attr = btn.get_attribute("id")
                print(f"   [{idx}] ID: {id_attr} | Texto: '{texto}'")
            
            # ANÁLISIS: Links / Imágenes clicables
            enlaces = driver.find_elements(By.TAG_NAME, "a")
            print(f"\n🔗 ENLACES: {len(enlaces)}")
            for idx, enlace in enumerate(enlaces):
                texto = enlace.text.strip()
                href = enlace.get_attribute("href")
                title = enlace.get_attribute("title")
                if texto or title:
                    print(f"   [{idx}] Texto: '{texto[:30]}' | Title: '{title}' | HREF: {href[:40]}")
            
            # ANÁLISIS: Imágenes (posible símbolo +)
            imgs = driver.find_elements(By.TAG_NAME, "img")
            print(f"\n🖼️  IMÁGENES: {len(imgs)}")
            for idx, img in enumerate(imgs):
                src = img.get_attribute("src")
                alt = img.get_attribute("alt")
                title = img.get_attribute("title")
                
                if alt or "+" in src.lower() or "add" in src.lower() or "new" in src.lower():
                    print(f"   [{idx}] Alt: '{alt}' | Title: '{title}'")
                    print(f"           Src: {src[-60:]}")
            
            driver.switch_to.default_content()
            
        except Exception as e:
            print(f"[-] Error: {e}")
            driver.switch_to.default_content()
        
        # PASO 4: PREGUNTAS
        print("\n" + "="*80)
        print("PASO 4: PREGUNTAS")
        print("="*80)
        
        print("""
Basándome en lo anterior, necesito saber:

1. ¿Ves un dropdown o campo llamado "Opción Beneficiario"?
   (probablemente en los DROPDOWNS listados arriba)
   
2. ¿Ves un símbolo "+" que sea un botón o imagen?
   (probablemente en los ENLACES o IMÁGENES listados arriba)

3. ¿Cuáles son sus NOMBRES o ÍNDICES?
   (ejemplo: "Dropdown [0]" o "Botón [2]")
""")
        
        respuesta = input("\nEscribe tus observaciones:\n▶ ")
        
        print(f"\nTu respuesta: {respuesta}")
        print("\nVoy a guardar esta información para actualizar el bot automático")
        
        # GUARDAR HTML
        print("\n[*] Guardando página HTML para análisis...")
        try:
            iframe = driver.find_element(By.ID, "frameContent")
            driver.switch_to.frame(iframe)
            html = driver.page_source
            driver.switch_to.default_content()
            
            with open("pagina_beneficiario_capturada.html", "w", encoding="utf-8") as f:
                f.write(html)
            print("[+] HTML guardado: pagina_beneficiario_capturada.html")
        except:
            pass
        
        # MANTENER ABIERTO
        print("\n" + "="*80)
        print("✓ PROCESO COMPLETADO")
        print("="*80)
        print("\nChrome permanecerá abierto por 60 segundos")
        print("Puedes seguir explorando manualmente si necesitas")
        print("Presiona Ctrl+C para cerrar")
        
        time.sleep(60)
        
    except KeyboardInterrupt:
        print("\n\n[*] Cerrado por usuario")
    except Exception as e:
        print(f"\n[-] Error: {e}")
        import traceback
        traceback.print_exc()
    finally:
        print("\n[*] Cerrando navegador...")
        driver.quit()
        print("[+] Hecho")

if __name__ == "__main__":
    bot_simple()
