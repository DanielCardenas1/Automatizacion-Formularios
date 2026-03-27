#!/usr/bin/env python3
"""
BOT ULTRA SIMPLE - Solo abre Chrome y espera
"""

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import time

print("="*80)
print("ABRIENDO RUB ONLINE")
print("="*80)

options = webdriver.ChromeOptions()
driver = webdriver.Chrome(
    service=Service(ChromeDriverManager().install()),
    options=options
)

wait = WebDriverWait(driver, 20)

try:
    print("[1] Abriendo página...")
    driver.get("https://rubonline.icbf.gov.co/DefaultF.aspx")
    time.sleep(2)
    
    print("[2] Ingresando usuario...")
    campo_usuario = wait.until(EC.presence_of_element_located((By.ID, "UserName")))
    campo_usuario.send_keys("angie.cardenas")
    time.sleep(1)
    
    print("[3] Ingresando contraseña...")
    campo_password = driver.find_element(By.ID, "Password")
    campo_password.send_keys("Celeste1020*")
    time.sleep(1)
    
    print("[4] Presionando LOGIN...")
    boton_login = driver.find_element(By.ID, "LoginButton")
    boton_login.click()
    time.sleep(6)
    
    print("\n" + "="*80)
    print("✓ LOGIN COMPLETADO")
    print("="*80)
    print("\nChrome está abierto. Ahora:")
    print("1. En el menú izquierdo, haz click en: RUB ONLINE")
    print("2. Luego: Rub online")
    print("3. Luego: Beneficiario")
    print("4. Luego: Beneficiario (el primer enlace de la lista)")
    print("\nCuando estés en la página de Beneficiario, presiona ENTER aquí...")
    
    input("\n")
    
    print("\n" + "="*80)
    print("ANALIZANDO PÁGINA...")
    print("="*80)
    
    # Entrar al iframe
    iframe = driver.find_element(By.ID, "frameContent")
    driver.switch_to.frame(iframe)
    
    # DROPDOWNS
    selects = driver.find_elements(By.TAG_NAME, "select")
    print(f"\n📋 DROPDOWNS: {len(selects)}")
    for idx, sel in enumerate(selects):
        name = sel.get_attribute("name") or "sin-nombre"
        id_attr = sel.get_attribute("id") or "sin-id"
        print(f"   [{idx}] Name: '{name}' | ID: '{id_attr}'")
    
    # INPUTS
    inputs = driver.find_elements(By.TAG_NAME, "input")
    visible_inputs = [i for i in inputs if i.get_attribute("type") not in ["hidden"]]
    print(f"\n📝 INPUTS VISIBLES: {len(visible_inputs)}")
    for idx, inp in enumerate(visible_inputs):
        name = inp.get_attribute("name") or "sin-nombre"
        id_attr = inp.get_attribute("id") or "sin-id"
        tipo = inp.get_attribute("type")
        print(f"   [{idx}] Name: '{name}' | ID: '{id_attr}' | Type: '{tipo}'")
    
    # BOTONES
    botones = driver.find_elements(By.TAG_NAME, "button")
    print(f"\n🔘 BOTONES: {len(botones)}")
    for idx, btn in enumerate(botones):
        texto = btn.text.strip() or "sin-texto"
        id_attr = btn.get_attribute("id") or "sin-id"
        title = btn.get_attribute("title") or ""
        print(f"   [{idx}] ID: '{id_attr}' | Texto: '{texto}' | Title: '{title}'")
    
    # ENLACES
    enlaces = driver.find_elements(By.TAG_NAME, "a")
    print(f"\n🔗 ENLACES: {len(enlaces)}")
    for idx, enlace in enumerate(enlaces):
        texto = enlace.text.strip() or "sin-texto"
        title = enlace.get_attribute("title") or ""
        id_attr = enlace.get_attribute("id") or "sin-id"
        if texto or title or "+" in id_attr:
            print(f"   [{idx}] Texto: '{texto}' | ID: '{id_attr}' | Title: '{title}'")
    
    # IMÁGENES con "+"
    imgs = driver.find_elements(By.TAG_NAME, "img")
    print(f"\n🖼️  IMÁGENES (total: {len(imgs)})")
    for idx, img in enumerate(imgs):
        src = img.get_attribute("src")
        alt = img.get_attribute("alt") or ""
        title = img.get_attribute("title") or ""
        
        # Solo mostrar las que tengan contenido relevante
        if "+" in alt or "plus" in src.lower() or "add" in src.lower() or alt or title:
            print(f"   [{idx}] Alt: '{alt}' | Title: '{title}'")
            print(f"          Src: {src[-40:]}")
    
    driver.switch_to.default_content()
    
    print("\n" + "="*80)
    print("ANÁLISIS COMPLETADO")
    print("="*80)
    print("\nAhora dime:")
    print("• ¿Ves el dropdown 'Opción Beneficiario'? ¿Cuál es su índice [x]?")
    print("• ¿Ves un símbolo '+'? ¿En qué sección? (BOTONES/ENLACES/IMÁGENES)")
    print("• ¿Cuál es su índice [x]?")
    
    print("\nChrome seguirá abierto por 300 segundos para que explores")
    print("Presiona Ctrl+C para cerrar")
    
    time.sleep(300)

except KeyboardInterrupt:
    print("\n✓ Cerrado")
except Exception as e:
    print(f"\n✗ Error: {e}")
    import traceback
    traceback.print_exc()
finally:
    driver.quit()
