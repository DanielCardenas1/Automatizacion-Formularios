#!/usr/bin/env python3
"""
BOT AUTOMÁTICO - Navega automáticamente y guarda análisis en archivo
"""

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import time

print("INICIANDO BOT AUTOMÁTICO...")

options = webdriver.ChromeOptions()
driver = webdriver.Chrome(
    service=Service(ChromeDriverManager().install()),
    options=options
)

wait = WebDriverWait(driver, 20)
resultado = []

try:
    # LOGIN
    print("[1/4] Login...")
    driver.get("https://rubonline.icbf.gov.co/DefaultF.aspx")
    time.sleep(2)
    
    campo_usuario = wait.until(EC.presence_of_element_located((By.ID, "UserName")))
    campo_usuario.send_keys("Usuario")
    
    campo_password = driver.find_element(By.ID, "Password")
    campo_password.send_keys("Contraseña")
    
    boton_login = driver.find_element(By.ID, "LoginButton")
    boton_login.click()
    time.sleep(6)
    
    # NAVEGAR A BENEFICIARIO
    print("[2/4] Navegando a Beneficiario...")
    try:
        enlace = wait.until(
            EC.element_to_be_clickable((By.XPATH, 
            "//a[@class='EstiloMenuUla' and contains(@href, 'BENEFICIARIO')]"))
        )
        enlace.click()
        print("    ✓ Click ejecutado")
        time.sleep(4)
    except Exception as e:
        resultado.append(f"ERROR navegando: {e}")
        print(f"    ERROR: {e}")
    
    # ANALIZAR
    print("[3/4] Analizando página...")
    
    try:
        iframe = driver.find_element(By.ID, "frameContent")
        driver.switch_to.frame(iframe)
        
        # DROPDOWNS
        selects = driver.find_elements(By.TAG_NAME, "select")
        resultado.append(f"\n📋 DROPDOWNS: {len(selects)}")
        for idx, sel in enumerate(selects):
            name = sel.get_attribute("name") or "sin-nombre"
            id_attr = sel.get_attribute("id") or "sin-id"
            resultado.append(f"   [{idx}] Name: '{name}' | ID: '{id_attr}'")
        
        # INPUTS
        inputs = driver.find_elements(By.TAG_NAME, "input")
        visible_inputs = [i for i in inputs if i.get_attribute("type") not in ["hidden"]]
        resultado.append(f"\n📝 INPUTS VISIBLES: {len(visible_inputs)}")
        for idx, inp in enumerate(visible_inputs):
            name = inp.get_attribute("name") or "sin-nombre"
            id_attr = inp.get_attribute("id") or "sin-id"
            tipo = inp.get_attribute("type")
            resultado.append(f"   [{idx}] Name: '{name}' | ID: '{id_attr}' | Type: '{tipo}'")
        
        # BOTONES
        botones = driver.find_elements(By.TAG_NAME, "button")
        resultado.append(f"\n🔘 BOTONES: {len(botones)}")
        for idx, btn in enumerate(botones):
            texto = btn.text.strip() or "sin-texto"
            id_attr = btn.get_attribute("id") or "sin-id"
            title = btn.get_attribute("title") or ""
            resultado.append(f"   [{idx}] ID: '{id_attr}' | Texto: '{texto}' | Title: '{title}'")
        
        # ENLACES
        enlaces = driver.find_elements(By.TAG_NAME, "a")
        resultado.append(f"\n🔗 ENLACES: {len(enlaces)}")
        for idx, enlace in enumerate(enlaces):
            texto = enlace.text.strip() or "sin-texto"
            title = enlace.get_attribute("title") or ""
            id_attr = enlace.get_attribute("id") or "sin-id"
            if texto or title or "+" in id_attr:
                resultado.append(f"   [{idx}] Texto: '{texto}' | ID: '{id_attr}' | Title: '{title}'")
        
        # IMÁGENES
        imgs = driver.find_elements(By.TAG_NAME, "img")
        resultado.append(f"\n🖼️  IMÁGENES (total: {len(imgs)})")
        for idx, img in enumerate(imgs):
            src = img.get_attribute("src")
            alt = img.get_attribute("alt") or ""
            title = img.get_attribute("title") or ""
            
            if "+" in alt or "plus" in src.lower() or "add" in src.lower() or alt or title:
                resultado.append(f"   [{idx}] Alt: '{alt}' | Title: '{title}'")
                resultado.append(f"          Src: {src[-40:]}")
        
        driver.switch_to.default_content()
        
    except Exception as e:
        resultado.append(f"\nERROR en análisis: {e}")
    
    # GUARDAR ARCHIVO
    print("[4/4] Guardando análisis...")
    contenido = "\n".join(resultado)
    
    with open("ANALISIS_BENEFICIARIO.txt", "w", encoding="utf-8") as f:
        f.write(contenido)
    
    print("    ✓ Guardado en: ANALISIS_BENEFICIARIO.txt")
    print("\n" + "="*80)
    print(contenido)
    print("="*80)
    
    print("\nChrome abierto por 60 segundos más...")
    time.sleep(60)

except Exception as e:
    print(f"ERROR GENERAL: {e}")
    import traceback
    traceback.print_exc()
finally:
    driver.quit()
    print("✓ Finalizado")
