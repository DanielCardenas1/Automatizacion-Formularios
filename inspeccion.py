#!/usr/bin/env python3
"""
Script de inspección - Encuentra los elementos del formulario
"""

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.action_chains import ActionChains
from webdriver_manager.chrome import ChromeDriverManager
import time

print("\n[*] Iniciando inspección de formulario...\n")

options = webdriver.ChromeOptions()
options.add_argument('--disable-notifications')
options.add_argument('--disable-popup-blocking')

driver = webdriver.Chrome(
    service=Service(ChromeDriverManager().install()),
    options=options
)

wait = WebDriverWait(driver, 20)

try:
    # LOGIN
    print("[1] Haciendo login...")
    driver.get("https://rubonline.icbf.gov.co/DefaultF.aspx")
    time.sleep(3)
    
    campo_usuario = wait.until(EC.presence_of_element_located((By.ID, "UserName")))
    campo_usuario.clear()
    campo_usuario.send_keys("angie.cardenas")
    
    campo_password = driver.find_element(By.ID, "Password")
    campo_password.clear()
    campo_password.send_keys("Celeste1020*")
    
    boton_login = driver.find_element(By.ID, "LoginButton")
    boton_login.click()
    time.sleep(8)
    print("[+] Login completado\n")
    
    # NAVEGAR
    print("[2] Navegando al formulario...")
    time.sleep(2)
    enlacesRUB = driver.find_elements(By.XPATH, "//a[contains(text(), 'Rub online')]")
    if enlacesRUB:
        driver.execute_script("arguments[0].scrollIntoView(true);", enlacesRUB[0])
        time.sleep(0.5)
        ActionChains(driver).move_to_element(enlacesRUB[0]).click().perform()
        time.sleep(3)
    
    time.sleep(1)
    enlacesBene = driver.find_elements(By.XPATH, "//a[contains(text(), 'Beneficiario')]")
    if enlacesBene:
        driver.execute_script("arguments[0].scrollIntoView(true);", enlacesBene[0])
        time.sleep(0.5)
        ActionChains(driver).move_to_element(enlacesBene[0]).click().perform()
        time.sleep(3)
    
    time.sleep(1)
    enlacesBene = driver.find_elements(By.XPATH, "//a[contains(text(), 'Beneficiario')]")
    if len(enlacesBene) > 1:
        driver.execute_script("arguments[0].scrollIntoView(true);", enlacesBene[1])
        time.sleep(0.5)
        ActionChains(driver).move_to_element(enlacesBene[1]).click().perform()
        time.sleep(4)
    
    print("[+] Navegación completada\n")
    
    # CLICKEAR BOTÓN +
    print("[3] Haciendo click en +...")
    time.sleep(2)
    iframe = wait.until(EC.presence_of_element_located((By.ID, "frameContent")))
    driver.switch_to.frame(iframe)
    
    boton_nuevo = wait.until(EC.element_to_be_clickable((By.ID, "btnNuevo")))
    driver.execute_script("arguments[0].scrollIntoView(true);", boton_nuevo)
    time.sleep(0.5)
    ActionChains(driver).move_to_element(boton_nuevo).click().perform()
    time.sleep(3)
    print("[+] Botón '+' clickeado\n")
    
    # INSPECCIONAR FORMULARIO
    print("="*70)
    print("ELEMENTOS DEL FORMULARIO ENCONTRADOS")
    print("="*70 + "\n")
    
    # Buscar campos de entrada
    print("[CAMPOS DE TEXTO]:")
    inputs = driver.find_elements(By.TAG_NAME, "input")
    for i, inp in enumerate(inputs):
        inp_id = inp.get_attribute("id")
        inp_type = inp.get_attribute("type")
        inp_name = inp.get_attribute("name")
        inp_placeholder = inp.get_attribute("placeholder")
        if inp_id or inp_name:
            print(f"  {i+1}. Type: {inp_type} | ID: {inp_id} | Name: {inp_name} | Placeholder: {inp_placeholder}")
    
    # Buscar botones
    print("\n[BOTONES]:")
    botones = driver.find_elements(By.TAG_NAME, "button")
    for i, btn in enumerate(botones):
        btn_id = btn.get_attribute("id")
        btn_text = btn.text.strip()
        if btn_id or btn_text:
            print(f"  {i+1}. ID: {btn_id} | Text: {btn_text}")
    
    # Buscar elementos con class "btnBuscar" o similar
    print("\n[BOTONES CON CLASE]:")
    elementos_especiales = driver.find_elements(By.XPATH, "//*[contains(@class, 'btn')] | //*[contains(@class, 'Buscar')] | //*[contains(@class, 'buscar')]")
    for i, elem in enumerate(elementos_especiales[:10]):
        elem_id = elem.get_attribute("id")
        elem_class = elem.get_attribute("class")
        elem_text = elem.text.strip()
        print(f"  {i+1}. ID: {elem_id} | Class: {elem_class} | Text: {elem_text}")
    
    # Buscar elementos img (pueden ser íconos de búsqueda)
    print("\n[ELEMENTOS IMG (POSIBLES LUPAS)]:")
    imgs = driver.find_elements(By.TAG_NAME, "img")
    for i, img in enumerate(imgs[:10]):
        img_id = img.get_attribute("id")
        img_src = img.get_attribute("src")
        img_title = img.get_attribute("title")
        img_alt = img.get_attribute("alt")
        print(f"  {i+1}. ID: {img_id} | Alt: {img_alt} | Title: {img_title}")
    
    # Buscar todo el HTML para ver estructura
    print("\n[HTML GENERAL DEL FORMULARIO]:")
    form_html = driver.find_element(By.TAG_NAME, "body").get_attribute("innerHTML")
    
    # Buscar menciones de "Documento" o "identificacion"
    if "Número" in form_html:
        print("  ✓ Encontrada palabra 'Número'")
    if "Documento" in form_html:
        print("  ✓ Encontrada palabra 'Documento'")
    if "identificacion" in form_html.lower():
        print("  ✓ Encontrada palabra 'Identificación'")
    if "Buscar" in form_html:
        print("  ✓ Encontrada palabra 'Buscar'")
    if "lupa" in form_html.lower():
        print("  ✓ Encontrada palabra 'lupa'")
    
    print("\n" + "="*70)
    print("INSPECCIÓN COMPLETADA - Navegador abierto")
    print("="*70 + "\n")
    
    while True:
        time.sleep(1)

except KeyboardInterrupt:
    print("\n[!] Cerrando navegador...")
    driver.quit()

except Exception as e:
    print(f"[-] Error: {str(e)}")
    import traceback
    traceback.print_exc()
    driver.quit()
