#!/usr/bin/env python3
"""
Bot para inspeccionar el formulario después de hacer clic en +
"""

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import time

print("\n" + "="*70)
print("BOT DE INSPECCIÓN - FORMULARIO")
print("="*70)

options = webdriver.ChromeOptions()
options.add_argument('--disable-notifications')
options.add_argument('--disable-popup-blocking')

driver = webdriver.Chrome(
    service=Service(ChromeDriverManager().install()),
    options=options
)

wait = WebDriverWait(driver, 15)

try:
    # NAVREGAR Y LLENAR HASTA EL CLICK EN +
    print("\n[*] Sistema de login automático...")
    driver.get("https://rubonline.icbf.gov.co/DefaultF.aspx")
    time.sleep(3)
    
    campo_usuario = wait.until(EC.presence_of_element_located((By.ID, "UserName")))
    campo_usuario.clear()
    campo_usuario.send_keys("Usuario")
    
    campo_password = driver.find_element(By.ID, "Password")
    campo_password.clear()
    campo_password.send_keys("Contraseña")
    
    boton_login = driver.find_element(By.ID, "LoginButton")
    boton_login.click()
    time.sleep(6)
    
    # Navegar
    from selenium.webdriver.common.action_chains import ActionChains
    
    enlaces = driver.find_elements(By.XPATH, "//a[contains(text(), 'Rub online')]")
    ActionChains(driver).move_to_element(enlaces[0]).click().perform()
    time.sleep(3)
    
    enlaces_beneficiario = driver.find_elements(By.XPATH, "//a[contains(text(), 'Beneficiario')]")
    ActionChains(driver).move_to_element(enlaces_beneficiario[0]).click().perform()
    time.sleep(3)
    
    enlaces_beneficiario = driver.find_elements(By.XPATH, "//a[contains(text(), 'Beneficiario')]")
    ActionChains(driver).move_to_element(enlaces_beneficiario[1]).click().perform()
    time.sleep(4)
    
    # Click en +
    print("[*] Haciendo click en +...")
    iframe = driver.find_element(By.ID, "frameContent")
    driver.switch_to.frame(iframe)
    
    boton_nuevo = wait.until(EC.presence_of_element_located((By.ID, "btnNuevo")))
    boton_nuevo.click()
    time.sleep(2)
    
    # INSPECCIONAR
    print("\n" + "="*70)
    print("ELEMENTOS DEL FORMULARIO")
    print("="*70)
    
    # Buscar todos los dropdowns/selects
    selects = driver.find_elements(By.TAG_NAME, "select")
    print(f"\n[+] Encontrados {len(selects)} elementos SELECT:")
    for i, sel in enumerate(selects):
        print(f"  {i}. ID: {sel.get_attribute('id')} | Name: {sel.get_attribute('name')}")
    
    # Buscar divs con clases específicas (podrían ser dropdowns customizados)
    divs_dropdown = driver.find_elements(By.XPATH, "//*[contains(@class, 'dropdown') or contains(@class, 'select') or contains(@class, 'option')]")
    print(f"\n[+] Encontrados {len(divs_dropdown)} elementos tipo dropdown customizado:")
    for i, div in enumerate(divs_dropdown[:10]):
        tag = div.tag_name
        class_name = div.get_attribute('class')
        id_attr = div.get_attribute('id')
        print(f"  {i}. {tag} | ID: {id_attr} | Class: {class_name[:50]}")
    
    # Buscar todos los inputs
    inputs = driver.find_elements(By.TAG_NAME, "input")
    print(f"\n[+] Encontrados {len(inputs)} elementos INPUT")
    
    # Búsqueda específica de "Área misional"
    print("\n[*] Buscando campo 'Área misional'...")
    try:
        area_misional_label = driver.find_element(By.XPATH, "//*[contains(text(), 'Área misional')]")
        print(f"[+] Label encontrado: {area_misional_label.text}")
        # Buscar elemento hermano o elemento padre/siguiente
        parent = area_misional_label.find_element(By.XPATH, ".//..")
        print(f"[+] Elemento padre tag: {parent.tag_name}")
        siguiente = area_misional_label.find_element(By.XPATH, "./following-sibling::*")
        print(f"[+] Elemento siguiente: {siguiente.tag_name} | {siguiente.get_attribute('class')}")
    except Exception as e:
        print(f"[-] No se encontró label 'Área misional': {str(e)}")
    
    print("\n[*] Esperando 30 segundos para inspeccionar visualmente...")
    time.sleep(30)
    
    driver.switch_to.default_content()
    driver.quit()
    print("\n[+] Inspección completada")
    
except Exception as e:
    print(f"\n[-] Error: {str(e)}")
    import traceback
    traceback.print_exc()
    try:
        driver.quit()
    except:
        pass
