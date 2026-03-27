#!/usr/bin/env python3
"""
Bot MANUAL INTERACTIVO - Inspecciona el formulario
Te muestra dónde está el campo de documento y la lupa
"""

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.action_chains import ActionChains
from webdriver_manager.chrome import ChromeDriverManager
import time

print("\n" + "="*70)
print("BOT MANUAL - INSPECCIÓN DEL FORMULARIO")
print("="*70 + "\n")

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
    
    driver.find_element(By.ID, "UserName").send_keys("Usuario")
    driver.find_element(By.ID, "Password").send_keys("Contraseña")
    driver.find_element(By.ID, "LoginButton").click()
    time.sleep(8)
    print("[+] Login completado\n")
    
    # NAVEGAR
    print("[2] Navegando...")
    time.sleep(2)
    driver.find_elements(By.XPATH, "//a[contains(text(), 'Rub online')]")[0].click()
    time.sleep(3)
    
    driver.find_elements(By.XPATH, "//a[contains(text(), 'Beneficiario')]")[0].click()
    time.sleep(3)
    
    driver.find_elements(By.XPATH, "//a[contains(text(), 'Beneficiario')]")[1].click()
    time.sleep(4)
    print("[+] Navegación completada\n")
    
    # CLICK EN +
    print("[3] Haciendo click en botón +...")
    iframe = wait.until(EC.presence_of_element_located((By.ID, "frameContent")))
    driver.switch_to.frame(iframe)
    
    boton_nuevo = wait.until(EC.element_to_be_clickable((By.ID, "btnNuevo")))
    ActionChains(driver).move_to_element(boton_nuevo).click().perform()
    print("[+] Botón + clickeado\n")
    time.sleep(3)
    
    # LLENAR FILTROS
    print("[4] Llenando filtros...")
    
    radioBtns = driver.find_elements(By.XPATH, "//input[@type='radio']")
    for radio in radioBtns:
        if 'Uno a uno' in (radio.get_attribute('value') or ''):
            radio.click()
            break
    time.sleep(1)
    
    Select(driver.find_element(By.ID, "cphCont_TabContainer1_tbnDatosB_BeneficiarioVincula_ddlIdDireccionesICBF")).select_by_visible_text("Dirección de Primera Infancia")
    time.sleep(1)
    
    Select(driver.find_element(By.ID, "cphCont_TabContainer1_tbnDatosB_BeneficiarioVincula_ddlRegional")).select_by_visible_text("Boyacá")
    time.sleep(1)
    
    Select(driver.find_element(By.ID, "cphCont_TabContainer1_tbnDatosB_BeneficiarioVincula_ddlIdVigencia")).select_by_visible_text("2026")
    time.sleep(1)
    
    Select(driver.find_element(By.ID, "cphCont_TabContainer1_tbnDatosB_BeneficiarioVincula_ddlNumeroContrato")).select_by_visible_text("OD 15 420272 00015 2026")
    time.sleep(1)
    
    Select(driver.find_element(By.ID, "cphCont_TabContainer1_tbnDatosB_BeneficiarioVincula_ddlNombreServicio")).select_by_visible_text("EDUCACIÓN INICIAL EN EL HOGAR - FAMILIAR Y COMUNITARIA - 420272-2026")
    time.sleep(2)
    
    selects = driver.find_elements(By.TAG_NAME, "select")
    for select_elem in selects:
        for opt in select_elem.find_elements(By.TAG_NAME, "option"):
            if "DUITAMA D2" in opt.text:
                Select(select_elem).select_by_value(opt.get_attribute("value"))
                break
    
    print("[+] Filtros completados\n")
    time.sleep(3)
    
    # INSPECCIONAR
    print("="*70)
    print("BUSCANDO CAMPO DE DOCUMENTO Y LUPA")
    print("="*70 + "\n")
    
    # Buscar TODOS los inputs
    todos_inputs = driver.find_elements(By.TAG_NAME, "input")
    print(f"[*] Total inputs encontrados: {len(todos_inputs)}\n")
    
    # Buscar específicamente por "documento"
    print("[INPUTS CON 'DOCUMENTO' EN EL NOMBRE]:")
    for i, inp in enumerate(todos_inputs):
        inp_id = inp.get_attribute("id") or ""
        inp_name = inp.get_attribute("name") or ""
        inp_type = inp.get_attribute("type") or ""
        inp_placeholder = inp.get_attribute("placeholder") or ""
        
        if 'documento' in inp_id.lower() or 'documento' in inp_name.lower() or 'documento' in inp_placeholder.lower():
            print(f"\n  #{i+1}: Found 'DOCUMENTO'")
            print(f"      ID: {inp_id}")
            print(f"      Name: {inp_name}")
            print(f"      Type: {inp_type}")
            print(f"      Placeholder: {inp_placeholder}")
            print(f"      Visible: {inp.is_displayed()}")
            
            # Resaltar el elemento
            driver.execute_script("""
                arguments[0].style.border = '3px solid red';
                arguments[0].style.backgroundColor = 'yellow';
            """, inp)
            
            # Buscar elementos cercanos (la lupa)
            try:
                padre = inp.find_element(By.XPATH, "..")
                botones_cercanos = padre.find_elements(By.TAG_NAME, "button")
                if botones_cercanos:
                    print(f"      Botones cercanos: {len(botones_cercanos)}")
                    for btn in botones_cercanos:
                        print(f"        - {btn.get_attribute('id')}: {btn.text}")
                        driver.execute_script("""
                            arguments[0].style.border = '2px solid blue';
                            arguments[0].style.backgroundColor = 'lightblue';
                        """, btn)
            except:
                pass
    
    # Buscar inputs visual sin etiqueta "documento"
    print(f"\n[INPUTS 'TEXT' SIN 'DOCUMENTO']:")
    contador = 0
    for i, inp in enumerate(todos_inputs):
        if inp.get_attribute("type") in ["text", ""]:
            inp_id = inp.get_attribute("id") or ""
            inp_name = inp.get_attribute("name") or ""
            inp_placeholder = inp.get_attribute("placeholder") or ""
            if 'documento' not in inp_id.lower() and 'documento' not in inp_name.lower() and 'documento' not in inp_placeholder.lower():
                if contador < 5:
                    print(f"\n  #{i+1}:")
                    print(f"      ID: {inp_id}")
                    print(f"      Name: {inp_name}")
                    print(f"      Placeholder: {inp_placeholder}")
                    print(f"      Visible: {inp.is_displayed()}")
                    contador += 1
    
    print("\n" + "="*70)
    print("NAVEGADOR ABIERTO - Revisa los elementos resaltados")
    print("Rojo = documento, Azul = botón cercano")
    print("Presiona Ctrl+C para cerrar")
    print("="*70 + "\n")
    
    while True:
        time.sleep(1)

except KeyboardInterrupt:
    print("\n[!] Cerrando...")
    driver.quit()

except Exception as e:
    print(f"\n[-] Error: {str(e)}")
    import traceback
    traceback.print_exc()
    driver.quit()
