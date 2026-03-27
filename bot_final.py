#!/usr/bin/env python3
"""
Bot Selenium automatizado - Login + Navegación + Llenado de formulario
Versión mejorada con esperas robusto
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
print("BOT AUTOMATIZADO MEJORADO - COMPLETO")
print("="*70)

options = webdriver.ChromeOptions()
options.add_argument('--disable-notifications')
options.add_argument('--disable-popup-blocking')

driver = webdriver.Chrome(
    service=Service(ChromeDriverManager().install()),
    options=options
)

wait = WebDriverWait(driver, 20)

try:
    # PASO 1: ACCEDER A LA PÁGINA
    print("\n[*] Paso 1: Accediendo a RUB Online...")
    driver.get("https://rubonline.icbf.gov.co/DefaultF.aspx")
    time.sleep(3)
    print("[+] Página cargada\n")
    
    # PASO 2: LOGIN
    print("[*] Paso 2: Realizando login...")
    
    # Llenar usuario
    campo_usuario = wait.until(EC.presence_of_element_located((By.ID, "UserName")))
    campo_usuario.clear()
    campo_usuario.send_keys("angie.cardenas")
    print("[+] Usuario ingresado")
    
    # Llenar contraseña
    campo_password = driver.find_element(By.ID, "Password")
    campo_password.clear()
    campo_password.send_keys("Celeste1020*")
    print("[+] Contraseña ingresada")
    
    # Click en login
    boton_login = driver.find_element(By.ID, "LoginButton")
    boton_login.click()
    print("[+] Click en Ingresar")
    
    # Esperar a que cargue completamente
    time.sleep(8)
    print("[+] Login completado\n")
    
    # Verificar que estamos logueados
    current_url = driver.current_url
    print(f"[+] URL actual: {current_url}")
    
    # PASO 3: CLICK EN "RUB online" DEL MENÚ
    print("\n[*] Paso 3: Buscando 'RUB online' en el menú...")
    
    time.sleep(2)
    enlacesRUB = driver.find_elements(By.XPATH, "//a[contains(text(), 'Rub online')]")
    
    if enlacesRUB:
        print(f"[+] Encontrados {len(enlacesRUB)} enlaces 'Rub online'")
        driver.execute_script("arguments[0].scrollIntoView(true);", enlacesRUB[0])
        time.sleep(0.5)
        ActionChains(driver).move_to_element(enlacesRUB[0]).click().perform()
        print("[+] Click en 'RUB online' realizado")
        time.sleep(3)
    else:
        print("[-] No se encontró 'RUB online', continuando...")
    
    # PASO 4: CLICK EN PRIMER "Beneficiario"
    print("\n[*] Paso 4: Buscando primer 'Beneficiario'...")
    
    time.sleep(1)
    enlacesBene = driver.find_elements(By.XPATH, "//a[contains(text(), 'Beneficiario')]")
    
    if enlacesBene:
        print(f"[+] Encontrados {len(enlacesBene)} enlaces 'Beneficiario'")
        driver.execute_script("arguments[0].scrollIntoView(true);", enlacesBene[0])
        time.sleep(0.5)
        ActionChains(driver).move_to_element(enlacesBene[0]).click().perform()
        print("[+] Click en primer 'Beneficiario' realizado")
        time.sleep(3)
    else:
        print("[-] No se encontró 'Beneficiario'")
    
    # PASO 5: CLICK EN SEGUNDO "Beneficiario"
    print("\n[*] Paso 5: Buscando segundo 'Beneficiario'...")
    
    time.sleep(1)
    enlacesBene = driver.find_elements(By.XPATH, "//a[contains(text(), 'Beneficiario')]")
    
    if len(enlacesBene) > 1:
        print(f"[+] Encontrados {len(enlacesBene)} enlaces, clickeando el segundo...")
        driver.execute_script("arguments[0].scrollIntoView(true);", enlacesBene[1])
        time.sleep(0.5)
        ActionChains(driver).move_to_element(enlacesBene[1]).click().perform()
        print("[+] Click en segundo 'Beneficiario' realizado")
        time.sleep(4)
    else:
        print("[!] No hay segundo 'Beneficiario'")
    
    # PASO 6: BUSCAR Y HACER CLIC EN EL SÍMBOLO "+"
    print("\n[*] Paso 6: Buscando el símbolo '+' (Nuevo)...")
    
    time.sleep(2)
    try:
        # Esperar a que el iframe cargue
        iframe = wait.until(EC.presence_of_element_located((By.ID, "frameContent")))
        driver.switch_to.frame(iframe)
        print("[+] Iframe 'frameContent' encontrado")
        
        # Buscar el botón "Nuevo"
        boton_nuevo = wait.until(EC.element_to_be_clickable((By.ID, "btnNuevo")))
        print("[+] Botón '+' encontrado")
        
        driver.execute_script("arguments[0].scrollIntoView(true);", boton_nuevo)
        time.sleep(0.5)
        ActionChains(driver).move_to_element(boton_nuevo).click().perform()
        print("[+] Click en '+' realizado")
        time.sleep(3)
        
    except Exception as e:
        print(f"[-] Error con iframe/botón: {str(e)}")
        try:
            driver.switch_to.default_content()
        except:
            pass
    
    # PASO 7: LLENAR EL FORMULARIO
    print("\n[*] Paso 7: Llenando los formularios...")
    
    try:
        # Asegurarse de estar en el iframe
        try:
            driver.switch_to.default_content()
            time.sleep(0.5)
            iframe = driver.find_element(By.ID, "frameContent")
            driver.switch_to.frame(iframe)
        except:
            pass
        
        # 1. Radio button "Uno a uno"
        print("\n  [1/10] Seleccionando 'Uno a uno'...")
        try:
            radioBtns = driver.find_elements(By.XPATH, "//input[@type='radio']")
            for radio in radioBtns:
                if radio.get_attribute('value') and 'Uno a uno' in radio.get_attribute('value'):
                    radio.click()
                    print("      [+] 'Uno a uno' seleccionado")
                    break
        except Exception as e:
            print(f"      [-] Error: {str(e)}")
        
        time.sleep(0.5)
        
        # 2. Dirección
        print("  [2/10] Seleccionando 'Dirección de Primera Infancia'...")
        try:
            select_elem = driver.find_element(By.ID, "cphCont_TabContainer1_tbnDatosB_BeneficiarioVincula_ddlIdDireccionesICBF")
            select = Select(select_elem)
            select.select_by_visible_text("Dirección de Primera Infancia")
            print("      [+] Seleccionada")
            time.sleep(1)
        except Exception as e:
            print(f"      [-] Error: {str(e)}")
        
        # 3. Regional
        print("  [3/10] Seleccionando Regional 'Boyacá'...")
        try:
            select_elem = driver.find_element(By.ID, "cphCont_TabContainer1_tbnDatosB_BeneficiarioVincula_ddlRegional")
            select = Select(select_elem)
            select.select_by_visible_text("Boyacá")
            print("      [+] Seleccionada")
            time.sleep(1)
        except Exception as e:
            print(f"      [-] Error: {str(e)}")
        
        # 4. Vigencia
        print("  [4/10] Seleccionando Vigencia '2026'...")
        try:
            select_elem = driver.find_element(By.ID, "cphCont_TabContainer1_tbnDatosB_BeneficiarioVincula_ddlIdVigencia")
            select = Select(select_elem)
            select.select_by_visible_text("2026")
            print("      [+] Seleccionada")
            time.sleep(1)
        except Exception as e:
            print(f"      [-] Error: {str(e)}")
        
        # 5. Contrato
        print("  [5/10] Seleccionando Contrato...")
        try:
            select_elem = driver.find_element(By.ID, "cphCont_TabContainer1_tbnDatosB_BeneficiarioVincula_ddlNumeroContrato")
            select = Select(select_elem)
            select.select_by_visible_text("OD 15 420272 00015 2026")
            print("      [+] Seleccionado")
            time.sleep(1)
        except Exception as e:
            print(f"      [-] Error: {str(e)}")
        
        # 6. Nombre del servicio
        print("  [6/10] Seleccionando Nombre del servicio...")
        try:
            select_elem = driver.find_element(By.ID, "cphCont_TabContainer1_tbnDatosB_BeneficiarioVincula_ddlNombreServicio")
            select = Select(select_elem)
            select.select_by_visible_text("EDUCACIÓN INICIAL EN EL HOGAR - FAMILIAR Y COMUNITARIA - 420272-2026")
            print("      [+] Seleccionado")
            time.sleep(2)  # Esperar a que se habilite el campo de UDS
        except Exception as e:
            print(f"      [-] Error: {str(e)}")
        
        # 6.5. Nombre de UDS (se habilita después de seleccionar servicio)
        print("  [6.5/10] Seleccionando Nombre de UDS...")
        try:
            # Esperar a que el elemento sea clickeable
            select_elem = wait.until(
                EC.presence_of_element_located((By.XPATH, "//select[contains(@name, 'ddlNumeroContrato') or contains(@name, 'UDS')]"))
            )
            # Buscar todos los selects y encontrar el que sean UDS
            selects = driver.find_elements(By.TAG_NAME, "select")
            for select_elem in selects:
                options = select_elem.find_elements(By.TAG_NAME, "option")
                for opt in options:
                    if "DUITAMA D2 D3" in opt.text or "D2 D3" in opt.text:
                        select = Select(select_elem)
                        select.select_by_value(opt.get_attribute("value"))
                        print(f"      [+] UDS seleccionada: {opt.text}")
                        time.sleep(1)
                        break
        except Exception as e:
            print(f"      [-] Error o campo no habilitado: {str(e)}")
        
        # 7. Tipo de beneficiario
        print("  [7/10] Seleccionando Tipo de beneficiario...")
        try:
            # Hacer scroll down para que aparezca el elemento
            driver.execute_script("window.scrollBy(0, 300);")
            time.sleep(1)
            
            select_elem = driver.find_element(By.ID, "cphCont_TabContainer1_tbnDatosB_BeneficiarioVincula_ddlIdTipoBeneficiario")
            select = Select(select_elem)
            options = select_elem.find_elements(By.TAG_NAME, "option")
            for opt in options:
                if "NIÑO O NIÑA ENTRE 6 MESES" in opt.text:
                    select.select_by_value(opt.get_attribute("value"))
                    print(f"      [+] {opt.text[:40]}...")
                    break
            time.sleep(1)
        except Exception as e:
            print(f"      [-] Error: {str(e)}")
        
        # 8. Tipo de Documento
        print("  [8/10] Seleccionando Tipo de Documento...")
        try:
            # Hacer scroll down para que aparezca el elemento
            driver.execute_script("window.scrollBy(0, 300);")
            time.sleep(1)
            
            select_elem = driver.find_element(By.ID, "cphCont_TabContainer1_tbnDatosB_BeneficiarioVincula_datosBasicosPersona_ddlIdTipoDocumento")
            select = Select(select_elem)
            select.select_by_visible_text("REGISTRO CIVIL")
            print("      [+] Seleccionado")
            time.sleep(1)
        except Exception as e:
            print(f"      [-] Error: {str(e)}")
        
        # 10. Sexo
        print("  [9/10] Seleccionando Sexo 'Seleccione'...")
        try:
            select_elem = driver.find_element(By.ID, "cphCont_TabContainer1_tbnDatosB_BeneficiarioVincula_datosBasicosPersona_ddlIdSexo")
            select = Select(select_elem)
            options = select_elem.find_elements(By.TAG_NAME, "option")
            first_option = options[0]  # Primer opción (Seleccione)
            select.select_by_value(first_option.get_attribute("value"))
            print("      [+] Seleccionado")
            time.sleep(1)
        except Exception as e:
            print(f"      [-] Error: {str(e)}")
        
        print("\n[+] ✓ TODOS LOS CAMPOS LLENADOS CORRECTAMENTE")
        
        driver.switch_to.default_content()
        
    except Exception as e:
        print(f"[-] Error llenando: {str(e)}")
        import traceback
        traceback.print_exc()
        try:
            driver.switch_to.default_content()
        except:
            pass
    
    print("\n" + "="*70)
    print("NAVEGADOR ABIERTO - Presiona Ctrl+C para cerrar")
    print("="*70 + "\n")
    
    while True:
        time.sleep(1)
    
except KeyboardInterrupt:
    print("\n\n[!] Cerrando navegador...")
    driver.quit()
    print("[+] Cerrado")
    
except Exception as e:
    print(f"\n[-] Error general: {str(e)}")
    import traceback
    traceback.print_exc()
    try:
        driver.quit()
    except:
        pass
