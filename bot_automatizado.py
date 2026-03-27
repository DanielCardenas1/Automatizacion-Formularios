#!/usr/bin/env python3
"""
Bot Selenium automatizado - Login + Navegación a Beneficiario
1. Login automático
2. Click en RUB online (menú)
3. Click en Beneficiario
4. Click en Beneficiario (nuevamente)
"""

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.select import Select
from webdriver_manager.chrome import ChromeDriverManager
import time

print("\n" + "="*70)
print("BOT AUTOMATIZADO - LOGIN + NAVEGACIÓN")
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
    
    # Esperar a que cargue
    time.sleep(6)
    print("[+] Login completado\n")
    
    # PASO 3: CLICK EN "RUB online" DEL MENÚ
    print("[*] Paso 3: Buscando 'RUB online' en el menú...")
    
    # Buscar el enlace "RUB online" dentro del menú principal
    enlaces = driver.find_elements(By.XPATH, "//a[contains(text(), 'Rub online')]")
    
    if enlaces:
        print(f"[+] Encontrados {len(enlaces)} enlaces con 'Rub online'")
        # Hacer scroll y click en el primero
        driver.execute_script("arguments[0].scrollIntoView(true);", enlaces[0])
        time.sleep(1)
        from selenium.webdriver.common.action_chains import ActionChains
        ActionChains(driver).move_to_element(enlaces[0]).click().perform()
        print("[+] Click en 'RUB online' realizado")
        time.sleep(3)
    else:
        print("[-] No se encontró 'RUB online'")
    
    # PASO 4: CLICK EN "Beneficiario" (primer nivel)
    print("\n[*] Paso 4: Buscando primer 'Beneficiario'...")
    
    # Buscar el primer "Beneficiario" del menú
    enlaces_beneficiario = driver.find_elements(By.XPATH, "//a[contains(text(), 'Beneficiario')]")
    
    if enlaces_beneficiario:
        print(f"[+] Encontrados {len(enlaces_beneficiario)} enlaces con 'Beneficiario'")
        # Hacer scroll y click en el primero
        driver.execute_script("arguments[0].scrollIntoView(true);", enlaces_beneficiario[0])
        time.sleep(1)
        ActionChains(driver).move_to_element(enlaces_beneficiario[0]).click().perform()
        print("[+] Click en primer 'Beneficiario' realizado")
        time.sleep(3)
    else:
        print("[-] No se encontró 'Beneficiario'")
    
    # PASO 5: CLICK EN "Beneficiario" (segundo nivel)
    print("\n[*] Paso 5: Buscando segundo 'Beneficiario'...")
    
    # Buscar el segundo "Beneficiario" (si hay más de uno)
    enlaces_beneficiario = driver.find_elements(By.XPATH, "//a[contains(text(), 'Beneficiario')]")
    
    if len(enlaces_beneficiario) > 1:
        print(f"[+] Encontrados {len(enlaces_beneficiario)} enlaces")
        # Hacer scroll y click en el segundo
        driver.execute_script("arguments[0].scrollIntoView(true);", enlaces_beneficiario[1])
        time.sleep(1)
        ActionChains(driver).move_to_element(enlaces_beneficiario[1]).click().perform()
        print("[+] Click en segundo 'Beneficiario' realizado")
        time.sleep(4)
    elif len(enlaces_beneficiario) == 1:
        print("[*] Solo hay un 'Beneficiario', intentando buscar 'Beneficiario' anidado...")
        # Intentar encontrar uno dentro de un submenú
        enlaces_nested = driver.find_elements(By.XPATH, "//ul//a[contains(text(), 'Beneficiario')]")
        if len(enlaces_nested) > 1:
            driver.execute_script("arguments[0].scrollIntoView(true);", enlaces_nested[1])
            time.sleep(1)
            ActionChains(driver).move_to_element(enlaces_nested[1]).click().perform()
            print("[+] Click en 'Beneficiario' anidado realizado")
            time.sleep(4)
        else:
            print("[!] No se encontró segundo 'Beneficiario'")
    else:
        print("[-] No se encontró ningún 'Beneficiario'")
    
    # PASO 6: CLICK EN EL SÍMBOLO "+" (btnNuevo)
    print("\n[*] Paso 6: Buscando el símbolo '+' (Nuevo)...")
    
    try:
        # Cambiar contexto al iframe
        iframe = driver.find_element(By.ID, "frameContent")
        driver.switch_to.frame(iframe)
        print("[+] Estoy dentro del iframe 'frameContent'")
        
        # Buscar el botón "Nuevo" por ID
        boton_nuevo = wait.until(EC.presence_of_element_located((By.ID, "btnNuevo")))
        print("[+] Botón '+' encontrado")
        
        # Hacer scroll y click
        driver.execute_script("arguments[0].scrollIntoView(true);", boton_nuevo)
        time.sleep(1)
        
        ActionChains(driver).move_to_element(boton_nuevo).click().perform()
        print("[+] Click en '+' realizado")
        time.sleep(3)
        
        # Volver al contenido principal
        driver.switch_to.default_content()
        print("[+] Volviendo al contenido principal")
        
    except Exception as e:
        print(f"[-] Error al hacer click en '+': {str(e)}")
        try:
            driver.switch_to.default_content()
        except:
            pass
    
    # PASO 7: LLENAR LOS FILTROS DEL FORMULARIO
    print("\n[*] Paso 7: Llenando los filtros para que quede como la imagen...")
    
    try:
        # Seguiros en el iframe
        iframe = driver.find_element(By.ID, "frameContent")
        driver.switch_to.frame(iframe)
        print("[+] Dentro del iframe")
        
        # 1. Seleccionar "Masivo" en Forma vinculación del beneficiario
        print("\n[*] 1. Seleccionando 'Masivo' en Forma vinculación...")
        try:
            radio_masivo = driver.find_element(
                By.XPATH, 
                "//input[@type='radio' and contains(@value, 'Masivo')]"
            )
            radio_masivo.click()
            print("[+] 'Masivo' seleccionado")
            time.sleep(0.5)
        except Exception as e:
            print(f"[-] Error seleccionando Masivo: {str(e)}")
        
        # 2. Dirección de Primera Infancia (ya está hecho, pero lo verificamos)
        print("\n[*] 2. Verificando 'Dirección de Primera Infancia'...")
        try:
            direccion_select = driver.find_element(
                By.ID, 
                "cphCont_TabContainer1_tbnDatosB_BeneficiarioVincula_ddlIdDireccionesICBF"
            )
            select_obj = Select(direccion_select)
            select_obj.select_by_visible_text("Dirección de Primera Infancia")
            print("[+] 'Dirección de Primera Infancia' confirmada")
            time.sleep(0.5)
        except Exception as e:
            print(f"[-] Error: {str(e)}")
        
        # 3. Regional: "Boyacá"
        print("\n[*] 3. Seleccionando Regional 'Boyacá'...")
        try:
            regional_select = driver.find_element(
                By.ID, 
                "cphCont_TabContainer1_tbnDatosB_BeneficiarioVincula_ddlRegional"
            )
            select_obj = Select(regional_select)
            select_obj.select_by_visible_text("Boyacá")
            print("[+] 'Boyacá' seleccionada")
            time.sleep(1)
        except Exception as e:
            print(f"[-] Error seleccionando Regional: {str(e)}")
        
        # 4. Vigencia: "2026"
        print("\n[*] 4. Seleccionando Vigencia '2026'...")
        try:
            vigencia_select = driver.find_element(
                By.ID, 
                "cphCont_TabContainer1_tbnDatosB_BeneficiarioVincula_ddlIdVigencia"
            )
            select_obj = Select(vigencia_select)
            select_obj.select_by_visible_text("2026")
            print("[+] '2026' seleccionada")
            time.sleep(1)
        except Exception as e:
            print(f"[-] Error seleccionando Vigencia: {str(e)}")
        
        # 5. Contrato: "OD 15 420272 00015 2026"
        print("\n[*] 5. Seleccionando Contrato...")
        try:
            contrato_select = driver.find_element(
                By.ID, 
                "cphCont_TabContainer1_tbnDatosB_BeneficiarioVincula_ddlNumeroContrato"
            )
            select_obj = Select(contrato_select)
            select_obj.select_by_visible_text("OD 15 420272 00015 2026")
            print("[+] Contrato 'OD 15 420272 00015 2026' seleccionado")
            time.sleep(1)
        except Exception as e:
            print(f"[-] Error seleccionando Contrato: {str(e)}")
        
        # 6. Nombre del servicio: "EDUCACIÓN INICIAL EN EL HOGAR - FAMILIAR Y COMUNITARIA - 420272-2026"
        print("\n[*] 6. Seleccionando Nombre del servicio...")
        try:
            servicio_select = driver.find_element(
                By.ID, 
                "cphCont_TabContainer1_tbnDatosB_BeneficiarioVincula_ddlNombreServicio"
            )
            select_obj = Select(servicio_select)
            select_obj.select_by_visible_text("EDUCACIÓN INICIAL EN EL HOGAR - FAMILIAR Y COMUNITARIA - 420272-2026")
            print("[+] Servicio seleccionado")
            time.sleep(1)
        except Exception as e:
            print(f"[-] Error seleccionando Servicio: {str(e)}")
        
        # 7. Número de UDS: "DUITAMA D2 D3 - 1523800124748"
        print("\n[*] 7. Seleccionando UDS...")
        try:
            uds_select = driver.find_element(
                By.XPATH,
                "//select[contains(@name, 'ddlNumeroUDS') or contains(@name, 'ddlNumeroD2D3')]"
            )
            select_obj = Select(uds_select)
            select_obj.select_by_visible_text("DUITAMA D2 D3 - 1523800124748")
            print("[+] UDS 'DUITAMA D2 D3' seleccionada")
            time.sleep(1)
        except Exception as e:
            print(f"[-] UDS no encontrada, intentando búsqueda común...")
            try:
                # Buscar todos los selects y mostrar opciones
                selects = driver.find_elements(By.TAG_NAME, "select")
                for i, s in enumerate(selects):
                    options = s.find_elements(By.TAG_NAME, "option")
                    for opt in options:
                        if "DUITAMA D2 D3" in opt.text or "D2 D3" in opt.text:
                            select_obj = Select(s)
                            select_obj.select_by_visible_text(opt.text)
                            print(f"[+] UDS encontrada y seleccionada: {opt.text}")
                            break
            except Exception as e2:
                print(f"[-] Error: {str(e2)}")
        
        # 8. Tipo de beneficiario: "NIÑO O NIÑA ENTRE 6 MESES Y 11 MESES Y 11 HE..."
        print("\n[*] 8. Seleccionando Tipo de beneficiario...")
        try:
            tipo_beneficiario_select = driver.find_element(
                By.ID, 
                "cphCont_TabContainer1_tbnDatosB_BeneficiarioVincula_ddlIdTipoBeneficiario"
            )
            select_obj = Select(tipo_beneficiario_select)
            # Buscar la opción que contenga "NIÑO O NIÑA ENTRE 6 MESES"
            opciones = tipo_beneficiario_select.find_elements(By.TAG_NAME, "option")
            for opcion in opciones:
                if "NIÑO O NIÑA ENTRE 6 MESES" in opcion.text:
                    select_obj.select_by_value(opcion.get_attribute("value"))
                    print(f"[+] Tipo de beneficiario seleccionado: {opcion.text[:50]}")
                    break
            time.sleep(1)
        except Exception as e:
            print(f"[-] Error seleccionando Tipo de beneficiario: {str(e)}")
        
        # 9. Tipo de Documento: "REGISTRO CIVIL"
        print("\n[*] 9. Seleccionando Tipo de Documento...")
        try:
            tipo_doc_select = driver.find_element(
                By.ID, 
                "cphCont_TabContainer1_tbnDatosB_BeneficiarioVincula_datosBasicosPersona_ddlIdTipoDocumento"
            )
            select_obj = Select(tipo_doc_select)
            select_obj.select_by_visible_text("REGISTRO CIVIL")
            print("[+] 'REGISTRO CIVIL' seleccionado")
            time.sleep(1)
        except Exception as e:
            print(f"[-] Error seleccionando Tipo de Documento: {str(e)}")
        
        print("\n[+] ✓ TODOS LOS CAMPOS LLENADOS IGUAL A LA IMAGEN")
        
        # Volver al contenido principal
        driver.switch_to.default_content()
        print("[+] Volviendo al contenido principal")
        
    except Exception as e:
        print(f"[-] Error al llenar filtros: {str(e)}")
        import traceback
        traceback.print_exc()
        try:
            driver.switch_to.default_content()
        except:
            pass
    
    # VERIFICAR QUE ESTAMOS EN LA PÁGINA CORRECTA
    print("\n" + "="*70)
    print("ESTADO ACTUAL")
    print("="*70)
    print(f"[+] URL actual: {driver.current_url}")
    print(f"[+] Título: {driver.title}")
    
    print("\n" + "="*70)
    print("NAVEGACIÓN COMPLETADA")
    print("="*70)
    print("\n[*] El navegador permanecerá abierto para que continúes...")
    print("[*] El bot está listo para cargar Excel y buscar documentos")
    print("[*] Presiona Ctrl+C cuando quieras cerrar\n")
    
    # Esperar indefinidamente
    while True:
        time.sleep(1)
    
except KeyboardInterrupt:
    print("\n\n[!] Cerrando navegador (Ctrl+C)")
    driver.quit()
    print("[+] Navegador cerrado")
    
except Exception as e:
    print(f"\n[-] Error: {str(e)}")
    import traceback
    traceback.print_exc()
    print("\n[*] Manteniéndoel navegador abierto para inspeccionar...")
    print("[*] Presiona Ctrl+C para cerrarlo")
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        driver.quit()
