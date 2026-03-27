#!/usr/bin/env python3
"""
Bot que graba los clicks del usuario para reproducirlos automáticamente
"""
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import time
import sys
import json

# Script JavaScript para grabar todos los clicks
SCRIPT_GRABADOR = """
window.clicksGrabados = [];
window.inputsLlenados = [];

document.addEventListener('click', function(e) {
    let elemento = e.target;
    let info = {
        tipo: 'click',
        tag: elemento.tagName,
        id: elemento.id,
        name: elemento.name,
        clase: elemento.className,
        texto: elemento.textContent.substring(0, 100),
        xpath: getXPath(elemento),
        timestamp: new Date().toLocaleTimeString()
    };
    window.clicksGrabados.push(info);
    console.log('Click grabado:', info);
}, true);

document.addEventListener('change', function(e) {
    let elemento = e.target;
    let info = {
        tipo: 'change',
        tag: elemento.tagName,
        id: elemento.id,
        name: elemento.name,
        valor: elemento.value,
        xpath: getXPath(elemento),
        timestamp: new Date().toLocaleTimeString()
    };
    window.inputsLlenados.push(info);
    console.log('Input cambió:', info);
}, true);

function getXPath(element) {
    if (element.id !== '')
        return "//*[@id='" + element.id + "']";
    if (element === document.body)
        return element.tagName.toLowerCase();

    var ix = 0;
    var siblings = element.parentNode.childNodes;
    for (var i = 0; i < siblings.length; i++) {
        var sibling = siblings[i];
        if (sibling === element)
            return getXPath(element.parentNode) + '/' + element.tagName.toLowerCase() + '[' + (ix + 1) + ']';
        if (sibling.nodeType === 1 && sibling.tagName.toLowerCase() === element.tagName.toLowerCase())
            ix++;
    }
}
"""

print("\n" + "="*70)
print("BOT GRABADOR DE CLICKS")
print("="*70)
print("\n[*] Abriendo navegador...")
time.sleep(1)

options = webdriver.ChromeOptions()
options.add_argument('--disable-notifications')
options.add_argument('--disable-popup-blocking')

try:
    driver = webdriver.Chrome(
        service=Service(ChromeDriverManager().install()),
        options=options
    )
    print("[+] Navegador abierto\n")
    
    print("[*] Accediendo a RUB Online...\n")
    driver.get("https://rubonline.icbf.gov.co/DefaultF.aspx")
    print("[+] Página cargada\n")
    
    # Inyectar script grabador
    driver.execute_script(SCRIPT_GRABADOR)
    print("[+] ✓ Sistema de grabación activado\n")
    
    # LOGIN AUTOMÁTICO
    print("="*70)
    print("PASO 1: LOGIN AUTOMÁTICO")
    print("="*70)
    
    try:
        wait = WebDriverWait(driver, 10)
        
        print("[*] Ingresando usuario: Usuario")
        campo_usuario = wait.until(EC.presence_of_element_located((By.ID, "UserName")))
        campo_usuario.clear()
        campo_usuario.send_keys("Usuario")
        print("[+] Usuario ingresado")
        
        print("[*] Ingresando contraseña...")
        campo_password = driver.find_element(By.ID, "Password")
        campo_password.clear()
        campo_password.send_keys("Contraseña")
        print("[+] Contraseña ingresada")
        
        print("[*] Haciendo clic en 'Ingresar'...")
        boton_login = driver.find_element(By.ID, "LoginButton")
        boton_login.click()
        print("[+] Clic realizado, esperando a que cargue...")
        
        time.sleep(6)
        print("[+] ✓ LOGIN COMPLETADO\n")
        
    except Exception as e:
        print(f"[-] Error en login: {str(e)}\n")
        raise
    
    print("="*70)
    print("PASO 2: CONTROLA EL NAVEGADOR (GRABANDO CLICKS)")
    print("="*70)
    print("Todos tus clicks y cambios se están GRABANDO automáticamente")
    print("\nDEBES HACER CLIC EN:")
    print("  1. RUB ONLINE (desde el menú izquierdo)")
    print("  2. Luego: Rub online")
    print("  3. Luego: Beneficiario")
    print("  4. Luego: Beneficiario (de nuevo)")
    print("  5. Luego: El símbolo '+' (botón Nuevo)")
    print("  6. Llenar los filtros")
    print("\n⏰ El navegador permanecerá abierto por 5 minutos")
    print("   Todos tus clicks se grabarán automáticamente")
    print("="*70 + "\n")
    
    print("[*] Esperando 5 minutos (300 segundos) para tu interacción...")
    print("[*] INTERACTÚA MANUALMENTE CON EL NAVEGADOR AHORA...\n")
    
    for i in range(30):
        time.sleep(10)
        try:
            clicks_result = driver.execute_script("return window.clicksGrabados;")
            inputs_result = driver.execute_script("return window.inputsLlenados;")
            clicks_count = len(clicks_result) if clicks_result else 0
            inputs_count = len(inputs_result) if inputs_result else 0
        except:
            clicks_count = 0
            inputs_count = 0
        tiempo_restante = 300 - (i+1)*10
        print(f"[*] {tiempo_restante}s restantes | Clicks: {clicks_count} | Cambios: {inputs_count}")
    
    # Capturar datos finales
    print("\n" + "="*70)
    print("INFORMACIÓN CAPTURADA")
    print("="*70)
    
    clicks_grabados = driver.execute_script("return window.clicksGrabados;") or []
    inputs_llenados = driver.execute_script("return window.inputsLlenados;") or []
    
    print(f"[+] Clicks grabados: {len(clicks_grabados)}")
    print(f"[+] Cambios de input: {len(inputs_llenados)}")
    
    # Guardar en archivo JSON
    archivo_grabacion = "grabacion_clicks.json"
    datos_grabados = {
        "clicks": clicks_grabados,
        "inputs": inputs_llenados,
        "url_final": driver.current_url,
        "titulo_final": driver.title
    }
    
    with open(archivo_grabacion, 'w', encoding='utf-8') as f:
        json.dump(datos_grabados, f, indent=2, ensure_ascii=False)
    
    print(f"[+] Grabación guardada en: {archivo_grabacion}\n")
    
    if clicks_grabados:
        print("[*] Mostrando primeros 5 clicks:")
        for i, click in enumerate(clicks_grabados[:5]):
            print(f"  {i+1}. {click['tag']} | ID: {click['id']} | Texto: {click['texto'][:30]}")
    else:
        print("[!] No se grabaron clicks")
    
    if inputs_llenados:
        print("\n[*] Mostrando primeros 5 cambios de input:")
        for i, inp in enumerate(inputs_llenados[:5]):
            print(f"  {i+1}. {inp['tag']} | ID: {inp['id']} | Valor: {inp['valor']}")
    else:
        print("[!] No se grabaron cambios de input")
    
    print("\n[*] Cerrando navegador en 3 segundos...")
    time.sleep(3)
    
    driver.quit()
    print("[+] ✓ Navegador cerrado")
    print("[+] ✓ GRABACIÓN COMPLETADA")
    print(f"\nAhora puedo reproducir automáticamente los {len(clicks_grabados)} clicks que hiciste\n")
    
except KeyboardInterrupt:
    print("\n\n[!] Detenido por el usuario (Ctrl+C)")
    try:
        driver.quit()
    except:
        pass
    sys.exit(0)
    
except Exception as e:
    print(f"\n[-] Error: {str(e)}")
    import traceback
    traceback.print_exc()
    try:
        driver.quit()
    except:
        pass
    sys.exit(1)
