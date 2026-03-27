#!/usr/bin/env python3
"""
Bot interactivo - Login automático, luego el usuario controla manualmente
"""
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import time
import sys

print("\n" + "="*70)
print("BOT INTERACTIVO - LOGIN AUTOMÁTICO + CONTROL MANUAL")
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
    
    # LOGIN AUTOMÁTICO
    print("="*70)
    print("PASO 1: LOGIN AUTOMÁTICO")
    print("="*70)
    
    try:
        wait = WebDriverWait(driver, 10)
        
        # Llenar usuario
        print("[*] Ingresando usuario: Usuario")
        campo_usuario = wait.until(EC.presence_of_element_located((By.ID, "UserName")))
        campo_usuario.clear()
        campo_usuario.send_keys("Usuario")
        print("[+] Usuario ingresado")
        
        # Llenar contraseña
        print("[*] Ingresando contraseña...")
        campo_password = driver.find_element(By.ID, "Password")
        campo_password.clear()
        campo_password.send_keys("Contraseña")
        print("[+] Contraseña ingresada")
        
        # Click en login
        print("[*] Haciendo clic en 'Ingresar'...")
        boton_login = driver.find_element(By.ID, "LoginButton")
        boton_login.click()
        print("[+] Clic realizado, esperando a que cargue...")
        
        # Esperar a que la página cargue
        time.sleep(6)
        print("[+] ✓ LOGIN COMPLETADO\n")
        
    except Exception as e:
        print(f"[-] Error en login: {str(e)}\n")
        raise
    
    print("="*70)
    print("PASO 2: CONTROLA EL NAVEGADOR MANUALMENTE")
    print("="*70)
    print("Ahora tienes control total del navegador")
    print("\nDEBES HACER CLIC EN:")
    print("  1. RUB ONLINE (desde el menú izquierdo)")
    print("  2. Luego: Rub online")
    print("  3. Luego: Beneficiario")
    print("  4. Luego: Beneficiario (de nuevo)")
    print("  5. Luego: El símbolo '+' (botón Nuevo)")
    print("\n⏰ El navegador permanecerá abierto por 5 minutos")
    print("   para que hagas clic en lo que necesites")
    print("="*70 + "\n")
    
    print("[*] Esperando 5 minutos (300 segundos) para tu interacción...")
    print("[*] Interactúa manualmente con el navegador ahora...\n")
    
    # Mantener abierto por 5 minutos
    for i in range(30):
        time.sleep(10)
        print(f"[*] Tiempo restante: {300 - (i+1)*10} segundos...")
    
    # Capturar info final
    print("\n" + "="*70)
    print("INFORMACIÓN CAPTURADA")
    print("="*70)
    
    url_actual = driver.current_url
    print(f"[+] URL actual: {url_actual}")
    
    titulo = driver.title
    print(f"[+] Título de página: {titulo}")
    
    print("\n[*] Cerrando navegador en 3 segundos...")
    time.sleep(3)
    
    driver.quit()
    print("[+] ✓ Navegador cerrado")
    print("\n[+] INTERACCIÓN COMPLETADA")
    print("[+] Ahora podemos analizar los pasos y automatizarlos\n")
    
except KeyboardInterrupt:
    print("\n\n[!] Navegador cerrado por el usuario (Ctrl+C)")
    try:
        driver.quit()
    except:
        pass
    sys.exit(0)
    
except Exception as e:
    print(f"\n[-] Error: {str(e)}")
    try:
        driver.quit()
    except:
        pass
    sys.exit(1)
