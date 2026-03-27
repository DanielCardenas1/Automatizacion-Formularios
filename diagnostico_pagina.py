"""
Script de Diagnóstico - Inspecciona la página para encontrar los selectores correctos
"""

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import time
import json


def diagnosticar_pagina():
    """Accede a la página y muestra la estructura del formulario de login"""
    
    print("[*] Inicializando WebDriver...")
    options = webdriver.ChromeOptions()
    
    driver = webdriver.Chrome(
        service=Service(ChromeDriverManager().install()),
        options=options
    )
    
    try:
        url = "https://rubonline.icbf.gov.co/DefaultF.aspx"
        print(f"\n[*] Accediendo a {url}...")
        driver.get(url)
        time.sleep(5)
        
        print("\n[+] Página cargada ✓")
        print(f"[+] Título: {driver.title}")
        
        # Buscar todos los inputs
        print("\n" + "="*60)
        print("CAMPOS DE ENTRADA ENCONTRADOS:")
        print("="*60)
        
        inputs = driver.find_elements(By.TAG_NAME, "input")
        for idx, input_elem in enumerate(inputs):
            id_attr = input_elem.get_attribute("id")
            name = input_elem.get_attribute("name")
            tipo = input_elem.get_attribute("type")
            placeholder = input_elem.get_attribute("placeholder")
            
            print(f"\nInput #{idx}")
            print(f"  ID: {id_attr}")
            print(f"  Name: {name}")
            print(f"  Type: {tipo}")
            print(f"  Placeholder: {placeholder}")
        
        # Buscar botones
        print("\n" + "="*60)
        print("BOTONES ENCONTRADOS:")
        print("="*60)
        
        botones = driver.find_elements(By.TAG_NAME, "button")
        botones += driver.find_elements(By.TAG_NAME, "input[type='button']")
        botones += driver.find_elements(By.TAG_NAME, "input[type='submit']")
        
        for idx, boton in enumerate(botones):
            id_attr = boton.get_attribute("id")
            name = boton.get_attribute("name")
            value = boton.get_attribute("value")
            texto = boton.text
            
            print(f"\nBotón #{idx}")
            print(f"  ID: {id_attr}")
            print(f"  Name: {name}")
            print(f"  Value: {value}")
            print(f"  Texto: {texto}")
        
        # Buscar elementos por label
        print("\n" + "="*60)
        print("LABELS Y TEXTO VISIBLES:")
        print("="*60)
        
        labels = driver.find_elements(By.TAG_NAME, "label")
        for lbl in labels:
            print(f"  • {lbl.text}")
        
        # Guardar HTML para inspección manual
        print("\n" + "="*60)
        print("Guardando HTML de la página...")
        
        with open("pagina_debug.html", "w", encoding="utf-8") as f:
            f.write(driver.page_source)
        
        print("[+] HTML guardado en: pagina_debug.html")
        
        # Mostrar JavaScript/formularios
        print("\n" + "="*60)
        print("FORMULARIOS:")
        print("="*60)
        
        forms = driver.find_elements(By.TAG_NAME, "form")
        for idx, form in enumerate(forms):
            id_form = form.get_attribute("id")
            name_form = form.get_attribute("name")
            metodo = form.get_attribute("method")
            accion = form.get_attribute("action")
            
            print(f"\nFormulario #{idx}")
            print(f"  ID: {id_form}")
            print(f"  Name: {name_form}")
            print(f"  Method: {metodo}")
            print(f"  Action: {accion}")
        
        # Intentar encontrar campos de usuario/contraseña por atributos comunes
        print("\n" + "="*60)
        print("BÚSQUEDA DE CAMPOS TÍPICOS:")
        print("="*60)
        
        # Buscar por contenido de label
        try:
            labels_dict = {}
            for lbl in driver.find_elements(By.TAG_NAME, "label"):
                texto = lbl.text.lower()
                label_id = lbl.get_attribute("for")
                labels_dict[texto] = label_id
            
            print("\nLabels y sus IDs asociados:")
            for texto, id_asociado in labels_dict.items():
                print(f"  {texto} → {id_asociado}")
        except:
            pass
        
        print("\n" + "="*60)
        print("✅ Diagnóstico completado")
        print("="*60)
        
        print("\n💡 PRÓXIMOS PASOS:")
        print("1. Abre 'pagina_debug.html' en un editor de texto")
        print("2. Busca los campos de 'usuario' y 'contraseña'")
        print("3. Encuentra sus atributos 'id' o 'name'")
        print("4. Comparte esa información")
        
    except Exception as e:
        print(f"\n[-] Error: {str(e)}")
    
    finally:
        print("\n[*] Cerrando navegador...")
        driver.quit()
        print("[+] WebDriver cerrado")


if __name__ == "__main__":
    diagnosticar_pagina()
