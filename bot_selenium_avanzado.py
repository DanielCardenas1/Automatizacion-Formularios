"""
Bot Avanzado Selenium - RUB Online
Funcionalidades para filtros, búsquedas en listas y validaciones complejas
"""

from bot_selenium import RUBBot
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
import time


class RUBBotAvanzado(RUBBot):
    """Extiende RUBBot con funcionalidades más avanzadas"""
    
    def seleccionar_dropdown(self, selector_id, valor):
        """
        Selecciona una opción de un dropdown
        
        Args:
            selector_id (str): ID del elemento select
            valor (str): Valor o texto de la opción a seleccionar
        """
        print(f"\n[*] Seleccionando '{valor}' en dropdown {selector_id}...")
        try:
            elemento = self.wait.until(
                EC.presence_of_element_located((By.ID, selector_id))
            )
            select = Select(elemento)
            
            # Intenta seleccionar por valor
            try:
                select.select_by_value(valor)
                print(f"[+] Seleccionado por valor: {valor}")
            except:
                # Si no funciona, intenta por texto visible
                select.select_by_visible_text(valor)
                print(f"[+] Seleccionado por texto: {valor}")
            
            return True
        except Exception as e:
            print(f"[-] Error al seleccionar dropdown: {str(e)}")
            return False
    
    def aplicar_filtro(self, filtros_dict):
        """
        Aplica múltiples filtros a la búsqueda
        
        Args:
            filtros_dict (dict): Diccionario con los filtros
                                  formato: {'campo': 'valor'}
        """
        print(f"\n[*] Aplicando {len(filtros_dict)} filtro(s)...")
        try:
            for campo, valor in filtros_dict.items():
                print(f"[*] Filtro: {campo} = {valor}")
                
                # Buscar campo de entrada o select
                try:
                    elemento = self.driver.find_element(By.NAME, campo)
                    
                    # Si es un select, usar método especial
                    if elemento.tag_name == 'select':
                        self.seleccionar_dropdown(campo, valor)
                    else:
                        # Si es un input regular
                        elemento.clear()
                        elemento.send_keys(valor)
                        print(f"[+] Filtro aplicado: {campo}")
                except:
                    print(f"[!] No se encontró campo: {campo}")
            
            return True
        except Exception as e:
            print(f"[-] Error aplicando filtros: {str(e)}")
            return False
    
    def buscar_en_tabla(self, termino_busqueda):
        """
        Busca en tabla usando campo de búsqueda rápida
        
        Args:
            termino_busqueda (str): Término a buscar
        """
        print(f"\n[*] Buscando '{termino_busqueda}' en tabla...")
        try:
            # Buscar campo de búsqueda rápida (usualmente en GridView de ASP.NET)
            campos_busqueda = [
                "gvDatos_GridViewFilterTextFieldItem",
                "txtFiltro",
                "searchBox",
                "search"
            ]
            
            for campo in campos_busqueda:
                try:
                    elemento = self.driver.find_element(By.ID, campo)
                    elemento.clear()
                    elemento.send_keys(termino_busqueda)
                    elemento.send_keys(Keys.RETURN)
                    print(f"[+] Búsqueda completada en campo: {campo}")
                    time.sleep(2)
                    return True
                except:
                    continue
            
            print("[-] No se encontró campo de búsqueda")
            return False
            
        except Exception as e:
            print(f"[-] Error en búsqueda: {str(e)}")
            return False
    
    def extraer_tabla_completa(self, id_tabla="gvResultados"):
        """
        Extrae todos los datos de una tabla
        
        Args:
            id_tabla (str): ID del elemento tabla
            
        Returns:
            list: Lista de diccionarios con los datos de la tabla
        """
        print(f"\n[*] Extrayendo datos de tabla: {id_tabla}")
        try:
            tabla = self.wait.until(
                EC.presence_of_element_located((By.ID, id_tabla))
            )
            
            datos = []
            
            # Obtener encabezados
            encabezados = []
            filas_header = tabla.find_elements(By.TAG_NAME, "thead")
            if filas_header:
                for th in filas_header[0].find_elements(By.TAG_NAME, "th"):
                    encabezados.append(th.text.strip())
            
            print(f"[+] Encabezados encontrados: {encabezados}")
            
            # Obtener filas de datos
            tbody = tabla.find_element(By.TAG_NAME, "tbody")
            filas = tbody.find_elements(By.TAG_NAME, "tr")
            
            print(f"[+] Se encontraron {len(filas)} filas de datos")
            
            for fila in filas:
                celdas = fila.find_elements(By.TAG_NAME, "td")
                if celdas:
                    fila_dict = {}
                    for idx, celda in enumerate(celdas):
                        if idx < len(encabezados):
                            fila_dict[encabezados[idx]] = celda.text.strip()
                    datos.append(fila_dict)
            
            return datos
            
        except Exception as e:
            print(f"[-] Error extrayendo tabla: {str(e)}")
            return []
    
    def validar_campos_especificos(self, id_elemento, validaciones):
        """
        Valida campos específicos del elemento
        
        Args:
            id_elemento (str): ID del elemento a validar
            validaciones (dict): Diccionario con las validaciones
                                formato: {'campo': 'valor_esperado'}
        """
        print(f"\n[*] Validando elemento: {id_elemento}")
        try:
            elemento = self.driver.find_element(By.ID, id_elemento)
            resultados = {}
            
            for campo, valor_esperado in validaciones.items():
                if campo == 'text':
                    valor_actual = elemento.text
                elif campo == 'value':
                    valor_actual = elemento.get_attribute('value')
                else:
                    valor_actual = elemento.get_attribute(campo)
                
                coincide = str(valor_esperado).strip() == str(valor_actual).strip()
                resultados[campo] = {
                    'esperado': valor_esperado,
                    'actual': valor_actual,
                    'coincide': coincide
                }
                
                estado = "✓" if coincide else "✗"
                print(f"  [{estado}] {campo}: {valor_actual}")
            
            return resultados
            
        except Exception as e:
            print(f"[-] Error validando elemento: {str(e)}")
            return {}
    
    def marcar_checkbox(self, id_checkbox, marcar=True):
        """
        Marca o desmarca un checkbox
        
        Args:
            id_checkbox (str): ID del checkbox
            marcar (bool): True para marcar, False para desmarcar
        """
        print(f"\n[*] {'Marcando' if marcar else 'Desmarcando'} checkbox: {id_checkbox}")
        try:
            checkbox = self.driver.find_element(By.ID, id_checkbox)
            
            if marcar and not checkbox.is_selected():
                checkbox.click()
                print(f"[+] Checkbox marcado")
            elif not marcar and checkbox.is_selected():
                checkbox.click()
                print(f"[+] Checkbox desmarcado")
            else:
                print(f"[!] Checkbox ya está {'marcado' if marcar else 'desmarcado'}")
            
            return True
        except Exception as e:
            print(f"[-] Error con checkbox: {str(e)}")
            return False
    
    def esperar_carga_tabla(self, id_tabla="gvResultados", timeout=20):
        """
        Espera a que una tabla se cargue completamente
        
        Args:
            id_tabla (str): ID de la tabla
            timeout (int): Segundos a esperar
        """
        print(f"\n[*] Esperando carga de tabla...")
        try:
            WebDriverWait(self.driver, timeout).until(
                EC.presence_of_element_located((By.ID, id_tabla))
            )
            print(f"[+] Tabla cargada correctamente")
            return True
        except:
            print(f"[-] Timeout esperando tabla")
            return False
    
    def hacer_clic_en_boton(self, id_boton):
        """
        Hace clic en un botón
        
        Args:
            id_boton (str): ID del botón
        """
        print(f"\n[*] Haciendo clic en botón: {id_boton}")
        try:
            boton = self.wait.until(
                EC.element_to_be_clickable((By.ID, id_boton))
            )
            boton.click()
            print(f"[+] Botón presionado")
            time.sleep(1)
            return True
        except Exception as e:
            print(f"[-] Error al hacer clic: {str(e)}")
            return False
    
    def obtener_reporte_filtrado(self, url, usuario, contraseña, 
                                excel_path, filtros=None):
        """
        Ejecuta un flujo completo con filtros y genera reporte
        
        Args:
            url (str): URL de acceso
            usuario (str): Usuario
            contraseña (str): Contraseña
            excel_path (str): Ruta Excel
            filtros (dict): Filtros a aplicar
        """
        print("="*60)
        print("INICIANDO VERIFICACIÓN CON FILTROS")
        print("="*60)
        
        try:
            # Inicializar
            self.inicializar_driver()
            
            # Acceder
            if not self.acceder_pagina(url):
                return
            
            # Login
            if not self.hacer_login(usuario, contraseña):
                return
            
            # Aplicar filtros si existen
            if filtros:
                self.aplicar_filtro(filtros)
                time.sleep(2)
            
            # Extraer tabla completa
            datos_pagina = self.extraer_tabla_completa()
            
            # Leer Excel
            datos_excel = self.leer_excel(excel_path)
            
            # Comparar
            print(f"\n[*] Comparando {len(datos_excel)} registros del Excel con {len(datos_pagina)} de la página...")
            
            diferencias = 0
            for registro_excel in datos_excel:
                encontrado = False
                for registro_pagina in datos_pagina:
                    # Buscar coincidencia por documento
                    doc_excel = str(registro_excel.get('Documento', '')).strip()
                    doc_pagina = str(registro_pagina.get('Documento', '')).strip()
                    
                    if doc_excel and doc_excel == doc_pagina:
                        encontrado = True
                        coincide, difs = self.comparar_datos(registro_excel, registro_pagina)
                        if not coincide:
                            diferencias += 1
                        break
                
                if not encontrado:
                    print(f"[-] Documento {registro_excel.get('Documento')} no encontrado en página")
                    diferencias += 1
            
            self.mostrar_resumen()
            
        except Exception as e:
            print(f"[-] Error: {str(e)}")
        finally:
            self.cerrar_driver()


# Importar WebDriverWait y EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


if __name__ == "__main__":
    # Ejemplo de uso del bot avanzado
    
    bot = RUBBotAvanzado()
    
    # Ejemplo con filtros
    filtros = {
        # 'ddlCiudad': 'DUITAMA',  # Descomenta si hay dropdown de ciudad
        # 'ddlEstado': 'ACTIVO',    # Descomenta si hay dropdown de estado
    }
    
    bot.obtener_reporte_filtrado(
        url="https://rubonline.icbf.gov.co/DefaultF.aspx",
        usuario="angie.cardenas",
        contraseña="Celeste1020*",
        excel_path="CARGUE MASIVO_DUITAMA A_ICBF_2026.xlsx",
        filtros=filtros
    )
