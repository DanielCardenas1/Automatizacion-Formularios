# ⚡ Guía de Inicio Rápido - Bot RUB Online

## 1️⃣ Instalación (una sola vez)

### En Terminal, ejecuta:
```bash
cd "/Users/stevenruiz/Downloads/CARGUE MASIVO EIH DUITAMA"
bash install.sh
```

O si prefieres manualmente:
```bash
pip3 install -r requirements.txt
```

## 2️⃣ Ejecutar el Bot

### Opción A: Bot Simple (lectura de Excel y búsqueda básica)
```bash
python3 bot_selenium.py
```

### Opción B: Bot Avanzado (con filtros y validaciones complejas)
```bash
python3 bot_selenium_avanzado.py
```

## 3️⃣ ¿Qué hará el bot?

1. Se abrirá Chrome automáticamente  
2. Accederá a https://rubonline.icbf.gov.co/DefaultF.aspx  
3. Hará login con: `Usuario / Contraseña`  
4. Leerá el archivo Excel  
5. Buscará cada documento en la página  
6. Comparará los datos  
7. Generará un reporte en archivo `.txt`  

## 4️⃣ Archivos Generados

- `reporte_verificacion_FECHA_HORA.txt` - Reporte de resultados

## 5️⃣ Configuración Personalizada

### Cambiar usuario/contraseña
Edita `bot_selenium.py` línea ~318:
```python
USUARIO = "Usuario"
CONTRASEÑA = "Contraseña"
```

### Cambiar archivo Excel
Edita `bot_selenium.py` línea ~330:
```python
excel_path = os.path.join(carpeta_actual, "TU_ARCHIVO.xlsx")
```

### Cambiar número de registros a verificar
Edita `bot_selenium.py` línea ~270:
```python
for idx, registro in enumerate(datos_excel[:5], 1):  # Cambiar 5 por otro número
```

### Usar modo headless (sin ventana visible)
Edita `bot_selenium.py` línea ~48:
```python
options.add_argument('--headless')  # Descomenta
```

## 6️⃣ Ejemplo de Uso - Bot Avanzado con Filtros

En `bot_selenium_avanzado.py` al final:

```python
# Aplica filtros antes de buscar
filtros = {
    'ddlCiudad': 'DUITAMA',
    'ddlEstado': 'ACTIVO',
}

bot.obtener_reporte_filtrado(
    url="https://rubonline.icbf.gov.co/DefaultF.aspx",
    usuario="Usuario",
    contraseña="Contraseña",
    excel_path="CARGUE MASIVO_DUITAMA A_ICBF_2026.xlsx",
    filtros=filtros
)
```

## ❓ Troubleshooting

### Error: "ModuleNotFoundError: No module named 'selenium'"
→ Ejecuta: `pip3 install -r requirements.txt`

### Chrome no abre
→ Asegúrate de tener Chrome instalado: https://www.google.com/chrome/

### Error: "Element not found"
→ Los selectores de la página probablemente cambiaron
→ Contacta para ajustar los selectores en el código

### Proceso muy lento
→ Aumenta el timeout en `bot_selenium.py` línea 57:
→ `self.wait = WebDriverWait(self.driver, 30)` (más de 20)

## 📊 Archivos Excel Disponibles

Se encontraron estos archivos:
```
CARGUE MASIVO_DUITAMA A_ICBF_2026.xlsx
CARGUE MASIVO 2026_DUITAMA F.xlsx
CARGUE MASIVO 2026 _ DUITAMA C_.xlsx
DUITAMA B1/CARGUE MASIVO_DUITAMA B1_152381130854.xlsx
DUITAMA B2/CARGUE MASIVO_DUITAMA B2_152381148302.xlsx
DUITAMA B3/CARGUE MASIVO_DUITAMA B3_152381148302.xlsx
DUITAMA D/CARQUE MASIVO 2026_DUITAMA D.xlsx
DUITAMA E/CARGUE MASIVO DUITAMA E 2026.xlsx
DUITAMA F/CARGUE MASIVO 2026_DUITAMA F_ACTUALIZADO.xlsx
DUITAMA G/CARGUE MASIVO_DUITAMA_G1_G2_G3_2026.xlsm
UDS OPERACION DIRECTA.xlsx
```

## 🔗 Enlaces útiles

- [Documentación Selenium](https://selenium.dev/documentation/)
- [GitHub Selenium Python](https://github.com/SeleniumHQ/selenium/tree/master/py)
- [RUB Online](https://rubonline.icbf.gov.co/DefaultF.aspx)

## 💡 Tips

✅ Los reportes se guardan en la misma carpeta que el script  
✅ Ejecuta en terminal/Powershell, no en Python IDE  
✅ Cierra Chrome completamente antes de ejecutar de nuevo  
✅ Usa modo headless si quieres que no veas la ventana  

---

**¿Necesitas ayuda?** Revisa el README.md para documentación completa.
