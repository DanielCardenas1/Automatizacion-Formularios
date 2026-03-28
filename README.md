# 🤖 Automatización de Cargue y Verificación de Formularios Web

Sistema de automatización RPA (Robotic Process Automation) desarrollado para eliminar el ingreso manual de información en plataformas web de alto volumen. El bot navega automáticamente el sistema, inicia sesión, lee los datos desde archivos Excel y completa los formularios de forma masiva, verificando la consistencia de cada registro.

> **Nota legal:** Este documento está redactado para portafolio profesional. No se incluyen nombres de instituciones, datos personales ni credenciales. Toda la información sensible se gestiona mediante variables de entorno.

---

## 📋 Stack Tecnológico

- **Automatización:** Python 3 + Selenium WebDriver
- **Lectura de datos:** openpyxl
- **Gestión de drivers:** webdriver-manager
- **Interfaz:** Chrome (modo visible o headless)
- **Reportes:** Archivos `.txt` generados automáticamente

---

## 🎯 Problema que resuelve

El proceso original requería que operadores ingresaran manualmente cientos de registros en un formulario web, uno por uno, copiando datos desde archivos Excel. Cada registro implicaba login, navegación, llenado de múltiples campos y verificación visual. El proceso era lento, propenso a errores de transcripción y no escalable.

Este sistema automatiza el flujo completo: lee los datos del Excel, accede al sistema, navega hasta el formulario correcto y completa cada registro de forma autónoma, generando al final un reporte de resultados.

---

## ⚡ Funcionalidades

### Cargue masivo
- ✅ Login automático al sistema web
- ✅ Lectura y validación de datos desde Excel
- ✅ Navegación automática hasta el formulario objetivo
- ✅ Llenado de campos por sección (datos básicos, ubicación, estado)
- ✅ Soporte para campos de texto, selectores y carga de imágenes
- ✅ Manejo de errores y reintentos automáticos

### Verificación y control de calidad
- ✅ Cruce entre datos del Excel y registros en plataforma
- ✅ Detección de inconsistencias por sección y campo
- ✅ Reporte detallado de diferencias encontradas
- ✅ Soporte para múltiples grupos de trabajo (lotes A, B, C, D...)

### Utilidades de diagnóstico
- ✅ Inspección de estructura de formularios
- ✅ Diagnóstico de sesión post-login
- ✅ Debug de selectores y elementos de página
- ✅ Grabador de clicks para mapeo de interacciones

---

## 📁 Estructura del Proyecto

```
Automatizacion-Formularios/
├── bot_carga_masiva_d2.py          # Bot principal de cargue masivo
├── bot_selenium_avanzado.py        # Bot con filtros y validaciones complejas
├── bot_selenium.py                 # Bot base con funciones reutilizables
├── bot_final.py                    # Versión estable de producción
├── bot_automatizado.py             # Flujo completamente automatizado
├── bot_interactivo.py              # Modo interactivo para casos especiales
├── bot_manual_interactivo.py       # Asistido por operador
│
├── verificar_excel_vs_formulario.py        # Verificación cruzada principal
├── verificar_excel_vs_formulario_A2_A3.py  # Verificación por grupos
├── verificar_registro.py                   # Verificación de registro individual
│
├── analizar_f2.py                  # Análisis de sección F2
├── analizar_uds.py                 # Análisis de unidades de servicio
├── buscar_doc.py                   # Búsqueda por documento
├── check_excel_data.py             # Validación previa del Excel
│
├── diagnostico_pagina.py           # Diagnóstico de estructura de página
├── diagnostico_post_login.py       # Diagnóstico de sesión activa
├── bot_inspeccion_formulario.py    # Inspección detallada del formulario
├── inspeccion.py                   # Inspección general
├── inspeccion_detallada_menu.py    # Inspección del menú de navegación
├── inspect_c.py                    # Inspección sección C
├── inspect_c_detailed.py           # Inspección detallada sección C
├── inspect_d_structure.py          # Inspección estructura sección D
│
├── bot_grabador_clicks.py          # Grabador de interacciones
├── explorar_menu.py                # Exploración de menú
├── leer_excel.py                   # Utilidad de lectura de Excel
├── limpiar_fotos_c.py              # Limpieza de imágenes sección C
├── limpiar_fotos_f.py              # Limpieza de imágenes sección F
├── limpiar_fotos_g.py              # Limpieza de imágenes sección G
│
├── requirements.txt                # Dependencias del proyecto
├── install.sh                      # Script de instalación automática
├── .env.example                    # Plantilla de variables de entorno
├── GUIA_RAPIDA.md                  # Guía de inicio rápido
└── README.md                       # Este archivo
```

---

## 🚀 Inicio Rápido

### Prerrequisitos

- Python 3.8 o superior
- Google Chrome instalado
- Archivo Excel con los registros a cargar

### Instalación

**Opción A — Script automático:**
```bash
bash install.sh
```

**Opción B — Manual:**
```bash
pip install -r requirements.txt
```

### Configuración de credenciales

Copia el archivo de ejemplo y completa tus datos:
```bash
cp .env.example .env
```

Edita el archivo `.env`:
```env
BOT_USUARIO=tu_usuario_aqui
BOT_PASSWORD=tu_contraseña_aqui
```

> ⚠️ **Nunca subas el archivo `.env` a GitHub.** Está incluido en `.gitignore`.

### Configurar la ruta del Excel

En el bot que vayas a usar, ajusta la ruta al archivo Excel:
```python
excel_path = os.path.join(carpeta_actual, "TU_ARCHIVO.xlsx")
```

---

## ▶️ Uso

### Cargue masivo básico

```bash
python bot_selenium.py
```

### Cargue masivo avanzado (con filtros por ciudad y estado)

```bash
python bot_selenium_avanzado.py
```

### Verificación de registros cargados

```bash
python verificar_excel_vs_formulario.py
```

### ¿Qué hace el bot al ejecutarse?

1. Abre Chrome automáticamente
2. Accede al sistema web objetivo
3. Inicia sesión con las credenciales del `.env`
4. Lee los registros del archivo Excel
5. Navega hasta el formulario correspondiente
6. Completa cada campo por sección
7. Carga las imágenes si aplica
8. Guarda el registro y pasa al siguiente
9. Genera un reporte de resultados al finalizar

---

## 📊 Scripts Disponibles

| Script | Descripción |
|--------|-------------|
| `bot_selenium.py` | Bot base. Punto de partida para nuevas implementaciones |
| `bot_selenium_avanzado.py` | Bot con filtros, validaciones y manejo de errores avanzado |
| `bot_carga_masiva_d2.py` | Cargue masivo para grupos de trabajo específicos |
| `bot_final.py` | Versión estable lista para producción |
| `verificar_excel_vs_formulario.py` | Verificación cruzada completa |
| `verificar_excel_vs_formulario_A2_A3.py` | Verificación por subgrupos |
| `check_excel_data.py` | Valida el Excel antes de ejecutar el bot |
| `diagnostico_pagina.py` | Diagnostica la estructura de la página si algo falla |

---

## ⚙️ Configuración Avanzada

### Modo headless (sin ventana visible)

En `bot_selenium.py`, descomenta la línea:
```python
options.add_argument('--headless')
```

### Ajustar el número de registros a procesar

```python
for idx, registro in enumerate(datos_excel[:50], 1):  # Cambia 50 por el número deseado
```

### Ajustar timeout de espera

```python
self.wait = WebDriverWait(self.driver, 30)  # Segundos de espera máxima
```

### Aplicar filtros antes del cargue

```python
filtros = {
    'ddlCiudad': 'NOMBRE_CIUDAD',
    'ddlEstado': 'ACTIVO',
}
```

---

## 📄 Archivos de Salida

| Archivo | Contenido |
|---------|-----------|
| `reporte_verificacion_FECHA_HORA.txt` | Resultados de verificación cruzada |
| `reporte_cargue_FECHA_HORA.txt` | Log del proceso de cargue |

---

## ❓ Solución de Problemas

**`ModuleNotFoundError: No module named 'selenium'`**
```bash
pip install -r requirements.txt
```

**Chrome no abre**
→ Verifica que tienes Google Chrome instalado. webdriver-manager descarga el driver automáticamente.

**`Element not found` o `TimeoutException`**
→ Los selectores de la página pueden haber cambiado. Usa `diagnostico_pagina.py` para inspeccionar la estructura actual.

**El proceso es muy lento**
→ Aumenta el timeout en la línea `WebDriverWait(driver, 20)` — cambia `20` por `30` o más.

**Error de login**
→ Verifica que las credenciales en `.env` son correctas y que la sesión no está bloqueada en el sistema.

---

## 🔧 Stack Completo

```
selenium>=4.10.0
webdriver-manager>=3.9.0
openpyxl>=3.10.0
```

---

## 📌 Notas de Desarrollo

- El proyecto fue desarrollado de forma iterativa, con versiones sucesivas del bot (`bot_v2`, `bot_final`, etc.) adaptadas a los cambios en la estructura del formulario objetivo.
- Los scripts de inspección (`inspeccion.py`, `diagnostico_pagina.py`) fueron clave para mapear los selectores correctos del formulario antes de automatizar.
- El módulo de verificación cruzada permite auditar el cargue sin intervención manual, comparando campo por campo entre el Excel de origen y los datos registrados en el sistema.

---

## 🔗 Recursos

- [Documentación oficial Selenium Python](https://selenium-python.readthedocs.io/)
- [webdriver-manager](https://github.com/SergeyPirogov/webdriver_manager)
- [openpyxl](https://openpyxl.readthedocs.io/)

---

*Desarrollado por Daniel Gerardo Cárdenas Espinosa · Colombia*
