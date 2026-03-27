#!/bin/bash

# Script de instalación del Bot Selenium para RUB Online
# Uso: bash install.sh

echo "=================================================="
echo "Instalación del Bot Selenium - RUB Online"
echo "=================================================="

# Verificar si Python está instalado
echo "[*] Verificando Python..."
if ! command -v python3 &> /dev/null; then
    echo "[-] Python3 no está instalado"
    echo "[*] Instalando Python usando Homebrew..."
    
    if ! command -v brew &> /dev/null; then
        echo "[-] Homebrew no está instalado"
        echo "[*] Instalando Homebrew..."
        /bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"
    fi
    
    brew install python3
else
    echo "[+] Python3 encontrado: $(python3 --version)"
fi

# Verificar si pip está instalado
echo "[*] Verificando pip..."
if ! command -v pip3 &> /dev/null; then
    echo "[-] pip3 no está instalado"
    python3 -m ensurepip --upgrade
else
    echo "[+] pip3 encontrado: $(pip3 --version)"
fi

# Crear entorno virtual (opcional pero recomendado)
echo ""
echo "[*] ¿Deseas crear un entorno virtual? (recomendado)"
read -p "Responde 's' para sí, 'n' para no: " create_venv

if [ "$create_venv" = "s" ] || [ "$create_venv" = "S" ]; then
    echo "[*] Creando entorno virtual..."
    python3 -m venv venv
    
    # Activar entorno virtual
    source venv/bin/activate
    echo "[+] Entorno virtual activado"
else
    echo "[!] Usando Python global"
fi

# Actualizar pip
echo ""
echo "[*] Actualizando pip..."
python3 -m pip install --upgrade pip

# Instalar dependencias
echo ""
echo "[*] Instalando dependencias del proyecto..."
pip install -r requirements.txt

echo ""
echo "[+] ¡Instalación completada!"
echo ""
echo "=================================================="
echo "Próximos pasos:"
echo "=================================================="
echo "1. Edita bot_selenium.py si necesitas cambiar la configuración"
echo "2. Ejecuta: python bot_selenium.py"
echo "3. El bot abrirá Chrome automáticamente"
echo ""
echo "Nota: Si creaste un entorno virtual, actvalo con:"
echo "  source venv/bin/activate"
echo "=================================================="
