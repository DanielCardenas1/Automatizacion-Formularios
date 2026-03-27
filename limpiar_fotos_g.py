#!/usr/bin/env python3
import os
import unicodedata
from pathlib import Path

def remover_tildes_y_enie(texto):
    """Remove accents and ñ/Ñ from text"""
    # Reemplazar ñ/Ñ primero
    texto = texto.replace('ñ', 'n').replace('Ñ', 'N')
    # Normalizar y remover acentos
    texto_nfd = unicodedata.normalize('NFD', texto)
    salida = ''.join(char for char in texto_nfd if unicodedata.category(char) != 'Mn')
    return unicodedata.normalize('NFC', salida)

def renombrar_archivos_duitama_g():
    base_dir = Path("/Users/stevenruiz/Downloads/CARGUE MASIVO EIH DUITAMA/DUITAMA G")
    renombrados = 0
    
    for foto_file in base_dir.rglob("*.jpg"):
        nombre_original = foto_file.name
        nombre_limpio = remover_tildes_y_enie(nombre_original)
        
        if nombre_original != nombre_limpio:
            nueva_ruta = foto_file.parent / nombre_limpio
            print(f"Renombrando: {nombre_original}")
            print(f"         a: {nombre_limpio}")
            foto_file.rename(nueva_ruta)
            renombrados += 1
    
    print(f"\n✓ Total archivos renombrados: {renombrados}")

if __name__ == "__main__":
    renombrar_archivos_duitama_g()
