#!/usr/bin/env python3
import os
import unicodedata
from pathlib import Path

def remover_tildes_y_enie(texto):
    """Remove accents and ñ/Ñ from text"""
    texto = texto.replace('ñ', 'n').replace('Ñ', 'N')
    texto_nfd = unicodedata.normalize('NFD', texto)
    salida = ''.join(char for char in texto_nfd if unicodedata.category(char) != 'Mn')
    return unicodedata.normalize('NFC', salida)

def renombrar_archivos_duitama_c():
    base_dir = Path("/Users/stevenruiz/Downloads/CARGUE MASIVO EIH DUITAMA/DUITAMA C")
    renombrados = 0
    extensiones_validas = {".jpg", ".jpeg", ".png"}
    
    for foto_file in base_dir.rglob("*"):
        if not foto_file.is_file() or foto_file.suffix.lower() not in extensiones_validas:
            continue

        nombre_original = foto_file.name
        nombre_limpio = remover_tildes_y_enie(nombre_original)
        
        if nombre_original != nombre_limpio:
            nueva_ruta = foto_file.parent / nombre_limpio
            if nueva_ruta.exists():
                print(f"[!] Se omite por colision: {nombre_original} -> {nombre_limpio}")
                continue
            print(f"Renombrando: {nombre_original}")
            print(f"         a: {nombre_limpio}")
            foto_file.rename(nueva_ruta)
            renombrados += 1
    
    print(f"\n✓ Total archivos renombrados: {renombrados}")

if __name__ == "__main__":
    renombrar_archivos_duitama_c()
