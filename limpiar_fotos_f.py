#!/usr/bin/env python3
import os
import unicodedata
from pathlib import Path

def remover_acentos(texto):
    """Remueve acentos y caracteres especiales"""
    # Normalizar a NFD (decomposición)
    texto_nfd = unicodedata.normalize('NFD', texto)
    # Filtrar caracteres acentuados
    texto_limpio = ''.join(c for c in texto_nfd if unicodedata.category(c) != 'Mn')
    # Reemplazar Ñ por N
    texto_limpio = texto_limpio.replace('ñ', 'n').replace('Ñ', 'N')
    return texto_limpio

duitama_f_path = Path('/Users/stevenruiz/Downloads/CARGUE MASIVO EIH DUITAMA/DUITAMA F')

total_cambios = 0

# Procesar F1, F2, F3 - buscar todas las fotos recursivamente
for uds_folder in ['DUITAMA F1', 'DUITAMA F2', 'DUITAMA F3']:
    base_path = duitama_f_path / uds_folder
    
    if not base_path.exists():
        print(f"⚠️  No existe: {base_path}")
        continue
    
    print(f"\n📁 Procesando: {uds_folder}")
    
    # Buscar TODAS las fotos Y DOCUMENTOS recursivamente desde la carpeta base
    fotos = list(base_path.rglob('*.jpg')) + list(base_path.rglob('*.JPG')) + \
            list(base_path.rglob('*.jpeg')) + list(base_path.rglob('*.JPEG')) + \
            list(base_path.rglob('*.png')) + list(base_path.rglob('*.PNG')) + \
            list(base_path.rglob('*.pdf')) + list(base_path.rglob('*.PDF'))
    
    if not fotos:
        print(f"  (Sin archivos de foto)")
        continue
    
    print(f"  Encontrados: {len(fotos)} archivos")
    
    cambios = 0
    for archivo in sorted(fotos):
        nuevo_nombre = remover_acentos(archivo.name)
        if nuevo_nombre != archivo.name:
            nuevo_path = archivo.parent / nuevo_nombre
            try:
                archivo.rename(nuevo_path)
                print(f"  ✓ {archivo.name}")
                print(f"    → {nuevo_nombre}")
                cambios += 1
                total_cambios += 1
            except Exception as e:
                print(f"  ✗ Error: {archivo.name}: {e}")
    
    if cambios == 0:
        print(f"  ✓ (Sin cambios necesarios)")

print(f"\n{'='*60}")
print(f"✅ Proceso completado: {total_cambios} archivos renombrados")
print(f"{'='*60}")
