#!/usr/bin/env python3
"""
generar.py — Script de generación automática del Dashboard SUPLES · CHUC
Lee el archivo CONTROL_SUPLES.ods y genera el index.html con los datos actualizados.
Se ejecuta automáticamente desde GitHub Actions cuando se sube un ODS nuevo.
"""

import pandas as pd
import json
import sys
import os
from datetime import datetime

ODS_FILE     = 'CONTROL_SUPLES.ods'
TEMPLATE_FILE = "plantilla.html"
OUTPUT_FILE  = 'index.html'

def fmt_date(v):
    """Convierte un valor de fecha a string ISO o None."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    try:
        if hasattr(v, 'strftime'):
            return v.strftime('%Y-%m-%d')
        s = str(v).strip()
        if not s or s in ('nan', 'NaT', 'None', ''):
            return None
        d = pd.to_datetime(s, errors='coerce')
        if pd.isna(d):
            return None
        return d.strftime('%Y-%m-%d')
    except Exception:
        return None

def safe_str(v):
    """Convierte a string seguro."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ''
    return str(v).strip()

def safe_num(v):
    """Convierte a número seguro."""
    try:
        x = float(v)
        return 0.0 if pd.isna(x) else x
    except Exception:
        return 0.0

def procesar_ods(ruta):
    """Lee el ODS y devuelve la lista de registros como dicts."""
    print(f"📂 Leyendo {ruta}...")

    df = pd.read_excel(ruta, engine='odf', sheet_name='1_Datos_generales')
    df = df.dropna(subset=['AÑO', 'ID'])
    df['AÑO'] = pd.to_numeric(df['AÑO'], errors='coerce')
    df = df[df['AÑO'].notna()].copy()
    df['AÑO'] = df['AÑO'].astype(int)

    print(f"✅ {len(df)} convocatorias encontradas")

    registros = []
    for _, r in df.iterrows():
        reg = {
            'AÑO': int(r['AÑO']),
            'ID': safe_str(r.get('ID', '')),
            'CATEGORIA': safe_str(r.get('CATEGORIA', '')),
            'AREA': safe_str(r.get('AREA', '')),
            'DIRECCION': safe_str(r.get('DIRECCION', '')),
            'PRESENTADOS': safe_num(r.get('PRESENTADOS', 0)),
            'ADMITIDOS PROV': safe_num(r.get('ADMITIDOS PROV', 0)),
            'EXCLUIDOS PROV': safe_num(r.get('EXCLUIDOS PROV', 0)),
            'ADMITIDOS DEF': safe_num(r.get('ADMITIDOS DEF', 0)),
            'EXCLUIDOS DEF': safe_num(r.get('EXCLUIDOS DEF', 0)),
            'N.º RECL 1': safe_num(r.get('N.º RECL 1', 0)),
            'N.º RECL 2': safe_num(r.get('N.º RECL 2', 0)),
            'TIEMPO EN RESOLVER (días)': safe_num(r.get('TIEMPO EN RESOLVER (días)', 0)) or None,
            'PUBL BASES': fmt_date(r.get('PUBL BASES')),
            'FECHA BASES': fmt_date(r.get('FECHA BASES')),
            'PUBL AD&EX PROV': fmt_date(r.get('PUBL AD&EX PROV')),
            'FECHA RG AD&EX PROV': fmt_date(r.get('FECHA RG AD&EX PROV')),
            'PUBL RG LISTA PROV': fmt_date(r.get('PUBL RG LISTA PROV')),
            'FECHA RG LISTA PROV': fmt_date(r.get('FECHA RG LISTA PROV')),
            'PUBL LISTA DEF': fmt_date(r.get('PUBL LISTA DEF')),
            'FECHA RG LISTA DEF': fmt_date(r.get('FECHA RG LISTA DEF')),
            'RG BASES': safe_str(r.get('RG BASES', '')),
            'RG LISTA DEF': safe_str(r.get('RG LISTA DEF', '')),
        }
        # Limpiar None en tiempo
        if reg['TIEMPO EN RESOLVER (días)'] == 0.0:
            reg['TIEMPO EN RESOLVER (días)'] = None
        registros.append(reg)

    return registros

def generar_html(registros, template_path, output_path):
    """Inyecta los datos en el template y genera el index.html."""
    print(f"📄 Leyendo template {template_path}...")

    with open(template_path, 'r', encoding='utf-8') as f:
        template = f.read()

    if '// __DATOS_SUPLES__' not in template:
        print("❌ ERROR: El template no contiene el marcador // __DATOS_SUPLES__")
        sys.exit(1)

    # Serializar datos a JSON compacto
    json_data = json.dumps(registros, ensure_ascii=False, separators=(',', ':'))
    n = len(registros)
    fecha = datetime.now().strftime('%d/%m/%Y %H:%M')

    bloque = f"""// ═══════ DATOS POR DEFECTO ═══════
// Generado automáticamente el {fecha} · {n} convocatorias
(function initDefault(){{
  const rows = {json_data};
  DATA = processRows(rows, 'CONTROL_SUPLES.ods');
  renderAll();
  setTimeout(()=>{{
    if(!sessionStorage.getItem('welcomed')){{
      sessionStorage.setItem('welcomed','1');
      toast('✅ {n} convocatorias cargadas · Actualizado {fecha}', 3500);
    }}
  }}, 800);
}})();

"""

    html = template.replace('// __DATOS_SUPLES__\n\n', bloque)

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html)

    size_kb = len(html) / 1024
    print(f"✅ {output_path} generado · {size_kb:.1f} KB · {n} convocatorias")

def main():
    # Verificar que existe el ODS
    if not os.path.exists(ODS_FILE):
        print(f"❌ ERROR: No se encuentra el archivo {ODS_FILE}")
        print(f"   Archivos en el directorio actual: {os.listdir('.')}")
        sys.exit(1)

    # Verificar que existe el template
    if not os.path.exists(TEMPLATE_FILE):
        print(f"❌ ERROR: No se encuentra el archivo {TEMPLATE_FILE}")
        sys.exit(1)

    print("🚀 Iniciando generación del Dashboard SUPLES · CHUC")
    print(f"   ODS:      {ODS_FILE}")
    print(f"   Template: {TEMPLATE_FILE}")
    print(f"   Output:   {OUTPUT_FILE}")
    print()

    registros = procesar_ods(ODS_FILE)
    generar_html(registros, TEMPLATE_FILE, OUTPUT_FILE)

    print()
    print("🎉 ¡Listo! El index.html ha sido actualizado.")

if __name__ == '__main__':
    main()
