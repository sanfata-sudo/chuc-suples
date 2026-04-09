#!/usr/bin/env python3
"""
generar.py — Generación automática del Dashboard CHUC RRHH
Lee CONTROL_SUPLES.ods (sección SUPLE) y ESTADO_VALORACIÓN.ods (sección Ayudas)
y genera el index.html con todos los datos actualizados.
"""

import pandas as pd
import json
import sys
import os
from datetime import datetime

ODS_SUPLES    = 'CONTROL_SUPLES.ods'
ODS_AYUDAS    = 'ESTADO_VALORACIÓN.ods'
TEMPLATE_FILE = 'template.html'
OUTPUT_FILE   = 'index.html'

def fmt_date(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    try:
        if hasattr(v, 'strftime'):
            return v.strftime('%Y-%m-%d')
        s = str(v).strip()
        if not s or s in ('nan', 'NaT', 'None', ''):
            return None
        d = pd.to_datetime(s, errors='coerce')
        return None if pd.isna(d) else d.strftime('%Y-%m-%d')
    except Exception:
        return None

def safe_str(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ''
    return str(v).strip()

def safe_num(v):
    try:
        x = float(v)
        return 0.0 if pd.isna(x) else x
    except Exception:
        return 0.0

def procesar_suples(ruta):
    print(f"Leyendo {ruta}...")
    df = pd.read_excel(ruta, engine='odf', sheet_name='1_Datos_generales')
    df = df.dropna(subset=['AÑO', 'ID'])
    df['AÑO'] = pd.to_numeric(df['AÑO'], errors='coerce')
    df = df[df['AÑO'].notna()].copy()
    df['AÑO'] = df['AÑO'].astype(int)
    print(f"  -> {len(df)} convocatorias")
    registros = []
    for _, r in df.iterrows():
        reg = {
            'AÑO': int(r['AÑO']),
            'ID': safe_str(r.get('ID','')),
            'CATEGORIA': safe_str(r.get('CATEGORIA','')),
            'AREA': safe_str(r.get('AREA','')),
            'DIRECCION': safe_str(r.get('DIRECCION','')),
            'PRESENTADOS': safe_num(r.get('PRESENTADOS',0)),
            'ADMITIDOS PROV': safe_num(r.get('ADMITIDOS PROV',0)),
            'EXCLUIDOS PROV': safe_num(r.get('EXCLUIDOS PROV',0)),
            'ADMITIDOS DEF': safe_num(r.get('ADMITIDOS DEF',0)),
            'EXCLUIDOS DEF': safe_num(r.get('EXCLUIDOS DEF',0)),
            'N.º RECL 1': safe_num(r.get('N.º RECL 1',0)),
            'N.º RECL 2': safe_num(r.get('N.º RECL 2',0)),
            'TIEMPO EN RESOLVER (días)': safe_num(r.get('TIEMPO EN RESOLVER (días)',0)) or None,
            'PUBL BASES': fmt_date(r.get('PUBL BASES')),
            'FECHA BASES': fmt_date(r.get('FECHA BASES')),
            'PUBL AD&EX PROV': fmt_date(r.get('PUBL AD&EX PROV')),
            'FECHA RG AD&EX PROV': fmt_date(r.get('FECHA RG AD&EX PROV')),
            'PUBL RG LISTA PROV': fmt_date(r.get('PUBL RG LISTA PROV')),
            'FECHA RG LISTA PROV': fmt_date(r.get('FECHA RG LISTA PROV')),
            'PUBL LISTA DEF': fmt_date(r.get('PUBL LISTA DEF')),
            'FECHA RG LISTA DEF': fmt_date(r.get('FECHA RG LISTA DEF')),
            'RG BASES': safe_str(r.get('RG BASES','')),
            'RG LISTA DEF': safe_str(r.get('RG LISTA DEF','')),
        }
        if reg['TIEMPO EN RESOLVER (días)'] == 0.0:
            reg['TIEMPO EN RESOLVER (días)'] = None
        registros.append(reg)
    return registros

def procesar_ayudas(ruta):
    """
    Estructura del ESTADO_VALORACION.ods (hoja A, sin header):
      Fila 2:  MIEMBRO CVAS-HCU | A | B | C | ... | AÑADIDOS | TOTAL  (cabeceras)
      Fila 3+: datos de cada miembro
      Fila vacía
      Fila "Total de solicitudes" | numero
      Fila "Valoradas"            | numero
      Fila "% realizado"          | numero
    """
    print(f"Leyendo {ruta}...")
    df = pd.read_excel(ruta, engine='odf', sheet_name='A', header=None)
    COLORES = ['#1A8080','#0B2D5F','#C89B2A','#8EA0BA','#D0DAE8','#7c3aed','#C0392B','#1A7850']

    # Letras: fila 2, columnas 1-16 (sin TOTAL)
    letras_raw = df.iloc[2, 1:17].tolist()
    letras = [str(x).strip() for x in letras_raw
              if str(x).strip() not in ('nan','None','TOTAL','')]

    # Miembros: filas 3 en adelante hasta fila vacía
    miembros = []
    for i in range(3, min(15, len(df))):
        nombre = safe_str(df.iloc[i, 0])
        if not nombre:
            break
        vals_raw = df.iloc[i, 1:17].tolist()  # 16 columnas de letras
        vals = [int(safe_num(v)) for v in vals_raw]
        total = sum(vals)
        partes = nombre.split()
        nombre_corto = partes[-2] if len(partes) >= 2 else nombre
        miembros.append({
            'nombre': nombre,
            'nombreCorto': nombre_corto,
            'vals': vals,
            'total': total,
            'color': COLORES[len(miembros) % len(COLORES)]
        })

    # Totales globales: buscar las filas por texto
    total_sol = 0
    valoradas  = 0
    pct        = 0.0
    for i in range(len(df)):
        cell = safe_str(df.iloc[i, 0]).lower()
        if 'total de solicitudes' in cell:
            total_sol = int(safe_num(df.iloc[i, 1]))
        elif 'valoradas' in cell and 'total' not in cell:
            valoradas = int(safe_num(df.iloc[i, 1]))
        elif '% realizado' in cell or 'realizado' in cell:
            pct = round(safe_num(df.iloc[i, 1]), 2)

    if pct == 0.0 and total_sol > 0:
        pct = round(valoradas / total_sol * 100, 2)

    print(f"  -> {total_sol} solicitudes, {valoradas} valoradas ({pct}%), {len(miembros)} miembros")
    return {'total': total_sol, 'valoradas': valoradas, 'pct': pct,
            'letras': letras, 'miembros': miembros}

def generar_html(reg_suples, datos_ayudas, template_path, output_path):
    print(f"Leyendo template {template_path}...")
    with open(template_path,'r',encoding='utf-8') as f:
        tmpl = f.read()
    for marker in ['// __DATOS_SUPLES__','// __DATOS_AYUDAS__']:
        if marker not in tmpl:
            print(f"ERROR: El template no contiene {marker}")
            sys.exit(1)
    fecha = datetime.now().strftime('%d/%m/%Y %H:%M')
    n = len(reg_suples)
    json_suples = json.dumps(reg_suples, ensure_ascii=False, separators=(',',':'))
    bloque_suples = f"""// Generado el {fecha} · {n} convocatorias
(function initDefault(){{
  const rows = {json_suples};
  DATA = processRows(rows, 'CONTROL_SUPLES.ods');
  renderAll();
  setTimeout(()=>{{if(!sessionStorage.getItem('welcomed')){{sessionStorage.setItem('welcomed','1');toast('\u2705 {n} convocatorias \u00b7 '+DATA.fechaODS, 3500);}}}}, 800);
}})();

"""
    json_ayudas = json.dumps(datos_ayudas, ensure_ascii=False, separators=(',',':'))
    bloque_ayudas = f"const AY = {json_ayudas};"
    html = tmpl.replace('// __DATOS_SUPLES__\n\n', bloque_suples)
    html = html.replace('// __DATOS_AYUDAS__', bloque_ayudas)
    with open(output_path,'w',encoding='utf-8') as f:
        f.write(html)
    print(f"OK: {output_path} generado ({len(html)/1024:.1f} KB)")

def main():
    missing = [f for f in [ODS_SUPLES, TEMPLATE_FILE] if not os.path.exists(f)]
    if missing:
        print(f"ERROR: Faltan archivos: {missing}")
        sys.exit(1)
    print("Generando Dashboard CHUC RRHH...")
    reg_suples = procesar_suples(ODS_SUPLES)
    if os.path.exists(ODS_AYUDAS):
        datos_ayudas = procesar_ayudas(ODS_AYUDAS)
    else:
        print(f"AVISO: {ODS_AYUDAS} no encontrado — seccion Ayudas sin datos")
        datos_ayudas = {'total':0,'valoradas':0,'pct':0.0,'letras':[],'miembros':[]}
    generar_html(reg_suples, datos_ayudas, TEMPLATE_FILE, OUTPUT_FILE)
    print("Listo!")

if __name__ == '__main__':
    main()
