#!/usr/bin/env python3
"""
generar.py - Script de generacion automatica del Dashboard SUPLES CHUC
Lee CONTROL_SUPLES.ods (convocatorias) y ESTADO_VALORACION.ods (ayudas)
y genera el index.html con todos los datos actualizados.
Se ejecuta automaticamente desde GitHub Actions al subir cualquiera de los dos ODS.
"""

import pandas as pd
import json
import sys
import os
from datetime import datetime

ODS_FILE       = 'CONTROL_SUPLES.ods'
ODS_VALORACION = 'ESTADO_VALORACIÓN.ods'
TEMPLATE_FILE  = "template.html"
OUTPUT_FILE    = 'index.html'

COLORES_MIEMBRO = [
    '#1A8080','#163A78','#D97706','#C0392B','#1A7850',
    '#7B3F00','#4A5568','#6B46C1','#2C7A7B','#B7791F',
]

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
    """Convierte a numero seguro."""
    try:
        x = float(v)
        return 0.0 if pd.isna(x) else x
    except Exception:
        return 0.0

def safe_int(v):
    """Convierte a entero seguro."""
    try:
        x = float(v)
        return 0 if pd.isna(x) else int(x)
    except Exception:
        return 0

def procesar_ods(ruta):
    """Lee CONTROL_SUPLES.ods y devuelve lista de registros."""
    print(f"Leyendo {ruta}...")
    df = pd.read_excel(ruta, engine='odf', sheet_name='1_Datos_generales')
    df = df.dropna(subset=['AÑO', 'ID'])
    df['AÑO'] = pd.to_numeric(df['AÑO'], errors='coerce')
    df = df[df['AÑO'].notna()].copy()
    df['AÑO'] = df['AÑO'].astype(int)
    print(f"  {len(df)} convocatorias encontradas")
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
        if reg['TIEMPO EN RESOLVER (días)'] == 0.0:
            reg['TIEMPO EN RESOLVER (días)'] = None
        registros.append(reg)
    return registros

def procesar_valoracion(ruta):
    """
    Lee ESTADO_VALORACION.ods y devuelve el bloque JS 'const AY = {...};'.
    Detecta automaticamente la estructura del archivo.
    """
    print(f"Leyendo {ruta}...")
    try:
        xl = pd.ExcelFile(ruta, engine='odf')
        sheet = xl.sheet_names[0]
        # Leer sin encabezados primero para detectar estructura
        df_raw = pd.read_excel(ruta, engine='odf', sheet_name=sheet, header=None)
        print(f"  Hoja: '{sheet}' | Shape: {df_raw.shape}")
        print(f"  Primeras 3 filas:")
        for i in range(min(3, len(df_raw))):
            print(f"    Fila {i}: {list(df_raw.iloc[i])}")
        # Detectar fila de encabezados: buscar la fila con mas texto
        header_row = 0
        max_text = 0
        for i in range(min(5, len(df_raw))):
            row_text = sum(1 for v in df_raw.iloc[i] if isinstance(v, str) and str(v).strip())
            if row_text > max_text:
                max_text = row_text
                header_row = i
        # Releer con la fila de encabezados correcta
        df = pd.read_excel(ruta, engine='odf', sheet_name=sheet, header=header_row)
        cols = [str(c).strip() for c in df.columns]
        df.columns = cols
        print(f"  Encabezados detectados en fila {header_row}: {cols}")

        LETRAS = ['A','B','C','D','E-F','G','H','I-J-K','L','M',
                  'N-Ñ-O','P-Q','R','S-T','U-Z','ÁÑADIDOS']

        # Detectar columna de nombre del miembro
        nombre_col = cols[0]
        for c in cols:
            cu = c.upper()
            if cu in ('NOMBRE', 'MIEMBRO', 'MIEMBROS', 'NOMBRE COMPLETO',
                      'APELLIDOS', 'APELLIDO', 'MIEMBRO CVAS-HCU',
                      'MIEMBRO CVAS HCU', 'MIEMBROS CVAS-HCU'):
                nombre_col = c
                break

        # Detectar columna de total
        total_col = None
        for c in cols:
            if c.upper() in ('TOTAL', 'TOTAL VALORADAS', 'TOTAL SOLICITUDES',
                             'SUMA', 'TOTAL GENERAL'):
                total_col = c
                break

        # Detectar columna de total global de solicitudes
        total_solic_col = None
        for c in cols:
            cu = c.upper()
            if 'TOTAL' in cu and ('SOLIC' in cu or 'PRES' in cu):
                total_solic_col = c
                break

        miembros = []
        total_solicitudes = 0
        valoradas_total = 0

        for _, row in df.iterrows():
            nombre = safe_str(row.get(nombre_col, ''))
            if not nombre or nombre.upper() in ('TOTAL','SUMA','TOTALES',
                                                 'NAN','','TOTAL SOLICITUDES'):
                # Podria ser la fila de totales globales
                if nombre.upper() in ('TOTAL','TOTALES','TOTAL SOLICITUDES'):
                    if total_solic_col:
                        total_solicitudes = safe_int(row.get(total_solic_col, 0))
                    elif total_col:
                        total_solicitudes = safe_int(row.get(total_col, 0))
                continue

            vals = []
            for letra in LETRAS:
                v = 0
                for c in cols:
                    if c.strip().upper() == letra.upper():
                        v = safe_int(row.get(c, 0))
                        break
                vals.append(v)

            if total_col:
                tot = safe_int(row.get(total_col, 0))
            else:
                tot = sum(vals)

            valoradas_total += tot
            partes = nombre.split()
            nombre_corto = partes[0] if partes else nombre
            color = COLORES_MIEMBRO[len(miembros) % len(COLORES_MIEMBRO)]
            miembros.append({
                'nombre': nombre,
                'nombreCorto': nombre_corto,
                'vals': vals,
                'total': tot,
                'color': color,
            })

        # Si no encontramos total de solicitudes en fila de totales,
        # intentar sumar una columna especifica
        if total_solicitudes == 0 and total_solic_col:
            total_solicitudes = int(df[total_solic_col].dropna().sum())
        if total_solicitudes == 0:
            total_solicitudes = valoradas_total

        pct = round(valoradas_total / total_solicitudes * 100, 2) if total_solicitudes > 0 else 0.0
        print(f"  Total solicitudes: {total_solicitudes} | Valoradas: {valoradas_total} | {pct}%")
        print(f"  Miembros procesados: {len(miembros)}")

        letras_js = json.dumps(LETRAS, ensure_ascii=False)
        miembros_lines = []
        for m in miembros:
            vals_str = json.dumps(m['vals'])
            ne = m['nombre'].replace("'", "\\'")
            ce2 = m['nombreCorto'].replace("'", "\\'")
            miembros_lines.append(
                f"    {{nombre:'{ne}', nombreCorto:'{ce2}', vals:{vals_str}, total:{m['total']}, color:'{m['color']}'}}"
            )
        miembros_js = ',\n'.join(miembros_lines)

        return f"""// ======== AYUDAS DE ACCION SOCIAL ========
const AY = {{
  total: {total_solicitudes},
  valoradas: {valoradas_total},
  pct: {pct},
  letras: {letras_js},
  miembros: [
{miembros_js}
  ]
}};"""

    except Exception as e:
        print(f"ERROR leyendo {ruta}: {e}")
        import traceback; traceback.print_exc()
        return """// ======== AYUDAS DE ACCION SOCIAL (SIN DATOS) ========
const AY = {
  total: 0, valoradas: 0, pct: 0,
  letras: ['A','B','C','D','E-F','G','H','I-J-K','L','M','N-\u00d1-O','P-Q','R','S-T','U-Z','\u00c1\u00d1ADIDOS'],
  miembros: []
};"""

def generar_html(registros, bloque_ay, template_path, output_path):
    """Inyecta ambos bloques de datos en el template y genera index.html."""
    print(f"Leyendo template {template_path}...")
    with open(template_path, 'r', encoding='utf-8') as f:
        template = f.read()

    if '// __DATOS_SUPLES__' not in template:
        print("ERROR: El template no contiene el marcador // __DATOS_SUPLES__")
        sys.exit(1)

    json_data = json.dumps(registros, ensure_ascii=False, separators=(',', ':'))
    n = len(registros)
    fecha = datetime.now().strftime('%d/%m/%Y %H:%M')

    bloque_suples = f"""// ======= DATOS POR DEFECTO =======
// Generado automaticamente el {fecha} - {n} convocatorias
(function initDefault(){{
  const rows = {json_data};
  DATA = processRows(rows, 'CONTROL_SUPLES.ods');
  renderAll();
  setTimeout(()=>{{
    if(!sessionStorage.getItem('welcomed')){{
      sessionStorage.setItem('welcomed','1');
      toast('\u2705 {n} convocatorias cargadas \u00b7 Actualizado {fecha}', 3500);
    }}
  }}, 800);
}})();

"""

    html = template.replace('// __DATOS_SUPLES__\n\n', bloque_suples)

    if '// __DATOS_VALORACION__' in html:
        html = html.replace('// __DATOS_VALORACION__', bloque_ay)
        print("  Seccion Ayudas de Accion Social actualizada.")
    else:
        print("  AVISO: marcador __DATOS_VALORACION__ no encontrado en template.")

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html)

    size_kb = len(html) / 1024
    print(f"Generado {output_path} - {size_kb:.1f} KB - {n} convocatorias")

def main():
    if not os.path.exists(ODS_FILE):
        print(f"ERROR: No se encuentra {ODS_FILE}")
        print(f"  Archivos en directorio: {os.listdir('.')}")
        sys.exit(1)
    if not os.path.exists(TEMPLATE_FILE):
        print(f"ERROR: No se encuentra {TEMPLATE_FILE}")
        sys.exit(1)

    print("Iniciando generacion del Dashboard SUPLES CHUC")
    print(f"  ODS principal  : {ODS_FILE}")
    print(f"  ODS valoracion : {ODS_VALORACION}")
    print(f"  Template       : {TEMPLATE_FILE}")
    print(f"  Output         : {OUTPUT_FILE}")
    print()

    registros = procesar_ods(ODS_FILE)

    if os.path.exists(ODS_VALORACION):
        bloque_ay = procesar_valoracion(ODS_VALORACION)
    else:
        print(f"AVISO: {ODS_VALORACION} no encontrado - Ayudas sin actualizar")
        bloque_ay = "const AY = {total:0,valoradas:0,pct:0,letras:[],miembros:[]};"

    generar_html(registros, bloque_ay, TEMPLATE_FILE, OUTPUT_FILE)

    print()
    print("Listo! index.html actualizado con datos de ambos ODS.")

if __name__ == '__main__':
    main()
