"""
DEHACHE Dashboard 2026 – Actualizador de datos
===============================================
Uso:
    python update.py

Lee POWER_BI.xlsx en la misma carpeta y genera index.html actualizado.
Requiere: pandas, openpyxl, Pillow
    pip install pandas openpyxl Pillow
"""

import pandas as pd
import json
import re
import os
import sys
import base64
import io
from datetime import datetime

# ──────────────────────────────────────────────────────────
# PATHS
# ──────────────────────────────────────────────────────────
BASE_DIR   = os.path.dirname(os.path.abspath(__file__))
EXCEL_FILE = os.path.join(BASE_DIR, 'POWER BI.xlsx')
TEMPLATE   = os.path.join(BASE_DIR, 'template.html')
LOGO_FILE  = os.path.join(BASE_DIR, 'LogoDHVertical.jpg')
OUTPUT     = os.path.join(BASE_DIR, 'index.html')

# ──────────────────────────────────────────────────────────
# CONFIG – objectives and display names
# ──────────────────────────────────────────────────────────
VEND_DISP = {
    'FANSINI GUSTAVO ANGEL':    'Gustavo Fansini',
    'SOMARE CRISTIAN LEONARDO': 'Cristian Somare',
    'TANUS JOSE ALBERTO':       'José Tanus',
    'DANIEL HERNANDEZ':         'Daniel Hernández',
    'GALETTO BAUTISTA':         'Bautista Galetto',
    'MOSTRADOR':                'Mostrador',
}
LVL_DISP = {
    'Impermeabilización':      'Impermeabilización',
    'Materiales p/ Construc.': 'Materiales construcción',
    'Aislante':                'Aislantes',
    'Const. En Seco':          'Construcción en seco',
    'Elementos de seguridad':  'Elementos de seguridad',
    'Herramientas':            'Herramientas',
    'Instalación Agua y Cloaca': 'Instalación agua y cloaca',
    'Instalación Gas':         'Instalación de gas',
    'Jardineria y Riego':      'Jardinería y riego',
    'Pintuteria':              'Pinturería',
    'Quimicos':                'Químicos',
    'Sanitarios':              'Sanitarios',
    'Zingueria':               'Zinguería',
    'Servicio de Fletes':      'Fletes',
    'Seccoplac':               'Seccoplac',
    'Oportunidades':           'Oportunidades',
    'Colocación Membrana':     'Colocación membrana',
    'Pisos':                   'Pisos',
}
LEVEL_TARGETS = {
    'Impermeabilización':      110_000_000,
    'Materiales p/ Construc.':  69_000_000,
    'Aislante':                  7_000_000,
    'Const. En Seco':           19_200_000,
    'Elementos de seguridad':      175_000,
    'Herramientas':              7_800_000,
    'Instalación Agua y Cloaca': 22_000_000,
    'Instalación Gas':            2_350_000,
    'Jardineria y Riego':          250_000,
    'Pintuteria':                 6_600_000,
    'Quimicos':                   4_750_000,
    'Sanitarios':                23_900_000,
    'Zingueria':                  4_500_000,
    'Servicio de Fletes':        14_000_000,
}
TOTAL_TARGET = 300_000_000

# ──────────────────────────────────────────────────────────
# LOGO
# ──────────────────────────────────────────────────────────
def build_logo_b64(path):
    from PIL import Image
    import numpy as np
    im = Image.open(path).convert('RGB')
    arr = np.array(im)
    nw = ~((arr[:,:,0]>245)&(arr[:,:,1]>245)&(arr[:,:,2]>245))
    ys, xs = np.where(nw)
    pad = 20
    box = (max(xs.min()-pad,0), max(ys.min()-pad,0),
           min(xs.max()+pad,im.width), min(ys.max()+pad,im.height))
    im2 = im.crop(box)
    h = 120; w = int(im2.width*h/im2.height)
    im2 = im2.resize((w,h), Image.LANCZOS)
    buf = io.BytesIO()
    im2.save(buf, format='PNG', optimize=True)
    return base64.b64encode(buf.getvalue()).decode()

# ──────────────────────────────────────────────────────────
# DATA PROCESSING
# ──────────────────────────────────────────────────────────
def process(excel_path):
    print(f'  Leyendo {excel_path}…')
    df = pd.read_excel(excel_path, sheet_name='Ventas total')
    df['Fecha'] = pd.to_datetime(df['Fecha'])
    df['Centro_str'] = df['Centro'].astype(str).str.zfill(4)

    # Global exclusions (applied to full history too)
    mask = (
        (df['Niveles'] != 'Descuentos') &
        (df['Niveles'] != 'Varios') &
        (df['Vendedor'] != 'SUCURSAL CENTRO') &
        (df['Centro_str'] != '0005')
    )
    fAll = df[mask].dropna(subset=['Niveles']).copy()

    ref = fAll['Fecha'].max()
    last_all   = fAll.groupby('Artículo')['Fecha'].max()
    art_level  = fAll.sort_values('Fecha').groupby('Artículo')['Niveles'].last()

    # 2026-only slice for dashboard rows
    f = fAll[fAll['Fecha'].dt.year == 2026].copy()
    f['m'] = f['Fecha'].dt.month
    f['Nombre Zona'] = f['Nombre Zona'].fillna('Sin zona')

    print(f'  Filas 2026 filtradas: {len(f):,}  |  Fecha máx: {ref.strftime("%d/%m/%Y")}')
    print(f'  Total facturado 2026: ${f["Importe"].sum():,.0f}')

    vendors = sorted(f['Vendedor'].unique())
    zones   = sorted(f['Nombre Zona'].unique())
    levels  = sorted(f['Niveles'].unique())
    cl      = f[['Cód. Cliente','Nombre Cliente']].drop_duplicates('Cód. Cliente')
    clients = list(cl.itertuples(index=False, name=None))
    codes   = [c[0] for c in clients]
    arts    = sorted(fAll['Artículo'].unique())

    vi = {v:i for i,v in enumerate(vendors)}
    zi = {z:i for i,z in enumerate(zones)}
    ni = {n:i for i,n in enumerate(levels)}
    ci = {c:i for i,c in enumerate(codes)}
    ai = {a:i for i,a in enumerate(arts)}

    f['invkey'] = (f['Centro_str'] + '|' + f['Tipo de Comprobante'].astype(str)
                   + '|' + f['Letra'].astype(str) + '|' + f['Número'].astype(str))
    invkeys = sorted(f['invkey'].unique())
    ii = {k:i for i,k in enumerate(invkeys)}

    f['vi']  = f['Vendedor'].map(vi)
    f['zi']  = f['Nombre Zona'].map(zi)
    f['ni']  = f['Niveles'].map(ni)
    f['ci']  = f['Cód. Cliente'].map(ci)
    f['ai']  = f['Artículo'].map(ai)
    f['ii']  = f['invkey'].map(ii)
    f['impr']= f['Importe'].round().astype(int)
    f['qr']  = f['Cantidad'].round(2)

    rows = {
        'm':   f['m'].tolist(),
        'v':   f['vi'].tolist(),
        'z':   f['zi'].tolist(),
        'n':   f['ni'].tolist(),
        'c':   f['ci'].tolist(),
        'a':   f['ai'].tolist(),
        'inv': f['ii'].tolist(),
        'imp': f['impr'].tolist(),
        'q':   f['qr'].tolist(),
    }

    def zone_parts(z):
        m = re.match(r'^\((\d+)\)\s*(.*)$', z)
        if m: return m.group(1), m.group(2).strip().title()
        return '', z.title()
    zmeta = [zone_parts(z) for z in zones]

    def months_since(dt):
        return ((ref.year - dt.year)*12 + (ref.month - dt.month)
                - (1 if ref.day < dt.day else 0))

    months_since_arr = [int(months_since(last_all[a])) for a in arts]
    art_lvl_idx      = [ni.get(art_level.get(a), -1) for a in arts]

    stale_total   = sum(1 for m in months_since_arr if m > 12)
    stale_oport   = sum(1 for a,m in zip(arts,months_since_arr)
                        if m>12 and art_level.get(a)=='Oportunidades')
    print(f'  Artículos +1 año sin movimiento: {stale_total} (excluye {stale_oport} de Oportunidades)')

    meta = {
        'vendors':   [VEND_DISP.get(v, v.title()) for v in vendors],
        'zones':     [zm[1] if zm[1] else z for zm,z in zip(zmeta,zones)],
        'zoneCodes': [zm[0] for zm in zmeta],
        'levels':    [LVL_DISP.get(l,l) for l in levels],
        'levelTargets': [int(LEVEL_TARGETS.get(l,0)) for l in levels],
        'clients':   [c[1].title() if isinstance(c[1],str) else str(c[1]) for c in clients],
        'clientCodes': [int(x) if not pd.isna(x) else 0 for x in codes],
        'articles':  arts,
        'articleMonthsSinceLastSale': months_since_arr,
        'articleLevelIdx':            art_lvl_idx,
        'totalTarget': TOTAL_TARGET,
        'months':     ['Ene','Feb','Mar','Abr','May','Jun'],
        'monthsFull': ['Enero','Febrero','Marzo','Abril','Mayo','Junio'],
        'lastDay':    ref.strftime('%d/%m/%Y'),
    }
    return {'meta': meta, 'rows': rows}

# ──────────────────────────────────────────────────────────
# MAIN
# ──────────────────────────────────────────────────────────
def main():
    print('=' * 55)
    print('  DEHACHE Dashboard 2026 — Actualizador')
    print('=' * 55)

    if not os.path.exists(EXCEL_FILE):
        print(f'\nERROR: No se encontró {EXCEL_FILE}')
        sys.exit(1)
    if not os.path.exists(TEMPLATE):
        print(f'\nERROR: No se encontró {TEMPLATE}')
        sys.exit(1)

    # 1. Process data
    data = process(EXCEL_FILE)
    last_day = data['meta']['lastDay']

    # 2. Logo
    logo_b64 = ''
    if os.path.exists(LOGO_FILE):
        print('  Procesando logo…')
        logo_b64 = build_logo_b64(LOGO_FILE)
    else:
        print('  AVISO: LogoDHVertical.jpg no encontrado, logo omitido.')

    # 3. Inject into template
    print('  Generando index.html…')
    with open(TEMPLATE, encoding='utf-8') as fh:
        html = fh.read()

    html = html.replace('__DATA_JSON__',  json.dumps(data, ensure_ascii=False, separators=(',',':')))
    html = html.replace('__LOGO_B64__',   logo_b64)
    html = html.replace('__LASTDAY__',    last_day)

    with open(OUTPUT, 'w', encoding='utf-8') as fh:
        fh.write(html)

    size_kb = round(os.path.getsize(OUTPUT) / 1024, 1)
    print(f'\n  ✓ index.html generado ({size_kb} KB)')
    print(f'  ✓ Datos al {last_day}')
    print(f'  ✓ {len(data["meta"]["clients"])} clientes | {len(data["meta"]["articles"])} artículos')
    print('=' * 55)

if __name__ == '__main__':
    main()
