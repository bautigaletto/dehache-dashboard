"""
DEHACHE Dashboard 2026 – Actualizador de datos
===============================================
Uso:
    python update.py

Lee 'POWER BI.xlsx' en la misma carpeta y genera index.html actualizado.
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
import calendar as cal
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
# MONTH NAMES
# ──────────────────────────────────────────────────────────
ALL_MONTHS_SHORT = ['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic']
ALL_MONTHS_FULL  = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto',
                    'Septiembre','Octubre','Noviembre','Diciembre']

# ──────────────────────────────────────────────────────────
# ARTICLE ALIASES  (same product, different names over time)
# Add entries here whenever a product code gets renamed.
# Format: 'CODE': 'CANONICAL NAME'
# ──────────────────────────────────────────────────────────
ART_ALIASES = {
    '1501': '1501 - MEMBRANA C/ALU MGX - 40 KGS. NO CRACK PLUS',
}

def normalize_art(name):
    if not isinstance(name, str):
        return name
    code = name.split(' - ')[0].strip()
    return ART_ALIASES.get(code, name)

# ──────────────────────────────────────────────────────────
# CONFIG
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
    'Griferias':               'Griferías',
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
    'Griferias':              2_200_000,
    # Oportunidades: target=0 → shown without break-even line
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

    # Normalize renamed articles
    df['Artículo'] = df['Artículo'].apply(normalize_art)

    # Global exclusions
    mask = (
        (df['Niveles'] != 'Descuentos') &
        (df['Niveles'] != 'Varios') &
        (df['Vendedor'] != 'SUCURSAL CENTRO') &
        (df['Centro_str'] != '0005')
    )
    fAll = df[mask].dropna(subset=['Niveles']).copy()

    ref        = fAll['Fecha'].max()
    last_all   = fAll.groupby('Artículo')['Fecha'].max()
    art_level  = fAll.sort_values('Fecha').groupby('Artículo')['Niveles'].last()

    fAll['m'] = fAll['Fecha'].dt.month
    fAll['y'] = fAll['Fecha'].dt.year

    # 2026-only slice
    f = fAll[fAll['y'] == 2026].copy()
    f['m'] = f['Fecha'].dt.month
    f['Nombre Zona'] = f['Nombre Zona'].fillna('Sin zona')

    lastDayNum        = ref.day
    daysInCurrentMonth = cal.monthrange(ref.year, ref.month)[1]
    currentMonth      = ref.month

    print(f'  Filas 2026 filtradas: {len(f):,}  |  Fecha máx: {ref.strftime("%d/%m/%Y")}')

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

    f['vi']   = f['Vendedor'].map(vi)
    f['zi']   = f['Nombre Zona'].map(zi)
    f['ni']   = f['Niveles'].map(ni)
    f['ci']   = f['Cód. Cliente'].map(ci)
    f['ai']   = f['Artículo'].map(ai)
    f['ii']   = f['invkey'].map(ii)
    f['impr'] = f['Importe'].round().astype(int)
    f['qr']   = f['Cantidad'].round(2)

    # Margen bruto & Linea (available from Aug 2026 onwards)
    if 'Margen bruto' in f.columns:
        f['mb_r'] = f['Margen bruto'].where(f['Margen bruto'].notna(), other=None)
        mb_list = [round(v, 2) if v is not None else None for v in f['mb_r'].tolist()]
    else:
        mb_list = [None] * len(f)

    if 'Linea' in f.columns:
        ln_list = [int(v) if pd.notna(v) else 0 for v in f['Linea'].tolist()]
    else:
        ln_list = [0] * len(f)

    rows = {
        'm':   f['m'].tolist(),   'v': f['vi'].tolist(),
        'z':   f['zi'].tolist(),  'n': f['ni'].tolist(),
        'c':   f['ci'].tolist(),  'a': f['ai'].tolist(),
        'inv': f['ii'].tolist(), 'imp': f['impr'].tolist(), 'q': f['qr'].tolist(),
        'mb':  mb_list,           'ln': ln_list,
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

    stale_total = sum(1 for m in months_since_arr if m > 12)
    stale_oport = sum(1 for a,m in zip(arts,months_since_arr)
                      if m>12 and art_level.get(a) in ('Oportunidades',))
    print(f'  Artículos +1 año sin movimiento: {stale_total} (excluye {stale_oport} de Oportunidades)')

    # ── 2025 & 2026 level monthly sales ──
    # 2025 is optional — returns zeros if the Excel only contains 2026 data
    f25 = fAll[fAll['y'] == 2025].copy()
    has2025 = len(f25) > 0
    if has2025:
        f25['m'] = f25['Fecha'].dt.month

    lvlMonthly2025 = []
    for l in levels:
        row = [0]*12
        if has2025:
            for m, g in f25[f25['Niveles']==l].groupby('m'):
                row[m-1] = int(g['Importe'].sum().round())
        lvlMonthly2025.append(row)

    lvlMonthly2026 = []
    for l in levels:
        row = [0]*currentMonth
        for m, g in f[f['Niveles']==l].groupby('m'):
            if m <= currentMonth:
                row[m-1] = int(g['Importe'].sum().round())
        lvlMonthly2026.append(row)

    totalMonthly2025 = [0]*12
    if has2025:
        for m, g in f25.groupby('m'):
            totalMonthly2025[m-1] = int(g['Importe'].sum().round())

    # ── Article monthly sales+qty 2026 ──
    art_grp = f.groupby('ai')
    artMonthly2026imp = []
    artMonthly2026qty = []
    for k in range(len(arts)):
        imps = [0]*currentMonth; qtys = [0]*currentMonth
        if k in art_grp.groups:
            g2 = art_grp.get_group(k)
            for m, mg in g2.groupby('m'):
                if m <= currentMonth:
                    imps[m-1] = int(mg['impr'].sum())
                    qtys[m-1] = round(float(mg['qr'].sum()), 2)
        artMonthly2026imp.append(imps)
        artMonthly2026qty.append(qtys)

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
        'months':     ALL_MONTHS_SHORT[:currentMonth],
        'monthsFull': ALL_MONTHS_FULL[:currentMonth],
        'months2025': ALL_MONTHS_SHORT,
        'lvlMonthly2025':    lvlMonthly2025,
        'lvlMonthly2026':    lvlMonthly2026,
        'artMonthly2026imp': artMonthly2026imp,
        'artMonthly2026qty': artMonthly2026qty,
        'totalMonthly2025':  totalMonthly2025,
        'lastDayNum':        lastDayNum,
        'daysInCurrentMonth': daysInCurrentMonth,
        'currentMonth':      currentMonth,
        'lastDay':           ref.strftime('%d/%m/%Y'),
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

    data     = process(EXCEL_FILE)
    last_day = data['meta']['lastDay']

    logo_b64 = ''
    if os.path.exists(LOGO_FILE):
        print('  Procesando logo…')
        logo_b64 = build_logo_b64(LOGO_FILE)
    else:
        print('  AVISO: LogoDHVertical.jpg no encontrado, logo omitido.')

    print('  Generando index.html…')
    with open(TEMPLATE, encoding='utf-8') as fh:
        html = fh.read()

    html = html.replace('__DATA_JSON__', json.dumps(data, ensure_ascii=False, separators=(',',':')))
    html = html.replace('__LOGO_B64__',  logo_b64)
    html = html.replace('__LASTDAY__',   last_day)

    with open(OUTPUT, 'w', encoding='utf-8') as fh:
        fh.write(html)

    size_kb = round(os.path.getsize(OUTPUT) / 1024, 1)
    print(f'\n  ✓ index.html generado ({size_kb} KB)')
    print(f'  ✓ Datos al {last_day}')
    print(f'  ✓ {len(data["meta"]["clients"])} clientes | {len(data["meta"]["articles"])} artículos')
    print(f'  ✓ {len(data["meta"]["levels"])} niveles: {", ".join(data["meta"]["levels"])}')
    print('=' * 55)

if __name__ == '__main__':
    main()
