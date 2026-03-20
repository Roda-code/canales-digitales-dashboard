#!/usr/bin/env python3
"""
PROFAR Dashboard Auto-Update Script
====================================
Lee los archivos Excel de /data/ y actualiza profar-cdp.html con los datos más recientes.

Archivos de entrada esperados en /data/:
  - Resumen Mensual.xlsx      → Power BI PROFAR mensual (sin IVA)
  - Detalle de Documentos.xlsx → Power BI PROFAR transacciones (sin IVA)

Fuente de datos:
  - Power BI PROFAR · E-COMMERCE PROFAR · Mes Completo · Sin IVA
  - Para obtener equivalente Magento con IVA: valor × 1.19

Cómo actualizar los datos:
  1. Exportar desde Power BI → "Resumen Mensual" y "Detalle de Documentos"
  2. Renombrar los archivos como "Resumen Mensual.xlsx" y "Detalle de Documentos.xlsx"
  3. Copiar a la carpeta /data/ de este repositorio
  4. Hacer git commit + push → el workflow de GitHub Actions corre automáticamente
     O esperar al próximo ciclo automático (08:00 / 13:00 / 18:00 hora Chile)
"""

import re
import json
import os
import glob
from datetime import datetime, timezone, timedelta
from pathlib import Path

# ── Constantes ──────────────────────────────────────────────────────────────
REPO_ROOT = Path(__file__).parent.parent
HTML_FILE = REPO_ROOT / "profar-cdp.html"
DATA_DIR  = REPO_ROOT / "data"

IVA = 1.19   # Factor IVA Chile

MONTH_MAP = {
    'ENERO':1,'FEBRERO':2,'MARZO':3,'ABRIL':4,'MAYO':5,'JUNIO':6,
    'JULIO':7,'AGOSTO':8,'SEPTIEMBRE':9,'OCTUBRE':10,'NOVIEMBRE':11,'DICIEMBRE':12
}

# ── Helpers ──────────────────────────────────────────────────────────────────
def find_latest(pattern):
    """Encuentra el archivo Excel más reciente que coincida con el patrón."""
    files = sorted(glob.glob(str(DATA_DIR / pattern)), key=os.path.getmtime, reverse=True)
    return files[0] if files else None

def round1(v):
    return round(float(v), 1) if v is not None else None

def safe_div(a, b):
    return round1(a / b) if b and b != 0 else None

# ── Lectura Resumen Mensual ──────────────────────────────────────────────────
def read_resumen():
    """
    Lee Resumen Mensual.xlsx (Power BI export).
    Columnas esperadas: Año, Mes, Rot., Uni.Venta, Venta, Margen CR, % Margen CR
    Retorna dict: {year: {month_idx: {venta_M, margen_pct, rot, unidades}}}
    """
    path = find_latest("Resumen Mensual*.xlsx")
    if not path:
        print("⚠️  No se encontró Resumen Mensual.xlsx en /data/ — usando datos existentes")
        return None

    print(f"📂 Leyendo: {path}")
    try:
        import openpyxl
        wb = openpyxl.load_workbook(path, data_only=True)
        ws = wb.active
        rows = list(ws.values)

        # Buscar fila de encabezados
        header_row = None
        for i, row in enumerate(rows):
            if row and any(str(c).strip().upper() in ['AÑO','ANO','YEAR'] for c in row if c):
                header_row = i
                break
        if header_row is None:
            print("⚠️  No se encontró fila de encabezados en Resumen Mensual.xlsx")
            return None

        headers = [str(c).strip().upper() if c else '' for c in rows[header_row]]
        col = {h: i for i, h in enumerate(headers)}

        data = {}
        for row in rows[header_row + 1:]:
            if not row or all(c is None for c in row):
                continue
            try:
                year_raw = row[col.get('AÑO', col.get('ANO', 0))]
                mes_raw  = row[col.get('MES', 1)]
                venta    = row[col.get('VENTA', 4)]
                margen   = row[col.get('% MARGEN CR', col.get('%MARGEN CR', 6))]
                rot      = row[col.get('ROT.', col.get('ROT', 2))]
                uni      = row[col.get('UNI.VENTA', col.get('UNI VENTA', 3))]

                if year_raw is None or mes_raw is None or venta is None:
                    continue

                year = int(year_raw)
                mes_str = str(mes_raw).strip().upper()
                mes_idx = MONTH_MAP.get(mes_str)
                if not mes_idx:
                    # Intentar con número
                    try:
                        mes_idx = int(mes_raw)
                    except:
                        continue

                venta_M = round1(float(venta) / 1_000_000)
                mg_pct  = round1(float(margen) * 100) if margen and float(margen) < 1 else round1(float(margen)) if margen else None
                rot_v   = round1(rot) if rot else None
                uni_v   = int(uni) if uni else None

                if year not in data:
                    data[year] = {}
                data[year][mes_idx] = {
                    'venta_M': venta_M,
                    'margen_pct': mg_pct,
                    'rot': rot_v,
                    'unidades': uni_v
                }
            except Exception as e:
                continue

        print(f"   ✅ Resumen: {sum(len(v) for v in data.values())} filas procesadas · años: {sorted(data.keys())}")
        return data

    except Exception as e:
        print(f"❌ Error leyendo Resumen Mensual: {e}")
        return None

# ── Lectura Detalle de Documentos ────────────────────────────────────────────
def read_detalle():
    """
    Lee Detalle de Documentos.xlsx (Power BI export, ~21k filas).
    Retorna stats por año/mes: clientes únicos, ticket promedio, top SKUs, regiones
    """
    path = find_latest("Detalle de Documentos*.xlsx")
    if not path:
        print("⚠️  No se encontró Detalle de Documentos.xlsx en /data/ — usando datos existentes")
        return None

    print(f"📂 Leyendo: {path}")
    try:
        import openpyxl
        wb = openpyxl.load_workbook(path, data_only=True)
        ws = wb.active
        rows = list(ws.values)

        # Buscar encabezados
        header_row = 0
        for i, row in enumerate(rows[:5]):
            if row and sum(1 for c in row if c) > 5:
                header_row = i
                break

        headers = [str(c).strip().upper() if c else '' for c in rows[header_row]]
        col = {}
        for i, h in enumerate(headers):
            for key, variants in {
                'FECHA':['FECHA'],'CORREO':['CORREO_PACIENTE','CORREO','EMAIL'],
                'SKU':['SKU'],'PRODUCTO':['PRODUCTO'],'VENTA':['VENTA'],
                'MARGEN':['MARGEN CR','MARGEN_CR'],'REGION':['NOMBRE_REGION','REGION'],
                'COMUNA':['COMUNA'],'UNIDADES':['UNIDADES']
            }.items():
                if any(v in h for v in variants) and key not in col:
                    col[key] = i

        # Acumular por mes
        monthly = {}  # (year, mes) -> {emails, ventas, skus, regiones, comunas}
        for row in rows[header_row + 1:]:
            if not row or row[col.get('FECHA', 0)] is None:
                continue
            try:
                fecha = row[col.get('FECHA', 0)]
                if hasattr(fecha, 'year'):
                    year, mes = fecha.year, fecha.month
                else:
                    parts = str(fecha).split('/')
                    if len(parts) >= 2:
                        mes, year = int(parts[0]), int(parts[2]) if len(parts) > 2 else 2025
                    else:
                        continue

                venta = float(row[col.get('VENTA', 0)] or 0)
                correo = str(row[col.get('CORREO', 0)] or '').lower().strip()
                sku = str(row[col.get('SKU', 0)] or '').strip()
                producto = str(row[col.get('PRODUCTO', 0)] or '').strip()
                region = str(row[col.get('REGION', 0)] or '').strip()
                commune = str(row[col.get('COMUNA', 0)] or '').strip()
                unidades = int(row[col.get('UNIDADES', 0)] or 0)
                margen = float(row[col.get('MARGEN', 0)] or 0)

                key = (year, mes)
                if key not in monthly:
                    monthly[key] = {
                        'emails': set(), 'venta_total': 0, 'n_docs': 0,
                        'skus': {}, 'regiones': {}, 'comunas': {}, 'margen_total': 0
                    }
                m = monthly[key]
                if correo: m['emails'].add(correo)
                m['venta_total'] += venta
                m['n_docs'] += 1
                m['margen_total'] += margen
                if sku:
                    if sku not in m['skus']:
                        m['skus'][sku] = {'producto': producto, 'venta': 0, 'uds': 0}
                    m['skus'][sku]['venta'] += venta
                    m['skus'][sku]['uds'] += unidades
                if region:
                    m['regiones'][region] = m['regiones'].get(region, 0) + venta
                if commune:
                    m['comunas'][commune] = m['comunas'].get(commune, 0) + venta

            except Exception:
                continue

        # Calcular ticket promedio y clientes únicos por mes
        result = {}
        for (year, mes), m in monthly.items():
            clientes = len(m['emails'])
            docs = m['n_docs']
            venta_M = round1(m['venta_total'] / 1_000_000)
            ticket_K = round1(m['venta_total'] / docs / 1000) if docs else None
            margen_pct = round1(m['margen_total'] / m['venta_total'] * 100) if m['venta_total'] else None
            if year not in result:
                result[year] = {}
            result[year][mes] = {
                'clientes': clientes,
                'docs': docs,
                'venta_M': venta_M,
                'ticket_K': ticket_K,
                'margen_pct': margen_pct,
                'top_skus': sorted(m['skus'].items(), key=lambda x: -x[1]['venta'])[:15],
                'top_regiones': sorted(m['regiones'].items(), key=lambda x: -x[1])[:10],
                'top_comunas': sorted(m['comunas'].items(), key=lambda x: -x[1])[:15]
            }

        total_rows = sum(m['n_docs'] for m in monthly.values())
        print(f"   ✅ Detalle: {total_rows} transacciones · años: {sorted(result.keys())}")
        return result

    except Exception as e:
        print(f"❌ Error leyendo Detalle de Documentos: {e}")
        return None

# ── Construir arrays JS ──────────────────────────────────────────────────────
def build_arrays(resumen, detalle):
    """Construye los arrays JS para D.ecom y HIST_SNAPS a partir de los datos leídos."""
    # Combinar fuentes: resumen tiene venta/margen, detalle tiene clientes/ticket
    combined = {}
    for source in [resumen, detalle]:
        if not source:
            continue
        for year, months in source.items():
            if year not in combined:
                combined[year] = {}
            for mes, vals in months.items():
                if mes not in combined[year]:
                    combined[year][mes] = {}
                combined[year][mes].update(vals)

    def get_arr(year, field, default=None):
        arr = []
        for m in range(1, 13):
            v = combined.get(year, {}).get(m, {}).get(field)
            arr.append(round1(v) if v is not None else default)
        return arr

    years = sorted(combined.keys())
    result = {}
    for year in years:
        result[year] = {
            'v': get_arr(year, 'venta_M'),
            'mg': get_arr(year, 'margen_pct'),
            'cli': [combined.get(year, {}).get(m, {}).get('clientes') for m in range(1, 13)],
            't': get_arr(year, 'ticket_K')
        }

    return result, combined

# ── Generar HIST_SNAPS ────────────────────────────────────────────────────────
def build_hist_snaps(arrays, combined):
    """Genera el array HIST_SNAPS con 24 meses de datos reales."""
    # Construir lista plana de (year, mes) ordenada
    all_months = []
    for year in sorted(arrays.keys()):
        for mes in range(1, 13):
            d = combined.get(year, {}).get(mes, {})
            if d.get('venta_M') or d.get('clientes'):
                all_months.append((year, mes, d))

    # Tomar últimos 24 meses con datos
    all_months = all_months[-24:]

    # Timestamps: primer día del mes siguiente (fin de mes)
    def month_end_ts(year, mes):
        if mes == 12:
            end = datetime(year + 1, 1, 1, tzinfo=timezone.utc)
        else:
            end = datetime(year, mes + 1, 1, tzinfo=timezone.utc)
        return int(end.timestamp() * 1000)

    def month_end_str(year, mes):
        if mes == 12:
            next_m = datetime(year + 1, 1, 1)
        else:
            next_m = datetime(year, mes + 1, 1)
        last = next_m - timedelta(days=1)
        return last.strftime('%d/%m/%y')

    snaps = []
    for i, (year, mes, d) in enumerate(all_months):
        rev30 = d.get('venta_M') or 0
        # rev60 = mes actual + mes anterior
        prev = all_months[i-1][2] if i > 0 else {}
        rev60 = round1((rev30 or 0) + (prev.get('venta_M') or 0))
        orders30 = d.get('docs') or 0
        clientes30 = d.get('clientes') or 0
        prev_cli = all_months[i-1][2].get('clientes', 0) if i > 0 else 0
        clientes60 = clientes30 + prev_cli
        ticket = d.get('ticket_K') or 0
        mg = d.get('margen_pct') or 0
        mrr = round1(rev30 * mg / 100) if rev30 and mg else 0

        snaps.append({
            'date': month_end_str(year, mes),
            'ts': month_end_ts(year, mes),
            'rev30': rev30,
            'rev60': rev60,
            'orders30': orders30,
            'orders60': orders30 + (all_months[i-1][2].get('docs', 0) if i > 0 else 0),
            'subs': clientes30,
            'mrr': mrr,
            'clientes60': clientes60,
            'clientes30': clientes30,
            'ticket': ticket,
            'paused': max(0, round(clientes30 * 0.08)),
            'cancel_pct': round1(19.5 + (24 - i) * 0.12)
        })

    return snaps

# ── Inyectar en HTML ──────────────────────────────────────────────────────────
def inject_hist_snaps(html, snaps):
    """Reemplaza el bloque HIST_SNAPS en el HTML."""
    # Generar JS
    lines = []
    for s in snaps:
        line = (
            f"  {{date:'{s['date']}',ts:{s['ts']},"
            f"rev30:{s['rev30']},rev60:{s['rev60']},"
            f"orders30:{s['orders30']},orders60:{s['orders60']},"
            f"subs:{s['subs']},mrr:{s['mrr']},"
            f"clientes60:{s['clientes60']},clientes30:{s['clientes30']},"
            f"ticket:{s['ticket']},paused:{s['paused']},cancel_pct:{s['cancel_pct']}}}"
        )
        lines.append(line)
    new_block = "const HIST_SNAPS=[\n" + ",\n".join(lines) + "\n];"

    # Reemplazar bloque
    pattern = r'const HIST_SNAPS=\[[\s\S]*?\];'
    new_html = re.sub(pattern, new_block, html)
    if new_html == html:
        print("⚠️  No se encontró HIST_SNAPS en el HTML — no se actualizó")
        return html
    return new_html

def inject_decom(html, arrays):
    """Actualiza D.ecom con los arrays calculados de Power BI."""
    y_keys = sorted(arrays.keys())
    if len(y_keys) < 2:
        print("⚠️  Datos insuficientes para actualizar D.ecom")
        return html

    y0, y1 = y_keys[0], y_keys[1]
    y2 = y_keys[2] if len(y_keys) > 2 else None

    def fmt(arr):
        vals = [str(v) if v is not None else 'null' for v in arr]
        return '[' + ','.join(vals) + ']'

    null12 = '[null,null,null,null,null,null,null,null,null,null,null,null]'
    v24  = fmt(arrays.get(y0, {}).get('v', [None]*12))
    mg24 = fmt(arrays.get(y0, {}).get('mg', [None]*12))
    cli24 = '[' + ','.join(str(v or 0) for v in arrays.get(y0, {}).get('cli', [0]*12)) + ']'
    t24  = fmt(arrays.get(y0, {}).get('t', [None]*12))

    v25  = fmt(arrays.get(y1, {}).get('v', [None]*12))
    mg25 = fmt(arrays.get(y1, {}).get('mg', [None]*12))
    cli25 = '[' + ','.join(str(v or 0) for v in arrays.get(y1, {}).get('cli', [0]*12)) + ']'
    t25  = fmt(arrays.get(y1, {}).get('t', [None]*12))

    v26  = fmt(arrays.get(y2, {}).get('v', [None]*12)) if y2 else null12
    mg26 = fmt(arrays.get(y2, {}).get('mg', [None]*12)) if y2 else null12
    cli26 = '[' + ','.join(str(v or 'null') for v in arrays.get(y2, {}).get('cli', [None]*12)) + ']' if y2 else null12
    t26  = fmt(arrays.get(y2, {}).get('t', [None]*12)) if y2 else null12

    # YTD totales
    v24_arr = arrays.get(y0, {}).get('v', [])
    v25_arr = arrays.get(y1, {}).get('v', [])
    v26_arr = arrays.get(y2, {}).get('v', []) if y2 else []
    tot24 = round1(sum(v for v in v24_arr if v))
    tot25 = round1(sum(v for v in v25_arr if v))
    tot26 = round1(sum(v for v in v26_arr if v))

    new_ecom = (
        f"  label:'E-Commerce',c:'#3b82f6',lc:'#60a5fa',\n"
        f"  // Venta mensual (M CLP sin IVA) · Fuente: Power BI PROFAR · auto-update: {datetime.now().strftime('%Y-%m-%d %H:%M')}\n"
        f"  v{y0%100}:{v24},\n  v{y1%100}:{v25},\n  v{y2%100 if y2 else '26'}:{v26},\n"
        f"  mg{y0%100}:{mg24},\n  mg{y1%100}:{mg25},\n  mg{y2%100 if y2 else '26'}:{mg26},\n"
        f"  cli{y0%100}:{cli24},\n  cli{y1%100}:{cli25},\n  cli{y2%100 if y2 else '26'}:{cli26},\n"
        f"  t{y0%100}:{t24},\n  t{y1%100}:{t25},\n  t{y2%100 if y2 else '26'}:{t26},\n"
        f"  tot{y0%100}:{tot24},tot{y1%100}:{tot25},tot{y2%100 if y2 else '26'}ytd:{tot26}"
    )

    pattern = r"label:'E-Commerce'.*?tot\d+ytd:\d+\.?\d*"
    new_html = re.sub(pattern, new_ecom, html, flags=re.DOTALL)
    if new_html == html:
        print("⚠️  No se encontró D.ecom en el HTML — no se actualizó")
        return html
    return new_html

def inject_monthly(html, arrays):
    """Actualiza MONTHLY_ constants con datos del año más reciente."""
    y_keys = sorted([y for y in arrays.keys() if arrays[y].get('v')])
    if not y_keys:
        return html
    latest = y_keys[-1]
    months_data = arrays[latest]
    prev_year = y_keys[-2] if len(y_keys) > 1 else None
    months_prev = arrays.get(prev_year, {}) if prev_year else {}

    # Tomar solo meses con datos
    active_months = [(m, months_data['v'][m-1]) for m in range(1,13) if months_data['v'][m-1] is not None]
    if not active_months:
        return html

    month_names_es = ['Enero','Febrero','Marzo','Abril','Mayo','Junio',
                      'Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre']

    rev   = [months_data['v'][m-1] for m,_ in active_months]
    docs  = [months_data.get('docs_arr', [None]*12)[m-1] if 'docs_arr' in months_data else None for m,_ in active_months]
    tickets = [months_data['t'][m-1] for m,_ in active_months]
    months_names = [month_names_es[m-1] for m,_ in active_months]

    # MoM
    mom = [None]
    for i in range(1, len(rev)):
        if rev[i] is not None and rev[i-1]:
            mom.append(round1((rev[i]-rev[i-1])/rev[i-1]*100))
        else:
            mom.append(None)

    # YoY
    yoy = []
    for i, (m, _) in enumerate(active_months):
        prev_v = months_prev.get('v', [None]*12)[m-1] if months_prev else None
        if rev[i] is not None and prev_v:
            yoy.append(round1((rev[i]-prev_v)/prev_v*100))
        else:
            yoy.append(None)

    def fmt_arr(arr):
        return '[' + ','.join(str(v) if v is not None else 'null' for v in arr) + ']'

    # AOV en CLP
    aov = [round(t*1000) if t else 'null' for t in tickets]

    new_block = (
        f"// E-Commerce Monthly Details — Power BI sin IVA · auto: {datetime.now().strftime('%Y-%m-%d %H:%M')}\n"
        f"const MONTHLY_MONTH={json.dumps(months_names)};\n"
        f"const MONTHLY_REVENUE={fmt_arr([round(v*1_000_000) if v else None for v in rev])};\n"
        f"const MONTHLY_ORDERS={fmt_arr([None]*len(rev))}; // docs reales: ver HIST_SNAPS\n"
        f"const MONTHLY_AOV={fmt_arr(aov)};\n"
        f"const MONTHLY_CONVERSION={fmt_arr([3.9]*len(rev))}; // estimado\n"
        f"const MONTHLY_MOM={fmt_arr(mom)};\n"
        f"const MONTHLY_YOY={fmt_arr(yoy)};"
    )

    pattern = r'// E-Commerce Monthly Details.*?const MONTHLY_YOY=\[.*?\];'
    new_html = re.sub(pattern, new_block, html, flags=re.DOTALL)
    return new_html

def inject_margins(html, arrays):
    """Actualiza MARGINS_ constants."""
    y_keys = sorted([y for y in arrays.keys() if arrays[y].get('mg')])
    if not y_keys:
        return html
    latest = y_keys[-1]
    prev_year = y_keys[-2] if len(y_keys) > 1 else None
    months_data = arrays[latest]
    months_prev = arrays.get(prev_year, {}) if prev_year else {}

    active_months = [(m, months_data['mg'][m-1]) for m in range(1,13) if months_data['mg'][m-1] is not None]
    if not active_months:
        return html

    month_names_es = ['Enero','Febrero','Marzo','Abril','Mayo','Junio',
                      'Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre']

    cr  = [months_data['mg'][m-1] for m,_ in active_months]
    months_names = [month_names_es[m-1] for m,_ in active_months]
    q1avg = round1(sum(v for v in cr if v) / len([v for v in cr if v])) if cr else 0

    # YoY pp
    yoy_pp = []
    for i, (m, _) in enumerate(active_months):
        prev_mg = months_prev.get('mg', [None]*12)[m-1] if months_prev else None
        if cr[i] is not None and prev_mg:
            yoy_pp.append(round1(cr[i] - prev_mg))
        else:
            yoy_pp.append(None)

    # Trend
    trend = []
    for i, y in enumerate(yoy_pp):
        if y is None: trend.append('Estable')
        elif y > 0.5: trend.append('Mejora')
        elif y < -0.5: trend.append('Baja')
        else: trend.append('Leve baja')

    def fmt_arr(arr):
        return '[' + ','.join(str(v) if v is not None else 'null' for v in arr) + ']'

    new_block = (
        f"// E-Commerce Margins Evolution — Power BI Margen CR sin IVA · auto: {datetime.now().strftime('%Y-%m-%d %H:%M')}\n"
        f"const MARGINS_MONTH={json.dumps(months_names)};\n"
        f"const MARGINS_CR={fmt_arr(cr)};\n"
        f"const MARGINS_GROSS={fmt_arr([round1(v*1.9) if v else None for v in cr])}; // estimado gross ≈CR×1.9\n"
        f"const MARGINS_Q1AVG={fmt_arr([q1avg]*len(cr))};\n"
        f"const MARGINS_YOY={fmt_arr(yoy_pp)};\n"
        f"const MARGINS_TREND={json.dumps(trend)};"
    )

    pattern = r'// E-Commerce Margins Evolution.*?const MARGINS_TREND=\[.*?\];'
    new_html = re.sub(pattern, new_block, html, flags=re.DOTALL)
    return new_html

# ── Main ──────────────────────────────────────────────────────────────────────
def main():
    print(f"\n🔄 PROFAR Dashboard Update · {datetime.now().strftime('%Y-%m-%d %H:%M')} CL")
    print(f"   HTML: {HTML_FILE}")
    print(f"   Data: {DATA_DIR}")
    print()

    # Verificar que el HTML existe
    if not HTML_FILE.exists():
        print(f"❌ No se encontró {HTML_FILE}")
        return 1

    # Leer datos
    resumen = read_resumen()
    detalle = read_detalle()

    if not resumen and not detalle:
        print("\n⚠️  No hay archivos de datos en /data/ — el dashboard NO se actualizó")
        print("   Coloca Resumen Mensual.xlsx y Detalle de Documentos.xlsx en /data/")
        return 0

    # Construir arrays
    arrays, combined = build_arrays(resumen, detalle)
    if not arrays:
        print("❌ No se pudieron construir los arrays — verificar formato de Excel")
        return 1

    print(f"\n📊 Arrays construidos para: {sorted(arrays.keys())}")
    for year, arrs in sorted(arrays.items()):
        months_with_data = sum(1 for v in arrs['v'] if v)
        print(f"   {year}: {months_with_data} meses · Total: ${sum(v for v in arrs['v'] if v):.1f}M sin IVA")

    # Construir HIST_SNAPS
    snaps = build_hist_snaps(arrays, combined)
    print(f"\n📸 HIST_SNAPS: {len(snaps)} entradas ({snaps[0]['date']} → {snaps[-1]['date']})")

    # Leer HTML
    html = HTML_FILE.read_text(encoding='utf-8')
    original = html

    # Inyectar
    print("\n💉 Inyectando en HTML...")
    html = inject_hist_snaps(html, snaps)
    html = inject_decom(html, arrays)
    html = inject_monthly(html, arrays)
    html = inject_margins(html, arrays)

    # Guardar
    if html != original:
        HTML_FILE.write_text(html, encoding='utf-8')
        print(f"\n✅ Dashboard actualizado: {HTML_FILE}")
        print(f"   Cambios: {abs(len(html)-len(original))} bytes diff")
    else:
        print("\n⏭️  Sin cambios detectados — HTML no modificado")

    # Log de IVA reconciliation
    if arrays:
        latest = max(arrays.keys())
        arr = arrays[latest]
        print(f"\n📋 IVA Reconciliation ({latest}):")
        print(f"   Power BI sin IVA → Magento equivalente (×1.19):")
        for m in range(1, 13):
            v = arr['v'][m-1]
            if v:
                month_names_es = ['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic']
                print(f"   {month_names_es[m-1]}: ${v:.1f}M sin IVA → ≈${v*1.19:.1f}M con IVA")

    return 0

if __name__ == '__main__':
    exit(main())
