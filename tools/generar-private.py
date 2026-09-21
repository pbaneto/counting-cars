#!/usr/bin/env python3
"""Genera src/private.js (NO se sube a git) a partir del Excel antiguo, el CSV de vehículos y los IDs.
Uso: python3 tools/generar-private.py "<super excel 2026.xlsx>" "<Vehículos.csv>" <ID hoja> <ID entrada> <ID procesados> <ID fras RM> <ID fras RM procesadas>
Requiere: pip install openpyxl"""
import csv, json, sys, openpyxl

xlsx, csv_path, sheet_id, c_ent, c_proc, c_fra, c_fra_proc = sys.argv[1:8]
wb = openpyxl.load_workbook(xlsx, data_only=True)
num = lambda v: v if isinstance(v, (int, float)) else None
norm = lambda s: ''.join(ch for ch in str(s).upper() if ch.isalnum())

# Coches: sólo filas con formato OK
coches = []
for r in csv.DictReader(open(csv_path, encoding='utf-8')):
    if r['Formato'].strip() == 'OK':
        coches.append([norm(r['Matrícula']), r['Cliente'].strip(), r['Vehículo'].strip()])

# Piloto: 1ª quincena de septiembre (columnas B..F de la hoja "sep") + tabla de órdenes (= trabajos)
s = wb['sep']
CORRIGE = {'557HYX': '0557HYX'}  # confirmado por el CSV y la factura RM
albaranes = []
for r in range(6, 400):
    mat, imp = s.cell(r, 5).value, s.cell(r, 6).value
    if mat is None or not isinstance(imp, (int, float)) or norm(mat) == 'FACTURA':
        continue
    m = norm(mat); m = CORRIGE.get(m, m)
    albaranes.append({'proveedor': 'Otros' if str(s.cell(r, 2).value).strip().lower() == 'otros' else 'RM', 'matricula': m, 'importe': imp})
placas = {a['matricula'] for a in albaranes}
trabajos, vistos = [], set()
for r in range(6, 400):
    m = s.cell(r, 15).value
    if not m: continue
    m = norm(m); m = CORRIGE.get(m, m)
    if m in placas and m not in vistos:
        vistos.add(m)
        trabajos.append({'matricula': m, 'pagado': s.cell(r, 13).value == 'Pagado', 'factura': num(s.cell(r, 17).value)})

# Resumen
t = wb['TOTAL']
y26 = [[num(t.cell(r, c).value) or 0 for c in (3, 4, 5, 6)] for r in range(3, 15)]
y25 = [[num(t.cell(r, c).value) for c in range(12, 18)] for r in range(19, 31)]
y24 = []
for r in range(37, 49):
    r1, r2 = num(t.cell(r, 12).value), num(t.cell(r, 13).value)
    rec = (r1 or 0) + (r2 or 0) if (r1 is not None or r2 is not None) else None
    y24.append([rec, num(t.cell(r, 14).value), num(t.cell(r, 15).value), num(t.cell(r, 16).value)])
fijos = {'banco': [[str(t.cell(r, 11).value).strip(), t.cell(r, 12).value] for r in range(4, 8)],
         'gastos': [[str(t.cell(r, 13).value).strip(), t.cell(r, 14).value] for r in range(4, 8)]}

priv = {'SPREADSHEET_ID': sheet_id,
        'CARPETAS': {'CARPETA_ENTRADA': c_ent, 'CARPETA_PROCESADOS': c_proc, 'CARPETA_FACTURAS_RM': c_fra, 'CARPETA_FACTURAS_RM_PROCESADAS': c_fra_proc},
        'COCHES': coches, 'PILOTO': {'fecha': '2026-09-15', 'albaranes': albaranes, 'trabajos': trabajos},
        'RESUMEN': {'anio': 2026, 'mesEnVivo': 9, 'y2026': y26, 'y2025': y25, 'y2024': y24, 'fijos': fijos}}
open('src/private.js', 'w', encoding='utf-8').write('/* Generado por tools/generar-private.py. NO SUBIR A GIT. */\nconst PRIVATE = ' + json.dumps(priv, ensure_ascii=False, indent=1) + ';\n')
print(f'coches={len(coches)} albaranes={len(albaranes)} trabajos={len(trabajos)} y26={len(y26)}')
