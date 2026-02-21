from flask import Flask, render_template, request, jsonify, send_file
import psycopg2
import psycopg2.extras
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import os
import io

app = Flask(__name__)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_DIR = os.path.join(BASE_DIR, 'uploads')
app.config['UPLOAD_FOLDER'] = UPLOAD_DIR
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024

DATABASE_URL = os.environ.get('DATABASE_URL', '')

def get_db():
    conn = psycopg2.connect(DATABASE_URL)
    return conn

def init_db():
    conn = get_db()
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS trabajadores (
        id SERIAL PRIMARY KEY,
        nomina TEXT UNIQUE,
        nombre TEXT,
        sucursal TEXT,
        puesto TEXT,
        area TEXT,
        nombre_suc TEXT
    )''')
    c.execute('''CREATE TABLE IF NOT EXISTS periodos (
        id SERIAL PRIMARY KEY,
        nombre TEXT UNIQUE,
        mes TEXT,
        anio INTEGER,
        fecha_inicio TEXT,
        fecha_fin TEXT
    )''')
    c.execute('''CREATE TABLE IF NOT EXISTS checklists (
        id SERIAL PRIMARY KEY,
        periodo_id INTEGER,
        fecha TEXT,
        sucursal TEXT,
        area TEXT,
        calificacion REAL,
        supervisor TEXT
    )''')
    c.execute('''CREATE TABLE IF NOT EXISTS afectaciones (
        id SERIAL PRIMARY KEY,
        periodo_id INTEGER,
        folio TEXT,
        sucursal TEXT,
        puesto TEXT,
        nomina TEXT,
        nombre TEXT,
        incidencia TEXT,
        fecha TEXT,
        porcentaje REAL,
        observacion TEXT
    )''')
    c.execute('''CREATE TABLE IF NOT EXISTS actas (
        id SERIAL PRIMARY KEY,
        periodo_id INTEGER,
        anio INTEGER,
        mes TEXT,
        almacen TEXT,
        area TEXT,
        puesto TEXT,
        nombre TEXT,
        fecha TEXT,
        procedimiento TEXT,
        folio TEXT,
        observaciones TEXT,
        nomina TEXT,
        porcentaje_afectacion REAL
    )''')
    c.execute('''CREATE TABLE IF NOT EXISTS bono_rotulos (
        id SERIAL PRIMARY KEY,
        periodo_id INTEGER,
        sucursal TEXT,
        material_pop REAL,
        limpieza_visual REAL,
        radio_dpp REAL,
        chequeo REAL,
        evidencias REAL,
        total REAL
    )''')
    conn.commit()
    c.close()
    conn.close()

def normalizar_sucursal(suc):
    if suc is None:
        return ''
    return str(suc).strip().lstrip('0').strip()

def detectar_area(puesto):
    puesto = (puesto or '').upper()
    if 'ROTUL' in puesto:
        return 'ROTULOS'
    elif 'RECIBO' in puesto or 'PROVEEDOR' in puesto:
        return 'RECIBO'
    else:
        return 'MESA DE CONTROL'

def calcular_afectacion_acta(observaciones):
    texto = (observaciones or '').upper()
    if 'REFERENCIA' in texto or 'REFERENCIAS' in texto:
        return -5.0
    palabras_20 = ['MAL APLICADO', 'MALA APLICACIÓN', 'MALA APLICACION',
                   'MOVIMIENTO', 'FAMILIA DISTINTA', 'CONVERSION', 'CONVERSIÓN',
                   'NEGATIVO', 'TIEMPO Y FORMA', 'EN TIEMPO']
    for p in palabras_20:
        if p in texto:
            return -20.0
    return -5.0

def row_to_dict(cursor, row):
    cols = [desc[0] for desc in cursor.description]
    return dict(zip(cols, row))

def rows_to_dicts(cursor, rows):
    cols = [desc[0] for desc in cursor.description]
    return [dict(zip(cols, row)) for row in rows]

# =================== CARGA DE CATÁLOGO ===================

def cargar_catalogo(filepath):
    wb = openpyxl.load_workbook(filepath, data_only=True)
    ws = wb["Hoja1"]
    conn = get_db()
    c = conn.cursor()
    insertados = 0
    actualizados = 0
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0] is None:
            continue
        nomina = str(int(row[0]))
        nombre = str(row[1]).strip() if row[1] else ''
        puesto = str(row[2]).strip() if row[2] else ''
        num_suc = str(row[4]).strip().lstrip('0').strip() if row[4] else ''
        nombre_suc = str(row[5]).strip() if row[5] else ''
        area = detectar_area(puesto)
        c.execute('SELECT id FROM trabajadores WHERE nomina=%s', (nomina,))
        existing = c.fetchone()
        if existing:
            c.execute('UPDATE trabajadores SET nombre=%s, sucursal=%s, puesto=%s, area=%s, nombre_suc=%s WHERE nomina=%s',
                      (nombre, num_suc, puesto, area, nombre_suc, nomina))
            actualizados += 1
        else:
            c.execute('INSERT INTO trabajadores (nomina, nombre, sucursal, puesto, area, nombre_suc) VALUES (%s,%s,%s,%s,%s,%s)',
                      (nomina, nombre, num_suc, puesto, area, nombre_suc))
            insertados += 1
    conn.commit()
    c.close()
    conn.close()
    return insertados, actualizados

# =================== PROCESAMIENTO DE ARCHIVOS ===================

def procesar_checklist(filepath, periodo_id):
    wb = openpyxl.load_workbook(filepath, data_only=True)
    ws = wb.active
    registros = []
    header_found = False
    for row in ws.iter_rows(values_only=True):
        if row[0] == 'Fecha':
            header_found = True
            continue
        if not header_found or row[0] is None:
            continue
        fecha, _, _, _, suc, supervisor, checklist, calificacion, mes, anio = row[:10]
        registros.append({
            'fecha': str(fecha),
            'sucursal': normalizar_sucursal(suc),
            'area': str(checklist).upper() if checklist else '',
            'calificacion': float(calificacion) if calificacion else 0,
            'supervisor': str(supervisor) if supervisor else ''
        })
    conn = get_db()
    c = conn.cursor()
    c.execute('DELETE FROM checklists WHERE periodo_id=%s', (periodo_id,))
    for r in registros:
        c.execute('INSERT INTO checklists (periodo_id, fecha, sucursal, area, calificacion, supervisor) VALUES (%s,%s,%s,%s,%s,%s)',
                  (periodo_id, r['fecha'], r['sucursal'], r['area'], r['calificacion'], r['supervisor']))
    conn.commit()
    c.close()
    conn.close()
    return len(registros)

def procesar_afectaciones(filepath, periodo_id):
    wb = openpyxl.load_workbook(filepath, data_only=True)
    registros = []
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        header_found = False
        for row in ws.iter_rows(values_only=True):
            if row[0] == 'Folio':
                header_found = True
                continue
            if not header_found or row[0] is None:
                continue
            folio, suc, puesto, nomina, nombre, incidencia, fecha, porcentaje, observacion = row[:9]
            registros.append({
                'folio': str(folio) if folio else '',
                'sucursal': normalizar_sucursal(suc),
                'puesto': str(puesto) if puesto else '',
                'nomina': str(int(nomina)) if nomina and str(nomina) != 'None' else '',
                'nombre': str(nombre) if nombre else '',
                'incidencia': str(incidencia) if incidencia else '',
                'fecha': str(fecha) if fecha else '',
                'porcentaje': float(porcentaje) if porcentaje else 0,
                'observacion': str(observacion) if observacion else ''
            })
    conn = get_db()
    c = conn.cursor()
    c.execute('DELETE FROM afectaciones WHERE periodo_id=%s', (periodo_id,))
    for r in registros:
        c.execute('INSERT INTO afectaciones (periodo_id, folio, sucursal, puesto, nomina, nombre, incidencia, fecha, porcentaje, observacion) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)',
                  (periodo_id, r['folio'], r['sucursal'], r['puesto'], r['nomina'],
                   r['nombre'], r['incidencia'], r['fecha'], r['porcentaje'], r['observacion']))
    conn.commit()
    c.close()
    conn.close()
    return len(registros)

def procesar_actas(filepath, periodo_id):
    wb = openpyxl.load_workbook(filepath, data_only=True)
    ws = wb.active
    registros = []
    header_found = False
    for row in ws.iter_rows(values_only=True):
        if row[0] == 'Año':
            header_found = True
            continue
        if not header_found or row[0] is None:
            continue
        anio, mes, almacen, area, puesto, nombre, fecha, procedimiento, folio, observaciones = row[:10]
        pct = calcular_afectacion_acta(str(observaciones))
        registros.append({
            'anio': anio, 'mes': mes,
            'almacen': normalizar_sucursal(almacen),
            'area': str(area) if area else '',
            'puesto': str(puesto) if puesto else '',
            'nombre': str(nombre) if nombre else '',
            'fecha': str(fecha) if fecha else '',
            'procedimiento': str(procedimiento) if procedimiento else '',
            'folio': str(folio) if folio else '',
            'observaciones': str(observaciones) if observaciones else '',
            'porcentaje_afectacion': pct
        })
    conn = get_db()
    c = conn.cursor()
    c.execute('DELETE FROM actas WHERE periodo_id=%s', (periodo_id,))
    for r in registros:
        c.execute('SELECT nomina FROM trabajadores WHERE nombre=%s', (r['nombre'],))
        nom = c.fetchone()
        nomina = nom[0] if nom else ''
        c.execute('INSERT INTO actas (periodo_id, anio, mes, almacen, area, puesto, nombre, fecha, procedimiento, folio, observaciones, nomina, porcentaje_afectacion) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)',
                  (periodo_id, r['anio'], r['mes'], r['almacen'], r['area'], r['puesto'],
                   r['nombre'], r['fecha'], r['procedimiento'], r['folio'],
                   r['observaciones'], nomina, r['porcentaje_afectacion']))
    conn.commit()
    c.close()
    conn.close()
    return len(registros)

def procesar_bono_rotulos(filepath, periodo_id):
    wb = openpyxl.load_workbook(filepath, data_only=True)
    registros = []
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        header_found = False
        for row in ws.iter_rows(values_only=True):
            if row[0] == 'Sucursal':
                header_found = True
                continue
            if not header_found or row[0] is None:
                continue
            suc = row[0]
            if not isinstance(suc, str):
                continue
            mat_pop, limp_vis, radio, chequeo, evidencias, total = row[1:7]
            registros.append({
                'sucursal': suc.strip(),
                'material_pop': float(mat_pop) if mat_pop else 0,
                'limpieza_visual': float(limp_vis) if limp_vis else 0,
                'radio_dpp': float(radio) if radio else 0,
                'chequeo': float(chequeo) if chequeo else 0,
                'evidencias': float(evidencias) if evidencias else 0,
                'total': float(total) if total else 0
            })
    conn = get_db()
    c = conn.cursor()
    c.execute('DELETE FROM bono_rotulos WHERE periodo_id=%s', (periodo_id,))
    for r in registros:
        c.execute('INSERT INTO bono_rotulos (periodo_id, sucursal, material_pop, limpieza_visual, radio_dpp, chequeo, evidencias, total) VALUES (%s,%s,%s,%s,%s,%s,%s,%s)',
                  (periodo_id, r['sucursal'], r['material_pop'], r['limpieza_visual'],
                   r['radio_dpp'], r['chequeo'], r['evidencias'], r['total']))
    conn.commit()
    c.close()
    conn.close()
    return len(registros)

# =================== LÓGICA DE CÁLCULO ===================

def get_personal_sucursal(sucursal, conn):
    c = conn.cursor()
    c.execute('SELECT * FROM trabajadores WHERE sucursal=%s', (sucursal,))
    rows = rows_to_dicts(c, c.fetchall())
    c.close()
    return {
        'encargado_mc': [r for r in rows if r['puesto'] == 'ENCARGADO DE MESA DE CONTROL'],
        'capturista': [r for r in rows if r['puesto'] == 'CAPTURISTA'],
        'rotulista': [r for r in rows if r['puesto'] == 'ROTULISTA'],
        'recibo': [r for r in rows if r['puesto'] == 'RECIBO DE PROVEEDORES'],
        'total': rows
    }

def get_checklist_sucursal(sucursal, area_keyword, periodo_id, conn):
    c = conn.cursor()
    c.execute('SELECT calificacion FROM checklists WHERE periodo_id=%s AND sucursal=%s AND area ILIKE %s ORDER BY fecha DESC LIMIT 1',
              (periodo_id, sucursal, f'%{area_keyword}%'))
    row = c.fetchone()
    c.close()
    return row

def calcular_bono_trabajador(nomina, sucursal, puesto, area, periodo_id):
    conn = get_db()

    resultado = {
        'nomina': nomina, 'sucursal': sucursal, 'puesto': puesto, 'area': area,
        'bono_base': 100.0, 'checklist_aplicado': False, 'checklist_calificacion': None,
        'afectaciones': [], 'actas': [], 'bono_rotulos_externo': None,
        'total_afectaciones': 0.0, 'bono_final': 100.0, 'checklist_heredado': None
    }

    suc_num = normalizar_sucursal(sucursal)
    personal = get_personal_sucursal(suc_num, conn)
    es_unico = len(personal['total']) == 1

    if es_unico:
        cls = []
        for kw in ['MESA', 'ROTUL', 'RECIBO']:
            cl = get_checklist_sucursal(suc_num, kw, periodo_id, conn)
            if cl:
                cls.append(float(cl[0]) * 100)
        if cls:
            cal = round(sum(cls) / len(cls), 2)
            resultado['checklist_aplicado'] = True
            resultado['checklist_calificacion'] = cal
            resultado['bono_base'] = cal

    elif area == 'ROTULOS':
        cl_rotulos = get_checklist_sucursal(suc_num, 'ROTUL', periodo_id, conn)
        c = conn.cursor()
        c.execute("SELECT total FROM bono_rotulos WHERE periodo_id=%s AND (sucursal ILIKE %s OR sucursal ILIKE %s) LIMIT 1",
                  (periodo_id, f'%{suc_num}%', f'{suc_num} %'))
        br = c.fetchone()
        c.close()
        checklist_pct = float(cl_rotulos[0]) * 100 if cl_rotulos else 100.0
        externo_pct = float(br[0]) * 100 if br else 50.0
        puntos_checklist = (checklist_pct / 100) * 50
        resultado['checklist_aplicado'] = True if cl_rotulos else False
        resultado['checklist_calificacion'] = round(checklist_pct, 2)
        resultado['bono_rotulos_externo'] = round(externo_pct, 2)
        resultado['bono_base'] = round(puntos_checklist + externo_pct, 2)

    elif area == 'RECIBO':
        cl = get_checklist_sucursal(suc_num, 'RECIBO', periodo_id, conn)
        if cl:
            resultado['checklist_aplicado'] = True
            resultado['checklist_calificacion'] = round(float(cl[0]) * 100, 2)
            resultado['bono_base'] = resultado['checklist_calificacion']

    elif area == 'MESA DE CONTROL':
        if puesto == 'CAPTURISTA':
            tiene_rotulista = len(personal['rotulista']) > 0
            if not tiene_rotulista:
                cl_rot = get_checklist_sucursal(suc_num, 'ROTUL', periodo_id, conn)
                if cl_rot:
                    resultado['checklist_aplicado'] = True
                    resultado['checklist_calificacion'] = round(float(cl_rot[0]) * 100, 2)
                    resultado['bono_base'] = resultado['checklist_calificacion']
                    resultado['checklist_heredado'] = 'ROTULOS'
                else:
                    cl_mc = get_checklist_sucursal(suc_num, 'MESA', periodo_id, conn)
                    if cl_mc:
                        resultado['checklist_aplicado'] = True
                        resultado['checklist_calificacion'] = round(float(cl_mc[0]) * 100, 2)
                        resultado['bono_base'] = resultado['checklist_calificacion']
                        resultado['checklist_heredado'] = 'MESA DE CONTROL'
        else:
            cl = get_checklist_sucursal(suc_num, 'MESA', periodo_id, conn)
            if cl:
                resultado['checklist_aplicado'] = True
                resultado['checklist_calificacion'] = round(float(cl[0]) * 100, 2)
                resultado['bono_base'] = resultado['checklist_calificacion']

    # Afectaciones
    c = conn.cursor()
    c.execute('SELECT * FROM afectaciones WHERE periodo_id=%s AND nomina=%s', (periodo_id, nomina))
    afects = rows_to_dicts(c, c.fetchall())
    c.close()
    total_afect = 0
    for a in afects:
        pct = float(a['porcentaje'])
        total_afect += pct
        resultado['afectaciones'].append({'folio': a['folio'], 'fecha': a['fecha'], 'porcentaje': pct, 'observacion': a['observacion']})

    # Actas
    c = conn.cursor()
    c.execute('SELECT nombre FROM trabajadores WHERE nomina=%s', (nomina,))
    trab = c.fetchone()
    nombre_trab = trab[0] if trab else ''
    c.execute('SELECT * FROM actas WHERE periodo_id=%s AND (nomina=%s OR nombre=%s)', (periodo_id, nomina, nombre_trab))
    actas_list = rows_to_dicts(c, c.fetchall())
    c.close()
    for a in actas_list:
        pct_acta = float(a['porcentaje_afectacion'])
        total_afect += pct_acta
        resultado['actas'].append({'folio': a['folio'], 'fecha': a['fecha'], 'procedimiento': a['procedimiento'], 'observaciones': a['observaciones'], 'porcentaje': pct_acta})

    resultado['total_afectaciones'] = round(total_afect, 2)
    resultado['bono_final'] = round(max(0, resultado['bono_base'] + total_afect), 2)
    conn.close()
    return resultado

# =================== EXPORTAR EXCEL ===================

def generar_excel_reporte(reporte, nombre_periodo):
    wb = Workbook()
    ws = wb.active
    ws.title = "Reporte Bonos"
    header_fill = PatternFill("solid", fgColor="1a2035")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    title_font = Font(bold=True, size=14, color="1a2035")
    border = Border(left=Side(style='thin',color='CCCCCC'), right=Side(style='thin',color='CCCCCC'), top=Side(style='thin',color='CCCCCC'), bottom=Side(style='thin',color='CCCCCC'))
    ws.merge_cells('A1:K1')
    ws['A1'] = f'REPORTE DE BONOS DE PRODUCTIVIDAD — {nombre_periodo}'
    ws['A1'].font = title_font
    ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 30
    headers = ['Nómina','Nombre','Sucursal','Nombre Suc.','Puesto','Área','Checklist','Ext. Rótulos','Total Afectaciones','Bono Final','Estado']
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=2, column=col, value=h)
        cell.fill = header_fill; cell.font = header_font
        cell.alignment = Alignment(horizontal='center'); cell.border = border
    for row_idx, t in enumerate(reporte, 3):
        bono = t['bono_final']
        estado = 'EXCELENTE' if bono >= 95 else 'BUENO' if bono >= 85 else 'REGULAR' if bono >= 70 else 'BAJO'
        row_fill = PatternFill("solid", fgColor="F8FFF8" if bono >= 85 else "FFFFF8" if bono >= 70 else "FFF8F8")
        valores = [t['nomina'], t.get('nombre',''), t['sucursal'], t.get('nombre_suc',''), t['puesto'], t['area'],
                   f"{t['checklist_calificacion']}%" if t['checklist_calificacion'] else '100% (base)',
                   f"{t['bono_rotulos_externo']}%" if t['bono_rotulos_externo'] is not None else '—',
                   f"{t['total_afectaciones']}%", f"{bono}%", estado]
        for col, val in enumerate(valores, 1):
            cell = ws.cell(row=row_idx, column=col, value=val)
            cell.fill = row_fill
            cell.alignment = Alignment(horizontal='center' if col > 5 else 'left')
            cell.border = border
    ws2 = wb.create_sheet("Detalle Afectaciones")
    detail_headers = ['Nómina','Nombre','Sucursal','Tipo','Folio','Fecha','% Afectación','Descripción']
    for col, h in enumerate(detail_headers, 1):
        cell = ws2.cell(row=1, column=col, value=h)
        cell.fill = header_fill; cell.font = header_font; cell.border = border
    row_d = 2
    for t in reporte:
        for a in t.get('afectaciones', []):
            for col, val in enumerate([t['nomina'], t.get('nombre',''), t['sucursal'], 'Afectación de Bono', a['folio'], a['fecha'], f"{a['porcentaje']}%", a['observacion']], 1):
                ws2.cell(row=row_d, column=col, value=val)
            row_d += 1
        for a in t.get('actas', []):
            for col, val in enumerate([t['nomina'], t.get('nombre',''), t['sucursal'], 'Acta de Incumplimiento', a['folio'], a['fecha'], f"{a['porcentaje']}%", a['observaciones']], 1):
                ws2.cell(row=row_d, column=col, value=val)
            row_d += 1
    for ws_obj in [ws, ws2]:
        for col in ws_obj.columns:
            max_len = max((len(str(cell.value or '')) for cell in col), default=0)
            ws_obj.column_dimensions[col[0].column_letter].width = min(max_len + 4, 50)
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# =================== RUTAS ===================

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/api/periodos', methods=['GET'])
def get_periodos():
    conn = get_db()
    c = conn.cursor()
    c.execute('SELECT * FROM periodos ORDER BY id DESC')
    data = rows_to_dicts(c, c.fetchall())
    c.close(); conn.close()
    return jsonify(data)

@app.route('/api/periodos', methods=['POST'])
def crear_periodo():
    data = request.json
    conn = get_db()
    c = conn.cursor()
    try:
        c.execute('INSERT INTO periodos (nombre, mes, anio, fecha_inicio, fecha_fin) VALUES (%s,%s,%s,%s,%s) RETURNING id',
                  (data['nombre'], data['mes'], data['anio'], data.get('fecha_inicio',''), data.get('fecha_fin','')))
        pid = c.fetchone()[0]
        conn.commit(); c.close(); conn.close()
        return jsonify({'success': True, 'id': pid})
    except Exception as e:
        conn.rollback(); c.close(); conn.close()
        return jsonify({'success': False, 'error': str(e)})

@app.route('/api/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({'success': False, 'error': 'No file'})
    file = request.files['file']
    tipo = request.form.get('tipo')
    periodo_id = request.form.get('periodo_id')
    os.makedirs(UPLOAD_DIR, exist_ok=True)
    filepath = os.path.join(UPLOAD_DIR, file.filename)
    file.save(filepath)
    try:
        if tipo == 'checklist':
            n = procesar_checklist(filepath, periodo_id)
        elif tipo == 'afectaciones':
            n = procesar_afectaciones(filepath, periodo_id)
        elif tipo == 'actas':
            n = procesar_actas(filepath, periodo_id)
        elif tipo == 'rotulos':
            n = procesar_bono_rotulos(filepath, periodo_id)
        elif tipo == 'catalogo':
            ins, act = cargar_catalogo(filepath)
            return jsonify({'success': True, 'registros': ins + act, 'insertados': ins, 'actualizados': act})
        else:
            return jsonify({'success': False, 'error': 'Tipo no reconocido'})
        return jsonify({'success': True, 'registros': n})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})

@app.route('/api/trabajadores', methods=['GET'])
def get_trabajadores():
    conn = get_db()
    c = conn.cursor()
    q = request.args.get('q', '')
    if q:
        c.execute('SELECT * FROM trabajadores WHERE nombre ILIKE %s OR nomina ILIKE %s OR sucursal ILIKE %s ORDER BY sucursal::integer, nombre',
                  (f'%{q}%', f'%{q}%', f'%{q}%'))
    else:
        c.execute('SELECT * FROM trabajadores ORDER BY sucursal::integer, nombre')
    data = rows_to_dicts(c, c.fetchall())
    c.close(); conn.close()
    return jsonify(data)

@app.route('/api/trabajadores', methods=['POST'])
def agregar_trabajador():
    data = request.json
    conn = get_db()
    c = conn.cursor()
    try:
        c.execute('INSERT INTO trabajadores (nomina, nombre, sucursal, puesto, area, nombre_suc) VALUES (%s,%s,%s,%s,%s,%s) ON CONFLICT (nomina) DO UPDATE SET nombre=EXCLUDED.nombre, sucursal=EXCLUDED.sucursal, puesto=EXCLUDED.puesto, area=EXCLUDED.area, nombre_suc=EXCLUDED.nombre_suc',
                  (data['nomina'], data['nombre'], data['sucursal'], data['puesto'], data['area'], data.get('nombre_suc','')))
        conn.commit(); c.close(); conn.close()
        return jsonify({'success': True})
    except Exception as e:
        conn.rollback(); c.close(); conn.close()
        return jsonify({'success': False, 'error': str(e)})

@app.route('/api/trabajadores/<int:id>', methods=['DELETE'])
def eliminar_trabajador(id):
    conn = get_db()
    c = conn.cursor()
    c.execute('DELETE FROM trabajadores WHERE id=%s', (id,))
    conn.commit(); c.close(); conn.close()
    return jsonify({'success': True})

@app.route('/api/reporte', methods=['GET'])
def get_reporte():
    periodo_id = request.args.get('periodo_id')
    sucursal = request.args.get('sucursal', '')
    area = request.args.get('area', '')
    buscar = request.args.get('buscar', '')
    if not periodo_id:
        return jsonify({'error': 'Periodo requerido'})
    conn = get_db()
    c = conn.cursor()
    query = 'SELECT * FROM trabajadores WHERE 1=1'
    params = []
    if sucursal:
        query += ' AND sucursal=%s'; params.append(sucursal)
    if area:
        query += ' AND area=%s'; params.append(area)
    if buscar:
        query += ' AND (nombre ILIKE %s OR nomina ILIKE %s)'; params += [f'%{buscar}%', f'%{buscar}%']
    query += ' ORDER BY sucursal::integer, nombre'
    c.execute(query, params)
    trabajadores = rows_to_dicts(c, c.fetchall())
    c.close(); conn.close()
    reporte = []
    for t in trabajadores:
        bono = calcular_bono_trabajador(t['nomina'], t['sucursal'], t['puesto'], t['area'], periodo_id)
        bono['nombre'] = t['nombre']
        bono['nombre_suc'] = t.get('nombre_suc', '')
        reporte.append(bono)
    return jsonify(reporte)

@app.route('/api/reporte/excel', methods=['GET'])
def exportar_excel():
    periodo_id = request.args.get('periodo_id')
    sucursal = request.args.get('sucursal', '')
    area = request.args.get('area', '')
    buscar = request.args.get('buscar', '')
    if not periodo_id:
        return jsonify({'error': 'Periodo requerido'})
    conn = get_db()
    c = conn.cursor()
    c.execute('SELECT nombre FROM periodos WHERE id=%s', (periodo_id,))
    periodo = c.fetchone()
    nombre_periodo = periodo[0] if periodo else 'Periodo'
    query = 'SELECT * FROM trabajadores WHERE 1=1'
    params = []
    if sucursal:
        query += ' AND sucursal=%s'; params.append(sucursal)
    if area:
        query += ' AND area=%s'; params.append(area)
    if buscar:
        query += ' AND (nombre ILIKE %s OR nomina ILIKE %s)'; params += [f'%{buscar}%', f'%{buscar}%']
    query += ' ORDER BY sucursal::integer, nombre'
    c.execute(query, params)
    trabajadores = rows_to_dicts(c, c.fetchall())
    c.close(); conn.close()
    reporte = []
    for t in trabajadores:
        bono = calcular_bono_trabajador(t['nomina'], t['sucursal'], t['puesto'], t['area'], periodo_id)
        bono['nombre'] = t['nombre']
        bono['nombre_suc'] = t.get('nombre_suc', '')
        reporte.append(bono)
    excel = generar_excel_reporte(reporte, nombre_periodo)
    return send_file(excel, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                     as_attachment=True, download_name=f'Bonos_{nombre_periodo.replace(" ","_")}.xlsx')

@app.route('/api/sucursales', methods=['GET'])
def get_sucursales():
    conn = get_db()
    c = conn.cursor()
    c.execute('SELECT DISTINCT sucursal, nombre_suc FROM trabajadores ORDER BY sucursal::integer')
    rows = c.fetchall()
    c.close(); conn.close()
    return jsonify([{'num': r[0], 'nombre': r[1]} for r in rows])

# Inicializar
os.makedirs(UPLOAD_DIR, exist_ok=True)
try:
    init_db()
except Exception as e:
    print(f"DB init error: {e}")

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
