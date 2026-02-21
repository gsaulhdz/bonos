from flask import Flask, render_template, request, jsonify
import sqlite3
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import os
import io
from flask import send_file

app = Flask(__name__)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_DIR = os.path.join(BASE_DIR, 'uploads')
DB_PATH = os.path.join(BASE_DIR, 'bonos.db')
app.config['UPLOAD_FOLDER'] = UPLOAD_DIR
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024

def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = get_db()
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS trabajadores (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nomina TEXT UNIQUE,
        nombre TEXT,
        sucursal TEXT,
        puesto TEXT,
        area TEXT,
        nombre_suc TEXT
    )''')
    c.execute('''CREATE TABLE IF NOT EXISTS periodos (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nombre TEXT UNIQUE,
        mes TEXT,
        anio INTEGER,
        fecha_inicio TEXT,
        fecha_fin TEXT
    )''')
    c.execute('''CREATE TABLE IF NOT EXISTS checklists (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        periodo_id INTEGER,
        fecha TEXT,
        sucursal TEXT,
        area TEXT,
        calificacion REAL,
        supervisor TEXT
    )''')
    c.execute('''CREATE TABLE IF NOT EXISTS afectaciones (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
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
        id INTEGER PRIMARY KEY AUTOINCREMENT,
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
        id INTEGER PRIMARY KEY AUTOINCREMENT,
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
    """Determina el % de afectación de un acta según su descripción."""
    texto = (observaciones or '').upper()
    # Prioridad 1: menciona referencia → 5%
    if 'REFERENCIA' in texto or 'REFERENCIAS' in texto:
        return -5.0
    # Prioridad 2: movimiento mal aplicado, tipo incorrecto, almacén incorrecto, fuera de tiempo
    palabras_20 = ['MAL APLICADO', 'MALA APLICACIÓN', 'MALA APLICACION',
                   'MOVIMIENTO', 'FAMILIA DISTINTA', 'CONVERSION', 'CONVERSIÓN',
                   'NEGATIVO', 'TIEMPO Y FORMA', 'EN TIEMPO']
    for p in palabras_20:
        if p in texto:
            return -20.0
    # Todo lo demás → 5%
    return -5.0

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
        existing = c.execute('SELECT id FROM trabajadores WHERE nomina=?', (nomina,)).fetchone()
        if existing:
            c.execute('UPDATE trabajadores SET nombre=?, sucursal=?, puesto=?, area=?, nombre_suc=? WHERE nomina=?',
                      (nombre, num_suc, puesto, area, nombre_suc, nomina))
            actualizados += 1
        else:
            c.execute('INSERT INTO trabajadores (nomina, nombre, sucursal, puesto, area, nombre_suc) VALUES (?,?,?,?,?,?)',
                      (nomina, nombre, num_suc, puesto, area, nombre_suc))
            insertados += 1
    conn.commit()
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
            'supervisor': supervisor
        })
    conn = get_db()
    c = conn.cursor()
    c.execute('DELETE FROM checklists WHERE periodo_id=?', (periodo_id,))
    for r in registros:
        c.execute('INSERT INTO checklists (periodo_id, fecha, sucursal, area, calificacion, supervisor) VALUES (?,?,?,?,?,?)',
                  (periodo_id, r['fecha'], r['sucursal'], r['area'], r['calificacion'], r['supervisor']))
    conn.commit()
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
    c.execute('DELETE FROM afectaciones WHERE periodo_id=?', (periodo_id,))
    for r in registros:
        c.execute('INSERT INTO afectaciones (periodo_id, folio, sucursal, puesto, nomina, nombre, incidencia, fecha, porcentaje, observacion) VALUES (?,?,?,?,?,?,?,?,?,?)',
                  (periodo_id, r['folio'], r['sucursal'], r['puesto'], r['nomina'],
                   r['nombre'], r['incidencia'], r['fecha'], r['porcentaje'], r['observacion']))
    conn.commit()
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
    c.execute('DELETE FROM actas WHERE periodo_id=?', (periodo_id,))
    for r in registros:
        nom = c.execute('SELECT nomina FROM trabajadores WHERE nombre=?', (r['nombre'],)).fetchone()
        nomina = nom['nomina'] if nom else ''
        c.execute('INSERT INTO actas (periodo_id, anio, mes, almacen, area, puesto, nombre, fecha, procedimiento, folio, observaciones, nomina, porcentaje_afectacion) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)',
                  (periodo_id, r['anio'], r['mes'], r['almacen'], r['area'], r['puesto'],
                   r['nombre'], r['fecha'], r['procedimiento'], r['folio'],
                   r['observaciones'], nomina, r['porcentaje_afectacion']))
    conn.commit()
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
    c.execute('DELETE FROM bono_rotulos WHERE periodo_id=?', (periodo_id,))
    for r in registros:
        c.execute('INSERT INTO bono_rotulos (periodo_id, sucursal, material_pop, limpieza_visual, radio_dpp, chequeo, evidencias, total) VALUES (?,?,?,?,?,?,?,?)',
                  (periodo_id, r['sucursal'], r['material_pop'], r['limpieza_visual'],
                   r['radio_dpp'], r['chequeo'], r['evidencias'], r['total']))
    conn.commit()
    conn.close()
    return len(registros)

# =================== LÓGICA DE CÁLCULO ===================

def get_personal_sucursal(sucursal, conn):
    """Devuelve el personal de una sucursal agrupado por rol."""
    rows = conn.execute('SELECT * FROM trabajadores WHERE sucursal=?', (sucursal,)).fetchall()
    personal = {
        'encargado_mc': [r for r in rows if r['puesto'] == 'ENCARGADO DE MESA DE CONTROL'],
        'capturista': [r for r in rows if r['puesto'] == 'CAPTURISTA'],
        'rotulista': [r for r in rows if r['puesto'] == 'ROTULISTA'],
        'recibo': [r for r in rows if r['puesto'] == 'RECIBO DE PROVEEDORES'],
        'total': rows
    }
    return personal

def get_checklist_sucursal(sucursal, area_keyword, periodo_id, conn):
    """Busca el checklist más reciente de una sucursal y área en el periodo."""
    return conn.execute(
        'SELECT calificacion FROM checklists WHERE periodo_id=? AND sucursal=? AND area LIKE ? ORDER BY fecha DESC LIMIT 1',
        (periodo_id, sucursal, f'%{area_keyword}%')
    ).fetchone()

def calcular_bono_trabajador(nomina, sucursal, puesto, area, periodo_id):
    conn = get_db()
    c = conn.cursor()

    resultado = {
        'nomina': nomina, 'sucursal': sucursal, 'puesto': puesto, 'area': area,
        'bono_base': 100.0, 'checklist_aplicado': False, 'checklist_calificacion': None,
        'afectaciones': [], 'actas': [], 'bono_rotulos_externo': None,
        'total_afectaciones': 0.0, 'bono_final': 100.0,
        'checklist_heredado': None
    }

    suc_num = normalizar_sucursal(sucursal)
    personal = get_personal_sucursal(suc_num, conn)
    es_unico = len(personal['total']) == 1

    # ---- CHECKLIST ----
    if es_unico:
        # Persona única: se le aplican todos los checklists, tomamos el promedio o el más bajo
        cls = []
        for kw in ['MESA', 'ROTUL', 'RECIBO']:
            cl = get_checklist_sucursal(suc_num, kw, periodo_id, c)
            if cl:
                cls.append(float(cl['calificacion']) * 100)
        if cls:
            cal = round(sum(cls) / len(cls), 2)
            resultado['checklist_aplicado'] = True
            resultado['checklist_calificacion'] = cal
            resultado['bono_base'] = cal

    elif area == 'ROTULOS':
        cl_rotulos = get_checklist_sucursal(suc_num, 'ROTUL', periodo_id, c)
        br = c.execute('SELECT total FROM bono_rotulos WHERE periodo_id=? AND (sucursal LIKE ? OR sucursal LIKE ?) LIMIT 1',
                       (periodo_id, f'%{suc_num}%', f'{suc_num} %')).fetchone()
        checklist_pct = float(cl_rotulos['calificacion']) * 100 if cl_rotulos else 100.0
        externo_pct = float(br['total']) * 100 if br else 50.0
        puntos_checklist = (checklist_pct / 100) * 50
        puntos_externo = externo_pct
        resultado['checklist_aplicado'] = True if cl_rotulos else False
        resultado['checklist_calificacion'] = round(checklist_pct, 2)
        resultado['bono_rotulos_externo'] = round(externo_pct, 2)
        resultado['bono_base'] = round(puntos_checklist + puntos_externo, 2)

    elif area == 'RECIBO':
        cl = get_checklist_sucursal(suc_num, 'RECIBO', periodo_id, c)
        if cl:
            resultado['checklist_aplicado'] = True
            resultado['checklist_calificacion'] = round(float(cl['calificacion']) * 100, 2)
            resultado['bono_base'] = resultado['checklist_calificacion']

    elif area == 'MESA DE CONTROL':
        if puesto == 'CAPTURISTA':
            # Capturista hereda rótulos si no hay rotulista en la sucursal
            tiene_rotulista = len(personal['rotulista']) > 0
            if not tiene_rotulista:
                cl_rot = get_checklist_sucursal(suc_num, 'ROTUL', periodo_id, c)
                if cl_rot:
                    resultado['checklist_aplicado'] = True
                    resultado['checklist_calificacion'] = round(float(cl_rot['calificacion']) * 100, 2)
                    resultado['bono_base'] = resultado['checklist_calificacion']
                    resultado['checklist_heredado'] = 'ROTULOS'
                else:
                    # No hay checklist de rótulos, hereda mesa de control
                    cl_mc = get_checklist_sucursal(suc_num, 'MESA', periodo_id, c)
                    if cl_mc:
                        resultado['checklist_aplicado'] = True
                        resultado['checklist_calificacion'] = round(float(cl_mc['calificacion']) * 100, 2)
                        resultado['bono_base'] = resultado['checklist_calificacion']
                        resultado['checklist_heredado'] = 'MESA DE CONTROL'
            # Si hay rotulista, el capturista no recibe checklist de rótulos ni de mesa
        else:
            # Encargado de Mesa de Control
            cl = get_checklist_sucursal(suc_num, 'MESA', periodo_id, c)
            if cl:
                resultado['checklist_aplicado'] = True
                resultado['checklist_calificacion'] = round(float(cl['calificacion']) * 100, 2)
                resultado['bono_base'] = resultado['checklist_calificacion']

    # ---- AFECTACIONES DE BONO ----
    afects = c.execute('SELECT * FROM afectaciones WHERE periodo_id=? AND nomina=?', (periodo_id, nomina)).fetchall()
    total_afect = 0
    for a in afects:
        pct = float(a['porcentaje'])
        total_afect += pct
        resultado['afectaciones'].append({
            'folio': a['folio'], 'fecha': a['fecha'],
            'porcentaje': pct, 'observacion': a['observacion']
        })

    # ---- ACTAS ----
    trab = c.execute('SELECT nombre FROM trabajadores WHERE nomina=?', (nomina,)).fetchone()
    nombre_trab = trab['nombre'] if trab else ''
    actas_list = c.execute('SELECT * FROM actas WHERE periodo_id=? AND (nomina=? OR nombre=?)',
                           (periodo_id, nomina, nombre_trab)).fetchall()
    for a in actas_list:
        pct_acta = float(a['porcentaje_afectacion'])
        total_afect += pct_acta
        resultado['actas'].append({
            'folio': a['folio'], 'fecha': a['fecha'],
            'procedimiento': a['procedimiento'],
            'observaciones': a['observaciones'],
            'porcentaje': pct_acta
        })

    resultado['total_afectaciones'] = round(total_afect, 2)
    resultado['bono_final'] = round(max(0, resultado['bono_base'] + total_afect), 2)
    conn.close()
    return resultado

# =================== EXPORTAR EXCEL ===================

def generar_excel_reporte(reporte, nombre_periodo):
    wb = Workbook()
    ws = wb.active
    ws.title = "Reporte Bonos"

    # Estilos
    header_fill = PatternFill("solid", fgColor="1a2035")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    title_font = Font(bold=True, size=14, color="1a2035")
    border = Border(
        left=Side(style='thin', color='CCCCCC'),
        right=Side(style='thin', color='CCCCCC'),
        top=Side(style='thin', color='CCCCCC'),
        bottom=Side(style='thin', color='CCCCCC')
    )

    # Título
    ws.merge_cells('A1:K1')
    ws['A1'] = f'REPORTE DE BONOS DE PRODUCTIVIDAD — {nombre_periodo}'
    ws['A1'].font = title_font
    ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 30

    # Headers
    headers = ['Nómina', 'Nombre', 'Sucursal', 'Puesto', 'Área',
               'Checklist Aplicado', '% Checklist', 'Ext. Rótulos',
               'Total Afectaciones', 'Bono Final', 'Estado']
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=2, column=col, value=h)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center')
        cell.border = border

    # Datos
    for row_idx, t in enumerate(reporte, 3):
        bono = t['bono_final']
        estado = 'EXCELENTE' if bono >= 95 else 'BUENO' if bono >= 85 else 'REGULAR' if bono >= 70 else 'BAJO'
        row_fill = PatternFill("solid", fgColor="F8FFF8" if bono >= 85 else "FFFFF8" if bono >= 70 else "FFF8F8")

        valores = [
            t['nomina'], t.get('nombre', ''), t['sucursal'], t['puesto'], t['area'],
            'Sí' if t['checklist_aplicado'] else 'No',
            f"{t['checklist_calificacion']}%" if t['checklist_calificacion'] else '100% (base)',
            f"{t['bono_rotulos_externo']}%" if t['bono_rotulos_externo'] is not None else '—',
            f"{t['total_afectaciones']}%",
            f"{bono}%", estado
        ]
        for col, val in enumerate(valores, 1):
            cell = ws.cell(row=row_idx, column=col, value=val)
            cell.fill = row_fill
            cell.alignment = Alignment(horizontal='center' if col > 5 else 'left')
            cell.border = border

    # Hoja de detalle
    ws2 = wb.create_sheet("Detalle Afectaciones")
    detail_headers = ['Nómina', 'Nombre', 'Sucursal', 'Tipo', 'Folio', 'Fecha', '% Afectación', 'Descripción']
    for col, h in enumerate(detail_headers, 1):
        cell = ws2.cell(row=1, column=col, value=h)
        cell.fill = header_fill
        cell.font = header_font
        cell.border = border

    row_d = 2
    for t in reporte:
        for a in t.get('afectaciones', []):
            ws2.cell(row=row_d, column=1, value=t['nomina'])
            ws2.cell(row=row_d, column=2, value=t.get('nombre', ''))
            ws2.cell(row=row_d, column=3, value=t['sucursal'])
            ws2.cell(row=row_d, column=4, value='Afectación de Bono')
            ws2.cell(row=row_d, column=5, value=a['folio'])
            ws2.cell(row=row_d, column=6, value=a['fecha'])
            ws2.cell(row=row_d, column=7, value=f"{a['porcentaje']}%")
            ws2.cell(row=row_d, column=8, value=a['observacion'])
            row_d += 1
        for a in t.get('actas', []):
            ws2.cell(row=row_d, column=1, value=t['nomina'])
            ws2.cell(row=row_d, column=2, value=t.get('nombre', ''))
            ws2.cell(row=row_d, column=3, value=t['sucursal'])
            ws2.cell(row=row_d, column=4, value='Acta de Incumplimiento')
            ws2.cell(row=row_d, column=5, value=a['folio'])
            ws2.cell(row=row_d, column=6, value=a['fecha'])
            ws2.cell(row=row_d, column=7, value=f"{a['porcentaje']}%")
            ws2.cell(row=row_d, column=8, value=a['observaciones'])
            row_d += 1

    # Ajustar anchos
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
    periodos = conn.execute('SELECT * FROM periodos ORDER BY id DESC').fetchall()
    conn.close()
    return jsonify([dict(p) for p in periodos])

@app.route('/api/periodos', methods=['POST'])
def crear_periodo():
    data = request.json
    conn = get_db()
    try:
        conn.execute('INSERT INTO periodos (nombre, mes, anio, fecha_inicio, fecha_fin) VALUES (?,?,?,?,?)',
                     (data['nombre'], data['mes'], data['anio'], data.get('fecha_inicio',''), data.get('fecha_fin','')))
        conn.commit()
        pid = conn.execute('SELECT id FROM periodos WHERE nombre=?', (data['nombre'],)).fetchone()['id']
        conn.close()
        return jsonify({'success': True, 'id': pid})
    except Exception as e:
        conn.close()
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
    q = request.args.get('q', '')
    if q:
        rows = conn.execute('SELECT * FROM trabajadores WHERE nombre LIKE ? OR nomina LIKE ? OR sucursal LIKE ?',
                            (f'%{q}%', f'%{q}%', f'%{q}%')).fetchall()
    else:
        rows = conn.execute('SELECT * FROM trabajadores ORDER BY sucursal+0, nombre').fetchall()
    conn.close()
    return jsonify([dict(t) for t in rows])

@app.route('/api/trabajadores', methods=['POST'])
def agregar_trabajador():
    data = request.json
    conn = get_db()
    try:
        conn.execute('INSERT OR REPLACE INTO trabajadores (nomina, nombre, sucursal, puesto, area, nombre_suc) VALUES (?,?,?,?,?,?)',
                     (data['nomina'], data['nombre'], data['sucursal'], data['puesto'], data['area'], data.get('nombre_suc','')))
        conn.commit()
        conn.close()
        return jsonify({'success': True})
    except Exception as e:
        conn.close()
        return jsonify({'success': False, 'error': str(e)})

@app.route('/api/trabajadores/<int:id>', methods=['DELETE'])
def eliminar_trabajador(id):
    conn = get_db()
    conn.execute('DELETE FROM trabajadores WHERE id=?', (id,))
    conn.commit()
    conn.close()
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
    query = 'SELECT * FROM trabajadores WHERE 1=1'
    params = []
    if sucursal:
        query += ' AND sucursal=?'
        params.append(sucursal)
    if area:
        query += ' AND area=?'
        params.append(area)
    if buscar:
        query += ' AND (nombre LIKE ? OR nomina LIKE ?)'
        params += [f'%{buscar}%', f'%{buscar}%']
    trabajadores = conn.execute(query, params).fetchall()
    conn.close()
    reporte = []
    for t in trabajadores:
        bono = calcular_bono_trabajador(t['nomina'], t['sucursal'], t['puesto'], t['area'], periodo_id)
        bono['nombre'] = t['nombre']
        bono['nombre_suc'] = t['nombre_suc'] if 'nombre_suc' in t.keys() else ''
        reporte.append(bono)
    reporte.sort(key=lambda x: (int(x['sucursal']) if x['sucursal'].isdigit() else 999, x['nombre'] or ''))
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
    periodo = conn.execute('SELECT nombre FROM periodos WHERE id=?', (periodo_id,)).fetchone()
    nombre_periodo = periodo['nombre'] if periodo else 'Periodo'
    query = 'SELECT * FROM trabajadores WHERE 1=1'
    params = []
    if sucursal:
        query += ' AND sucursal=?'; params.append(sucursal)
    if area:
        query += ' AND area=?'; params.append(area)
    if buscar:
        query += ' AND (nombre LIKE ? OR nomina LIKE ?)'; params += [f'%{buscar}%', f'%{buscar}%']
    trabajadores = conn.execute(query, params).fetchall()
    conn.close()
    reporte = []
    for t in trabajadores:
        bono = calcular_bono_trabajador(t['nomina'], t['sucursal'], t['puesto'], t['area'], periodo_id)
        bono['nombre'] = t['nombre']
        reporte.append(bono)
    reporte.sort(key=lambda x: (int(x['sucursal']) if x['sucursal'].isdigit() else 999, x['nombre'] or ''))
    excel = generar_excel_reporte(reporte, nombre_periodo)
    return send_file(excel, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                     as_attachment=True, download_name=f'Bonos_{nombre_periodo.replace(" ","_")}.xlsx')

@app.route('/api/sucursales', methods=['GET'])
def get_sucursales():
    conn = get_db()
    rows = conn.execute('SELECT DISTINCT sucursal, nombre_suc FROM trabajadores ORDER BY sucursal+0').fetchall()
    conn.close()
    return jsonify([{'num': r['sucursal'], 'nombre': r['nombre_suc']} for r in rows])

# Inicializar al arrancar
os.makedirs(UPLOAD_DIR, exist_ok=True)
init_db()

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
