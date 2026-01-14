# --- STOCK REGISTRO ENDPOINT ---
# (Movido después de la inicialización de Flask y Firebase)

from flask import Flask, request, jsonify, redirect, render_template
import firebase_admin
from firebase_admin import credentials, auth
from google.cloud import firestore
import os
import json

print('Iniciando Flask...')

# Inicializar Firebase con credenciales desde variable de entorno o archivo local
if os.environ.get('FIREBASE_CREDENTIALS'):
    # En producción (Railway): leer desde variable de entorno
    print('Usando credenciales de Firebase desde variable de entorno')
    firebase_creds = json.loads(os.environ.get('FIREBASE_CREDENTIALS'))
    cred = credentials.Certificate(firebase_creds)
    if not firebase_admin._apps:
        firebase_admin.initialize_app(cred)
    db = firestore.Client.from_service_account_info(firebase_creds)
else:
    # En desarrollo local: leer desde archivo
    print('Usando credenciales de Firebase desde archivo local')
    cred = credentials.Certificate('supratechweb-firebase-adminsdk-fbsvc-8d4aa68a75.json')
    if not firebase_admin._apps:
        firebase_admin.initialize_app(cred)
    db = firestore.Client.from_service_account_json('supratechweb-firebase-adminsdk-fbsvc-8d4aa68a75.json')

app = Flask(__name__)

# Función auxiliar para obtener credenciales de Google API
def get_google_credentials():
    """Retorna credenciales de Google API desde variable de entorno o archivo local"""
    from google.oauth2 import service_account
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

    if os.environ.get('FIREBASE_CREDENTIALS'):
        # En producción: usar credenciales desde variable de entorno
        firebase_creds = json.loads(os.environ.get('FIREBASE_CREDENTIALS'))
        return service_account.Credentials.from_service_account_info(firebase_creds, scopes=SCOPES)
    else:
        # En desarrollo: usar archivo local
        SERVICE_ACCOUNT_FILE = 'supratechweb-firebase-adminsdk-fbsvc-8d4aa68a75.json'
        return service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)

# --- ENDPOINTS PARA BDMARCAS ---
@app.route('/api/bdmarcas_campos', methods=['GET'])
def bdmarcas_campos():
    auth_header = request.headers.get('Authorization')
    if not auth_header or not auth_header.startswith('Bearer '):
        return jsonify({'error': 'No token provided'}), 401
    id_token = auth_header.split(' ')[1]
    try:
        decoded_token = auth.verify_id_token(id_token)
        uid = decoded_token['uid']
        hoja_ref = db.collection('Areas').document(uid).collection('Hojas').document('BDMarcas')
        hoja_doc = hoja_ref.get()
        if not hoja_doc.exists:
            return jsonify({'error': 'No se encontró la configuración de hoja/rango'}), 404
        hoja_data = hoja_doc.to_dict()
        campos = [k for k in hoja_data.keys() if k != 'Hoja']
        return jsonify({'campos': campos})
    except Exception as e:
        return jsonify({'error': 'Error al obtener campos', 'details': str(e)}), 500

@app.route('/api/bdmarcas_registro', methods=['POST'])
def bdmarcas_registro():
    auth_header = request.headers.get('Authorization')
    if not auth_header or not auth_header.startswith('Bearer '):
        return jsonify({'error': 'No token provided'}), 401
    id_token = auth_header.split(' ')[1]
    import datetime, re, traceback
    try:
        decoded_token = auth.verify_id_token(id_token)
        uid = decoded_token['uid']
        data = request.get_json()
        area_doc = db.collection('Areas').document(uid).get()
        if not area_doc.exists:
            return jsonify({'error': 'No se encontró el área para este usuario'}), 404
        area_data = area_doc.to_dict()
        spreadsheet_url = area_data.get('BDMarcas')
        if not spreadsheet_url:
            return jsonify({'error': 'No se encontró la URL del spreadsheet'}), 404
        match = re.search(r"/d/([a-zA-Z0-9-_]+)", spreadsheet_url)
        if not match:
            return jsonify({'error': 'URL de spreadsheet inválida'}), 400
        spreadsheet_id = match.group(1)
        hoja_ref = db.collection('Areas').document(uid).collection('Hojas').document('BDMarcas')
        hoja_doc = hoja_ref.get()
        if not hoja_doc.exists:
            return jsonify({'error': 'No se encontró la configuración de hoja/rango'}), 404
        hoja_data = hoja_doc.to_dict()
        nombre_hoja = hoja_data.get('Hoja', 'Sheet1')
        ubicaciones = {k: v for k, v in hoja_data.items() if k != 'Hoja'}
        from googleapiclient.discovery import build
        creds = get_google_credentials()
        service = build('sheets', 'v4', credentials=creds)
        sheet = service.spreadsheets()
        for campo, rango in ubicaciones.items():
            rango_col = f"{nombre_hoja}!{rango}"
            sheet.values().clear(
                spreadsheetId=spreadsheet_id,
                range=rango_col,
                body={}
            ).execute()
        for campo, rango in ubicaciones.items():
            # Extraer columna y fila inicial correctamente (ej: A2:A, B5:B, etc.)
            import re
            partes = rango.split(':')[0]
            col_match = re.match(r"^([A-Z]+)([0-9]+)?$", partes)
            if col_match:
                col = col_match.group(1)
                fila_inicial = int(col_match.group(2)) if col_match.group(2) else 1
            else:
                col = partes
                fila_inicial = 1
            values = [[data.get(campo, '')]]
            rango_celda = f"{nombre_hoja}!{col}{fila_inicial}:{col}{fila_inicial}"
            sheet.values().update(
                spreadsheetId=spreadsheet_id,
                range=rango_celda,
                valueInputOption='USER_ENTERED',
                body={'values': values}
            ).execute()
        return jsonify({'status': 'ok'})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/bdmarcas_bulk', methods=['POST'])
def bdmarcas_bulk():
    auth_header = request.headers.get('Authorization')
    if not auth_header or not auth_header.startswith('Bearer '):
        return jsonify({'error': 'No token provided'}), 401
    id_token = auth_header.split(' ')[1]
    import traceback, sys, re, io, csv, openpyxl
    try:
        decoded_token = auth.verify_id_token(id_token)
        uid = decoded_token['uid']

        hoja_ref = db.collection('Areas').document(uid).collection('Hojas').document('BDMarcas')
        hoja_doc = hoja_ref.get()
        if not hoja_doc.exists:
            return jsonify({'error': 'No se encontró la configuración de hoja/rango'}), 404
        hoja_data = hoja_doc.to_dict()
        nombre_hoja = hoja_data.get('Hoja', 'Sheet1')
        ubicaciones = {k: v for k, v in hoja_data.items() if k != 'Hoja'}
        area_doc = db.collection('Areas').document(uid).get()
        if not area_doc.exists:
            return jsonify({'error': 'No se encontró el área para este usuario'}), 404
        area_data = area_doc.to_dict()
        spreadsheet_url = area_data.get('BDMarcas')
        match = re.search(r"/d/([a-zA-Z0-9-_]+)", spreadsheet_url)
        if not match:
            return jsonify({'error': 'URL de spreadsheet inválida'}), 400
        spreadsheet_id = match.group(1)
        if 'file' not in request.files:
            return jsonify({'error': 'No file uploaded'}), 400
        file = request.files['file']
        filename = file.filename.lower()
        campos_config = list(ubicaciones.keys())
        rows = []
        if filename.endswith('.csv'):
            try:
                stream = io.StringIO(file.stream.read().decode('utf-8'))
            except Exception:
                stream = io.StringIO(file.stream.read().decode('latin-1'))
            reader = csv.DictReader(stream)
            if reader.fieldnames is None:
                return jsonify({'error': 'El archivo CSV no tiene encabezados.'}), 400
            missing = [campo for campo in campos_config if campo not in reader.fieldnames]
            if missing:
                return jsonify({'error': f'Faltan columnas requeridas: {missing}'}), 400
            rows = list(reader)
        elif filename.endswith('.xlsx'):
            wb = openpyxl.load_workbook(file, read_only=True)
            ws = wb.active
            headers = [cell.value for cell in next(ws.iter_rows(min_row=1, max_row=1))]
            missing = [campo for campo in campos_config if campo not in headers]
            if missing:
                return jsonify({'error': f'Faltan columnas requeridas: {missing}'}), 400
            for row in ws.iter_rows(min_row=2, values_only=True):
                rows.append(dict(zip(headers, row)))
        else:
            return jsonify({'error': 'Solo se permiten archivos CSV o XLSX'}), 400
        rows_preview = rows[:100]
        from googleapiclient.discovery import build
        creds = get_google_credentials()
        service = build('sheets', 'v4', credentials=creds)
        sheet = service.spreadsheets()
        for campo, rango in ubicaciones.items():
            rango_col = f"{nombre_hoja}!{rango}"
            sheet.values().clear(
                spreadsheetId=spreadsheet_id,
                range=rango_col,
                body={}
            ).execute()
        import re
        for campo, rango in ubicaciones.items():
            # Extraer columna y fila inicial correctamente (ej: G7:G)
            match = re.match(r"([A-Z]+)(\d+):", rango)
            if match:
                col = match.group(1)
                fila_inicial = int(match.group(2))
            else:
                # fallback: solo columna, empieza en 1
                col = rango.split(':')[0]
                fila_inicial = 1
            values = [[row.get(campo, '')] for row in rows]
            fila_final = fila_inicial + len(values) - 1
            sheet.values().update(
                spreadsheetId=spreadsheet_id,
                range=f"{nombre_hoja}!{col}{fila_inicial}:{col}{fila_final}",
                valueInputOption='USER_ENTERED',
                body={'values': values}
            ).execute()
        return jsonify({'status': 'ok', 'rows': len(rows), 'preview': rows_preview})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

# Endpoint de login para recibir el token de Firebase y devolver info básica del usuario
@app.route('/api/login', methods=['POST'])
def api_login():
    data = request.get_json()
    id_token = data.get('idToken')
    if not id_token:
        return jsonify({'error': 'No token provided'}), 401
    try:
        decoded_token = auth.verify_id_token(id_token)
        uid = decoded_token['uid']
        email = decoded_token.get('email', '')
        # Puedes personalizar la consulta a Firestore para obtener más info del usuario
        user_ref = db.collection('users').document(uid)
        user_doc = user_ref.get()
        user_data = user_doc.to_dict() if user_doc.exists else {}
        return jsonify({'status': 'ok', 'uid': uid, 'email': email, 'rol': user_data.get('rol', 'usuario')})
    except Exception as e:
        return jsonify({'error': 'Invalid token', 'details': str(e)}), 401


@app.route('/api/baseplus_campos', methods=['GET'])
def baseplus_campos():
    auth_header = request.headers.get('Authorization')
    if not auth_header or not auth_header.startswith('Bearer '):
        return jsonify({'error': 'No token provided'}), 401
    id_token = auth_header.split(' ')[1]
    try:
        decoded_token = auth.verify_id_token(id_token)
        uid = decoded_token['uid']
        hoja_ref = db.collection('Areas').document(uid).collection('Hojas').document('BASEPLUS')
        hoja_doc = hoja_ref.get()
        if not hoja_doc.exists:
            return jsonify({'error': 'No se encontró la configuración de hoja/rango'}), 404
        hoja_data = hoja_doc.to_dict()
        campos = [k for k in hoja_data.keys() if k != 'Hoja']
        return jsonify({'campos': campos})
    except Exception as e:
        return jsonify({'error': 'Error al obtener campos', 'details': str(e)}), 500

@app.route('/')
def home():
    return render_template('login.html')

@app.route('/panel')
def panel():
    # Ejemplo básico: datos de usuario de prueba
    user = {
        'nombre': 'Usuario',
        'email': 'usuario@email.com',
        'rol': 'usuario'
    }
    esquema_raw = 'N/A'
    uid = 'uid_de_ejemplo'
    return render_template('panel.html', user=user, esquema_raw=esquema_raw, uid=uid)

@app.route('/api/get_spreadsheet', methods=['POST'])
def get_spreadsheet():
    auth_header = request.headers.get('Authorization')
    if not auth_header or not auth_header.startswith('Bearer '):
        return jsonify({'error': 'No token provided'}), 401
    id_token = auth_header.split(' ')[1]
    try:
        decoded_token = auth.verify_id_token(id_token)
        uid = decoded_token['uid']
        user_ref = db.collection('users').document(uid)
        user_doc = user_ref.get()
        if user_doc.exists:
            user_data = user_doc.to_dict()
            spreadsheet_id = user_data.get('spreadsheetId')
            if spreadsheet_id:
                return jsonify({'spreadsheetId': spreadsheet_id})
            else:
                return jsonify({'error': 'No Software asignado aún'}), 404
        else:
            return jsonify({'error': 'Usuario no encontrado'}), 404
    except Exception as e:
        return jsonify({'error': 'Invalid token', 'details': str(e)}), 401




@app.route('/panel_opciones')
def panel_opciones():
    return render_template('panel_opciones.html')

# Nueva ruta para BASEPLUS
@app.route('/baseplus')
def baseplus():
    return render_template('BASEPLUS.html')


@app.route('/Llenar_BDS')
def llenar_bds():
    return render_template('Llenar_BDS.html')

# Nueva ruta para BD MARCAS
@app.route('/BDMARCAS')
def bdmarcas():
    return render_template('BDMARCAS.html')

@app.route('/api/userinfo', methods=['GET'])
def userinfo():
    auth_header = request.headers.get('Authorization')
    if not auth_header or not auth_header.startswith('Bearer '):
        return jsonify({'error': 'No token provided'}), 401
    id_token = auth_header.split(' ')[1]
    try:
        decoded_token = auth.verify_id_token(id_token)
        uid = decoded_token['uid']
        user_ref = db.collection('users').document(uid)
        user_doc = user_ref.get()
        if user_doc.exists:
            user_data = user_doc.to_dict()
            user_data['uid'] = uid
            return jsonify(user_data)
        else:
            return jsonify({'error': 'Usuario no encontrado'}), 404
    except Exception as e:
        return jsonify({'error': 'Invalid token', 'details': str(e)}), 401


# Endpoint para registrar datos en BASEPLUS (spreadsheets personalizados por usuario)

@app.route('/api/baseplus_registro', methods=['POST'])
def baseplus_registro():
    auth_header = request.headers.get('Authorization')
    if not auth_header or not auth_header.startswith('Bearer '):
        return jsonify({'error': 'No token provided'}), 401
    id_token = auth_header.split(' ')[1]
    try:
        import traceback
        decoded_token = auth.verify_id_token(id_token)
        uid = decoded_token['uid']
        data = request.get_json()
        print('[BASEPLUS_REGISTRO] UID:', uid)
        print('[BASEPLUS_REGISTRO] Data recibida:', data)
        # Buscar la URL del spreadsheet en la colección Areas, campo BASEPLUS
        area_doc = db.collection('Areas').document(uid).get()
        if not area_doc.exists:
            print('[BASEPLUS_REGISTRO] No se encontró el área para este usuario')
            return jsonify({'error': 'No se encontró el área para este usuario'}), 404
        area_data = area_doc.to_dict()
        spreadsheet_url = area_data.get('BASEPLUS')
        print('[BASEPLUS_REGISTRO] Spreadsheet URL:', spreadsheet_url)
        if not spreadsheet_url:
            print('[BASEPLUS_REGISTRO] No hay spreadsheet asignado para BASEPLUS')
            return jsonify({'error': 'No hay spreadsheet asignado para BASEPLUS'}), 404
        import re
        match = re.search(r"/d/([a-zA-Z0-9-_]+)", spreadsheet_url)
        if not match:
            print('[BASEPLUS_REGISTRO] URL de spreadsheet inválida')
            return jsonify({'error': 'URL de spreadsheet inválida'}), 400
        spreadsheet_id = match.group(1)
        print('[BASEPLUS_REGISTRO] Spreadsheet ID:', spreadsheet_id)

        # Leer hoja y ubicaciones desde /Areas/{uid}/Hojas/BASEPLUS
        hoja_ref = db.collection('Areas').document(uid).collection('Hojas').document('BASEPLUS')
        hoja_doc = hoja_ref.get()
        if not hoja_doc.exists:
            print('[BASEPLUS_REGISTRO] No se encontró la configuración de hoja/rango')
            return jsonify({'error': 'No se encontró la configuración de hoja/rango'}), 404
        hoja_data = hoja_doc.to_dict()
        print('[BASEPLUS_REGISTRO] Hoja data:', hoja_data)
        nombre_hoja = hoja_data.get('Hoja', 'Sheet1')
        # Quitar el campo 'Hoja' para solo dejar los campos de ubicaciones
        ubicaciones = {k: v for k, v in hoja_data.items() if k != 'Hoja'}
        print('[BASEPLUS_REGISTRO] Ubicaciones:', ubicaciones)
        # Construir los valores a insertar
        from google.oauth2 import service_account
        from googleapiclient.discovery import build
        import datetime
        SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
        SERVICE_ACCOUNT_FILE = 'supratechweb-firebase-adminsdk-fbsvc-8d4aa68a75.json'
        creds = service_account.Credentials.from_service_account_file(
            SERVICE_ACCOUNT_FILE, scopes=SCOPES)
        service = build('sheets', 'v4', credentials=creds)
        sheet = service.spreadsheets()
        fecha = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        # --- NUEVA LÓGICA: Limpiar todos los rangos configurados antes de registrar y escribir SIEMPRE en la primera fila del rango configurado ---
        print(f'[BASEPLUS_REGISTRO] Limpiando todos los rangos configurados antes de registrar...')
        for campo, rango in ubicaciones.items():
            rango_col = f"{nombre_hoja}!{rango}"
            print(f'[BASEPLUS_REGISTRO] Limpiando rango {rango_col}')
            try:
                sheet.values().clear(
                    spreadsheetId=spreadsheet_id,
                    range=rango_col,
                    body={}
                ).execute()
            except Exception as e:
                print(f'[BASEPLUS_REGISTRO] Error limpiando rango {rango_col}:', str(e))
                print(traceback.format_exc())

        # Escribir cada campo en su rango configurado individualmente
        for campo, rango in ubicaciones.items():
            valor = data.get(campo, '')
            rango_celda = f"{nombre_hoja}!{rango.split(':')[0]}"
            print(f'[BASEPLUS_REGISTRO] Insertando {campo} en {rango_celda}:', valor)
            try:
                sheet.values().update(
                    spreadsheetId=spreadsheet_id,
                    range=rango_celda,
                    valueInputOption='USER_ENTERED',
                    body={'values': [[valor]]}
                ).execute()
            except Exception as e:
                print(f'[BASEPLUS_REGISTRO] Error insertando {campo} en {rango_celda}:', str(e))
                print(traceback.format_exc())
                return jsonify({'error': f'Error insertando {campo} en {rango_celda}', 'details': str(e)}), 500
        print('[BASEPLUS_REGISTRO] Registro exitoso')
        return jsonify({'status': 'ok'})
    except Exception as e:
        import traceback
        print('[BASEPLUS_REGISTRO] Error general:', str(e))
        print(traceback.format_exc())
        return jsonify({'error': 'Error al registrar en BASEPLUS', 'details': str(e)}), 500


# --- BASEPLUS BULK ENDPOINT ---
from werkzeug.utils import secure_filename
import os
import csv
import io
import openpyxl
from google.oauth2 import service_account
from googleapiclient.discovery import build

@app.route('/api/baseplus_bulk', methods=['POST'])
def baseplus_bulk():
    auth_header = request.headers.get('Authorization')
    if not auth_header or not auth_header.startswith('Bearer '):
        return jsonify({'error': 'No token provided'}), 401
    id_token = auth_header.split(' ')[1]
    import traceback
    import sys
    def log(msg, *args):
        print('[BULK]', msg, *args)
        sys.stdout.flush()
    try:
        decoded_token = auth.verify_id_token(id_token)
        uid = decoded_token['uid']
        log('UID:', uid)
        db = firestore.Client.from_service_account_json('supratechweb-firebase-adminsdk-fbsvc-8d4aa68a75.json')
        hoja_ref = db.collection('Areas').document(uid).collection('Hojas').document('BASEPLUS')
        hoja_doc = hoja_ref.get()
        if not hoja_doc.exists:
            log('No se encontró la configuración de hoja/rango')
            return jsonify({'error': 'No se encontró la configuración de hoja/rango'}), 404
        hoja_data = hoja_doc.to_dict()
        nombre_hoja = hoja_data.get('Hoja', 'Sheet1')
        ubicaciones = {k: v for k, v in hoja_data.items() if k != 'Hoja'}
        log('Ubicaciones:', ubicaciones)
        area_doc = db.collection('Areas').document(uid).get()
        if not area_doc.exists:
            log('No se encontró el área para este usuario')
            return jsonify({'error': 'No se encontró el área para este usuario'}), 404
        area_data = area_doc.to_dict()
        spreadsheet_url = area_data.get('BASEPLUS')
        import re
        match = re.search(r"/d/([a-zA-Z0-9-_]+)", spreadsheet_url)
        if not match:
            log('URL de spreadsheet inválida:', spreadsheet_url)
            return jsonify({'error': 'URL de spreadsheet inválida'}), 400
        spreadsheet_id = match.group(1)
        if 'file' not in request.files:
            log('No file uploaded')
            return jsonify({'error': 'No file uploaded'}), 400
        file = request.files['file']
        filename = file.filename.lower()
        campos_config = list(ubicaciones.keys())
        log('Campos config:', campos_config)
        rows = []
        if filename.endswith('.csv'):
            try:
                content = file.stream.read().decode('utf-8')
            except UnicodeDecodeError as e:
                log('Fallo utf-8, intentando latin-1:', str(e))
                file.stream.seek(0)
                content = file.stream.read().decode('latin-1')
            stream = io.StringIO(content)
            reader = csv.DictReader(stream)
            log('CSV headers:', reader.fieldnames)
            if reader.fieldnames is None:
                log('El archivo CSV no tiene encabezados.')
                return jsonify({'error': 'El archivo CSV no tiene encabezados.'}), 400
            missing = [campo for campo in campos_config if campo not in reader.fieldnames]
            if missing:
                log('Faltan columnas requeridas:', missing)
                return jsonify({'error': f'Faltan columnas requeridas: {missing}'}), 400
            rows = list(reader)
            log('Primeras filas:', rows[:3])
        elif filename.endswith('.xlsx'):
            wb = openpyxl.load_workbook(file, read_only=True)
            ws = wb.active
            headers = [cell.value for cell in next(ws.iter_rows(min_row=1, max_row=1))]
            log('XLSX headers:', headers)
            missing = [campo for campo in campos_config if campo not in headers]
            if missing:
                log('Faltan columnas requeridas:', missing)
                return jsonify({'error': f'Faltan columnas requeridas: {missing}'}), 400
            for row in ws.iter_rows(min_row=2, values_only=True):
                rows.append(dict(zip(headers, row)))
            log('Primeras filas:', rows[:3])
        else:
            log('Formato de archivo no permitido:', filename)
            return jsonify({'error': 'Solo se permiten archivos CSV o XLSX'}), 400
        rows_preview = rows[:100]
        SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
        SERVICE_ACCOUNT_FILE = 'supratechweb-firebase-adminsdk-fbsvc-8d4aa68a75.json'
        creds = service_account.Credentials.from_service_account_file(
            SERVICE_ACCOUNT_FILE, scopes=SCOPES)
        service = build('sheets', 'v4', credentials=creds)
        sheet = service.spreadsheets()
        for campo, rango in ubicaciones.items():
            rango_col = f"{nombre_hoja}!{rango}"
            log(f'Limpiando rango {rango_col}')
            sheet.values().clear(
                spreadsheetId=spreadsheet_id,
                range=rango_col,
                body={}
            ).execute()
        import re
        for campo, rango in ubicaciones.items():
            # Extraer columna y fila inicial correctamente (ej: G7:G)
            match = re.match(r"([A-Z]+)(\d+):", rango)
            if match:
                col = match.group(1)
                fila_inicial = int(match.group(2))
            else:
                # fallback: solo columna, empieza en 1
                col = rango.split(':')[0]
                fila_inicial = 1
            values = [[row.get(campo, '')] for row in rows]
            fila_final = fila_inicial + len(values) - 1
            log(f'Escribiendo campo {campo} en {col}{fila_inicial}:{col}{fila_final} con valores:', values[:3])
            rango_celda = f"{nombre_hoja}!{col}{fila_inicial}:{col}{fila_final}"
            sheet.values().update(
                spreadsheetId=spreadsheet_id,
                range=rango_celda,
                valueInputOption='USER_ENTERED',
                body={'values': values}
            ).execute()
        log('Escritura masiva terminada')
        return jsonify({'status': 'ok', 'rows': len(rows), 'preview': rows_preview})
    except Exception as e:
        log('ERROR:', str(e))
        log(traceback.format_exc())
        return jsonify({'error': str(e)}), 500

# Endpoint para registrar compras en el spreadsheet del usuario
from google.oauth2 import service_account
from googleapiclient.discovery import build
import datetime

@app.route('/api/registrar_compra', methods=['POST'])
def registrar_compra():
    auth_header = request.headers.get('Authorization')
    if not auth_header or not auth_header.startswith('Bearer '):
        return jsonify({'error': 'No token provided'}), 401
    id_token = auth_header.split(' ')[1]
    try:
        decoded_token = auth.verify_id_token(id_token)
        uid = decoded_token['uid']
        # Obtener datos de la compra
        data = request.get_json()
        producto = data.get('producto')
        cantidad = data.get('cantidad')
        precio = data.get('precio')
        if not producto or not cantidad or not precio:
            return jsonify({'error': 'Datos incompletos'}), 400
        # Buscar la URL del spreadsheet en la colección Areas, campo Compras
        area_doc = db.collection('Areas').document(uid).get()
        if not area_doc.exists:
            return jsonify({'error': 'No se encontró el área para este usuario'}), 404
        area_data = area_doc.to_dict()
        spreadsheet_url = area_data.get('Compras')
        if not spreadsheet_url:
            return jsonify({'error': 'No hay spreadsheet asignado para Compras'}), 404
        # Extraer el ID del spreadsheet desde la URL
        import re
        match = re.search(r"/d/([a-zA-Z0-9-_]+)", spreadsheet_url)
        if not match:
            return jsonify({'error': 'URL de spreadsheet inválida'}), 400
        spreadsheet_id = match.group(1)

        # Leer hoja y rango desde /Areas/{uid}/Hojas/Desgloce
        desgloce_ref = db.collection('Areas').document(uid).collection('Hojas').document('Desgloce')
        desgloce_doc = desgloce_ref.get()
        if not desgloce_doc.exists:
            return jsonify({'error': 'No se encontró la configuración de hoja/rango'}), 404
        desgloce_data = desgloce_doc.to_dict()
        nombre_hoja = desgloce_data.get('Hoja', 'Sheet1')
        rango = desgloce_data.get('Rango', 'A1')
        # El rango final será "{nombre_hoja}!{rango}", por ejemplo "Compras2!A2:B"
        rango_final = f"{nombre_hoja}!{rango}"

        # Autenticación con Google Sheets API
        SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
        SERVICE_ACCOUNT_FILE = 'supratechweb-firebase-adminsdk-fbsvc-8d4aa68a75.json'
        creds = service_account.Credentials.from_service_account_file(
            SERVICE_ACCOUNT_FILE, scopes=SCOPES)
        service = build('sheets', 'v4', credentials=creds)
        sheet = service.spreadsheets()
        # Registrar la compra (agregar fila)
        fecha = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        # Solo enviar los campos que correspondan al rango
        # Si el rango es A2:B, solo enviar dos columnas: producto y cantidad
        columnas = rango.split(':')
        num_cols = 1
        if len(columnas) == 2:
            # Calcular número de columnas entre A y B, etc.
            def col_to_num(col):
                num = 0
                for c in col:
                    num = num * 26 + (ord(c.upper()) - ord('A') + 1)
                return num
            num_cols = col_to_num(columnas[1].rstrip('0123456789')) - col_to_num(columnas[0].rstrip('0123456789')) + 1
        # Preparar los valores según el número de columnas
        all_values = [fecha, producto, cantidad, precio]
        values = [all_values[:num_cols]]
        body = {'values': values}
        result = sheet.values().append(
            spreadsheetId=spreadsheet_id,
            range=rango_final,
            valueInputOption='USER_ENTERED',
            insertDataOption='INSERT_ROWS',
            body=body
        ).execute()
        return jsonify({'status': 'ok', 'updatedRange': result.get('updates', {}).get('updatedRange', '')})
    except Exception as e:
        return jsonify({'error': 'Error al registrar la compra', 'details': str(e)}), 500

# Nueva ruta para STOCK
@app.route('/stock')
def stock():
    return render_template('Stock.html')

@app.route('/api/stock', methods=['GET'])
def stock_api():
    auth_header = request.headers.get('Authorization')
    if not auth_header or not auth_header.startswith('Bearer '):
        return jsonify({'error': 'No token provided'}), 401
    id_token = auth_header.split(' ')[1]
    try:
        decoded_token = auth.verify_id_token(id_token)
        uid = decoded_token['uid']
        hoja_ref = db.collection('Areas').document(uid).collection('Hojas').document('STOCK')
        hoja_doc = hoja_ref.get()
        if not hoja_doc.exists:
            return jsonify({'error': 'No se encontró la configuración de hoja/rango'}), 404
        hoja_data = hoja_doc.to_dict()
        campos = [k for k in hoja_data.keys() if k != 'Hoja']
        return jsonify({'campos': campos})
    except Exception as e:
        return jsonify({'error': 'Error al obtener campos', 'details': str(e)}), 500

@app.route('/api/stock_campos', methods=['GET'])
def stock_campos():
    auth_header = request.headers.get('Authorization')
    if not auth_header or not auth_header.startswith('Bearer '):
        return jsonify({'error': 'No token provided'}), 401
    id_token = auth_header.split(' ')[1]
    try:
        decoded_token = auth.verify_id_token(id_token)
        uid = decoded_token['uid']
        hoja_ref = db.collection('Areas').document(uid).collection('Hojas').document('Stock')  # Igual que en Firebase
        hoja_doc = hoja_ref.get()
        if not hoja_doc.exists:
            return jsonify({'error': 'No se encontró la configuración de hoja/rango'}), 404
        hoja_data = hoja_doc.to_dict()
        campos = [k for k in hoja_data.keys() if k != 'Hoja']
        return jsonify({'campos': campos})
    except Exception as e:
        return jsonify({'error': 'Error al obtener campos', 'details': str(e)}), 500


# --- STOCK BULK ENDPOINT ---
from werkzeug.utils import secure_filename
import os
import csv
import io
import openpyxl
from google.oauth2 import service_account
from googleapiclient.discovery import build

@app.route('/api/stock_bulk', methods=['POST'])
def stock_bulk():
    auth_header = request.headers.get('Authorization')
    if not auth_header or not auth_header.startswith('Bearer '):
        return jsonify({'error': 'No token provided'}), 401
    id_token = auth_header.split(' ')[1]
    import traceback
    try:
        decoded_token = auth.verify_id_token(id_token)
        uid = decoded_token['uid']
        import sys
        def log(msg, *args):
            print('[STOCK_BULK]', msg, *args)
            sys.stdout.flush()
        log('UID:', uid)

        hoja_ref = db.collection('Areas').document(uid).collection('Hojas').document('Stock')
        hoja_doc = hoja_ref.get()
        if not hoja_doc.exists:
            log('No se encontró la configuración de hoja/rango')
            return jsonify({'error': 'No se encontró la configuración de hoja/rango'}), 404
        hoja_data = hoja_doc.to_dict()
        nombre_hoja = hoja_data.get('Hoja', 'Sheet1')
        ubicaciones = {k: v for k, v in hoja_data.items() if k != 'Hoja'}
        log('Ubicaciones:', ubicaciones)
        area_doc = db.collection('Areas').document(uid).get()
        if not area_doc.exists:
            log('No se encontró el área para este usuario')
            return jsonify({'error': 'No se encontró el área para este usuario'}), 404
        area_data = area_doc.to_dict()
        spreadsheet_url = area_data.get('Stock')
        import re
        match = re.search(r"/d/([a-zA-Z0-9-_]+)", spreadsheet_url)
        if not match:
            log('URL de spreadsheet inválida:', spreadsheet_url)
            return jsonify({'error': 'URL de spreadsheet inválida'}), 400
        spreadsheet_id = match.group(1)
        if 'file' not in request.files:
            log('No file uploaded')
            return jsonify({'error': 'No file uploaded'}), 400
        file = request.files['file']
        filename = file.filename.lower()
        campos_config = list(ubicaciones.keys())
        log('Campos config:', campos_config)
        rows = []
        if filename.endswith('.csv'):
            try:
                content = file.stream.read().decode('utf-8')
            except UnicodeDecodeError as e:
                log('Fallo utf-8, intentando latin-1:', str(e))
                file.stream.seek(0)
                content = file.stream.read().decode('latin-1')
            stream = io.StringIO(content)
            reader = csv.DictReader(stream)
            log('CSV headers:', reader.fieldnames)
            if reader.fieldnames is None:
                log('El archivo CSV no tiene encabezados.')
                return jsonify({'error': 'El archivo CSV no tiene encabezados.'}), 400
            missing = [campo for campo in campos_config if campo not in reader.fieldnames]
            if missing:
                log('Faltan columnas requeridas:', missing)
                return jsonify({'error': f'Faltan columnas requeridas: {missing}'}), 400
            rows = list(reader)
            log('Primeras filas:', rows[:3])
        elif filename.endswith('.xlsx'):
            wb = openpyxl.load_workbook(file, read_only=True)
            ws = wb.active
            headers = [cell.value for cell in next(ws.iter_rows(min_row=1, max_row=1))]
            log('XLSX headers:', headers)
            missing = [campo for campo in campos_config if campo not in headers]
            if missing:
                log('Faltan columnas requeridas:', missing)
                return jsonify({'error': f'Faltan columnas requeridas: {missing}'}), 400
            for row in ws.iter_rows(min_row=2, values_only=True):
                rows.append(dict(zip(headers, row)))
            log('Primeras filas:', rows[:3])
        else:
            log('Formato de archivo no permitido:', filename)
            return jsonify({'error': 'Solo se permiten archivos CSV o XLSX'}), 400
        rows_preview = rows[:100]
        SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
        SERVICE_ACCOUNT_FILE = 'supratechweb-firebase-adminsdk-fbsvc-8d4aa68a75.json'
        creds = service_account.Credentials.from_service_account_file(
            SERVICE_ACCOUNT_FILE, scopes=SCOPES)
        service = build('sheets', 'v4', credentials=creds)
        sheet = service.spreadsheets()
        # Limpiar todos los rangos configurados antes de registrar y escribir SIEMPRE en la fila 2
        import re
        for campo, rango in ubicaciones.items():
            # Extraer columna correctamente (ej: A22:A, B5:B, etc.) y forzar fila 2
            match = re.match(r"([A-Z]+)", rango)
            if match:
                col = match.group(1)
            else:
                col = rango.split(':')[0]
            clear_range = f"{nombre_hoja}!{col}2:{col}1000"
            log(f'Limpiando rango {clear_range}')
            sheet.values().clear(
                spreadsheetId=spreadsheet_id,
                range=clear_range,
                body={}
            ).execute()
        for campo, rango in ubicaciones.items():
            # Extraer columna correctamente y forzar fila 2
            match = re.match(r"([A-Z]+)", rango)
            if match:
                col = match.group(1)
            else:
                col = rango.split(':')[0]
            fila = 2  # Siempre iniciar en la fila 2
            values = [[row.get(campo, '')] for row in rows]
            fila_final = fila + len(values) - 1
            log(f'Escribiendo campo {campo} en {col}{fila}:{col}{fila_final} con valores:', values[:3])
            rango_celda = f"{nombre_hoja}!{col}{fila}:{col}{fila_final}"
            sheet.values().update(
                spreadsheetId=spreadsheet_id,
                range=rango_celda,
                valueInputOption='USER_ENTERED',
                body={'values': values}
            ).execute()
        log('Escritura masiva terminada')
        return jsonify({'status': 'ok', 'rows': len(rows), 'preview': rows_preview})
    except Exception as e:
        log('ERROR:', str(e))
        log(traceback.format_exc())
        return jsonify({'error': str(e)}), 500

# --- STOCK REGISTRO ENDPOINT ---
@app.route('/api/stock_registro', methods=['POST'])
def stock_registro():
    auth_header = request.headers.get('Authorization')
    if not auth_header or not auth_header.startswith('Bearer '):
        return jsonify({'error': 'No token provided'}), 401
    id_token = auth_header.split(' ')[1]
    try:
        import traceback
        decoded_token = auth.verify_id_token(id_token)
        uid = decoded_token['uid']
        data = request.get_json()
        print('[STOCK_REGISTRO] UID:', uid)
        print('[STOCK_REGISTRO] Data recibida:', data)
        # Buscar la URL del spreadsheet en la colección Areas, campo Stock
        area_doc = db.collection('Areas').document(uid).get()
        if not area_doc.exists:
            print('[STOCK_REGISTRO] No se encontró el área para este usuario')
            return jsonify({'error': 'No se encontró el área para este usuario'}), 404
        area_data = area_doc.to_dict()
        spreadsheet_url = area_data.get('Stock')
        print('[STOCK_REGISTRO] Spreadsheet URL:', spreadsheet_url)
        if not spreadsheet_url:
            print('[STOCK_REGISTRO] No hay spreadsheet asignado para Stock')
            return jsonify({'error': 'No hay spreadsheet asignado para Stock'}), 404
        import re
        match = re.search(r"/d/([a-zA-Z0-9-_]+)", spreadsheet_url)
        if not match:
            print('[STOCK_REGISTRO] URL de spreadsheet inválida')
            return jsonify({'error': 'URL de spreadsheet inválida'}), 400
        spreadsheet_id = match.group(1)
        print('[STOCK_REGISTRO] Spreadsheet ID:', spreadsheet_id)

        # Leer hoja y ubicaciones desde /Areas/{uid}/Hojas/Stock
        hoja_ref = db.collection('Areas').document(uid).collection('Hojas').document('Stock')
        hoja_doc = hoja_ref.get()
        if not hoja_doc.exists:
            print('[STOCK_REGISTRO] No se encontró la configuración de hoja/rango')
            return jsonify({'error': 'No se encontró la configuración de hoja/rango'}), 404
        hoja_data = hoja_doc.to_dict()
        print('[STOCK_REGISTRO] Hoja data:', hoja_data)
        nombre_hoja = hoja_data.get('Hoja', 'Sheet1')
        ubicaciones = {k: v for k, v in hoja_data.items() if k != 'Hoja'}
        print('[STOCK_REGISTRO] Ubicaciones:', ubicaciones)
        from google.oauth2 import service_account
        from googleapiclient.discovery import build
        import datetime
        SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
        SERVICE_ACCOUNT_FILE = 'supratechweb-firebase-adminsdk-fbsvc-8d4aa68a75.json'
        creds = service_account.Credentials.from_service_account_file(
            SERVICE_ACCOUNT_FILE, scopes=SCOPES)
        service = build('sheets', 'v4', credentials=creds)
        sheet = service.spreadsheets()
        fecha = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        print(f'[STOCK_REGISTRO] Limpiando todos los rangos configurados antes de registrar...')
        for campo, rango in ubicaciones.items():
            rango_col = f"{nombre_hoja}!{rango}"
            print(f'[STOCK_REGISTRO] Limpiando rango {rango_col}')
            try:
                sheet.values().clear(
                    spreadsheetId=spreadsheet_id,
                    range=rango_col,
                    body={}
                ).execute()
            except Exception as e:
                print(f'[STOCK_REGISTRO] Error limpiando rango {rango_col}:', str(e))
                print(traceback.format_exc())
        for campo, rango in ubicaciones.items():
            valor = data.get(campo, '')
            # Extraer columna (ej: A2:A, B5:B, etc.) y forzar fila 2
            import re
            partes = rango.split(':')[0]
            col_match = re.match(r"^([A-Z]+)([0-9]+)?$", partes)
            if col_match:
                col = col_match.group(1)
            else:
                col = partes
            fila_inicial = 2
            rango_celda = f"{nombre_hoja}!{col}{fila_inicial}:{col}{fila_inicial}"
            print(f'[STOCK_REGISTRO] Insertando {campo} en {rango_celda}:', valor)
            try:
                sheet.values().update(
                    spreadsheetId=spreadsheet_id,
                    range=rango_celda,
                    valueInputOption='USER_ENTERED',
                    body={'values': [[valor]]}
                ).execute()
            except Exception as e:
                print(f'[STOCK_REGISTRO] Error insertando {campo} en {rango_celda}:', str(e))
                print(traceback.format_exc())
                return jsonify({'error': f'Error insertando {campo} en {rango_celda}', 'details': str(e)}), 500
        print('[STOCK_REGISTRO] Registro exitoso')
        return jsonify({'status': 'ok'})
    except Exception as e:
        import traceback
        print('[STOCK_REGISTRO] Error general:', str(e))
        print(traceback.format_exc())
        return jsonify({'error': 'Error al registrar en Stock', 'details': str(e)}), 500

# Nueva ruta para VENTAS
@app.route('/ventas')
def ventas():
    return render_template('Ventas.html')

@app.route('/api/ventas_campos', methods=['GET'])
def ventas_campos():
    auth_header = request.headers.get('Authorization')
    if not auth_header or not auth_header.startswith('Bearer '):
        return jsonify({'error': 'No token provided'}), 401
    id_token = auth_header.split(' ')[1]
    try:
        decoded_token = auth.verify_id_token(id_token)
        uid = decoded_token['uid']
        hoja_ref = db.collection('Areas').document(uid).collection('Hojas').document('Ventas')
        hoja_doc = hoja_ref.get()
        if not hoja_doc.exists:
            return jsonify({'error': 'No se encontró la configuración de hoja/rango'}), 404
        hoja_data = hoja_doc.to_dict()
        campos = [k for k in hoja_data.keys() if k != 'Hoja']
        return jsonify({'campos': campos})
    except Exception as e:
        return jsonify({'error': 'Error al obtener campos', 'details': str(e)}), 500

# --- VENTAS REGISTRO ENDPOINT ---
@app.route('/api/ventas_registro', methods=['POST'])
def ventas_registro():
    auth_header = request.headers.get('Authorization')
    if not auth_header or not auth_header.startswith('Bearer '):
        return jsonify({'error': 'No token provided'}), 401
    id_token = auth_header.split(' ')[1]
    try:
        import datetime, re, traceback
        decoded_token = auth.verify_id_token(id_token)
        uid = decoded_token['uid']
        data = request.get_json()
        print('[VENTAS_REGISTRO] UID:', uid)
        print('[VENTAS_REGISTRO] Data recibida:', data)
        # Buscar la URL del spreadsheet en la colección Areas, campo Ventas
        area_doc = db.collection('Areas').document(uid).get()
        if not area_doc.exists:
            return jsonify({'error': 'No se encontró el área para este usuario'}), 404
        area_data = area_doc.to_dict()
        spreadsheet_url = area_data.get('Ventas')
        print('[VENTAS_REGISTRO] Spreadsheet URL:', spreadsheet_url)
        if not spreadsheet_url:
            return jsonify({'error': 'No se encontró la URL del spreadsheet'}), 404
        import re
        match = re.search(r"/d/([a-zA-Z0-9-_]+)", spreadsheet_url)
        if not match:
            return jsonify({'error': 'URL de spreadsheet inválida'}), 400
        spreadsheet_id = match.group(1)
        print('[VENTAS_REGISTRO] Spreadsheet ID:', spreadsheet_id)

        # Leer hoja y ubicaciones desde /Areas/{uid}/Hojas/Ventas
        hoja_ref = db.collection('Areas').document(uid).collection('Hojas').document('Ventas')
        hoja_doc = hoja_ref.get()
        if not hoja_doc.exists:
            return jsonify({'error': 'No se encontró la configuración de hoja/rango'}), 404
        hoja_data = hoja_doc.to_dict()
        print('[VENTAS_REGISTRO] Hoja data:', hoja_data)
        nombre_hoja = hoja_data.get('Hoja', 'Sheet1')
        ubicaciones = {k: v for k, v in hoja_data.items() if k != 'Hoja'}
        print('[VENTAS_REGISTRO] Ubicaciones:', ubicaciones)
        from google.oauth2 import service_account
        from googleapiclient.discovery import build
        import datetime
        SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
        SERVICE_ACCOUNT_FILE = 'supratechweb-firebase-adminsdk-fbsvc-8d4aa68a75.json'
        creds = service_account.Credentials.from_service_account_file(
            SERVICE_ACCOUNT_FILE, scopes=SCOPES)
        service = build('sheets', 'v4', credentials=creds)
        sheet = service.spreadsheets()
        fecha = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        print(f'[VENTAS_REGISTRO] Limpiando todos los rangos configurados antes de registrar...')
        for campo, rango in ubicaciones.items():
            # Extraer columna correctamente y forzar fila 2
            match = re.match(r"([A-Z]+)", rango)
            if match:
                col = match.group(1)
            else:
                col = rango.split(':')[0]
            clear_range = f"{nombre_hoja}!{col}2:{col}1000"
            print(f'[VENTAS_REGISTRO] Limpiando rango {clear_range}')
            sheet.values().clear(
                spreadsheetId=spreadsheet_id,
                range=clear_range,
                body={}
            ).execute()
        for campo, rango in ubicaciones.items():
            # Extraer columna correctamente y forzar fila 2
            match = re.match(r"([A-Z]+)", rango)
            if match:
                col = match.group(1)
            else:
                col = rango.split(':')[0]
            valor = data.get(campo, '')
            if campo.lower() == 'fecha':
                valor = fecha
            values = [[valor]]
            rango_celda = f"{nombre_hoja}!{col}2:{col}2"
            print(f'[VENTAS_REGISTRO] Escribiendo campo {campo} en {rango_celda} con valor: {valor}')
            try:
                sheet.values().update(
                    spreadsheetId=spreadsheet_id,
                    range=rango_celda,
                    valueInputOption='USER_ENTERED',
                    body={'values': values}
                ).execute()
            except Exception as e:
                print(f'[VENTAS_REGISTRO] Error insertando {campo} en {rango_celda}:', str(e))
                print(traceback.format_exc())
                return jsonify({'error': f'Error insertando {campo} en {rango_celda}', 'details': str(e)}), 500
        print('[VENTAS_REGISTRO] Registro exitoso')
        return jsonify({'status': 'ok'})
    except Exception as e:
        import traceback
        print('[VENTAS_REGISTRO] Error general:', str(e))
        print(traceback.format_exc())
        return jsonify({'error': 'Error al registrar en Ventas', 'details': str(e)}), 500

# --- VENTAS BULK ENDPOINT ---
@app.route('/api/ventas_bulk', methods=['POST'])
def ventas_bulk():
    auth_header = request.headers.get('Authorization')
    if not auth_header or not auth_header.startswith('Bearer '):
        return jsonify({'error': 'No token provided'}), 401
    id_token = auth_header.split(' ')[1]
    import traceback
    try:
        decoded_token = auth.verify_id_token(id_token)
        uid = decoded_token['uid']
        import sys
        def log(msg, *args):
            print('[VENTAS_BULK]', msg, *args)
            sys.stdout.flush()
        log('UID:', uid)

        hoja_ref = db.collection('Areas').document(uid).collection('Hojas').document('Ventas')
        hoja_doc = hoja_ref.get()
        if not hoja_doc.exists:
            log('No se encontró la configuración de hoja/rango')
            return jsonify({'error': 'No se encontró la configuración de hoja/rango'}), 404
        hoja_data = hoja_doc.to_dict()
        nombre_hoja = hoja_data.get('Hoja', 'Sheet1')
        ubicaciones = {k: v for k, v in hoja_data.items() if k != 'Hoja'}
        log('Ubicaciones:', ubicaciones)
        area_doc = db.collection('Areas').document(uid).get()
        if not area_doc.exists:
            log('No se encontró el área para este usuario')
            return jsonify({'error': 'No se encontró el área para este usuario'}), 404
        area_data = area_doc.to_dict()
        spreadsheet_url = area_data.get('Ventas')
        if not spreadsheet_url:
            log('No se encontró la URL del spreadsheet para Ventas')
            return jsonify({'error': 'No se encontró la URL del spreadsheet para Ventas. Configura el campo "Ventas" en Firestore.'}), 404
        import re
        match = re.search(r"/d/([a-zA-Z0-9-_]+)", spreadsheet_url)
        if not match:
            log('URL de spreadsheet inválida:', spreadsheet_url)
            return jsonify({'error': 'URL de spreadsheet inválida'}), 400
        spreadsheet_id = match.group(1)
        if 'file' not in request.files:
            log('No file uploaded')
            return jsonify({'error': 'No file uploaded'}), 400
        file = request.files['file']
        filename = file.filename.lower()
        campos_config = list(ubicaciones.keys())
        log('Campos config:', campos_config)
        rows = []
        if filename.endswith('.csv'):
            import csv, io
            try:
                content = file.stream.read().decode('utf-8')
            except UnicodeDecodeError as e:
                log('Fallo utf-8, intentando latin-1:', str(e))
                file.stream.seek(0)
                content = file.stream.read().decode('latin-1')
            stream = io.StringIO(content)
            reader = csv.DictReader(stream)
            log('CSV headers:', reader.fieldnames)
            if reader.fieldnames is None:
                log('El archivo CSV no tiene encabezados.')
                return jsonify({'error': 'El archivo CSV no tiene encabezados.'}), 400
            missing = [campo for campo in campos_config if campo not in reader.fieldnames]
            if missing:
                log('Faltan columnas requeridas:', missing)
                return jsonify({'error': f'Faltan columnas requeridas: {missing}'}), 400
            rows = list(reader)
            log('Primeras filas:', rows[:3])
        elif filename.endswith('.xlsx'):
            import openpyxl
            wb = openpyxl.load_workbook(file, read_only=True)
            ws = wb.active
            headers = [cell.value for cell in next(ws.iter_rows(min_row=1, max_row=1))]
            log('XLSX headers:', headers)
            missing = [campo for campo in campos_config if campo not in headers]
            if missing:
                log('Faltan columnas requeridas:', missing)
                return jsonify({'error': f'Faltan columnas requeridas: {missing}'}), 400
            for row in ws.iter_rows(min_row=2, values_only=True):
                rows.append(dict(zip(headers, row)))
            log('Primeras filas:', rows[:3])
        else:
            log('Formato de archivo no permitido:', filename)
            return jsonify({'error': 'Solo se permiten archivos CSV o XLSX'}), 400
        rows_preview = rows[:100]
        from googleapiclient.discovery import build
        creds = get_google_credentials()
        service = build('sheets', 'v4', credentials=creds)
        sheet = service.spreadsheets()
        # Limpiar todos los rangos configurados antes de registrar y escribir SIEMPRE en la fila 2
        import re
        for campo, rango in ubicaciones.items():
            # Extraer columna correctamente (ej: A22:A, B5:B, etc.) y forzar fila 2
            match = re.match(r"([A-Z]+)", rango)
            if match:
                col = match.group(1)
            else:
                col = rango.split(':')[0]
            clear_range = f"{nombre_hoja}!{col}2:{col}1000"
            log(f'Limpiando rango {clear_range}')
            sheet.values().clear(
                spreadsheetId=spreadsheet_id,
                range=clear_range,
                body={}
            ).execute()
        for campo, rango in ubicaciones.items():
            # Extraer columna correctamente y forzar fila 2
            match = re.match(r"([A-Z]+)", rango)
            if match:
                col = match.group(1)
            else:
                col = rango.split(':')[0]
            fila = 2  # Siempre iniciar en la fila 2
            values = [[row.get(campo, '')] for row in rows]
            fila_final = fila + len(values) - 1
            log(f'Escribiendo campo {campo} en {col}{fila}:{col}{fila_final} con valores:', values[:3])
            rango_celda = f"{nombre_hoja}!{col}{fila}:{col}{fila_final}"
            sheet.values().update(
                spreadsheetId=spreadsheet_id,
                range=rango_celda,
                valueInputOption='USER_ENTERED',
                body={'values': values}
            ).execute()
        log('Escritura masiva terminada')
        return jsonify({'status': 'ok', 'rows': len(rows), 'preview': rows_preview})
    except Exception as e:
        log('ERROR:', str(e))
        log(traceback.format_exc())
        return jsonify({'error': str(e)}), 500

# Nueva ruta para BLACKLIST
@app.route('/blacklist')
def blacklist():
    return render_template('Blacklist.html')

# Nueva ruta para PEDIDOS 9.0
@app.route('/pedidos')
def pedidos():
    return render_template('Pedidos.html')

# Rutas para submódulos de Pedidos 9.0
@app.route('/pedidos/calendario')
def pedidos_calendario():
    return render_template('Pedidos_Calendario.html')

@app.route('/pedidos/bd')
def pedidos_bd():
    return render_template('Pedidos_BD.html')

@app.route('/pedidos/bdqty')
def pedidos_bdqty():
    return render_template('Pedidos_BDQTY.html')

@app.route('/pedidos/clasificaciones')
def pedidos_clasificaciones():
    return render_template('Pedidos_Clasificaciones.html')

@app.route('/api/blacklist_campos', methods=['GET'])
def blacklist_campos():
    auth_header = request.headers.get('Authorization')
    if not auth_header or not auth_header.startswith('Bearer '):
        return jsonify({'error': 'No token provided'}), 401
    id_token = auth_header.split(' ')[1]
    try:
        decoded_token = auth.verify_id_token(id_token)
        uid = decoded_token['uid']
        hoja_ref = db.collection('Areas').document(uid).collection('Hojas').document('Blacklist')
        hoja_doc = hoja_ref.get()
        if not hoja_doc.exists:
            return jsonify({'error': 'No se encontró la configuración de hoja/rango'}), 404
        hoja_data = hoja_doc.to_dict()
        campos = [k for k in hoja_data.keys() if k != 'Hoja']
        return jsonify({'campos': campos})
    except Exception as e:
        return jsonify({'error': 'Error al obtener campos', 'details': str(e)}), 500

# --- BLACKLIST REGISTRO ENDPOINT ---
@app.route('/api/blacklist_registro', methods=['POST'])
def blacklist_registro():
    auth_header = request.headers.get('Authorization')
    if not auth_header or not auth_header.startswith('Bearer '):
        return jsonify({'error': 'No token provided'}), 401
    id_token = auth_header.split(' ')[1]
    try:
        import datetime, re, traceback
        decoded_token = auth.verify_id_token(id_token)
        uid = decoded_token['uid']
        data = request.get_json()
        print('[BLACKLIST_REGISTRO] UID:', uid)
        print('[BLACKLIST_REGISTRO] Data recibida:', data)
        area_doc = db.collection('Areas').document(uid).get()
        if not area_doc.exists:
            return jsonify({'error': 'No se encontró el área para este usuario'}), 404
        area_data = area_doc.to_dict()
        spreadsheet_url = area_data.get('Blacklist')
        print('[BLACKLIST_REGISTRO] Spreadsheet URL:', spreadsheet_url)
        if not spreadsheet_url:
            return jsonify({'error': 'No se encontró la URL del spreadsheet'}), 404
        import re
        match = re.search(r"/d/([a-zA-Z0-9-_]+)", spreadsheet_url)
        if not match:
            return jsonify({'error': 'URL de spreadsheet inválida'}), 400
        spreadsheet_id = match.group(1)
        print('[BLACKLIST_REGISTRO] Spreadsheet ID:', spreadsheet_id)
        hoja_ref = db.collection('Areas').document(uid).collection('Hojas').document('Blacklist')
        hoja_doc = hoja_ref.get()
        if not hoja_doc.exists:
            return jsonify({'error': 'No se encontró la configuración de hoja/rango'}), 404
        hoja_data = hoja_doc.to_dict()
        print('[BLACKLIST_REGISTRO] Hoja data:', hoja_data)
        nombre_hoja = hoja_data.get('Hoja', 'Sheet1')
        ubicaciones = {k: v for k, v in hoja_data.items() if k != 'Hoja'}
        print('[BLACKLIST_REGISTRO] Ubicaciones:', ubicaciones)
        from google.oauth2 import service_account
        from googleapiclient.discovery import build
        import datetime
        SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
        SERVICE_ACCOUNT_FILE = 'supratechweb-firebase-adminsdk-fbsvc-8d4aa68a75.json'
        creds = service_account.Credentials.from_service_account_file(
            SERVICE_ACCOUNT_FILE, scopes=SCOPES)
        service = build('sheets', 'v4', credentials=creds)
        sheet = service.spreadsheets()
        fecha = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        
        # Encontrar la siguiente fila vacía usando 'Marca' como referencia
        primer_campo = 'Marca' if 'Marca' in ubicaciones else list(ubicaciones.keys())[0]
        primer_rango = ubicaciones[primer_campo]
        match = re.match(r"([A-Z]+)", primer_rango)
        if match:
            col = match.group(1)
        else:
            col = primer_rango.split(':')[0]
        
        rango_lectura = f"{nombre_hoja}!{col}2:{col}"
        result = sheet.values().get(
            spreadsheetId=spreadsheet_id,
            range=rango_lectura
        ).execute()
        valores_existentes = result.get('values', [])
        siguiente_fila = len(valores_existentes) + 2
        print(f'[BLACKLIST_REGISTRO] Siguiente fila disponible: {siguiente_fila}')
        
        # Escribir en la siguiente fila vacía
        for campo, rango in ubicaciones.items():
            match = re.match(r"([A-Z]+)", rango)
            if match:
                col = match.group(1)
            else:
                col = rango.split(':')[0]
            valor = data.get(campo, '')
            if campo.lower() == 'fecha':
                valor = fecha
            values = [[valor]]
            rango_celda = f"{nombre_hoja}!{col}{siguiente_fila}:{col}{siguiente_fila}"
            print(f'[BLACKLIST_REGISTRO] Escribiendo campo {campo} en {rango_celda} con valor: {valor}')
            try:
                sheet.values().update(
                    spreadsheetId=spreadsheet_id,
                    range=rango_celda,
                    valueInputOption='USER_ENTERED',
                    body={'values': values}
                ).execute()
            except Exception as e:
                print(f'[BLACKLIST_REGISTRO] Error insertando {campo} en {rango_celda}:', str(e))
                print(traceback.format_exc())
                return jsonify({'error': f'Error insertando {campo} en {rango_celda}', 'details': str(e)}), 500
        print('[BLACKLIST_REGISTRO] Registro exitoso')
        return jsonify({'status': 'ok'})
    except Exception as e:
        import traceback
        print('[BLACKLIST_REGISTRO] Error general:', str(e))
        print(traceback.format_exc())
        return jsonify({'error': 'Error al registrar en Blacklist', 'details': str(e)}), 500

# --- BLACKLIST BULK ENDPOINT ---
@app.route('/api/blacklist_bulk', methods=['POST'])
def blacklist_bulk():
    auth_header = request.headers.get('Authorization')
    if not auth_header or not auth_header.startswith('Bearer '):
        return jsonify({'error': 'No token provided'}), 401
    id_token = auth_header.split(' ')[1]
    import traceback
    try:
        decoded_token = auth.verify_id_token(id_token)
        uid = decoded_token['uid']
        import sys
        def log(msg, *args):
            print('[BLACKLIST_BULK]', msg, *args)
            sys.stdout.flush()
        log('UID:', uid)

        hoja_ref = db.collection('Areas').document(uid).collection('Hojas').document('Blacklist')
        hoja_doc = hoja_ref.get()
        if not hoja_doc.exists:
            log('No se encontró la configuración de hoja/rango')
            return jsonify({'error': 'No se encontró la configuración de hoja/rango'}), 404
        hoja_data = hoja_doc.to_dict()
        nombre_hoja = hoja_data.get('Hoja', 'Sheet1')
        ubicaciones = {k: v for k, v in hoja_data.items() if k != 'Hoja'}
        log('Ubicaciones:', ubicaciones)
        area_doc = db.collection('Areas').document(uid).get()
        if not area_doc.exists:
            log('No se encontró el área para este usuario')
            return jsonify({'error': 'No se encontró el área para este usuario'}), 404
        area_data = area_doc.to_dict()
        spreadsheet_url = area_data.get('Blacklist')
        if not spreadsheet_url:
            log('No se encontró la URL del spreadsheet para Blacklist')
            return jsonify({'error': 'No se encontró la URL del spreadsheet para Blacklist. Configura el campo "Blacklist" en Firestore.'}), 404
        import re
        match = re.search(r"/d/([a-zA-Z0-9-_]+)", spreadsheet_url)
        if not match:
            log('URL de spreadsheet inválida:', spreadsheet_url)
            return jsonify({'error': 'URL de spreadsheet inválida'}), 400
        spreadsheet_id = match.group(1)
        if 'file' not in request.files:
            log('No file uploaded')
            return jsonify({'error': 'No file uploaded'}), 400
        file = request.files['file']
        filename = file.filename.lower()
        campos_config = list(ubicaciones.keys())
        log('Campos config:', campos_config)
        rows = []
        if filename.endswith('.csv'):
            import csv, io
            try:
                content = file.stream.read().decode('utf-8-sig')  # utf-8-sig elimina BOM automáticamente
            except UnicodeDecodeError as e:
                log('Fallo utf-8, intentando latin-1:', str(e))
                file.stream.seek(0)
                content = file.stream.read().decode('latin-1')
            stream = io.StringIO(content)
            reader = csv.DictReader(stream)
            log('CSV headers:', reader.fieldnames)
            if reader.fieldnames is None:
                log('El archivo CSV no tiene encabezados.')
                return jsonify({'error': 'El archivo CSV no tiene encabezados.'}), 400
            missing = [campo for campo in campos_config if campo not in reader.fieldnames]
            if missing:
                log('Faltan columnas requeridas:', missing)
                return jsonify({'error': f'Faltan columnas requeridas: {missing}'}), 400
            rows = list(reader)
            log('Primeras filas:', rows[:3])
        elif filename.endswith('.xlsx'):
            import openpyxl
            wb = openpyxl.load_workbook(file, read_only=True)
            ws = wb.active
            headers = [cell.value for cell in next(ws.iter_rows(min_row=1, max_row=1))]
            log('XLSX headers:', headers)
            missing = [campo for campo in campos_config if campo not in headers]
            if missing:
                log('Faltan columnas requeridas:', missing)
                return jsonify({'error': f'Faltan columnas requeridas: {missing}'}), 400
            for row in ws.iter_rows(min_row=2, values_only=True):
                rows.append(dict(zip(headers, row)))
            log('Primeras filas:', rows[:3])
        else:
            log('Formato de archivo no permitido:', filename)
            return jsonify({'error': 'Solo se permiten archivos CSV o XLSX'}), 400
        rows_preview = rows[:100]
        from googleapiclient.discovery import build
        creds = get_google_credentials()
        service = build('sheets', 'v4', credentials=creds)
        sheet = service.spreadsheets()
        import re
        rows_preview = rows[:100]
        
        # Encontrar la siguiente fila vacía usando 'Marca' como referencia
        primer_campo = 'Marca' if 'Marca' in ubicaciones else list(ubicaciones.keys())[0]
        primer_rango = ubicaciones[primer_campo]
        match = re.match(r"([A-Z]+)", primer_rango)
        if match:
            col = match.group(1)
        else:
            col = primer_rango.split(':')[0]
        
        rango_lectura = f"{nombre_hoja}!{col}2:{col}"
        result = sheet.values().get(
            spreadsheetId=spreadsheet_id,
            range=rango_lectura
        ).execute()
        valores_existentes = result.get('values', [])
        siguiente_fila = len(valores_existentes) + 2
        log(f'Siguiente fila disponible para carga masiva: {siguiente_fila}')
        
        # Escribir desde la siguiente fila vacía
        for campo, rango in ubicaciones.items():
            match = re.match(r"([A-Z]+)", rango)
            if match:
                col = match.group(1)
            else:
                col = rango.split(':')[0]
            values = [[row.get(campo, '')] for row in rows]
            fila_final = siguiente_fila + len(values) - 1
            log(f'Escribiendo campo {campo} en {col}{siguiente_fila}:{col}{fila_final} con valores:', values[:3])
            rango_celda = f"{nombre_hoja}!{col}{siguiente_fila}:{col}{fila_final}"
            sheet.values().update(
                spreadsheetId=spreadsheet_id,
                range=rango_celda,
                valueInputOption='USER_ENTERED',
                body={'values': values}
            ).execute()
        log('Escritura masiva terminada')
        return jsonify({'status': 'ok', 'rows': len(rows), 'preview': rows_preview})
    except Exception as e:
        log('ERROR:', str(e))
        log(traceback.format_exc())
        return jsonify({'error': str(e)}), 500

# --- BLACKLIST DATOS (OBTENER TODOS) ENDPOINT ---
@app.route('/api/blacklist_datos', methods=['GET'])
def blacklist_datos():
    auth_header = request.headers.get('Authorization')
    if not auth_header or not auth_header.startswith('Bearer '):
        return jsonify({'error': 'No token provided'}), 401
    id_token = auth_header.split(' ')[1]
    try:
        decoded_token = auth.verify_id_token(id_token)
        uid = decoded_token['uid']
        print('[BLACKLIST_DATOS] UID:', uid)
        

        
        # Obtener configuración de hoja/rango
        hoja_ref = db.collection('Areas').document(uid).collection('Hojas').document('Blacklist')
        hoja_doc = hoja_ref.get()
        if not hoja_doc.exists:
            return jsonify({'error': 'No se encontró la configuración de hoja/rango'}), 404
        
        hoja_data = hoja_doc.to_dict()
        nombre_hoja = hoja_data.get('Hoja', 'Sheet1')
        ubicaciones = {k: v for k, v in hoja_data.items() if k != 'Hoja'}
        print('[BLACKLIST_DATOS] Ubicaciones:', ubicaciones)
        
        # Obtener URL del spreadsheet
        area_doc = db.collection('Areas').document(uid).get()
        if not area_doc.exists:
            return jsonify({'error': 'No se encontró el área para este usuario'}), 404
        
        area_data = area_doc.to_dict()
        spreadsheet_url = area_data.get('Blacklist')
        if not spreadsheet_url:
            return jsonify({'error': 'No se encontró la URL del spreadsheet'}), 404
        
        import re
        match = re.search(r"/d/([a-zA-Z0-9-_]+)", spreadsheet_url)
        if not match:
            return jsonify({'error': 'URL de spreadsheet inválida'}), 400
        spreadsheet_id = match.group(1)
        
        from googleapiclient.discovery import build
        creds = get_google_credentials()
        service = build('sheets', 'v4', credentials=creds)
        sheet = service.spreadsheets()
        
        # Leer datos de todas las columnas configuradas
        datos = []
        campos_orden = list(ubicaciones.keys())
        
        # Obtener el campo 'Marca' para determinar cuántas filas hay (o el primer campo si no existe)
        primer_campo = 'Marca' if 'Marca' in ubicaciones else campos_orden[0]
        primer_rango = ubicaciones[primer_campo]
        match = re.match(r"([A-Z]+)", primer_rango)
        if match:
            col = match.group(1)
        else:
            col = primer_rango.split(':')[0]
        
        # Usar rango abierto para leer todas las filas disponibles
        rango_lectura = f"{nombre_hoja}!{col}2:{col}"
        print(f'[BLACKLIST_DATOS] Leyendo {rango_lectura} para contar filas (campo: {primer_campo})')
        
        result = sheet.values().get(
            spreadsheetId=spreadsheet_id,
            range=rango_lectura
        ).execute()
        
        primera_columna = result.get('values', [])
        num_filas = len(primera_columna)
        print(f'[BLACKLIST_DATOS] Filas detectadas: {num_filas}')
        
        if num_filas == 0:
            return jsonify({'datos': [], 'campos': campos_orden})
        
        # Leer todas las columnas usando rangos abiertos
        columnas_datos = {}
        for campo in campos_orden:
            rango = ubicaciones[campo]
            match = re.match(r"([A-Z]+)", rango)
            if match:
                col = match.group(1)
            else:
                col = rango.split(':')[0]
            
            rango_lectura = f"{nombre_hoja}!{col}2:{col}"
            print(f'[BLACKLIST_DATOS] Leyendo {campo} desde {rango_lectura}')
            
            result = sheet.values().get(
                spreadsheetId=spreadsheet_id,
                range=rango_lectura
            ).execute()
            
            valores = result.get('values', [])
            columnas_datos[campo] = [v[0] if v else '' for v in valores]
        
        # Construir array de objetos
        for i in range(num_filas):
            fila = {'_row_index': i + 2}  # +2 porque empezamos en la fila 2
            for campo in campos_orden:
                fila[campo] = columnas_datos[campo][i] if i < len(columnas_datos[campo]) else ''
            datos.append(fila)
        
        print(f'[BLACKLIST_DATOS] Retornando {len(datos)} filas')
        return jsonify({'datos': datos, 'campos': campos_orden})
        
    except Exception as e:
        import traceback
        print('[BLACKLIST_DATOS] Error:', str(e))
        print(traceback.format_exc())
        return jsonify({'error': str(e)}), 500

# --- BLACKLIST ELIMINAR FILA ENDPOINT ---
@app.route('/api/blacklist_eliminar', methods=['POST'])
def blacklist_eliminar():
    auth_header = request.headers.get('Authorization')
    if not auth_header or not auth_header.startswith('Bearer '):
        return jsonify({'error': 'No token provided'}), 401
    id_token = auth_header.split(' ')[1]
    try:
        decoded_token = auth.verify_id_token(id_token)
        uid = decoded_token['uid']
        data = request.get_json()
        row_index = data.get('row_index')
        
        if not row_index:
            return jsonify({'error': 'row_index requerido'}), 400
        
        print(f'[BLACKLIST_ELIMINAR] UID: {uid}, Fila a eliminar: {row_index}')
        

        
        # Obtener configuración de hoja/rango
        hoja_ref = db.collection('Areas').document(uid).collection('Hojas').document('Blacklist')
        hoja_doc = hoja_ref.get()
        if not hoja_doc.exists:
            return jsonify({'error': 'No se encontró la configuración de hoja/rango'}), 404
        
        hoja_data = hoja_doc.to_dict()
        nombre_hoja = hoja_data.get('Hoja', 'Sheet1')
        ubicaciones = {k: v for k, v in hoja_data.items() if k != 'Hoja'}
        
        # Obtener URL del spreadsheet
        area_doc = db.collection('Areas').document(uid).get()
        if not area_doc.exists:
            return jsonify({'error': 'No se encontró el área para este usuario'}), 404
        
        area_data = area_doc.to_dict()
        spreadsheet_url = area_data.get('Blacklist')
        if not spreadsheet_url:
            return jsonify({'error': 'No se encontró la URL del spreadsheet'}), 404
        
        import re
        match = re.search(r"/d/([a-zA-Z0-9-_]+)", spreadsheet_url)
        if not match:
            return jsonify({'error': 'URL de spreadsheet inválida'}), 400
        spreadsheet_id = match.group(1)
        
        from googleapiclient.discovery import build
        creds = get_google_credentials()
        service = build('sheets', 'v4', credentials=creds)
        
        # Obtener el sheetId del spreadsheet
        spreadsheet = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        sheet_id = None
        for sheet in spreadsheet.get('sheets', []):
            if sheet['properties']['title'] == nombre_hoja:
                sheet_id = sheet['properties']['sheetId']
                break
        
        if sheet_id is None:
            return jsonify({'error': f'No se encontró la hoja {nombre_hoja}'}), 404
        
        print(f'[BLACKLIST_ELIMINAR] SheetId: {sheet_id}, Eliminando fila {row_index}')
        
        # Eliminar la fila usando batchUpdate
        requests = [{
            'deleteDimension': {
                'range': {
                    'sheetId': sheet_id,
                    'dimension': 'ROWS',
                    'startIndex': row_index - 1,  # 0-indexed
                    'endIndex': row_index  # exclusive
                }
            }
        }]
        
        body = {'requests': requests}
        service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body=body
        ).execute()
        
        print(f'[BLACKLIST_ELIMINAR] Fila {row_index} eliminada exitosamente')
        return jsonify({'status': 'ok', 'deleted_row': row_index})
        
    except Exception as e:
        import traceback
        print('[BLACKLIST_ELIMINAR] Error:', str(e))
        print(traceback.format_exc())
        return jsonify({'error': str(e)}), 500

# --- ENDPOINTS PARA PEDIDOS CALENDARIO ---
@app.route('/api/pedidos_calendario_campos', methods=['GET'])
def pedidos_calendario_campos():
    auth_header = request.headers.get('Authorization')
    if not auth_header or not auth_header.startswith('Bearer '):
        return jsonify({'error': 'No token provided'}), 401
    id_token = auth_header.split(' ')[1]
    try:
        decoded_token = auth.verify_id_token(id_token)
        uid = decoded_token['uid']
        hoja_ref = db.collection('Areas').document(uid).collection('Hojas').document('CALENDARIO')
        hoja_doc = hoja_ref.get()
        if not hoja_doc.exists:
            return jsonify({'error': 'No se encontró la configuración de hoja/rango'}), 404
        hoja_data = hoja_doc.to_dict()
        campos = [k for k in hoja_data.keys() if k != 'Hoja']
        return jsonify({'campos': campos})
    except Exception as e:
        return jsonify({'error': 'Error al obtener campos', 'details': str(e)}), 500

@app.route('/api/pedidos_calendario_registro', methods=['POST'])
def pedidos_calendario_registro():
    auth_header = request.headers.get('Authorization')
    if not auth_header or not auth_header.startswith('Bearer '):
        return jsonify({'error': 'No token provided'}), 401
    id_token = auth_header.split(' ')[1]
    try:
        import datetime, re, traceback
        decoded_token = auth.verify_id_token(id_token)
        uid = decoded_token['uid']
        data = request.get_json()
        print('[PEDIDOS_CALENDARIO_REGISTRO] UID:', uid)
        print('[PEDIDOS_CALENDARIO_REGISTRO] Data recibida:', data)
        area_doc = db.collection('Areas').document(uid).get()
        if not area_doc.exists:
            return jsonify({'error': 'No se encontró el área para este usuario'}), 404
        area_data = area_doc.to_dict()
        spreadsheet_url = area_data.get('CALENDARIO')
        print('[PEDIDOS_CALENDARIO_REGISTRO] Spreadsheet URL:', spreadsheet_url)
        if not spreadsheet_url:
            return jsonify({'error': 'No se encontró la URL del spreadsheet'}), 404
        import re
        match = re.search(r"/d/([a-zA-Z0-9-_]+)", spreadsheet_url)
        if not match:
            return jsonify({'error': 'URL de spreadsheet inválida'}), 400
        spreadsheet_id = match.group(1)
        print('[PEDIDOS_CALENDARIO_REGISTRO] Spreadsheet ID:', spreadsheet_id)
        hoja_ref = db.collection('Areas').document(uid).collection('Hojas').document('CALENDARIO')
        hoja_doc = hoja_ref.get()
        if not hoja_doc.exists:
            return jsonify({'error': 'No se encontró la configuración de hoja/rango'}), 404
        hoja_data = hoja_doc.to_dict()
        print('[PEDIDOS_CALENDARIO_REGISTRO] Hoja data:', hoja_data)
        nombre_hoja = hoja_data.get('Hoja', 'Sheet1')
        ubicaciones = {k: v for k, v in hoja_data.items() if k != 'Hoja'}
        print('[PEDIDOS_CALENDARIO_REGISTRO] Ubicaciones:', ubicaciones)
        from google.oauth2 import service_account
        from googleapiclient.discovery import build
        import datetime
        SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
        SERVICE_ACCOUNT_FILE = 'supratechweb-firebase-adminsdk-fbsvc-8d4aa68a75.json'
        creds = service_account.Credentials.from_service_account_file(
            SERVICE_ACCOUNT_FILE, scopes=SCOPES)
        service = build('sheets', 'v4', credentials=creds)
        sheet = service.spreadsheets()
        fecha = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        print(f'[PEDIDOS_CALENDARIO_REGISTRO] Limpiando todos los rangos configurados antes de registrar...')
        for campo, rango in ubicaciones.items():
            match = re.match(r"([A-Z]+)", rango)
            if match:
                col = match.group(1)
            else:
                col = rango.split(':')[0]
            clear_range = f"{nombre_hoja}!{col}2:{col}1000"
            print(f'[PEDIDOS_CALENDARIO_REGISTRO] Limpiando rango {clear_range}')
            sheet.values().clear(
                spreadsheetId=spreadsheet_id,
                range=clear_range,
                body={}
            ).execute()
        for campo, rango in ubicaciones.items():
            match = re.match(r"([A-Z]+)", rango)
            if match:
                col = match.group(1)
            else:
                col = rango.split(':')[0]
            valor = data.get(campo, '')
            if campo.lower() == 'fecha':
                valor = fecha
            values = [[valor]]
            rango_celda = f"{nombre_hoja}!{col}2:{col}2"
            print(f'[PEDIDOS_CALENDARIO_REGISTRO] Escribiendo campo {campo} en {rango_celda} con valor: {valor}')
            try:
                sheet.values().update(
                    spreadsheetId=spreadsheet_id,
                    range=rango_celda,
                    valueInputOption='USER_ENTERED',
                    body={'values': values}
                ).execute()
            except Exception as e:
                print(f'[PEDIDOS_CALENDARIO_REGISTRO] Error insertando {campo} en {rango_celda}:', str(e))
                print(traceback.format_exc())
                return jsonify({'error': f'Error insertando {campo} en {rango_celda}', 'details': str(e)}), 500
        print('[PEDIDOS_CALENDARIO_REGISTRO] Registro exitoso')
        return jsonify({'status': 'ok'})
    except Exception as e:
        import traceback
        print('[PEDIDOS_CALENDARIO_REGISTRO] Error general:', str(e))
        print(traceback.format_exc())
        return jsonify({'error': 'Error al registrar en Pedidos Calendario', 'details': str(e)}), 500

@app.route('/api/pedidos_calendario_bulk', methods=['POST'])
def pedidos_calendario_bulk():
    auth_header = request.headers.get('Authorization')
    if not auth_header or not auth_header.startswith('Bearer '):
        return jsonify({'error': 'No token provided'}), 401
    id_token = auth_header.split(' ')[1]
    import traceback
    try:
        decoded_token = auth.verify_id_token(id_token)
        uid = decoded_token['uid']
        import sys
        def log(msg, *args):
            print('[PEDIDOS_CALENDARIO_BULK]', msg, *args)
            sys.stdout.flush()
        log('UID:', uid)

        hoja_ref = db.collection('Areas').document(uid).collection('Hojas').document('CALENDARIO')
        hoja_doc = hoja_ref.get()
        if not hoja_doc.exists:
            log('No se encontró la configuración de hoja/rango')
            return jsonify({'error': 'No se encontró la configuración de hoja/rango'}), 404
        hoja_data = hoja_doc.to_dict()
        nombre_hoja = hoja_data.get('Hoja', 'Sheet1')
        ubicaciones = {k: v for k, v in hoja_data.items() if k != 'Hoja'}
        log('Ubicaciones:', ubicaciones)
        area_doc = db.collection('Areas').document(uid).get()
        if not area_doc.exists:
            log('No se encontró el área para este usuario')
            return jsonify({'error': 'No se encontró el área para este usuario'}), 404
        area_data = area_doc.to_dict()
        spreadsheet_url = area_data.get('CALENDARIO')
        if not spreadsheet_url:
            log('No se encontró la URL del spreadsheet para CALENDARIO')
            return jsonify({'error': 'No se encontró la URL del spreadsheet para CALENDARIO. Configura el campo "CALENDARIO" en Firestore.'}), 404
        import re
        match = re.search(r"/d/([a-zA-Z0-9-_]+)", spreadsheet_url)
        if not match:
            log('URL de spreadsheet inválida:', spreadsheet_url)
            return jsonify({'error': 'URL de spreadsheet inválida'}), 400
        spreadsheet_id = match.group(1)
        if 'file' not in request.files:
            log('No file uploaded')
            return jsonify({'error': 'No file uploaded'}), 400
        file = request.files['file']
        filename = file.filename.lower()
        campos_config = list(ubicaciones.keys())
        log('Campos config:', campos_config)
        rows = []
        if filename.endswith('.csv'):
            import csv, io
            try:
                content = file.stream.read().decode('utf-8')
            except UnicodeDecodeError as e:
                log('Fallo utf-8, intentando latin-1:', str(e))
                file.stream.seek(0)
                content = file.stream.read().decode('latin-1')
            stream = io.StringIO(content)
            reader = csv.DictReader(stream)
            log('CSV headers:', reader.fieldnames)
            if reader.fieldnames is None:
                log('El archivo CSV no tiene encabezados.')
                return jsonify({'error': 'El archivo CSV no tiene encabezados.'}), 400
            missing = [campo for campo in campos_config if campo not in reader.fieldnames]
            if missing:
                log('Faltan columnas requeridas:', missing)
                return jsonify({'error': f'Faltan columnas requeridas: {missing}'}), 400
            rows = list(reader)
            log('Primeras filas:', rows[:3])
        elif filename.endswith('.xlsx'):
            import openpyxl
            wb = openpyxl.load_workbook(file, read_only=True)
            ws = wb.active
            headers = [cell.value for cell in next(ws.iter_rows(min_row=1, max_row=1))]
            log('XLSX headers:', headers)
            missing = [campo for campo in campos_config if campo not in headers]
            if missing:
                log('Faltan columnas requeridas:', missing)
                return jsonify({'error': f'Faltan columnas requeridas: {missing}'}), 400
            for row in ws.iter_rows(min_row=2, values_only=True):
                rows.append(dict(zip(headers, row)))
            log('Primeras filas:', rows[:3])
        else:
            log('Formato de archivo no permitido:', filename)
            return jsonify({'error': 'Solo se permiten archivos CSV o XLSX'}), 400
        rows_preview = rows[:100]
        from googleapiclient.discovery import build
        creds = get_google_credentials()
        service = build('sheets', 'v4', credentials=creds)
        sheet = service.spreadsheets()
        import re
        for campo, rango in ubicaciones.items():
            match = re.match(r"([A-Z]+)", rango)
            if match:
                col = match.group(1)
            else:
                col = rango.split(':')[0]
            clear_range = f"{nombre_hoja}!{col}2:{col}1000"
            log(f'Limpiando rango {clear_range}')
            sheet.values().clear(
                spreadsheetId=spreadsheet_id,
                range=clear_range,
                body={}
            ).execute()
        for campo, rango in ubicaciones.items():
            match = re.match(r"([A-Z]+)", rango)
            if match:
                col = match.group(1)
            else:
                col = rango.split(':')[0]
            fila = 2
            values = [[row.get(campo, '')] for row in rows]
            fila_final = fila + len(values) - 1
            log(f'Escribiendo campo {campo} en {col}{fila}:{col}{fila_final} con valores:', values[:3])
            rango_celda = f"{nombre_hoja}!{col}{fila}:{col}{fila_final}"
            sheet.values().update(
                spreadsheetId=spreadsheet_id,
                range=rango_celda,
                valueInputOption='USER_ENTERED',
                body={'values': values}
            ).execute()
        log('Escritura masiva terminada')
        return jsonify({'status': 'ok', 'rows': len(rows), 'preview': rows_preview})
    except Exception as e:
        log('ERROR:', str(e))
        log(traceback.format_exc())
        return jsonify({'error': str(e)}), 500

# --- ENDPOINT PARA OBTENER DATOS DE CALENDARIO (TABLA A2:G14) ---
@app.route('/api/pedidos_calendario_datos', methods=['GET'])
def pedidos_calendario_datos():
    auth_header = request.headers.get('Authorization')
    if not auth_header or not auth_header.startswith('Bearer '):
        return jsonify({'error': 'No token provided'}), 401
    id_token = auth_header.split(' ')[1]
    try:
        decoded_token = auth.verify_id_token(id_token)
        uid = decoded_token['uid']

        # Obtener URL del spreadsheet
        area_doc = db.collection('Areas').document(uid).get()
        if not area_doc.exists:
            return jsonify({'error': 'No se encontró el área para este usuario'}), 404

        area_data = area_doc.to_dict()
        spreadsheet_url = area_data.get('CALENDARIO')
        if not spreadsheet_url:
            return jsonify({'error': 'No se encontró la URL del spreadsheet'}), 404

        import re
        match = re.search(r"/d/([a-zA-Z0-9-_]+)", spreadsheet_url)
        if not match:
            return jsonify({'error': 'URL de spreadsheet inválida'}), 400
        spreadsheet_id = match.group(1)

        # Obtener nombre de hoja
        hoja_ref = db.collection('Areas').document(uid).collection('Hojas').document('CALENDARIO')
        hoja_doc = hoja_ref.get()
        nombre_hoja = 'Sheet1'
        if hoja_doc.exists:
            hoja_data = hoja_doc.to_dict()
            nombre_hoja = hoja_data.get('Hoja', 'Sheet1')

        from googleapiclient.discovery import build
        creds = get_google_credentials()
        service = build('sheets', 'v4', credentials=creds)
        sheet = service.spreadsheets()

        # Leer encabezados de la fila 1
        rango_headers = f"{nombre_hoja}!A1:G1"
        result_headers = sheet.values().get(
            spreadsheetId=spreadsheet_id,
            range=rango_headers
        ).execute()

        encabezados_raw = result_headers.get('values', [[]])[0] if result_headers.get('values') else []
        encabezados = []
        for i in range(7):
            if i < len(encabezados_raw):
                encabezados.append(encabezados_raw[i])
            else:
                encabezados.append(chr(65 + i))  # A, B, C, etc como fallback

        # Leer rango A2:G14
        rango = f"{nombre_hoja}!A2:G14"
        result = sheet.values().get(
            spreadsheetId=spreadsheet_id,
            range=rango
        ).execute()

        valores = result.get('values', [])

        # Asegurar que tenemos 13 filas con 7 columnas cada una
        datos = []
        for i in range(13):
            if i < len(valores):
                fila = valores[i]
                # Completar con valores vacíos si faltan columnas
                while len(fila) < 7:
                    fila.append('')
                datos.append(fila[:7])  # Solo tomar primeras 7 columnas
            else:
                datos.append(['', '', '', '', '', '', ''])

        return jsonify({'encabezados': encabezados, 'datos': datos})

    except Exception as e:
        import traceback
        print('[CALENDARIO_DATOS] Error:', str(e))
        print(traceback.format_exc())
        return jsonify({'error': str(e)}), 500

# --- ENDPOINT PARA GUARDAR DATOS DE CALENDARIO (TABLA A2:G14) ---
@app.route('/api/pedidos_calendario_guardar', methods=['POST'])
def pedidos_calendario_guardar():
    auth_header = request.headers.get('Authorization')
    if not auth_header or not auth_header.startswith('Bearer '):
        return jsonify({'error': 'No token provided'}), 401
    id_token = auth_header.split(' ')[1]
    try:
        decoded_token = auth.verify_id_token(id_token)
        uid = decoded_token['uid']
        data = request.get_json()
        tabla_datos = data.get('datos', [])

        # Obtener URL del spreadsheet
        area_doc = db.collection('Areas').document(uid).get()
        if not area_doc.exists:
            return jsonify({'error': 'No se encontró el área para este usuario'}), 404

        area_data = area_doc.to_dict()
        spreadsheet_url = area_data.get('CALENDARIO')
        if not spreadsheet_url:
            return jsonify({'error': 'No se encontró la URL del spreadsheet'}), 404

        import re
        match = re.search(r"/d/([a-zA-Z0-9-_]+)", spreadsheet_url)
        if not match:
            return jsonify({'error': 'URL de spreadsheet inválida'}), 400
        spreadsheet_id = match.group(1)

        # Obtener nombre de hoja
        hoja_ref = db.collection('Areas').document(uid).collection('Hojas').document('CALENDARIO')
        hoja_doc = hoja_ref.get()
        nombre_hoja = 'Sheet1'
        if hoja_doc.exists:
            hoja_data = hoja_doc.to_dict()
            nombre_hoja = hoja_data.get('Hoja', 'Sheet1')

        from googleapiclient.discovery import build
        creds = get_google_credentials()
        service = build('sheets', 'v4', credentials=creds)
        sheet = service.spreadsheets()

        # Escribir en rango A2:G14
        rango = f"{nombre_hoja}!A2:G14"
        sheet.values().update(
            spreadsheetId=spreadsheet_id,
            range=rango,
            valueInputOption='USER_ENTERED',
            body={'values': tabla_datos}
        ).execute()

        return jsonify({'status': 'ok'})

    except Exception as e:
        import traceback
        print('[CALENDARIO_GUARDAR] Error:', str(e))
        print(traceback.format_exc())
        return jsonify({'error': str(e)}), 500

# --- ENDPOINTS PARA PEDIDOS BD ---
@app.route('/api/pedidos_bd_campos', methods=['GET'])
def pedidos_bd_campos():
    auth_header = request.headers.get('Authorization')
    if not auth_header or not auth_header.startswith('Bearer '):
        return jsonify({'error': 'No token provided'}), 401
    id_token = auth_header.split(' ')[1]
    try:
        decoded_token = auth.verify_id_token(id_token)
        uid = decoded_token['uid']
        hoja_ref = db.collection('Areas').document(uid).collection('Hojas').document('BDsi')
        hoja_doc = hoja_ref.get()
        if not hoja_doc.exists:
            return jsonify({'error': 'No se encontró la configuración de hoja/rango'}), 404
        hoja_data = hoja_doc.to_dict()
        campos = [k for k in hoja_data.keys() if k != 'Hoja']
        return jsonify({'campos': campos})
    except Exception as e:
        return jsonify({'error': 'Error al obtener campos', 'details': str(e)}), 500

@app.route('/api/pedidos_bd_registro', methods=['POST'])
def pedidos_bd_registro():
    auth_header = request.headers.get('Authorization')
    if not auth_header or not auth_header.startswith('Bearer '):
        return jsonify({'error': 'No token provided'}), 401
    id_token = auth_header.split(' ')[1]
    try:
        import datetime, re, traceback
        decoded_token = auth.verify_id_token(id_token)
        uid = decoded_token['uid']
        data = request.get_json()
        area_doc = db.collection('Areas').document(uid).get()
        if not area_doc.exists:
            return jsonify({'error': 'No se encontró el área para este usuario'}), 404
        area_data = area_doc.to_dict()
        spreadsheet_url = area_data.get('BDsi')
        if not spreadsheet_url:
            return jsonify({'error': 'No se encontró la URL del spreadsheet'}), 404
        match = re.search(r"/d/([a-zA-Z0-9-_]+)", spreadsheet_url)
        if not match:
            return jsonify({'error': 'URL de spreadsheet inválida'}), 400
        spreadsheet_id = match.group(1)
        hoja_ref = db.collection('Areas').document(uid).collection('Hojas').document('BDsi')
        hoja_doc = hoja_ref.get()
        if not hoja_doc.exists:
            return jsonify({'error': 'No se encontró la configuración de hoja/rango'}), 404
        hoja_data = hoja_doc.to_dict()
        nombre_hoja = hoja_data.get('Hoja', 'Sheet1')
        ubicaciones = {k: v for k, v in hoja_data.items() if k != 'Hoja'}
        from googleapiclient.discovery import build
        creds = get_google_credentials()
        service = build('sheets', 'v4', credentials=creds)
        sheet = service.spreadsheets()
        fecha = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        for campo, rango in ubicaciones.items():
            match = re.match(r"([A-Z]+)", rango)
            col = match.group(1) if match else rango.split(':')[0]
            clear_range = f"{nombre_hoja}!{col}2:{col}1000"
            sheet.values().clear(spreadsheetId=spreadsheet_id, range=clear_range, body={}).execute()
        for campo, rango in ubicaciones.items():
            match = re.match(r"([A-Z]+)", rango)
            col = match.group(1) if match else rango.split(':')[0]
            valor = data.get(campo, '')
            if campo.lower() == 'fecha':
                valor = fecha
            values = [[valor]]
            rango_celda = f"{nombre_hoja}!{col}2:{col}2"
            sheet.values().update(spreadsheetId=spreadsheet_id, range=rango_celda, valueInputOption='USER_ENTERED', body={'values': values}).execute()
        return jsonify({'status': 'ok'})
    except Exception as e:
        import traceback
        print('[PEDIDOS_BD_REGISTRO] Error:', str(e))
        print(traceback.format_exc())
        return jsonify({'error': 'Error al registrar en Pedidos BD', 'details': str(e)}), 500

@app.route('/api/pedidos_bd_bulk', methods=['POST'])
def pedidos_bd_bulk():
    auth_header = request.headers.get('Authorization')
    if not auth_header or not auth_header.startswith('Bearer '):
        return jsonify({'error': 'No token provided'}), 401
    id_token = auth_header.split(' ')[1]
    import traceback
    try:
        decoded_token = auth.verify_id_token(id_token)
        uid = decoded_token['uid']

        hoja_ref = db.collection('Areas').document(uid).collection('Hojas').document('BDsi')
        hoja_doc = hoja_ref.get()
        if not hoja_doc.exists:
            return jsonify({'error': 'No se encontró la configuración de hoja/rango'}), 404
        hoja_data = hoja_doc.to_dict()
        nombre_hoja = hoja_data.get('Hoja', 'Sheet1')
        ubicaciones = {k: v for k, v in hoja_data.items() if k != 'Hoja'}
        area_doc = db.collection('Areas').document(uid).get()
        if not area_doc.exists:
            return jsonify({'error': 'No se encontró el área para este usuario'}), 404
        area_data = area_doc.to_dict()
        spreadsheet_url = area_data.get('BDsi')
        if not spreadsheet_url:
            return jsonify({'error': 'No se encontró la URL del spreadsheet para BDsi'}), 404
        import re
        match = re.search(r"/d/([a-zA-Z0-9-_]+)", spreadsheet_url)
        if not match:
            return jsonify({'error': 'URL de spreadsheet inválida'}), 400
        spreadsheet_id = match.group(1)
        if 'file' not in request.files:
            return jsonify({'error': 'No file uploaded'}), 400
        file = request.files['file']
        filename = file.filename.lower()
        campos_config = list(ubicaciones.keys())
        rows = []
        if filename.endswith('.csv'):
            import csv, io
            try:
                content = file.stream.read().decode('utf-8')
            except UnicodeDecodeError:
                file.stream.seek(0)
                content = file.stream.read().decode('latin-1')
            stream = io.StringIO(content)
            reader = csv.DictReader(stream)
            if reader.fieldnames is None:
                return jsonify({'error': 'El archivo CSV no tiene encabezados.'}), 400

            # Normalizar nombres de columnas (case-insensitive)
            csv_headers_lower = {h.lower(): h for h in reader.fieldnames}
            config_to_csv = {}
            missing = []
            for campo in campos_config:
                campo_lower = campo.lower()
                if campo_lower in csv_headers_lower:
                    config_to_csv[campo] = csv_headers_lower[campo_lower]
                else:
                    missing.append(campo)

            if missing:
                return jsonify({'error': f'Faltan columnas requeridas: {missing}'}), 400

            # Leer filas y mapear nombres de columnas
            rows = []
            for row in reader:
                normalized_row = {}
                for campo_config, campo_csv in config_to_csv.items():
                    normalized_row[campo_config] = row.get(campo_csv, '')
                rows.append(normalized_row)
        elif filename.endswith('.xlsx'):
            import openpyxl
            wb = openpyxl.load_workbook(file, read_only=True)
            ws = wb.active
            headers = [cell.value for cell in next(ws.iter_rows(min_row=1, max_row=1))]

            # Normalizar nombres de columnas (case-insensitive)
            xlsx_headers_lower = {h.lower(): h for h in headers if h}
            config_to_xlsx = {}
            missing = []
            for campo in campos_config:
                campo_lower = campo.lower()
                if campo_lower in xlsx_headers_lower:
                    config_to_xlsx[campo] = xlsx_headers_lower[campo_lower]
                else:
                    missing.append(campo)

            if missing:
                return jsonify({'error': f'Faltan columnas requeridas: {missing}'}), 400

            # Leer filas y mapear nombres de columnas
            for row in ws.iter_rows(min_row=2, values_only=True):
                row_dict = dict(zip(headers, row))
                normalized_row = {}
                for campo_config, campo_xlsx in config_to_xlsx.items():
                    normalized_row[campo_config] = row_dict.get(campo_xlsx, '')
                rows.append(normalized_row)
        else:
            return jsonify({'error': 'Solo se permiten archivos CSV o XLSX'}), 400
        rows_preview = rows[:100]
        from googleapiclient.discovery import build
        creds = get_google_credentials()
        service = build('sheets', 'v4', credentials=creds)
        sheet = service.spreadsheets()
        for campo, rango in ubicaciones.items():
            match = re.match(r"([A-Z]+)", rango)
            col = match.group(1) if match else rango.split(':')[0]
            clear_range = f"{nombre_hoja}!{col}2:{col}1000"
            sheet.values().clear(spreadsheetId=spreadsheet_id, range=clear_range, body={}).execute()
        for campo, rango in ubicaciones.items():
            match = re.match(r"([A-Z]+)", rango)
            col = match.group(1) if match else rango.split(':')[0]
            fila = 2
            values = [[row.get(campo, '')] for row in rows]
            fila_final = fila + len(values) - 1
            rango_celda = f"{nombre_hoja}!{col}{fila}:{col}{fila_final}"
            sheet.values().update(spreadsheetId=spreadsheet_id, range=rango_celda, valueInputOption='USER_ENTERED', body={'values': values}).execute()
        return jsonify({'status': 'ok', 'rows': len(rows), 'preview': rows_preview})
    except Exception as e:
        print('[PEDIDOS_BD_BULK] Error:', str(e))
        print(traceback.format_exc())
        return jsonify({'error': str(e)}), 500

# ==================== PEDIDOS - BDQTY ====================
@app.route('/api/pedidos_bdqty_campos', methods=['GET'])
def pedidos_bdqty_campos():
    try:
        auth_header = request.headers.get('Authorization')
        if not auth_header or not auth_header.startswith('Bearer '):
            return jsonify({'error': 'Token no proporcionado'}), 401
        id_token = auth_header.split('Bearer ')[1]
        decoded_token = auth.verify_id_token(id_token)
        uid = decoded_token['uid']
        hoja_ref = db.collection('Areas').document(uid).collection('Hojas').document('BDQTY')
        hoja_doc = hoja_ref.get()
        if not hoja_doc.exists:
            return jsonify({'error': 'No se encontró la configuración de hoja/rango'}), 404
        hoja_data = hoja_doc.to_dict()
        campos = [k for k in hoja_data.keys() if k != 'Hoja']
        return jsonify({'campos': campos})
    except Exception as e:
        import traceback
        print('[PEDIDOS_BDQTY_CAMPOS] Error:', str(e))
        print(traceback.format_exc())
        return jsonify({'error': str(e)}), 500

@app.route('/api/pedidos_bdqty_registro', methods=['POST'])
def pedidos_bdqty_registro():
    auth_header = request.headers.get('Authorization')
    if not auth_header or not auth_header.startswith('Bearer '):
        return jsonify({'error': 'No token provided'}), 401
    id_token = auth_header.split(' ')[1]
    try:
        import datetime, re, traceback
        decoded_token = auth.verify_id_token(id_token)
        uid = decoded_token['uid']
        data = request.get_json()
        area_doc = db.collection('Areas').document(uid).get()
        if not area_doc.exists:
            return jsonify({'error': 'No se encontró el área para este usuario'}), 404
        area_data = area_doc.to_dict()
        spreadsheet_url = area_data.get('BDQTY')
        if not spreadsheet_url:
            return jsonify({'error': 'No se encontró la URL del spreadsheet'}), 404
        match = re.search(r"/d/([a-zA-Z0-9-_]+)", spreadsheet_url)
        if not match:
            return jsonify({'error': 'URL de spreadsheet inválida'}), 400
        spreadsheet_id = match.group(1)
        hoja_ref = db.collection('Areas').document(uid).collection('Hojas').document('BDQTY')
        hoja_doc = hoja_ref.get()
        if not hoja_doc.exists:
            return jsonify({'error': 'No se encontró la configuración de hoja/rango'}), 404
        hoja_data = hoja_doc.to_dict()
        nombre_hoja = hoja_data.get('Hoja', 'Sheet1')
        ubicaciones = {k: v for k, v in hoja_data.items() if k != 'Hoja'}
        from googleapiclient.discovery import build
        creds = get_google_credentials()
        service = build('sheets', 'v4', credentials=creds)
        sheet = service.spreadsheets()
        fecha = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        for campo, rango in ubicaciones.items():
            match = re.match(r"([A-Z]+)", rango)
            col = match.group(1) if match else rango.split(':')[0]
            clear_range = f"{nombre_hoja}!{col}2:{col}1000"
            sheet.values().clear(spreadsheetId=spreadsheet_id, range=clear_range, body={}).execute()
        for campo, rango in ubicaciones.items():
            match = re.match(r"([A-Z]+)", rango)
            col = match.group(1) if match else rango.split(':')[0]
            valor = data.get(campo, '')
            if campo.lower() == 'fecha':
                valor = fecha
            values = [[valor]]
            rango_celda = f"{nombre_hoja}!{col}2:{col}2"
            sheet.values().update(spreadsheetId=spreadsheet_id, range=rango_celda, valueInputOption='USER_ENTERED', body={'values': values}).execute()
        return jsonify({'status': 'ok'})
    except Exception as e:
        import traceback
        print('[PEDIDOS_BDQTY_REGISTRO] Error:', str(e))
        print(traceback.format_exc())
        return jsonify({'error': 'Error al registrar en Pedidos BDQTY', 'details': str(e)}), 500

@app.route('/api/pedidos_bdqty_bulk', methods=['POST'])
def pedidos_bdqty_bulk():
    try:
        auth_header = request.headers.get('Authorization')
        if not auth_header or not auth_header.startswith('Bearer '):
            return jsonify({'error': 'Token no proporcionado'}), 401
        id_token = auth_header.split('Bearer ')[1]
        decoded_token = auth.verify_id_token(id_token)
        uid = decoded_token['uid']

        hoja_ref = db.collection('Areas').document(uid).collection('Hojas').document('BDQTY')
        hoja_doc = hoja_ref.get()
        if not hoja_doc.exists:
            return jsonify({'error': 'No se encontró la configuración de hoja/rango'}), 404
        hoja_data = hoja_doc.to_dict()
        nombre_hoja = hoja_data.get('Hoja', 'Sheet1')
        ubicaciones = {k: v for k, v in hoja_data.items() if k != 'Hoja'}
        area_doc = db.collection('Areas').document(uid).get()
        if not area_doc.exists:
            return jsonify({'error': 'No se encontró el área para este usuario'}), 404
        area_data = area_doc.to_dict()
        spreadsheet_url = area_data.get('BDQTY')
        if not spreadsheet_url:
            return jsonify({'error': 'No se encontró la URL del spreadsheet para BDQTY'}), 404
        import re
        match = re.search(r"/d/([a-zA-Z0-9-_]+)", spreadsheet_url)
        if not match:
            return jsonify({'error': 'URL de spreadsheet inválida'}), 400
        spreadsheet_id = match.group(1)
        if 'file' not in request.files:
            return jsonify({'error': 'No file uploaded'}), 400
        file = request.files['file']
        filename = file.filename.lower()
        campos_config = list(ubicaciones.keys())
        rows = []
        if filename.endswith('.csv'):
            import csv, io
            try:
                content = file.stream.read().decode('utf-8')
            except UnicodeDecodeError:
                file.stream.seek(0)
                content = file.stream.read().decode('latin-1')
            stream = io.StringIO(content)
            reader = csv.DictReader(stream)
            if reader.fieldnames is None:
                return jsonify({'error': 'El archivo CSV no tiene encabezados.'}), 400

            # Normalizar nombres de columnas (case-insensitive)
            csv_headers_lower = {h.lower(): h for h in reader.fieldnames}
            config_to_csv = {}
            missing = []
            for campo in campos_config:
                campo_lower = campo.lower()
                if campo_lower in csv_headers_lower:
                    config_to_csv[campo] = csv_headers_lower[campo_lower]
                else:
                    missing.append(campo)

            if missing:
                return jsonify({'error': f'Faltan columnas requeridas: {missing}'}), 400

            # Leer filas y mapear nombres de columnas
            rows = []
            for row in reader:
                normalized_row = {}
                for campo_config, campo_csv in config_to_csv.items():
                    normalized_row[campo_config] = row.get(campo_csv, '')
                rows.append(normalized_row)
        elif filename.endswith('.xlsx'):
            import openpyxl
            wb = openpyxl.load_workbook(file, read_only=True)
            ws = wb.active
            headers = [cell.value for cell in next(ws.iter_rows(min_row=1, max_row=1))]

            # Normalizar nombres de columnas (case-insensitive)
            xlsx_headers_lower = {h.lower(): h for h in headers if h}
            config_to_xlsx = {}
            missing = []
            for campo in campos_config:
                campo_lower = campo.lower()
                if campo_lower in xlsx_headers_lower:
                    config_to_xlsx[campo] = xlsx_headers_lower[campo_lower]
                else:
                    missing.append(campo)

            if missing:
                return jsonify({'error': f'Faltan columnas requeridas: {missing}'}), 400

            # Leer filas y mapear nombres de columnas
            for row in ws.iter_rows(min_row=2, values_only=True):
                row_dict = dict(zip(headers, row))
                normalized_row = {}
                for campo_config, campo_xlsx in config_to_xlsx.items():
                    normalized_row[campo_config] = row_dict.get(campo_xlsx, '')
                rows.append(normalized_row)
        else:
            return jsonify({'error': 'Solo se permiten archivos CSV o XLSX'}), 400
        rows_preview = rows[:100]
        from googleapiclient.discovery import build
        creds = get_google_credentials()
        service = build('sheets', 'v4', credentials=creds)
        sheet = service.spreadsheets()
        for campo, rango in ubicaciones.items():
            match = re.match(r"([A-Z]+)", rango)
            col = match.group(1) if match else rango.split(':')[0]
            clear_range = f"{nombre_hoja}!{col}2:{col}1000"
            sheet.values().clear(spreadsheetId=spreadsheet_id, range=clear_range, body={}).execute()
        for campo, rango in ubicaciones.items():
            match = re.match(r"([A-Z]+)", rango)
            col = match.group(1) if match else rango.split(':')[0]
            fila = 2
            values = [[row.get(campo, '')] for row in rows]
            fila_final = fila + len(values) - 1
            rango_celda = f"{nombre_hoja}!{col}{fila}:{col}{fila_final}"
            sheet.values().update(spreadsheetId=spreadsheet_id, range=rango_celda, valueInputOption='USER_ENTERED', body={'values': values}).execute()
        return jsonify({'status': 'ok', 'rows': len(rows), 'preview': rows_preview})
    except Exception as e:
        print('[PEDIDOS_BDQTY_BULK] Error:', str(e))
        print(traceback.format_exc())
        return jsonify({'error': str(e)}), 500

# --- RUTAS PARA COMPRAS ---
@app.route('/compras')
def compras():
    return render_template('compras.html')

@app.route('/cotizaciones')
def cotizaciones():
    """Redirige a la hoja de cálculo de cotizaciones generada ese día"""
    try:
        # Obtener spreadsheet ID desde Firebase usando el string COTIZACIONES
        doc_ref = db.collection('configuracion').document('COTIZACIONES')
        doc = doc_ref.get()

        if doc.exists:
            spreadsheet_url = doc.to_dict().get('url', '')
            if spreadsheet_url:
                return redirect(spreadsheet_url)

        return jsonify({'error': 'No se encontró la configuración de COTIZACIONES'}), 404
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/pedidos_anteriores')
def pedidos_anteriores():
    """Muestra la información de pedidos anteriores"""
    try:
        # Redirigir directamente a la hoja de cálculo de pedidos anteriores
        pedidos_url = 'https://docs.google.com/spreadsheets/d/1LLksaPv6x_7YS3XmeY9S2j1I0sIDNRKBpShtfqe3_MA/edit?gid=0#gid=0'

        # Opcionalmente, puedes obtener la URL desde Firebase
        doc_ref = db.collection('configuracion').document('PEDIDOSANT')
        doc = doc_ref.get()

        if doc.exists:
            spreadsheet_url = doc.to_dict().get('url', '')
            if spreadsheet_url:
                pedidos_url = spreadsheet_url

        return redirect(pedidos_url)
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/indicadores')
def indicadores():
    """Muestra la página de indicadores con los 3 botones"""
    return render_template('indicadores.html')

@app.route('/api/indicadores/<section>')
def get_indicadores_data(section):
    """Obtiene datos filtrados de la hoja de cálculo de indicadores"""
    try:
        from googleapiclient.discovery import build

        # Obtener credenciales
        creds = get_google_credentials()
        service = build('sheets', 'v4', credentials=creds)

        # ID de la hoja de cálculo de indicadores
        spreadsheet_id = '1w8OOsQ-N9XkxD84xYIN6XJtjKvgvFBNjIm1U21-ksJk'

        # Opcionalmente, obtener desde Firebase
        doc_ref = db.collection('configuracion').document('INDICADORES')
        doc = doc_ref.get()

        if doc.exists:
            config_data = doc.to_dict()
            if 'spreadsheet_id' in config_data:
                spreadsheet_id = config_data['spreadsheet_id']

        # Leer todos los datos de la hoja
        result = service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range='A:Z'  # Leer todas las columnas
        ).execute()

        values = result.get('values', [])

        if not values:
            return jsonify({'error': 'No se encontraron datos'}), 404

        # Buscar las palabras clave en la primera columna
        cotizados_start = -1
        surtido_start = -1
        procesado_start = -1
        mesas_start = -1

        for i, row in enumerate(values):
            if row and len(row) > 0:
                cell_value = str(row[0]).upper().strip()
                if 'COTIZADOS' in cell_value or cell_value == 'COTIZADOS':
                    cotizados_start = i
                elif 'SURTIDO' in cell_value or cell_value == 'SURTIDO':
                    surtido_start = i
                elif 'PROCESADO' in cell_value or cell_value == 'PROCESADO':
                    procesado_start = i
                elif 'MESAS' in cell_value or cell_value == 'MESAS':
                    mesas_start = i

        # Determinar qué sección devolver
        if section == 'cotizados':
            if cotizados_start == -1:
                return jsonify({'error': 'No se encontró la sección COTIZADOS'}), 404

            end_index = surtido_start - 1 if surtido_start > cotizados_start else len(values)
            section_data = values[cotizados_start + 1:end_index]

        elif section == 'surtido':
            if surtido_start == -1:
                return jsonify({'error': 'No se encontró la sección SURTIDO'}), 404

            end_index = procesado_start - 1 if procesado_start > surtido_start else len(values)
            section_data = values[surtido_start + 1:end_index]

        elif section == 'procesado':
            if procesado_start == -1:
                return jsonify({'error': 'No se encontró la sección PROCESADO'}), 404

            end_index = mesas_start - 1 if mesas_start > procesado_start else len(values)
            section_data = values[procesado_start + 1:end_index]
        else:
            return jsonify({'error': 'Sección no válida'}), 400

        # Filtrar filas vacías
        section_data = [row for row in section_data if any(cell.strip() for cell in row if cell)]

        # Separar encabezados de datos
        headers = section_data[0] if section_data else []
        rows = section_data[1:] if len(section_data) > 1 else []

        return jsonify({
            'headers': headers,
            'rows': rows,
            'section': section
        })

    except Exception as e:
        print('[INDICADORES] Error:', str(e))
        import traceback
        print(traceback.format_exc())
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    import os
    port = int(os.environ.get('PORT', 5001))
    print(f'Servidor Flask corriendo en puerto {port}...')
    app.run(host='0.0.0.0', port=port, debug=False)