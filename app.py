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
        return service_account.Credentials.from_service_account_file(
            'supratechweb-firebase-adminsdk-fbsvc-8d4aa68a75.json',
            scopes=SCOPES
        )

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
        # Usar el cliente db global en lugar de crear uno nuevo
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
        # Usar el cliente db global en lugar de crear uno nuevo
        hoja_ref = db.collection('Areas').document(uid).collection('Hojas').document('BDQTY')
        hoja_doc = hoja_ref.get()
        if not hoja_doc.exists:
            return jsonify({'error': 'No se encontró la configuración de hoja/rango'}), 404
        hoja_data = hoja_doc.to_dict()
        campos = [k for k in hoja_data.keys() if k != 'Hoja']
        return jsonify({'campos': campos})
    except Exception as e:
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
        # Usar el cliente db global en lugar de crear uno nuevo
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

if __name__ == '__main__':
    import os
    port = int(os.environ.get('PORT', 5001))
    print(f'Servidor Flask corriendo en puerto {port}...')
    app.run(host='0.0.0.0', port=port, debug=False)