
# LIMPIEZA DE DUPLICADOS Y DEFINICIÓN ÚNICA DE RUTAS
from flask import Flask, request, jsonify, redirect, render_template
import firebase_admin
from firebase_admin import credentials, auth
from google.cloud import firestore

app = Flask(__name__)

# Usar el nombre real del archivo de credenciales subido como Secret File en Render
cred = credentials.Certificate('supratechweb-firebase-adminsdk-fbsvc-8d4aa68a75.json')
firebase_admin.initialize_app(cred)
db = firestore.Client.from_service_account_json('supratechweb-firebase-adminsdk-fbsvc-8d4aa68a75.json')

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

@app.route('/api/login', methods=['POST'])
def login():
    auth_header = request.headers.get('Authorization')
    if not auth_header or not auth_header.startswith('Bearer '):
        return jsonify({'error': 'No token provided'}), 401
    id_token = auth_header.split(' ')[1]
    try:
        decoded_token = auth.verify_id_token(id_token)
        uid = decoded_token['uid']
        email = decoded_token.get('email')
        # Busca o crea el usuario en Firestore
        user_ref = db.collection('users').document(uid)
        user_doc = user_ref.get()
        if not user_doc.exists:
            user_ref.set({'email': email, 'rol': 'usuario'})
            user_data = {'email': email, 'rol': 'usuario'}
        else:
            user_data = user_doc.to_dict()
        return jsonify({'status': 'ok', 'uid': uid, 'email': email, 'rol': user_data.get('rol', 'usuario')})
    except Exception as e:
        return jsonify({'error': 'Invalid token', 'details': str(e)}), 401

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

if __name__ == '__main__':
    app.run(debug=True)

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