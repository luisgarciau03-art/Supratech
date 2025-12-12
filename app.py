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
from flask import Flask, request, jsonify, redirect, render_template
@app.route('/')
def home():
    return render_template('login.html')
from flask import Flask, request, jsonify
import firebase_admin
from firebase_admin import credentials, auth
from google.cloud import firestore

app = Flask(__name__)

# Usar el nombre real del archivo de credenciales subido como Secret File en Render
cred = credentials.Certificate('supratechweb-firebase-adminsdk-fbsvc-8d4aa68a75.json')
firebase_admin.initialize_app(cred)
db = firestore.Client.from_service_account_json('supratechweb-firebase-adminsdk-fbsvc-8d4aa68a75.json')

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

if __name__ == '__main__':
    app.run(debug=True)