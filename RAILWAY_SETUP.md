# Configuración de Railway - Variable de Entorno Firebase

## Problema resuelto

El código ahora lee las credenciales de Firebase desde una variable de entorno llamada `FIREBASE_CREDENTIALS` en producción (Railway), o desde el archivo local en desarrollo.

## Pasos para configurar en Railway:

### 1. Obtener las credenciales

En tu computadora local, abre el archivo:
```
C:\Users\PC 1\SUPRATECHWEB\supratechweb-firebase-adminsdk-fbsvc-8d4aa68a75.json
```

### 2. Convertir a una sola línea (minificar)

Copia TODO el contenido del archivo JSON y elimina los saltos de línea y espacios extra.

**Método fácil:** Ve a https://codebeautify.org/jsonminifier y pega el JSON, luego copia el resultado minificado.

### 3. Configurar en Railway

1. Ve a tu proyecto en [railway.app](https://railway.app)
2. Click en tu servicio
3. Ve a la pestaña "Variables"
4. Click en "New Variable"
5. **Nombre:** `FIREBASE_CREDENTIALS`
6. **Valor:** Pega el JSON minificado (debe ser una sola línea sin saltos)
7. Click en "Add" o "Save"

### 4. Verificar el despliegue

Railway redesplegará automáticamente. En los logs deberías ver:
```
Usando credenciales de Firebase desde variable de entorno
```

Si ves esto, significa que funcionó correctamente.

## Formato esperado de la variable

La variable debe ser un JSON válido en una sola línea, como:
```
{"type":"service_account","project_id":"supratechweb","private_key_id":"...","private_key":"-----BEGIN PRIVATE KEY-----\n...","client_email":"...","client_id":"..."}
```

**IMPORTANTE:** No incluyas saltos de línea dentro de los valores de las llaves, excepto en `private_key` donde los `\n` son parte del string.

## Troubleshooting

### Error: "Invalid service account credentials"
- Verifica que copiaste TODO el JSON completo
- Asegúrate de que no haya caracteres extra al inicio o final
- Verifica que los `\n` en la `private_key` estén presentes

### Error: "Unexpected token" o JSON parse error
- El JSON debe estar en una sola línea
- No debe tener comillas extra alrededor
- Usa un validador JSON para verificar: https://jsonlint.com/

### La app sigue sin funcionar
- Verifica los logs de Railway
- Asegúrate de que la variable se llame exactamente `FIREBASE_CREDENTIALS`
- Railway redespliega automáticamente al agregar variables, espera unos segundos
