# Despliegue en Railway - SupratechWEB

## Archivos de configuración creados

- `Procfile` - Define el comando para iniciar la aplicación
- `railway.json` - Configuración específica de Railway
- `runtime.txt` - Especifica la versión de Python
- `requirements.txt` - Dependencias de Python actualizadas
- `.gitignore` - Archivos a ignorar en git

## Pasos para desplegar en Railway

### 1. Preparar el repositorio Git

Si aún no tienes un repositorio git, inicialízalo:

```bash
cd C:\Users\PC 1\SUPRATECHWEB
git init
git add .
git commit -m "Preparar para despliegue en Railway"
```

Si ya tienes un repositorio, actualiza los cambios:

```bash
git add .
git commit -m "Actualizar configuración para Railway"
git push
```

### 2. Crear proyecto en Railway

1. Ve a [railway.app](https://railway.app)
2. Inicia sesión con GitHub
3. Click en "New Project"
4. Selecciona "Deploy from GitHub repo"
5. Selecciona tu repositorio SUPRATECHWEB
6. Railway detectará automáticamente que es una aplicación Python

### 3. Subir el archivo de credenciales de Firebase

**IMPORTANTE:** El archivo `supratechweb-firebase-adminsdk-fbsvc-8d4aa68a75.json` NO debe estar en git por seguridad.

En Railway:
1. Ve a tu proyecto
2. Click en "Variables"
3. Click en "RAW Editor"
4. Copia el contenido completo del archivo JSON de credenciales
5. O mejor aún, usa Railway CLI para subir el archivo:

```bash
# Instalar Railway CLI
npm i -g @railway/cli

# Login
railway login

# Link al proyecto
railway link

# Agregar el archivo como variable
railway run echo "Subir archivo manualmente desde el dashboard"
```

**Método alternativo - Subir archivo directamente:**
1. En el dashboard de Railway, ve a "Settings"
2. Busca la sección de "Volumes" o usa variables de entorno
3. Sube el archivo JSON directamente

**O usar variables de entorno (más seguro):**
Convierte el contenido del JSON a una variable de entorno:
1. Copia todo el contenido de `supratechweb-firebase-adminsdk-fbsvc-8d4aa68a75.json`
2. En Railway, crea una variable: `FIREBASE_CREDENTIALS`
3. Pega el contenido JSON completo

Luego modifica `app.py` para leer de la variable de entorno (opcional).

### 4. Variables de entorno necesarias

Railway automáticamente proporciona:
- `PORT` - Puerto donde correrá la app

No necesitas configurar nada más si subes el archivo JSON directamente.

### 5. Despliegue

Railway automáticamente:
1. Detecta Python usando `requirements.txt`
2. Instala las dependencias
3. Ejecuta el comando en `Procfile`: `gunicorn app:app`
4. Asigna un dominio público

### 6. Verificar el despliegue

1. Railway te dará una URL como: `https://tu-proyecto.up.railway.app`
2. Verifica que la app esté corriendo visitando la URL
3. Revisa los logs en Railway dashboard para cualquier error

### 7. Dominio personalizado (Opcional)

1. En Railway, ve a "Settings"
2. En "Domains", añade tu dominio personalizado
3. Configura los DNS según las instrucciones

## Troubleshooting

### Error: No module named 'X'
- Asegúrate de que `requirements.txt` tenga todas las dependencias
- Verifica los logs de build en Railway

### Error: Firebase credentials not found
- Verifica que el archivo JSON esté en el root del proyecto
- O que la variable de entorno `FIREBASE_CREDENTIALS` esté configurada

### Error: Port already in use
- Railway maneja esto automáticamente con la variable `PORT`
- No necesitas configurar nada

### App no inicia
- Revisa los logs en Railway dashboard
- Verifica que `Procfile` tenga el comando correcto: `web: gunicorn app:app`

## Notas importantes

- El archivo `supratechweb-firebase-adminsdk-fbsvc-8d4aa68a75.json` NO debe estar en git
- Railway tiene capa gratuita con $5 USD de crédito mensual
- El despliegue es automático cada vez que hagas push a GitHub
- Los logs están disponibles en tiempo real en el dashboard

## Comandos útiles

```bash
# Ver logs en tiempo real
railway logs

# Abrir la app en el navegador
railway open

# Ver variables de entorno
railway variables

# Ejecutar comando en el servidor
railway run [comando]
```
