# 📹 Guía para Agregar Videos a la Página Demo

## 🎯 Resumen Rápido

1. **Descarga ScreenToGif**: https://www.screentogif.com/
2. **Graba 6 videos** de 10-15 segundos cada uno
3. **Exporta como WebM** (File > Save as > WebM)
4. **Sube a Cloudinary**: https://cloudinary.com/console
5. **Copia las URLs** de tus videos
6. **Busca los comentarios `<!-- 📹 REEMPLAZA -->`** en `demo.html`
7. **Reemplaza los placeholders** con el código de video

---

## 🎬 Paso 1: Grabar tus Videos

### Herramientas Recomendadas (Gratis):

#### Opción A: ScreenToGif (⭐ Recomendado)
1. Descarga: https://www.screentogif.com/
2. Abre la aplicación
3. Clic en "Recorder"
4. Selecciona el área que quieres grabar (tu navegador con el sistema abierto)
5. Clic en "Record" (F7)
6. Realiza las acciones que quieres mostrar (10-15 segundos máximo)
7. Clic en "Stop" (F8)
8. Edita si es necesario (puedes eliminar frames innecesarios)
9. Exporta como:
   - **WebM** (más ligero, mejor calidad) - File > Save as > WebM
   - **GIF** (compatible con todo) - File > Save as > GIF

#### Opción B: OBS Studio
1. Descarga: https://obsproject.com/
2. Configura una escena con "Window Capture" de tu navegador
3. Graba en MP4
4. Convierte a WebM usando: https://cloudconvert.com/mp4-to-webm

### Configuración Recomendada:
- **Duración**: 10-20 segundos por video
- **Resolución**: 1920x1080 o 1280x720
- **FPS**: 30 fps
- **Tamaño**: Menos de 5MB por video

---

## 📤 Paso 2: Subir Videos a Cloudinary

### Método 1: Interfaz Web
1. Ve a https://cloudinary.com/console
2. Inicia sesión con tu cuenta
3. Clic en "Media Library"
4. Arrastra y suelta tus videos
5. Copia la URL del video (clic derecho > "Copy URL")

### Método 2: URL directa
Tu URL base de Cloudinary: `https://res.cloudinary.com/dipt3jq6r/`

Ejemplo de URL de video:
```
https://res.cloudinary.com/dipt3jq6r/video/upload/v1234567890/panel-demo.webm
```

---

## 🎨 Paso 3: Reemplazar Placeholders en demo.html

### Formato para VIDEO (WebM/MP4):

Busca este código en `templates/demo.html`:

```html
<!-- ANTES (Placeholder) -->
<div class="preview-screenshot">
  <div class="preview-badge">Panel de Usuario</div>
  <div class="preview-placeholder">
    <div class="preview-placeholder-content">
      <div>🏠</div>
      <p>Panel Principal</p>
    </div>
  </div>
</div>
```

Reemplaza con:

```html
<!-- DESPUÉS (Video) -->
<div class="preview-screenshot">
  <div class="preview-badge">Panel de Usuario</div>
  <video autoplay loop muted playsinline>
    <source src="https://res.cloudinary.com/dipt3jq6r/video/upload/v1234567890/panel-demo.webm" type="video/webm">
    <source src="https://res.cloudinary.com/dipt3jq6r/video/upload/v1234567890/panel-demo.mp4" type="video/mp4">
  </video>
</div>
```

### Formato para GIF:

```html
<!-- DESPUÉS (GIF) -->
<div class="preview-screenshot">
  <div class="preview-badge">Panel de Usuario</div>
  <img src="https://res.cloudinary.com/dipt3jq6r/image/upload/v1234567890/panel-demo.gif" alt="Panel Principal">
</div>
```

### Formato para IMAGEN (Screenshot estática):

```html
<!-- DESPUÉS (Imagen) -->
<div class="preview-screenshot">
  <div class="preview-badge">Panel de Usuario</div>
  <img src="https://res.cloudinary.com/dipt3jq6r/image/upload/v1234567890/panel-screenshot.png" alt="Panel Principal">
</div>
```

---

## 📋 Lista de Videos a Grabar

Graba estos 6 videos/capturas de tu sistema:

### 1. 🏠 Panel Principal
- **Qué mostrar**: Login → Panel principal con todos los botones
- **Duración**: 10-15 segundos
- **Archivo**: `panel-demo.webm`

### 2. 🛒 Módulo de Compras
- **Qué mostrar**: Abrir Compras → Mostrar las opciones (Cotizaciones, Pedidos, Indicadores)
- **Duración**: 10-15 segundos
- **Archivo**: `compras-demo.webm`

### 3. 📊 Bases de Datos
- **Qué mostrar**: Abrir BASE+ o BD MARCAS → Mostrar la tabla editable
- **Duración**: 10-15 segundos
- **Archivo**: `bases-datos-demo.webm`

### 4. 💰 Sistema de Descuentos
- **Qué mostrar**: Abrir Descuentos → Mostrar Errores o Promocionables
- **Duración**: 10-15 segundos
- **Archivo**: `descuentos-demo.webm`

### 5. 🔄 Automatización
- **Qué mostrar**: Abrir ACTUALIZAR → Mostrar los botones de automatización
- **Duración**: 10-15 segundos
- **Archivo**: `automatizacion-demo.webm`

### 6. 📈 Indicadores
- **Qué mostrar**: Abrir Indicadores → Mostrar los gráficos/datos
- **Duración**: 10-15 segundos
- **Archivo**: `indicadores-demo.webm`

---

## 💡 Consejos para Grabar

1. **Limpia tu navegador**: Cierra pestañas innecesarias
2. **Pantalla completa**: Usa F11 para ocultar la barra de direcciones
3. **Movimientos lentos**: Mueve el mouse despacio para que se vea bien
4. **Sin datos sensibles**: Asegúrate de no mostrar información confidencial
5. **Buenos datos de ejemplo**: Usa datos de prueba que se vean profesionales

---

## 🚀 Ejemplo Completo

Aquí está un ejemplo completo de cómo se vería la sección del Panel Principal con video:

```html
<div class="preview-item">
  <h3>🏠 Panel Principal</h3>
  <p>Interfaz intuitiva con acceso rápido a todos los módulos. Visualiza tu información de usuario y navega entre las diferentes secciones del sistema.</p>
  <div class="preview-screenshot">
    <div class="preview-badge">Panel de Usuario</div>
    <video autoplay loop muted playsinline>
      <source src="https://res.cloudinary.com/dipt3jq6r/video/upload/v1737586000/panel-demo.webm" type="video/webm">
      <source src="https://res.cloudinary.com/dipt3jq6r/video/upload/v1737586000/panel-demo.mp4" type="video/mp4">
      Tu navegador no soporta videos HTML5.
    </video>
  </div>
</div>
```

---

## 🔍 Ubicación en demo.html

Los placeholders a reemplazar están entre las líneas **479-567** en `templates/demo.html`

Busca el comentario: `<!-- Preview Section -->`

---

## ✅ Checklist Final

- [ ] Grabar 6 videos (10-15 segundos cada uno)
- [ ] Convertir a WebM (o dejar como GIF)
- [ ] Subir a Cloudinary
- [ ] Copiar URLs
- [ ] Reemplazar placeholders en demo.html
- [ ] Probar en navegador local
- [ ] ¡Listo para mostrar a clientes! 🎉

---

## ❓ Preguntas Frecuentes

**P: ¿GIF o WebM?**
R: WebM es mejor (más ligero, mejor calidad), pero GIF funciona en todos lados.

**P: ¿Puedo usar MP4?**
R: Sí, pero WebM es más ligero. Incluye ambos formatos para compatibilidad.

**P: ¿Y si el video es muy pesado?**
R: Reduce la resolución a 720p, baja los FPS a 24, o acorta la duración.

**P: ¿Se reproduce automáticamente?**
R: Sí, con los atributos `autoplay loop muted playsinline`.

**P: ¿Puedo mezclar videos e imágenes?**
R: ¡Claro! Algunas secciones pueden tener videos y otras imágenes estáticas.
