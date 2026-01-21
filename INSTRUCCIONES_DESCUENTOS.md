# Instrucciones para Implementar Sistema de Descuentos

## Resumen
Este documento contiene las instrucciones completas para implementar el sistema de descuentos con 6 páginas HTML y sus respectivos endpoints del backend.

## Estructura de Archivos

### HTMLs Necesarios:
1. ✅ `ventas_semanales.html` - YA CREADO
2. ⏳ `para_impulsar.html` - PENDIENTE
3. ⏳ `para_descartar.html` - PENDIENTE
4. ⏳ `para_poner_en_venta.html` - PENDIENTE
5. ⏳ `promocionables.html` - PENDIENTE (tabla de lectura)
6. ⏳ `errores.html` - PENDIENTE (tabla de lectura + plantilla con errores marcados)

---

## Configuración de Google Sheets

### Hoja BDPROMOTE
- **URL:** https://docs.google.com/spreadsheets/d/14F6ZSyrhp9_f6tHYz6GYaIVqoAEdZo6UICJP0_GR7ew

#### Sub-hojas:
1. **VENTAS SEMANALES**
   - Columnas a llenar:
     - `SKU de la publicación` (D2:D)
     - `Unidades vendidas` (I2:I)
   - También duplicar SKU en: `Numero de publicacion` (A2:A)

2. **PARA IMPULSAR VENTAS**
   - Columnas a llenar:
     - `SKU` (C2:C)
     - `Unidades para impulsar ventas` (L2:L)
     - `Ventas últimos 30 días (u.)` (I2:I)
   - También duplicar SKU en: `# Publicación` (D2:D)

3. **PARA EVITAR DESCARTE**
   - Columnas a llenar:
     - `SKU` (C3:C)
   - También duplicar SKU en: `# Publicación` (D3:D)

4. **PARA PONER EN VENTA**
   - Columnas a llenar:
     - `SKU` (C2:C)
   - También duplicar SKU en: `# Publicación` (D2:D)

### Hoja PROMOTE5.0
- **URL:** https://docs.google.com/spreadsheets/d/1nOB3Lr07FqOcFKj-1dxNCRnUxGOjUoqAf_WterqH3rg

#### Sub-hojas:
1. **PROMOTE 5.0** (Para tabla PROMOCIONABLES - SOLO LECTURA)
   - Columnas a mostrar:
     - %DESCUENTO (G2:G)
     - CATEGORIA (A2:A)
     - ID (B2:B)
     - MARCA (C2:C)
     - PRECIO (H2:H)
     - PRECIOOFERTA (I2:I)
     - RANGO (L2:L)
     - UTILIDAD (K2:K)
     - UTILIDADMONEDA (J2:J)

2. **ID ERROR** (Para tabla ERRORES - SOLO LECTURA)
   - Columnas a mostrar:
     - COMISION (C2:C)
     - COSTO (B2:B)
     - ENVIO (D2:D)
     - ID (A2:A)
   - Plantilla de descarga debe tener:
     - `SKU, MARCA, COSTO, PRECIO, ¿ENVIO?, ENVIO`
   - Marcar en ROJO las celdas con "ERROR" en Excel
   - En CSV remarcar de alguna forma

---

## Endpoints del Backend Necesarios

### Para VENTAS SEMANALES
```python
@app.route('/api/ventas_semanales/add', methods=['POST'])
@app.route('/api/ventas_semanales/bulk', methods=['POST'])
```

### Para PARA IMPULSAR
```python
@app.route('/api/para_impulsar/add', methods=['POST'])
@app.route('/api/para_impulsar/bulk', methods=['POST'])
```

### Para PARA DESCARTAR
```python
@app.route('/api/para_descartar/add', methods=['POST'])
@app.route('/api/para_descartar/bulk', methods=['POST'])
```

### Para PARA PONER EN VENTA
```python
@app.route('/api/para_poner_en_venta/add', methods=['POST'])
@app.route('/api/para_poner_en_venta/bulk', methods=['POST'])
```

### Para PROMOCIONABLES (Solo lectura)
```python
@app.route('/api/promocionables/data', methods=['GET'])
```

### Para ERRORES (Solo lectura + plantilla)
```python
@app.route('/api/errores/data', methods=['GET'])
@app.route('/api/errores/plantilla_csv', methods=['GET'])
@app.route('/api/errores/plantilla_excel', methods=['GET'])
```

---

## Funcionalidad de cada Página

### 1. VENTAS SEMANALES ✅
- **Formulario:** SKU + Unidades vendidas
- **CSV:** Descargable con ejemplo
- **Carga masiva:** CSV/XLSX
- **Backend:** Escribir en D2:D, I2:I y duplicar SKU en A2:A

### 2-4. PARA IMPULSAR / DESCARTAR / PONER EN VENTA
- Similar a VENTAS SEMANALES pero con sus campos específicos
- Cada uno duplica SKU en su respectiva columna

### 5. PROMOCIONABLES
- **Tabla de solo lectura** con 9 columnas
- **Botones de descarga:** CSV y Excel
- No permite edición

### 6. ERRORES
- **Tabla de solo lectura** con 4 columnas
- **Plantilla descargable:**
  - CSV: con marcas especiales en datos faltantes
  - Excel: celdas en ROJO donde hay "ERROR"
- No permite edición

---

## Próximos Pasos

Necesito que me confirmes si:
1. ¿Quieres que continúe creando TODOS los archivos HTML restantes?
2. ¿Quieres que también te cree los endpoints del backend en Python?
3. ¿Prefieres hacerlo tú mismo siguiendo este documento?

Dime cómo prefieres proceder y continuaré con lo que necesites.
