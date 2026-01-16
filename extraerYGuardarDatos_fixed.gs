function doGet(e) {
  try {
    var result = extraerYGuardarDatos();

    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      message: 'Datos extraídos y guardados correctamente',
      data: result
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: error.toString(),
      stack: error.stack
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function extraerYGuardarDatos() {
  // IMPORTANTE: Reemplaza con el ID del spreadsheet principal (donde está la hoja "historial")
  // Encuentra el ID en la URL: https://docs.google.com/spreadsheets/d/ESTE_ES_EL_ID/edit
  const SPREADSHEET_ID = 'TU_SPREADSHEET_ID_AQUI'; // <--- CAMBIAR ESTO

  const TIEMPO_MAX = 280000;
  const inicio = Date.now();

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let hojaHistorial = ss.getSheetByName("historial");
  let hojaErrores = ss.getSheetByName("errores");

  if (!hojaHistorial) hojaHistorial = ss.insertSheet("historial");
  if (!hojaErrores) hojaErrores = ss.insertSheet("errores");
  else hojaErrores.clear();

  hojaErrores.getRange("A1:C1").setValues([["Enlace", "Error", "Fila"]]);
  hojaHistorial.getRange("A2:D").clearContent();

  const ultimaFila = hojaHistorial.getLastRow();
  const links = hojaHistorial.getRange("E2:E" + ultimaFila).getValues().flat().filter(String);
  const hojas = hojaHistorial.getRange("F2:F" + ultimaFila).getValues().flat();

  if (links.length === 0 && hojas.length === 1) {
    return {
      mensaje: "No hay links para procesar",
      datosExtraidos: 0,
      errores: 0
    };
  }

  if (links.length !== hojas.length) {
    Logger.log(`Diferente cantidad de links (${links.length}) y hojas (${hojas.length})`);
    hojaErrores.getRange(2, 1, 1, 3).setValues([["", "Error: El número de filas no coincide", ""]]);
    return {
      error: "El número de filas no coincide",
      links: links.length,
      hojas: hojas.length
    };
  }

  const datos = [];
  const errores = [];

  for (let index = 0; index < links.length; index++) {
    if (Date.now() - inicio > TIEMPO_MAX) {
      errores.push(["", "Tiempo máximo alcanzado. Proceso interrumpido.", ""]);
      Logger.log("Proceso interrumpido por límite de tiempo.");
      break;
    }

    const link = links[index];
    const nombreHoja = hojas[index];
    const filaOrigen = index + 2;

    try {
      const docId = extraerIdDeUrl(link);
      const documento = SpreadsheetApp.openById(docId);

      if (!nombreHoja) throw new Error("Nombre de hoja vacío");
      const hoja = documento.getSheetByName(nombreHoja);
      if (!hoja) throw new Error(`No existe hoja: ${nombreHoja}`);

      const valores = hoja.getRange("A8:E").getValues();

      for (let i = 0; i < valores.length; i++) {
        const fila = valores[i];
        if (!fila[0]) break;

        datos.push([fila[1] || "", fila[0] || "", fila[4] || "", nombreHoja]);
      }

    } catch (e) {
      Logger.log(`Error en fila ${filaOrigen}: ${e.message}`);
      errores.push([link, e.message, filaOrigen]);
    }
  }

  if (datos.length > 0) {
    hojaHistorial.getRange(2, 1, datos.length, 4).setValues(datos);
  }

  if (errores.length > 0) {
    hojaErrores.getRange(2, 1, errores.length, 3).setValues(errores);
  }

  return {
    datosExtraidos: datos.length,
    erroresEncontrados: errores.length,
    linksProcessados: links.length,
    tiempoEjecucion: (Date.now() - inicio) / 1000 + " segundos"
  };
}

function extraerIdDeUrl(url) {
  const match = url.match(/\/d\/([a-zA-Z0-9-_]+)/);
  if (match && match[1]) return match[1];
  throw new Error("URL no válida: " + url);
}
