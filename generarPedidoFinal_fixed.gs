function doGet(e) {
  try {
    var result = generarPedidoFinal();

    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      message: 'Pedido final generado correctamente',
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

function generarPedidoFinal() {
  // IMPORTANTE: Reemplaza 'TU_SPREADSHEET_ID' con el ID real de tu hoja de cálculo
  // El ID está en la URL: https://docs.google.com/spreadsheets/d/AQUI_ESTA_EL_ID/edit
  var SPREADSHEET_ID = 'TU_SPREADSHEET_ID'; // <--- CAMBIAR ESTO

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var hojaMarcas = ss.getSheetByName("MARCAS A PEDIR");
  var hojaPedidos = ss.getSheetByName("PEDIDOS PRECARGADOS");
  var hojaFinal = ss.getSheetByName("PEDIDO FINAL");

  hojaFinal.clear();

  var marcasAPedir = hojaMarcas.getRange("A2:A").getValues()
    .flat()
    .map(marca => marca ? marca.trim() : "")
    .filter(String);

  var dataPedidos = hojaPedidos.getDataRange().getValues();
  var numColumnas = dataPedidos[0].length;
  var numFilas = dataPedidos.length;

  var resultado = [];

  for (var col = 0; col < numColumnas; col += 5) {
    var marca = dataPedidos[0][col] ? dataPedidos[0][col].toString().trim() : "";

    if (marcasAPedir.includes(marca)) {
      for (var fila = 2; fila < numFilas; fila++) {
        var filaDatos = dataPedidos[fila].slice(col, col + 4);

        if (filaDatos.every(celda => !celda || celda.toString().trim() === "")) {
          break;
        }

        resultado.push(filaDatos);
      }
      resultado.push(["", "", "", ""]);
    }
  }

  if (resultado.length > 0) {
    hojaFinal.getRange(1, 1, resultado.length, 4).setValues(resultado);
  }

  return {
    filasGeneradas: resultado.length,
    marcasProcesadas: marcasAPedir.length
  };
}
