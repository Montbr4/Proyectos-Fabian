const SPREADSHEET_ID = "";
const HOJA_VENTAS = "Ventas";
const HOJA_DELIVERY = "Delivery";
const HOJA_ANALISIS = "Analisis";
const HOJA_PRODUCTOS = "Productos";
const HOJA_MOVIMIENTOS = "Movimientos";

function doGet(e) {
  var userAgent = e.parameter.ua || "PC";
  return HtmlService.createTemplateFromFile("Loader")
    .evaluate()
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setTitle("Bobbastik POS")
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getView(viewName) {
  return HtmlService.createHtmlOutputFromFile(viewName).getContent();
}

function inicializarHoja(nombreHoja, tipo) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(nombreHoja);
  
  if (!sheet) {
    sheet = ss.insertSheet(nombreHoja);
    let headers, color;

    if (tipo === "RESUMEN") {
      headers = [["ID", "Fecha", "Hora", "Cliente", "Observaciones", "Detalle Orden", "Monto Total", "Método Pago", "Estado"]];
      color = "#2E86AB"; 
      sheet.setColumnWidth(6, 300); 
      sheet.setColumnWidth(5, 200); 
    } else if (tipo === "ANALISIS") {
      headers = [["ID Venta", "Fecha", "Canal", "Sabor", "Unidades", "Total"]];
      color = "#27AE60"; 
    } else if (tipo === "MOVIMIENTOS") { 
      headers = [["FECHA", "CONCEPTO", "MONTO"]];
      color = "#F1C40F"; 
    }

    sheet.getRange(1, 1, 1, headers[0].length).setValues(headers);
    sheet.getRange(1, 1, 1, headers[0].length).setBackground(color).setFontColor("white").setFontWeight("bold");
  }
  return sheet;
}

function guardarVenta(venta) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const fechaActual = new Date();
    const fecha = Utilities.formatDate(fechaActual, Session.getScriptTimeZone(), "dd/MM/yyyy");
    const hora = Utilities.formatDate(fechaActual, Session.getScriptTimeZone(), "HH:mm:ss");

    const nombreHojaDestino = (venta.canal === "DELIVERY") ? HOJA_DELIVERY : HOJA_VENTAS;
    let sheetResumen = ss.getSheetByName(nombreHojaDestino);
    if (!sheetResumen) sheetResumen = inicializarHoja(nombreHojaDestino, "RESUMEN");
    
    const detalleStr = venta.items.map(item => `${item.cantidad}x ${item.nombre}`).join(", ");
    
    sheetResumen.appendRow([
      venta.id, fecha, hora, venta.cliente, venta.observaciones, detalleStr, venta.total, venta.metodoPago, venta.estadoPago
    ]);

    let sheetAnalisis = ss.getSheetByName(HOJA_ANALISIS);
    if (!sheetAnalisis) sheetAnalisis = inicializarHoja(HOJA_ANALISIS, "ANALISIS");

    venta.items.forEach(item => {
      const subtotalLinea = item.cantidad * item.precio;
      sheetAnalisis.appendRow([
        venta.id, fecha, venta.canal, item.nombre, item.cantidad, subtotalLinea
      ]);
    });
    
    return { success: true, message: `Guardado en ${nombreHojaDestino}` };
  } catch (error) {
    return { success: false, message: "Error: " + error.message };
  }
}

function guardarMovimiento(mov) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(HOJA_MOVIMIENTOS);
    if (!sheet) sheet = inicializarHoja(HOJA_MOVIMIENTOS, "MOVIMIENTOS");

    const fechaFormat = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM");
    
    let montoFinal = parseFloat(mov.monto);
    if (mov.tipo === "SALIDA") {
      montoFinal = montoFinal * -1;
    }

    sheet.appendRow([fechaFormat, mov.concepto, montoFinal]);

    const lastRow = sheet.getLastRow();
    const cellMonto = sheet.getRange(lastRow, 3);
    
    if (mov.tipo === "SALIDA") {
      cellMonto.setBackground("#ea9999");
      cellMonto.setFontColor("black");   
    } else {
      cellMonto.setBackground("#b6d7a8");
      cellMonto.setFontColor("black");   
    }

    return { success: true };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function inicializarProductos() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(HOJA_PRODUCTOS);
  if (!sheet) {
    sheet = ss.insertSheet(HOJA_PRODUCTOS);
    sheet.appendRow(["ID", "Nombre", "Precio"]);
    sheet.getRange(1, 1, 1, 3).setBackground("#2E86AB").setFontColor("white").setFontWeight("bold");
    const productosBase = [
      [1, 'Maracuya', 8.00], [2, 'Fresa', 8.00], [3, 'Mango', 8.00],
      [4, 'Limon', 8.00], [5, 'Te con leche', 8.00], [6, 'Taro', 8.50],
      [7, 'Te negro', 4.00], [8, 'Te jazmin', 4.00], [9, 'Mango con Leche', 8.50],
      [10, 'Fresa con Leche', 8.50], [11, 'Bobba Extra', 2.00]
    ];
    sheet.getRange(2, 1, productosBase.length, 3).setValues(productosBase);
  }
}

function obtenerProductos() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(HOJA_PRODUCTOS);
  if (!sheet) { inicializarProductos(); sheet = ss.getSheetByName(HOJA_PRODUCTOS); }
  const data = sheet.getDataRange().getValues();
  const productos = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] !== "") {
      productos.push({ id: data[i][0], nombre: data[i][1], precio: parseFloat(data[i][2]) });
    }
  }
  return productos;
}

function guardarProducto(producto) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(HOJA_PRODUCTOS);
  if (!sheet) { inicializarProductos(); sheet = ss.getSheetByName(HOJA_PRODUCTOS); }
  const data = sheet.getDataRange().getValues();
  let rowToUpdate = -1;
  if (producto.id) {
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == producto.id) { rowToUpdate = i + 1; break; }
    }
  }
  if (rowToUpdate > 0) {
    sheet.getRange(rowToUpdate, 2).setValue(producto.nombre);
    sheet.getRange(rowToUpdate, 3).setValue(producto.precio);
  } else {
    const newId = Date.now(); 
    sheet.appendRow([newId, producto.nombre, producto.precio]);
  }
  return { success: true };
}

function eliminarProducto(id) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(HOJA_PRODUCTOS);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == id) { sheet.deleteRow(i + 1); return { success: true }; }
  }
  return { success: false, message: "Producto no encontrado" };
}
