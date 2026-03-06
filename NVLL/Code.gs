function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('NVLL')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (name === 'Categorias') {
      sheet.appendRow(['Nombre']);
      ['GENERAL', 'SKINCARE', 'MAQUILLAJE', 'CABELLO', 'HOGAR', 'BEBÉS'].forEach(c => sheet.appendRow([c]));
    }
    if (name === 'Productos') {
       sheet.appendRow(['ID', 'Nombre', 'Precio', 'ImagenURL', 'Descripcion', 'Agotado', 'Orden', 'Categoria', 'Oculto']);
    }
  }
  return sheet;
}

function obtenerCategorias() {
  const sheet = getSheet('Categorias');
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return ['GENERAL', 'SKINCARE', 'MAQUILLAJE', 'CABELLO', 'HOGAR'];
  return sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat().filter(c => c !== "");
}

function guardarCategoria(nombre) {
  const sheet = getSheet('Categorias');
  const currentCats = obtenerCategorias();
  if (!currentCats.includes(nombre.toUpperCase())) {
    sheet.appendRow([nombre.toUpperCase()]);
  }
  return obtenerCategorias();
}

function eliminarCategoria(nombre) {
  const sheet = getSheet('Categorias');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === nombre) {
      sheet.deleteRow(i + 1);
      break;
    }
  }
  return obtenerCategorias();
}

function obtenerProductos() {
  try {
    const sheet = getSheet('Productos');
    const lastRow = sheet.getLastRow();
    
    if (lastRow <= 1) return []; 
    
    const numCols = sheet.getLastColumn();
    const colsToRead = numCols < 9 ? numCols : 9;
    
    const data = sheet.getRange(2, 1, lastRow - 1, colsToRead).getValues();
    
    const productos = data
      .filter(row => row[0] !== "")
      .map((row) => {
        return {
          id: row[0],
          nombre: row[1] || "Sin Nombre",
          precio: row[2],
          imagen: row[3], 
          descripcion: row[4] || "",
          agotado: row[5],
          orden: row[6],
          categoria: row[7] ? row[7] : "GENERAL",
          oculto: row[8] ? row[8] : false 
        };
      });
    
    return productos.sort((a, b) => {
      if (a.agotado !== b.agotado) return a.agotado ? 1 : -1;
      
      const ordenA = (a.orden === "" || a.orden === null) ? 999999 : Number(a.orden);
      const ordenB = (b.orden === "" || b.orden === null) ? 999999 : Number(b.orden);
      
      if (ordenA !== ordenB) return ordenA - ordenB;
      return b.id - a.id;
    });

  } catch (e) {
    Logger.log(e);
    return []; 
  }
}

function guardarOrdenMasivo(listaIds) {
  try {
    const sheet = getSheet('Productos');
    const data = sheet.getDataRange().getValues();
    const mapIdFila = {};
    for (let i = 1; i < data.length; i++) { mapIdFila[data[i][0]] = i + 1; }

    listaIds.forEach((id, index) => {
      const fila = mapIdFila[id];
      if (fila) {
        sheet.getRange(fila, 7).setValue(index + 1);
      }
    });
    return { success: true };
  } catch(e) { return { success: false, error: e.toString() }; }
}

function procesarProducto(info) {
  try {
    const sheet = getSheet('Productos');
    const FOLDER_ID = ""; 
    
    let fileIdOrUrl = info.imagenActual;
    
    if (info.imagenBase64) {
      const partes = info.imagenBase64.split(","); 
      const tipoMime = partes[0].split(":")[1].split(";")[0];
      const dataPura = partes[1];
      const blob = Utilities.newBlob(Utilities.base64Decode(dataPura), tipoMime, "producto_" + new Date().getTime());
      const folder = DriveApp.getFolderById(FOLDER_ID);
      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      fileIdOrUrl = file.getId();
    }

    const catFinal = info.categoria || "GENERAL";

    if (info.id) {
      const data = sheet.getDataRange().getValues();
      let rowToEdit = -1;
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] == info.id) {
          rowToEdit = i + 1;
          break;
        }
      }
      if (rowToEdit > 0) {
        sheet.getRange(rowToEdit, 2).setValue(info.nombre);
        sheet.getRange(rowToEdit, 3).setValue(info.precio);
        sheet.getRange(rowToEdit, 4).setValue(fileIdOrUrl);
        sheet.getRange(rowToEdit, 5).setValue(info.descripcion);
        sheet.getRange(rowToEdit, 6).setValue(info.agotado);
        sheet.getRange(rowToEdit, 8).setValue(catFinal);
        sheet.getRange(rowToEdit, 9).setValue(info.oculto);
      }
    } else {
      let newId = 1;
      const lastRow = sheet.getLastRow();
      let maxOrden = 0;
      if (lastRow > 1) {
        const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
        const ordenes = sheet.getRange(2, 7, lastRow - 1, 1).getValues().flat();
        
        const lastId = Math.max(...ids.filter(n => !isNaN(n)));
        if (lastId > 0) newId = lastId + 1;

        const validOrdenes = ordenes.filter(n => !isNaN(parseFloat(n)) && isFinite(n));
        if (validOrdenes.length > 0) maxOrden = Math.max(...validOrdenes);
      }
      sheet.appendRow([newId, info.nombre, info.precio, fileIdOrUrl, info.descripcion, info.agotado, maxOrden + 1, catFinal, info.oculto]);
    }
    return { success: true };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function borrarProducto(id) {
  try {
    const sheet = getSheet('Productos');
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == id) {
        sheet.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, error: "No encontrado" };
  } catch(e) { return { success: false, error: e.toString() }; }
}
