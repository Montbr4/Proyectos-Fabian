function doGet(e) {
  let page = (e && e.parameter) ? e.parameter.p : null;
  
  if (page === 'biografias') {
    try { return HtmlService.createHtmlOutputFromFile('Biografias').setTitle('Free Ride - Nuestro Equipo').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL).addMetaTag('viewport', 'width=device-width, initial-scale=1'); } 
    catch(err) { return HtmlService.createHtmlOutputFromFile('biografias').setTitle('Free Ride - Nuestro Equipo').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL).addMetaTag('viewport', 'width=device-width, initial-scale=1'); }
  } else if (page === 'filmmaker') {
    try { return HtmlService.createHtmlOutputFromFile('Filmmaker').setTitle('Free Ride - Portal Filmmaker').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL).addMetaTag('viewport', 'width=device-width, initial-scale=1'); } 
    catch(err) { return HtmlService.createHtmlOutputFromFile('filmmaker').setTitle('Free Ride - Portal Filmmaker').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL).addMetaTag('viewport', 'width=device-width, initial-scale=1'); }
  } else if (page === 'surfistas') {
    try { return HtmlService.createHtmlOutputFromFile('Surfistas').setTitle('Free Ride - Surfistas').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL).addMetaTag('viewport', 'width=device-width, initial-scale=1'); } 
    catch(err) { return HtmlService.createHtmlOutputFromFile('surfistas').setTitle('Free Ride - Surfistas').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL).addMetaTag('viewport', 'width=device-width, initial-scale=1'); }
  }
  
  try { return HtmlService.createHtmlOutputFromFile('index').setTitle('Free Ride').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL).addMetaTag('viewport', 'width=device-width, initial-scale=1'); } 
  catch(err) { return HtmlService.createHtmlOutputFromFile('Index').setTitle('Free Ride').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL).addMetaTag('viewport', 'width=device-width, initial-scale=1'); }
}

function getScriptUrl() { return ScriptApp.getService().getUrl(); }

function getSheet(name) {
  try {
    const ss = SpreadsheetApp.openById('');
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
      if (name === 'Categorias') {
        sheet.appendRow(['Nombre']);
        ['MAKAHA', 'PUNTA ROQUITAS', 'WAIKIKI', 'PAMPILLA', 'HERRADURA', 'SAN BARTOLO', 'PUNTA HERMOSA'].forEach(c => sheet.appendRow([c]));
      }
      if (name === 'Productos') { sheet.appendRow(['ID', 'Surfista', 'Filmmaker', 'ImagenURL', 'Descripcion', 'Orden', 'Playa', 'Oculto']); }
      if (name === 'Equipo') { sheet.appendRow(['ID', 'Nombre', 'Rol', 'Biografia', 'ImagenURL', 'Instagram']); }
      if (name === 'Merch') { sheet.appendRow(['ID', 'Nombre', 'Precio', 'ImagenURL', 'Descripcion', 'Oculto', 'Agotado']); }
      if (name === 'Usuarios') { sheet.appendRow(['ID', 'Email', 'Password', 'NombreFilmmaker']); }
      if (name === 'Surfistas') { sheet.appendRow(['ID', 'Nombre', 'Numero', 'ImagenURL', 'Posicion', 'Redes', 'Nivel', 'PlayaFavorita']); }
    }
    return sheet;
  } catch (e) { throw new Error("Error accediendo al Sheet: " + e.message); }
}

function validarAcceso(email, pass) {
  if (!email || !pass) return { success: false, error: 'Campos vacíos' };
  if (email.toLowerCase() === 'freeride@cv.com' && pass === 'brach.1311') return { success: true, role: 'MASTER', nombre: 'MASTER' };
  
  try {
    const sheet = getSheet('Usuarios');
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][1].toString().toLowerCase() === email.toLowerCase() && data[i][2].toString() === pass) {
        return { success: true, role: 'FILMMAKER', nombre: data[i][3] };
      }
    }
    return { success: false, error: 'Credenciales incorrectas' };
  } catch (e) { return { success: false, error: e.message }; }
}

function obtenerUsuarios() {
  try {
    const sheet = getSheet('Usuarios');
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return [];
    const data = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
    return data.filter(row => row[0] !== "").map(row => { return { id: row[0], email: row[1], pass: row[2], nombre: row[3] }; });
  } catch(e) { throw new Error(e.message); }
}

function guardarUsuario(id, email, pass, nombre) {
  try {
    const sheet = getSheet('Usuarios');
    if (id) {
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] == id) {
          sheet.getRange(i + 1, 2).setValue(email.toLowerCase());
          sheet.getRange(i + 1, 3).setValue(pass);
          sheet.getRange(i + 1, 4).setValue(nombre.toUpperCase());
          break;
        }
      }
    } else {
      let newId = 1;
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
        const lastId = Math.max(...ids.filter(n => !isNaN(n)));
        if (lastId > 0) newId = lastId + 1;
      }
      sheet.appendRow([newId, email.toLowerCase(), pass, nombre.toUpperCase()]);
    }
    return obtenerUsuarios();
  } catch(e) { throw new Error(e.message); }
}

function borrarUsuario(id) {
  try {
    const sheet = getSheet('Usuarios');
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) { if (data[i][0] == id) { sheet.deleteRow(i + 1); break; } }
    return obtenerUsuarios();
  } catch(e) { throw new Error(e.message); }
}

function obtenerCategorias() {
  try {
    const sheet = getSheet('Categorias');
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return ['GENERAL', 'MAKAHA', 'PUNTA ROQUITAS', 'WAIKIKI', 'PAMPILLA', 'HERRADURA', 'SAN BARTOLO', 'PUNTA HERMOSA'];
    return sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat().filter(c => c !== "");
  } catch (e) { throw new Error(e.message); }
}

function guardarCategoria(nombre) {
  const sheet = getSheet('Categorias');
  const currentCats = obtenerCategorias();
  if (!currentCats.includes(nombre.toUpperCase())) sheet.appendRow([nombre.toUpperCase()]);
  return obtenerCategorias();
}

function eliminarCategoria(nombre) {
  const sheet = getSheet('Categorias');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) { if (data[i][0] === nombre) { sheet.deleteRow(i + 1); break; } }
  return obtenerCategorias();
}

function obtenerProductos() {
  try {
    const sheet = getSheet('Productos');
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return []; 
    const colsToRead = sheet.getLastColumn() < 8 ? sheet.getLastColumn() : 8;
    const data = sheet.getRange(2, 1, lastRow - 1, colsToRead).getValues();
    
    return data.filter(row => row[0] !== "").map((row) => {
        return {
          id: row[0], nombre: row[1] || "Sin Nombre", filmmaker: row[2] || "",
          imagen: row[3], descripcion: row[4] || "", orden: row[5],
          categoria: row[6] ? row[6] : "GENERAL", oculto: row[7] ? row[7] : false 
        };
      }).sort((a, b) => {
        const ordenA = (a.orden === "" || a.orden === null) ? 999999 : Number(a.orden);
        const ordenB = (b.orden === "" || b.orden === null) ? 999999 : Number(b.orden);
        if (ordenA !== ordenB) return ordenA - ordenB;
        return b.id - a.id;
      });
  } catch (e) { throw new Error(e.message); }
}

function guardarOrdenMasivo(listaIds) {
  try {
    const sheet = getSheet('Productos');
    const data = sheet.getDataRange().getValues();
    const mapIdFila = {};
    for (let i = 1; i < data.length; i++) { mapIdFila[data[i][0]] = i + 1; }
    listaIds.forEach((id, index) => {
      const fila = mapIdFila[id];
      if (fila) sheet.getRange(fila, 6).setValue(index + 1); 
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
      const blob = Utilities.newBlob(Utilities.base64Decode(partes[1]), tipoMime, "foto_" + new Date().getTime() + ".jpg");
      const folder = DriveApp.getFolderById(FOLDER_ID);
      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      fileIdOrUrl = file.getId();
    }

    const catFinal = info.categoria || "GENERAL";

    if (info.id) {
      const data = sheet.getDataRange().getValues();
      let rowToEdit = -1;
      for (let i = 1; i < data.length; i++) { if (data[i][0] == info.id) { rowToEdit = i + 1; break; } }
      if (rowToEdit > 0) {
        sheet.getRange(rowToEdit, 2).setValue(info.nombre); sheet.getRange(rowToEdit, 3).setValue(info.filmmaker);
        sheet.getRange(rowToEdit, 4).setValue(fileIdOrUrl); sheet.getRange(rowToEdit, 5).setValue(info.descripcion);
        sheet.getRange(rowToEdit, 7).setValue(catFinal); sheet.getRange(rowToEdit, 8).setValue(info.oculto);
      }
    } else {
      let newId = 1;
      const lastRow = sheet.getLastRow();
      let maxOrden = 0;
      if (lastRow > 1) {
        const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
        const ordenes = sheet.getRange(2, 6, lastRow - 1, 1).getValues().flat();
        const lastId = Math.max(...ids.filter(n => !isNaN(n)));
        if (lastId > 0) newId = lastId + 1;
        const validOrdenes = ordenes.filter(n => !isNaN(parseFloat(n)) && isFinite(n));
        if (validOrdenes.length > 0) maxOrden = Math.max(...validOrdenes);
      }
      sheet.appendRow([newId, info.nombre, info.filmmaker, fileIdOrUrl, info.descripcion, maxOrden + 1, catFinal, info.oculto]);
    }
    return { success: true };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function borrarProducto(id) {
  try {
    const sheet = getSheet('Productos');
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) { if (data[i][0] == id) { sheet.deleteRow(i + 1); return { success: true }; } }
    return { success: false, error: "No encontrado" };
  } catch(e) { return { success: false, error: e.toString() }; }
}

function obtenerEquipo() {
  try {
    const sheet = getSheet('Equipo');
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return []; 
    const data = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
    return data.filter(row => row[0] !== "").map((row) => {
        return { id: row[0], nombre: row[1], rol: row[2], biografia: row[3], imagen: row[4], instagram: row[5] };
    });
  } catch (e) { throw new Error(e.message); }
}

function procesarMiembro(info) {
  try {
    const sheet = getSheet('Equipo');
    const FOLDER_ID = "16nxEBk6_LpDPztROIgR689Mj4YUFRn5P"; 
    let fileIdOrUrl = info.imagenActual;
    if (info.imagenBase64) {
      const partes = info.imagenBase64.split(","); 
      const tipoMime = partes[0].split(":")[1].split(";")[0];
      const blob = Utilities.newBlob(Utilities.base64Decode(partes[1]), tipoMime, "miembro_" + new Date().getTime() + ".jpg");
      const file = DriveApp.getFolderById(FOLDER_ID).createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      fileIdOrUrl = file.getId();
    }
    if (info.id) {
      const data = sheet.getDataRange().getValues();
      let rowToEdit = -1;
      for (let i = 1; i < data.length; i++) { if (data[i][0] == info.id) { rowToEdit = i + 1; break; } }
      if (rowToEdit > 0) {
        sheet.getRange(rowToEdit, 2).setValue(info.nombre); sheet.getRange(rowToEdit, 3).setValue(info.rol);
        sheet.getRange(rowToEdit, 4).setValue(info.biografia); sheet.getRange(rowToEdit, 5).setValue(fileIdOrUrl);
        sheet.getRange(rowToEdit, 6).setValue(info.instagram);
      }
    } else {
      let newId = 1;
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
        const lastId = Math.max(...ids.filter(n => !isNaN(n)));
        if (lastId > 0) newId = lastId + 1;
      }
      sheet.appendRow([newId, info.nombre, info.rol, info.biografia, fileIdOrUrl, info.instagram]);
    }
    return { success: true };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function borrarMiembro(id) {
  try {
    const sheet = getSheet('Equipo');
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) { if (data[i][0] == id) { sheet.deleteRow(i + 1); return { success: true }; } }
    return { success: false, error: "No encontrado" };
  } catch(e) { return { success: false, error: e.toString() }; }
}

function obtenerMerch() {
  try {
    const sheet = getSheet('Merch');
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return []; 
    const colsActuales = sheet.getLastColumn();
    const data = sheet.getRange(2, 1, lastRow - 1, colsActuales).getValues();
    return data.filter(row => row[0] !== "").map((row) => {
        return { 
          id: row[0], nombre: row[1], precio: row[2], imagen: row[3], descripcion: row[4], oculto: row[5], agotado: row[6] ? true : false 
        };
    }).sort((a,b) => b.id - a.id);
  } catch (e) { throw new Error(e.message); }
}

function procesarMerch(info) {
  try {
    const sheet = getSheet('Merch');
    const FOLDER_ID = ""; 
    let fileIdOrUrl = info.imagenActual;
    if (info.imagenBase64) {
      const partes = info.imagenBase64.split(","); 
      const tipoMime = partes[0].split(":")[1].split(";")[0];
      const blob = Utilities.newBlob(Utilities.base64Decode(partes[1]), tipoMime, "merch_" + new Date().getTime() + ".jpg");
      const file = DriveApp.getFolderById(FOLDER_ID).createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      fileIdOrUrl = file.getId();
    }
    if (info.id) {
      const data = sheet.getDataRange().getValues();
      let rowToEdit = -1;
      for (let i = 1; i < data.length; i++) { if (data[i][0] == info.id) { rowToEdit = i + 1; break; } }
      if (rowToEdit > 0) {
        sheet.getRange(rowToEdit, 2).setValue(info.nombre); sheet.getRange(rowToEdit, 3).setValue(info.precio);
        sheet.getRange(rowToEdit, 4).setValue(fileIdOrUrl); sheet.getRange(rowToEdit, 5).setValue(info.descripcion);
        sheet.getRange(rowToEdit, 6).setValue(info.oculto); sheet.getRange(rowToEdit, 7).setValue(info.agotado); 
      }
    } else {
      let newId = 1;
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
        const lastId = Math.max(...ids.filter(n => !isNaN(n)));
        if (lastId > 0) newId = lastId + 1;
      }
      sheet.appendRow([newId, info.nombre, info.precio, fileIdOrUrl, info.descripcion, info.oculto, info.agotado]);
    }
    return { success: true };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function borrarMerch(id) {
  try {
    const sheet = getSheet('Merch');
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) { if (data[i][0] == id) { sheet.deleteRow(i + 1); return { success: true }; } }
    return { success: false, error: "No encontrado" };
  } catch(e) { return { success: false, error: e.toString() }; }
}

function obtenerSurfistas() {
  try {
    const sheet = getSheet('Surfistas');
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return []; 
    const colsActuales = sheet.getLastColumn();
    const data = sheet.getRange(2, 1, lastRow - 1, colsActuales).getValues();
    return data.filter(row => row[0] !== "").map((row) => {
        return { 
          id: row[0], nombre: row[1], numero: row[2], imagen: row[3], posicion: row[4], redes: row[5], nivel: row[6], playa: row[7] 
        };
    }).sort((a,b) => b.id - a.id);
  } catch (e) { throw new Error(e.message); }
}

function procesarSurfista(info) {
  try {
    const sheet = getSheet('Surfistas');
    const FOLDER_ID = "1Fye-rz_yGrFw26x2sEFH2pdKCDNGQYXh";
    let fileIdOrUrl = info.imagenActual;
    
    if (info.imagenBase64) {
      const partes = info.imagenBase64.split(","); 
      const tipoMime = partes[0].split(":")[1].split(";")[0];
      const blob = Utilities.newBlob(Utilities.base64Decode(partes[1]), tipoMime, "surfista_" + new Date().getTime() + ".jpg");
      const file = DriveApp.getFolderById(FOLDER_ID).createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      fileIdOrUrl = file.getId();
    }

    if (info.id) {
      const data = sheet.getDataRange().getValues();
      let rowToEdit = -1;
      for (let i = 1; i < data.length; i++) { if (data[i][0] == info.id) { rowToEdit = i + 1; break; } }
      if (rowToEdit > 0) {
        sheet.getRange(rowToEdit, 2).setValue(info.nombre); sheet.getRange(rowToEdit, 3).setValue(info.numero);
        sheet.getRange(rowToEdit, 4).setValue(fileIdOrUrl); sheet.getRange(rowToEdit, 5).setValue(info.posicion);
        sheet.getRange(rowToEdit, 6).setValue(info.redes); sheet.getRange(rowToEdit, 7).setValue(info.nivel); 
        sheet.getRange(rowToEdit, 8).setValue(info.playa); 
      }
    } else {
      let newId = 1;
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
        const lastId = Math.max(...ids.filter(n => !isNaN(n)));
        if (lastId > 0) newId = lastId + 1;
      }
      sheet.appendRow([newId, info.nombre, info.numero, fileIdOrUrl, info.posicion, info.redes, info.nivel, info.playa]);
    }
    return { success: true };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function borrarSurfista(id) {
  try {
    const sheet = getSheet('Surfistas');
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) { if (data[i][0] == id) { sheet.deleteRow(i + 1); return { success: true }; } }
    return { success: false, error: "No encontrado" };
  } catch(e) { return { success: false, error: e.toString() }; }
}
