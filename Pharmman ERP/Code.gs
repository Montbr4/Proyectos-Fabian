const ID_HOJA = "";

const DB_USUARIOS = {
  "davis":     { pass: "Pharman091510.", roles: ["M", "S", "SMP", "OC", "IR", "PC"] },
  "grace":     { pass: "Pharman091510.", roles: ["M", "S", "SMP", "OC", "IR", "PC"] },
  "marissa":   { pass: "nico",           roles: ["M", "S", "SMP", "OC", "IR", "PC"] },
  "fabian":    { pass: "masha",          roles: ["M", "S", "SMP", "OC", "IR", "PC"] },
  "jhair":     { pass: "Pharman2025",    roles: ["S","OC", "IR", "PC"] },
  "kimberly":  { pass: "Pharman2025",    roles: ["S","OC", "IR"] },
  "rafael":    { pass: "BartoloMew123#", roles: ["M","OC", "PC"] },
  "alvaro":    { pass: "Pharman2025",    roles: ["SMP", "M"] },
  "sebastian": { pass: "Pharman2025",    roles: ["S"] },
  "paola":     { pass: "Pharman2025",    roles: ["S"] }
};

function normalizarTexto(texto) {
  if (!texto) return "";
  return texto.toString().trim().toUpperCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
}

function doGet(e) {
  let page = e.parameter.page;
  let sede = e.parameter.sede;
  let usuario = e.parameter.usuario;

  if (!page) {
    return HtmlService.createTemplateFromFile("Login").evaluate().setTitle("Pharman SR - Acceso")
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
  }

  if (!usuario || !DB_USUARIOS[usuario]) {
    return HtmlService.createTemplateFromFile("Login").evaluate()
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  let rolesUsuario = DB_USUARIOS[usuario].roles;
  let template;
  
  if (page === "inicio") {
    template = HtmlService.createTemplateFromFile("Inicio");
  } else {
    if (sede && !rolesUsuario.includes(sede)) return HtmlService.createHtmlOutput("<h2 style='text-align:center;margin-top:50px;'>Acceso Denegado a esta Sede</h2>");
    if (page === "oc" && !rolesUsuario.includes("OC")) return HtmlService.createHtmlOutput("<h2 style='text-align:center;margin-top:50px;'>Acceso Denegado a Compras</h2>");
    if (page === "ir" && !rolesUsuario.includes("IR")) return HtmlService.createHtmlOutput("<h2 style='text-align:center;margin-top:50px;'>Acceso Denegado a Recepción</h2>");
    if (page === "bd_pc" && !rolesUsuario.includes("PC")) return HtmlService.createHtmlOutput("<h2 style='text-align:center;margin-top:50px;'>Acceso Denegado a Gestión de Costos</h2>");

    switch(page) {
      case "ventas":      template = HtmlService.createTemplateFromFile("Ventas"); break;
      case "inventario":  template = HtmlService.createTemplateFromFile("Inventario"); break;
      case "oc":          template = HtmlService.createTemplateFromFile("OC"); break;
      case "ir":          template = HtmlService.createTemplateFromFile("IR"); break;
      case "anulacion":   template = HtmlService.createTemplateFromFile("Anulacion"); break;
      case "bd_pc":       template = HtmlService.createTemplateFromFile("BD_Precio_Costo"); break;
      default:            template = HtmlService.createTemplateFromFile("Inicio");
    }
  }

  template.sede = sede || ""; 
  template.usuario = usuario || "";
  template.permisos = JSON.stringify(rolesUsuario); 

  return template.evaluate().setTitle("Pharman SR - " + page.toUpperCase())
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
}

function getUrl() { return ScriptApp.getService().getUrl(); }

function validarLogin(email, password) {
  let user = email.trim().split("@")[0].toLowerCase();
  if (email.trim().endsWith("@pharman.com") && DB_USUARIOS[user]) {
    if (DB_USUARIOS[user].pass === password) {
      try { registrarLogUsuario(user, "ENTRADA"); } catch(e) {}
      return { exito: true, usuario: user };
    }
  }
  return { exito: false };
}

function registrarLogUsuario(usuario, accion) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000); 
    const ss = SpreadsheetApp.openById(ID_HOJA);
    let sheet = ss.getSheetByName("USUARIOS_LOG");
    if(!sheet) { 
      sheet = ss.insertSheet("USUARIOS_LOG");
      sheet.appendRow(["Usuario", "Acción", "Fecha", "Hora"]);
    }
    const fecha = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");
    const hora = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "HH:mm:ss");
    sheet.appendRow([usuario, accion, fecha, hora]);
    SpreadsheetApp.flush();
  } catch (e) { console.error(e); } finally { lock.releaseLock(); }
}

function obtenerProductosInventario(sede) {
  const ss = SpreadsheetApp.openById(ID_HOJA);
  const sheet = ss.getSheetByName("INVENTARIO_" + sede);
  if (!sheet || sheet.getLastRow() < 2) return []; 
  const datos = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
  
  return datos.filter(r => r[0] && r[0].toString().trim() !== "").map(r => ({
    nombre: r[0].toString().trim(), 
    stock: parseInt(r[1]) || 0, 
    precioVenta: parseFloat(r[2]) || 0, 
    precioCosto: parseFloat(r[3]) || 0 
  }));
}

function obtenerBaseProductosOC() {
  const ss = SpreadsheetApp.openById(ID_HOJA);
  const sheet = ss.getSheetByName("BD_OC");
  if (!sheet || sheet.getLastRow() < 2) return [];
  const datos = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
  return datos.filter(r => r[1] && r[1].toString().trim() !== "").map(fila => ({ 
    codigo: fila[0], 
    producto: fila[1].toString().trim(), 
    proveedor: fila[2], 
    costo: parseFloat(fila[3]) || 0 
  }));
}

function obtenerListaPreciosCosto() {
  const ss = SpreadsheetApp.openById(ID_HOJA);
  let sBD = ss.getSheetByName("BD_PC");
  if(!sBD) { inicializarTablas(); sBD = ss.getSheetByName("BD_PC"); }
  if (sBD.getLastRow() < 2) return [];
  const datos = sBD.getRange(2, 1, sBD.getLastRow() - 1, 2).getValues();
  return datos.filter(r => r[0] && r[0].toString().trim() !== "").map(r => ({ 
    producto: r[0].toString().trim(), 
    precioCosto: r[1] 
  }));
}

function actualizarCostoGlobal(producto, nuevoCosto) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(30000); } catch(e) { return "Error: Sistema ocupado."; }

  try {
    producto = producto ? producto.toString().trim() : "";
    if (producto === "") return "Error: Producto inválido.";

    let prodBuscado = normalizarTexto(producto);

    const ss = SpreadsheetApp.openById(ID_HOJA);
    const costo = parseFloat(nuevoCosto) || 0;
    let sBD = ss.getSheetByName("BD_PC");
    if(!sBD) { inicializarTablas(); sBD = ss.getSheetByName("BD_PC"); }
    
    let dataBD = sBD.getDataRange().getValues();
    let encontradoBD = false;
    for(let i=1; i<dataBD.length; i++) {
        if(normalizarTexto(dataBD[i][0]) === prodBuscado) { 
          sBD.getRange(i+1, 2).setValue(costo); 
          encontradoBD = true; 
          break; 
        }
    }
    if(!encontradoBD) sBD.appendRow([producto, costo]);
    
    ["M", "S", "SMP"].forEach(sede => {
        let sInv = ss.getSheetByName("INVENTARIO_" + sede);
        if(sInv) {
            let dataInv = sInv.getDataRange().getValues();
            for(let i=1; i<dataInv.length; i++) {
                if(normalizarTexto(dataInv[i][0]) === prodBuscado) { 
                  sInv.getRange(i+1, 4).setValue(costo); 
                  break; 
                }
            }
        }
    });
    return "Precio Costo actualizado y sincronizado en todas las sedes.";
  } catch(e) { return "Error al actualizar: " + e.toString(); } finally { lock.releaseLock(); }
}

function registrarVentaCompleta(data, sedeRecibida, usuario) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(30000); } catch (e) { return "Timeout del servidor."; }

  try {
      const ss = SpreadsheetApp.openById(ID_HOJA);
      const sede = sedeRecibida.toString().trim();
      let sVenta = ss.getSheetByName("VENTA_" + sede);
      let sInventario = ss.getSheetByName("INVENTARIO_" + sede);
      let sClientes = ss.getSheetByName("CLIENTES_" + sede);
      let sDetallePacks = ss.getSheetByName("DETALLE_PACKS_" + sede);

      if(!sVenta || !sInventario || !sClientes) {
          inicializarTablas(); SpreadsheetApp.flush();
          sVenta = ss.getSheetByName("VENTA_" + sede);
          sInventario = ss.getSheetByName("INVENTARIO_" + sede);
          sClientes = ss.getSheetByName("CLIENTES_" + sede);
          sDetallePacks = ss.getSheetByName("DETALLE_PACKS_" + sede);
      }
      
      const fecha = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");
      let totalCalculado = 0;
      
      const dataInv = sInventario.getDataRange().getValues();
      let mapaInv = {};
      for(let i=1; i<dataInv.length; i++) {
        if(dataInv[i][0]) mapaInv[normalizarTexto(dataInv[i][0])] = i; 
      }

      let deducciones = {};

      data.carrito.forEach(item => {
        let nombreProducto = item.nombre ? item.nombre.toString().trim() : "";
        if (nombreProducto === "") return; 

        let totalLinea = parseFloat(item.cantidad) * parseFloat(item.precioVenta);
        totalCalculado += totalLinea;
        
        sVenta.appendRow([fecha, data.idVenta, data.tienda, item.cantidad, nombreProducto, item.precioVenta, totalLinea, usuario, ""]);

        if (item.esPack && item.componentes) {
            item.componentes.forEach(comp => {
                let compNombre = comp.nombre ? comp.nombre.toString().trim() : "";
                if(sDetallePacks && compNombre !== "") {
                    sDetallePacks.appendRow([data.idVenta, nombreProducto, compNombre, comp.cantidad]);
                }
                let cantDesc = comp.cantidad * item.cantidad;
                let normComp = normalizarTexto(compNombre);
                if (normComp) deducciones[normComp] = (deducciones[normComp] || 0) + cantDesc;
            });
        } else {
            let normProd = normalizarTexto(nombreProducto);
            if (normProd) deducciones[normProd] = (deducciones[normProd] || 0) + parseInt(item.cantidad);
        }
      });
      
      for (let prodNorm in deducciones) {
          let idx = mapaInv[prodNorm];
          if (idx !== undefined) {
              let celdaStock = sInventario.getRange(idx + 1, 2);
              let stockReal = parseInt(celdaStock.getValue()) || 0;
              let cantADescontar = deducciones[prodNorm];
              celdaStock.setValue(stockReal - cantADescontar);
          }
      }
      
      let prodResumen = data.carrito
          .filter(p => p.nombre && p.nombre.toString().trim() !== "")
          .map(p => `${p.nombre.toString().trim()} (${p.cantidad})`)
          .join(", ");

      sClientes.appendRow([data.cliente.nombre, data.cliente.dni, data.cliente.celular, data.cliente.correo, data.cliente.distrito, prodResumen, data.cliente.categoria, totalCalculado, fecha, data.tienda, usuario]);
      
      SpreadsheetApp.flush();
      return "Venta registrada exitosamente. Total: S/. " + totalCalculado.toFixed(2);
  } catch (e) { return "Error: " + e.toString(); } finally { lock.releaseLock(); }
}

function procesarAnulacionVenta(idVenta, motivo, sedeRecibida, usuario) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(30000); } catch(e) { return { exito:false, msg:"Sistema ocupado."}; }
  try {
    const ss = SpreadsheetApp.openById(ID_HOJA);
    const sede = sedeRecibida.toString().trim();
    const sVenta = ss.getSheetByName("VENTA_" + sede);
    const sInventario = ss.getSheetByName("INVENTARIO_" + sede);
    const sAnulaciones = ss.getSheetByName("ANULACIONES_" + sede);
    const sDetallePacks = ss.getSheetByName("DETALLE_PACKS_" + sede);
    if (!sAnulaciones) return { exito: false, msg: "Error: Hoja anulaciones falta." };
    
    const dataVenta = sVenta.getDataRange().getValues();
    const dataInv = sInventario.getDataRange().getValues();
    
    let productosRestaurados = [];
    let tiendaOrigen = "";
    let encontrado = false;
    let filasVentaAEliminar = [];
    
    let mapaInv = {};
    for(let i=1; i<dataInv.length; i++) {
        if (dataInv[i][0]) mapaInv[normalizarTexto(dataInv[i][0])] = i;
    }

    let sumasRestauracion = {};
    let descripcionesRestauradas = [];

    for (let i = 1; i < dataVenta.length; i++) {
      if (dataVenta[i][1].toString() === idVenta.toString()) {
        encontrado = true;
        if (!tiendaOrigen) tiendaOrigen = dataVenta[i][2]; 
        
        let producto = dataVenta[i][4] ? dataVenta[i][4].toString().trim() : "";
        if (producto === "") continue; 
        
        let cantidadVenta = parseInt(dataVenta[i][3]);
        let esPack = producto.startsWith("PACK:");
        
        if(esPack) {
           if (sDetallePacks) {
               const dataPacks = sDetallePacks.getDataRange().getValues();
               let filasPackEliminar = [];
               for (let p = 1; p < dataPacks.length; p++) {
                   if (dataPacks[p][0].toString() === idVenta.toString() && dataPacks[p][1].toString().trim() === producto) {
                       let compNombre = dataPacks[p][2] ? dataPacks[p][2].toString().trim() : "";
                       if (compNombre === "") continue;
                       let compCantUnit = parseInt(dataPacks[p][3]);
                       let totalRestaurar = compCantUnit * cantidadVenta;
                       
                       let normComp = normalizarTexto(compNombre);
                       sumasRestauracion[normComp] = (sumasRestauracion[normComp] || 0) + totalRestaurar;
                       filasPackEliminar.push(p + 1);
                   }
               }
               descripcionesRestauradas.push(`${producto} (Pack Anulado)`);
               filasPackEliminar.sort((a,b) => b-a).forEach(r => sDetallePacks.deleteRow(r));
           }
        } else {
           let normProd = normalizarTexto(producto);
           sumasRestauracion[normProd] = (sumasRestauracion[normProd] || 0) + cantidadVenta;
           descripcionesRestauradas.push(`${producto} (+${cantidadVenta})`);
        }
        filasVentaAEliminar.push(i + 1);
      }
    }
    
    if (!encontrado) return { exito: false, msg: "ID no encontrado." };

    for (let prodNorm in sumasRestauracion) {
        let idx = mapaInv[prodNorm];
        if (idx !== undefined) {
            let celdaStock = sInventario.getRange(idx + 1, 2);
            let stockReal = parseInt(celdaStock.getValue()) || 0;
            let aRestaurar = sumasRestauracion[prodNorm];
            celdaStock.setValue(stockReal + aRestaurar);
            productosRestaurados.push(`${dataInv[idx][0]} (+${aRestaurar})`);
        }
    }

    const fecha = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
    sAnulaciones.appendRow([idVenta, fecha, motivo, productosRestaurados.join(" | "), tiendaOrigen, usuario]);
    filasVentaAEliminar.sort((a,b) => b-a).forEach(r => sVenta.deleteRow(r));
    SpreadsheetApp.flush();
    return { exito: true, msg: "Venta anulada." };
  } catch(e) { return { exito: false, msg: "Error: " + e.message }; } finally { lock.releaseLock(); }
}

function registrarListaMovimientos(data, sedeRecibida, usuario) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch(e) { return "Servidor ocupado."; }

  try {
    const ss = SpreadsheetApp.openById(ID_HOJA);
    const sede = sedeRecibida.toString().trim();
    const sInv = ss.getSheetByName("INVENTARIO_" + sede);
    const sMov = ss.getSheetByName("M_INVENTARIO_" + sede);
    if(!sInv) throw new Error("No se encuentra inventario de " + sede);

    const fecha = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");
    const dataInv = sInv.getDataRange().getValues();
    let mapaInv = {};
    for(let i=1; i<dataInv.length; i++) {
       if (dataInv[i][0]) mapaInv[normalizarTexto(dataInv[i][0])] = i; 
    }
    
    let operacionesSeguras = [];
    for (let i = 0; i < data.lista.length; i++) {
      let item = data.lista[i];
      let nombreProd = item.producto ? item.producto.toString().trim() : "";
      if (nombreProd === "") continue; 

      let idx = mapaInv[normalizarTexto(nombreProd)];
      if (idx === undefined) return `Error: El producto ${nombreProd} no existe en inventario.`;
      
      let celdaStock = sInv.getRange(idx + 1, 2);
      let stockActual = parseInt(celdaStock.getValue()) || 0;
      let cantidadOp = parseInt(item.cantidad);

      if (data.tipo !== "INGRESO" && (stockActual - cantidadOp < 0)) {
          return `Error: Stock insuficiente para ${nombreProd}. Stock actual: ${stockActual}, Intentó descontar: ${cantidadOp}. No se procesó ningún movimiento.`;
      }

      operacionesSeguras.push({
          nombreProd: nombreProd,
          celdaStock: celdaStock,
          stockActual: stockActual,
          cantidadOp: cantidadOp,
          precioCosto: item.precioCosto,
          proveedor: item.proveedor,
          destino: item.destino,
          idx: idx
      });
    }

    let actualizacionesCosto = [];
    operacionesSeguras.forEach(op => {
      sMov.appendRow([data.tipo, op.nombreProd, op.cantidadOp, fecha, op.proveedor || "-", usuario, op.destino || "-", ""]);
      
      let nuevoStock = (data.tipo === "INGRESO") ? op.stockActual + op.cantidadOp : op.stockActual - op.cantidadOp;
      op.celdaStock.setValue(nuevoStock);

      if (data.tipo === "INGRESO" && op.precioCosto && parseFloat(op.precioCosto) > 0) {
           let costo = parseFloat(op.precioCosto);
           sInv.getRange(op.idx + 1, 4).setValue(costo); 
           actualizacionesCosto.push({producto: op.nombreProd, costo: costo}); 
      }
    });

    SpreadsheetApp.flush();
    if(actualizacionesCosto.length > 0) {
        actualizacionesCosto.forEach(upd => { actualizarCostoGlobal(upd.producto, upd.costo); });
    }
    return "Movimientos registrados correctamente.";
  } catch(e) { return "Error: " + e.toString(); } finally { lock.releaseLock(); }
}

function procesarListaAnulacionTraslado(data, sedeRecibida, usuario) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch(e) { return "Ocupado."; }
  try {
    const ss = SpreadsheetApp.openById(ID_HOJA);
    const sede = sedeRecibida.toString().trim();
    const sInv = ss.getSheetByName("INVENTARIO_" + sede);
    const sMov = ss.getSheetByName("M_INVENTARIO_" + sede);
    const fecha = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");
    const dataInv = sInv.getDataRange().getValues();
    let mapaInv = {};
    for(let i=1; i<dataInv.length; i++) {
        if (dataInv[i][0]) mapaInv[normalizarTexto(dataInv[i][0])] = i;
    }
    
    let tipoAccion = data.tipoAnulacion || "RESTAURAR";
    let motivo = data.motivo || "Sin detalle";
    let etiquetaMov = (tipoAccion === "RESTAURAR") ? "ANULACION DE SALIDAS (+)" : "ANULACION DE INGRESOS (-)";
    
    data.lista.forEach(item => {
      let nombreProd = item.producto ? item.producto.toString().trim() : "";
      if (nombreProd === "") return;

      let idx = mapaInv[normalizarTexto(nombreProd)];
      if (idx !== undefined) {
        let celdaStock = sInv.getRange(idx + 1, 2);
        let stockActual = parseInt(celdaStock.getValue()) || 0;
        let cantOperacion = parseInt(item.cantidad);
        let nuevoStock = (tipoAccion === "RESTAURAR") ? stockActual + cantOperacion : stockActual - cantOperacion;
        
        celdaStock.setValue(nuevoStock);
      }
      sMov.appendRow([etiquetaMov, nombreProd, parseInt(item.cantidad), fecha, "-", usuario, "CORRECCION", motivo]);
    });
    return "Corrección procesada.";
  } catch(e) { return "Error: " + e.message; } finally { lock.releaseLock(); }
}

function modificarProductoInventario(nombreAntiguo, nombreNuevo, precioNuevo, sedeRecibida) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch(e) { return "Ocupado."; }
  try {
    nombreAntiguo = nombreAntiguo ? nombreAntiguo.toString().trim() : "";
    nombreNuevo = nombreNuevo ? nombreNuevo.toString().trim() : "";
    if (nombreAntiguo === "" || nombreNuevo === "") return "Error: El nombre no puede estar en blanco.";
    let nomAntiguoNorm = normalizarTexto(nombreAntiguo);

    const ss = SpreadsheetApp.openById(ID_HOJA);
    const sede = sedeRecibida.toString().trim();
    const sInv = ss.getSheetByName("INVENTARIO_" + sede);
    if(!sInv) return "Error inventario.";
    const data = sInv.getDataRange().getValues();
    for(let i = 1; i < data.length; i++) {
      if(normalizarTexto(data[i][0]) === nomAntiguoNorm) {
          sInv.getRange(i + 1, 1).setValue(nombreNuevo);
          sInv.getRange(i + 1, 3).setValue(parseFloat(precioNuevo));
          return "Actualizado.";
      }
    }
    return "No encontrado.";
  } catch(e) { return "Error: " + e.message; } finally { lock.releaseLock(); }
}

function crearNuevoProducto(nombre, stockInicial, precioVenta, precioCosto, sedeRecibida, usuario) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch(e) { return "Ocupado."; }

  try {
    nombre = nombre ? nombre.toString().trim() : "";
    if (nombre === "") return "Error: El nombre del producto no puede estar en blanco.";
    let nomNormalizado = normalizarTexto(nombre);

    const ss = SpreadsheetApp.openById(ID_HOJA);
    const sede = sedeRecibida.toString().trim();
    const sInv = ss.getSheetByName("INVENTARIO_" + sede);
    const sMov = ss.getSheetByName("M_INVENTARIO_" + sede);
    let sBD = ss.getSheetByName("BD_PC");
    if(!sBD) { inicializarTablas(); sBD = ss.getSheetByName("BD_PC"); }

    const fecha = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");

    if (sInv.getLastRow() > 0) {
      const data = sInv.getDataRange().getValues();
      const existe = data.some(r => r[0] && normalizarTexto(r[0]) === nomNormalizado);
      if (existe) return "Error: El producto ya existe en esta sede.";
    }
    
    let pCosto = parseFloat(precioCosto) || 0;
    sInv.appendRow([nombre, parseInt(stockInicial), parseFloat(precioVenta), pCosto, 0]);
    sMov.appendRow(["INGRESO (INICIAL)", nombre, stockInicial, fecha, "-", usuario, "-", "Creación Producto"]);
    
    const dataBD = sBD.getDataRange().getValues();
    let existeEnGlobal = false;
    for(let i = 1; i < dataBD.length; i++) {
        if(dataBD[i][0] && normalizarTexto(dataBD[i][0]) === nomNormalizado) { existeEnGlobal = true; break; }
    }
    if (!existeEnGlobal) sBD.appendRow([nombre, pCosto]);
    
    SpreadsheetApp.flush();
    if (pCosto > 0) actualizarCostoGlobal(nombre, pCosto);
    
    return "Producto creado correctamente.";
  } catch(e) { return "Error: " + e.message; } finally { lock.releaseLock(); }
}

function guardarOC(data, usuario) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch(e) { return "Sistema ocupado."; }
  try {
    const ss = SpreadsheetApp.openById(ID_HOJA);
    let sOC = ss.getSheetByName("OC");
    if(!sOC) { inicializarTablas(); sOC = ss.getSheetByName("OC"); }
    const fecha = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");
    let resumenItems = data.items.map(i => `${i.codigo} - ${i.producto} (${i.cantidad})`).join(" | ");
    sOC.appendRow([data.idOrden, data.proveedor, fecha, resumenItems, data.totalEstimado, usuario, "PENDIENTE"]);
    return "Orden de Compra " + data.idOrden + " guardada.";
  } catch(e) { return "Error OC: " + e.message; } finally { lock.releaseLock(); }
}

function guardarIR(data, usuario) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch(e) { return "Sistema ocupado."; }
  try {
    const ss = SpreadsheetApp.openById(ID_HOJA);
    let sIR = ss.getSheetByName("IR");
    if(!sIR) { inicializarTablas(); sIR = ss.getSheetByName("IR"); }
    const fecha = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");
    sIR.appendRow([data.idRecepcion, data.refOC, data.proveedor, data.notas, data.listaProductos, data.montoPagar, "PENDIENTE PAGO", fecha, usuario]);
    return "Recepción registrada.";
  } catch(e) { return "Error IR: " + e.message; } finally { lock.releaseLock(); }
}

function inicializarTablas() {
  const ss = SpreadsheetApp.openById(ID_HOJA);
  const sedes = ["M", "S", "SMP"];
  
  const hojasGlobales = [
    { nombre: "USUARIOS_LOG", headers: ["Usuario", "Acción", "Fecha", "Hora"] },
    { nombre: "OC", headers: ["ID Orden", "Proveedor", "Fecha Solicitud", "Items Resumen", "Total Estimado", "Usuario", "Estado"] },
    { nombre: "IR", headers: ["ID Recepción", "Referencia OC", "Proveedor", "Notas", "Lista Productos", "Monto a Pagar", "Estado Pago", "Fecha Llegada", "Usuario"] },
    { nombre: "BD_OC", headers: ["Código", "Producto", "Proveedor", "Costo"] },
    { nombre: "BD_PC", headers: ["NOMBRE", "PRECIO_C"] }
  ];

  hojasGlobales.forEach(h => {
    let sheet = ss.getSheetByName(h.nombre);
    if (!sheet) {
      sheet = ss.insertSheet(h.nombre);
      sheet.getRange(1, 1, 1, h.headers.length).setValues([h.headers]);
      sheet.getRange(1, 1, 1, h.headers.length).setFontWeight("bold").setBackground("#cccccc");
    }
  });

  const estructurasSede = [
    { nombre: "INVENTARIO", headers: ["Producto", "Stock", "PRECIO_V", "PRECIO_C", "COSTO REAL MERCADERIA"] }, 
    { nombre: "VENTA", headers: ["Fecha", "ID Venta", "Tienda", "Cantidad", "Producto", "PRECIO_U", "Total Línea", "Usuario", "Estado"] },
    { nombre: "CLIENTES", headers: ["Nombre", "DNI/RUC", "Celular", "Correo", "Distrito", "Productos Resumen", "Categoría", "Total Venta", "Fecha", "Tienda", "Usuario"] },
    { nombre: "M_INVENTARIO", headers: ["Tipo Movimiento", "Producto", "Cantidad", "Fecha", "Proveedor", "Usuario", "Destino", "Detalle"] },
    { nombre: "ANULACIONES", headers: ["ID Venta", "Fecha Anulación", "Motivo", "Items Restaurados", "Tienda", "Usuario"] },
    { nombre: "DETALLE_PACKS", headers: ["ID Venta", "Nombre Pack", "Producto Componente", "Cantidad Unit"] }
  ];

  sedes.forEach(sede => {
    estructurasSede.forEach(est => {
      let nombreHoja = est.nombre + "_" + sede;
      let sheet = ss.getSheetByName(nombreHoja);
      if (!sheet) {
        try { sheet = ss.insertSheet(nombreHoja); } catch (e) { sheet = ss.getSheetByName(nombreHoja); }
        if (sheet.getLastRow() === 0) {
          sheet.getRange(1, 1, 1, est.headers.length).setValues([est.headers]);
          sheet.getRange(1, 1, 1, est.headers.length).setFontWeight("bold").setBackground("#cccccc");
        }
      }
    });
  });
}


function probarReporteSemanalManual() {
  const hoy = new Date();
  let diaSemana = hoy.getDay();
  let diasParaDomingo = (diaSemana === 0) ? 7 : diaSemana;
  let fin = new Date(hoy.getFullYear(), hoy.getMonth(), hoy.getDate() - diasParaDomingo);
  let inicio = new Date(fin.getFullYear(), fin.getMonth(), fin.getDate() - 6);
  inicio.setHours(0, 0, 0, 0); fin.setHours(23, 59, 59, 999);

  generarTopProductosSemanal(inicio, fin);
  limpiarReportesAntiguos("ANALISIS_SEMANAL");
}

function probarReporteMensualManual() {
  const hoy = new Date();
  const fechaMesPasado = new Date(hoy.getFullYear(), hoy.getMonth() - 1, 1);
  let mes = fechaMesPasado.getMonth();
  let anio = fechaMesPasado.getFullYear();

  generarTopProductosMensual(mes, anio);
  generarTopClientesMensual(mes, anio, "TODOS", "A_CLIENTES");
  generarTopClientesMensual(mes, anio, "Bebés", "A_CLIENTES_B");

  limpiarReportesAntiguos("ANALISIS_MENSUAL");
  limpiarReportesAntiguos("A_CLIENTES");
  limpiarReportesAntiguos("A_CLIENTES_B");
}

function triggerReporteSemanal() {
  const hoy = new Date();
  let diaSemana = hoy.getDay();
  let diasParaDomingo = (diaSemana === 0) ? 7 : diaSemana;
  let fin = new Date(hoy.getFullYear(), hoy.getMonth(), hoy.getDate() - diasParaDomingo);
  let inicio = new Date(fin.getFullYear(), fin.getMonth(), fin.getDate() - 6);
  inicio.setHours(0, 0, 0, 0); fin.setHours(23, 59, 59, 999);

  generarTopProductosSemanal(inicio, fin);
  limpiarReportesAntiguos("ANALISIS_SEMANAL");
}

function triggerReporteMensual() {
  const hoy = new Date();
  const fechaMesPasado = new Date(hoy.getFullYear(), hoy.getMonth() - 1, 1);
  let mes = fechaMesPasado.getMonth();
  let anio = fechaMesPasado.getFullYear();

  generarTopProductosMensual(mes, anio);
  generarTopClientesMensual(mes, anio, "TODOS", "A_CLIENTES");
  generarTopClientesMensual(mes, anio, "Bebés", "A_CLIENTES_B");

  limpiarReportesAntiguos("ANALISIS_MENSUAL");
  limpiarReportesAntiguos("A_CLIENTES");
  limpiarReportesAntiguos("A_CLIENTES_B");
}

function generarTopProductosSemanal(fechaInicio, fechaFin) {
  const ss = SpreadsheetApp.openById(ID_HOJA);
  let sheetDestino = ss.getSheetByName("ANALISIS_SEMANAL");
  if (!sheetDestino) {
    sheetDestino = ss.insertSheet("ANALISIS_SEMANAL");
    sheetDestino.appendRow(["FECHA", "PRODUCTO", "SEDE", "NRO DE VENTAS", "TOTAL"]);
    sheetDestino.getRange(1, 1, 1, 5).setFontWeight("bold").setBackground("#FFF2CC");
  }

  const sedes = ["M", "S", "SMP"];
  const nombresSedes = { "M": "MAGDALENA", "S": "SURCO", "SMP": "SAN MARTIN DE PORRES" };
  const strInicio = Utilities.formatDate(fechaInicio, Session.getScriptTimeZone(), "dd/MM/yyyy");
  const strFin = Utilities.formatDate(fechaFin, Session.getScriptTimeZone(), "dd/MM/yyyy");
  const etiquetaReporte = `Semana: ${strInicio} - ${strFin}`;
  let filasParaGuardar = [];

  sedes.forEach(sede => {
    const hojaVentas = ss.getSheetByName("VENTA_" + sede);
    if (!hojaVentas) return;
    const dataVentas = hojaVentas.getDataRange().getValues();
    let stats = {};

    for (let i = 1; i < dataVentas.length; i++) {
      let fila = dataVentas[i];
      let fechaObj = parseFecha(fila[0]);

      if (fechaObj) {
        fechaObj.setHours(12, 0, 0, 0); 
        if (fechaObj >= fechaInicio && fechaObj <= fechaFin) {
          procesarFilaVentasParaTop(fila, stats);
        }
      }
    }
    ordenarYGuardarTop(stats, 10, etiquetaReporte, nombresSedes[sede], filasParaGuardar);
  });

  if (filasParaGuardar.length > 0) {
    sheetDestino.getRange(sheetDestino.getLastRow() + 1, 1, filasParaGuardar.length, 5).setValues(filasParaGuardar);
  }
}

function generarTopProductosMensual(mes, anio) {
  const ss = SpreadsheetApp.openById(ID_HOJA);
  let sheetDestino = ss.getSheetByName("ANALISIS_MENSUAL");
  if (!sheetDestino) {
    sheetDestino = ss.insertSheet("ANALISIS_MENSUAL");
    sheetDestino.appendRow(["FECHA", "PRODUCTO", "SEDE", "NRO DE VENTAS", "TOTAL"]);
    sheetDestino.getRange(1, 1, 1, 5).setFontWeight("bold").setBackground("#D9EAD3");
  }

  const sedes = ["M", "S", "SMP"];
  const nombresSedes = { "M": "MAGDALENA", "S": "SURCO", "SMP": "SAN MARTIN DE PORRES" };
  const etiquetaReporte = `${obtenerNombreMes(mes)} ${anio}`;
  let filasParaGuardar = [];

  sedes.forEach(sede => {
    const hojaVentas = ss.getSheetByName("VENTA_" + sede);
    if (!hojaVentas) return;
    const dataVentas = hojaVentas.getDataRange().getValues();
    let stats = {};

    for (let i = 1; i < dataVentas.length; i++) {
      let fila = dataVentas[i];
      let fechaObj = parseFecha(fila[0]);

      if (fechaObj && fechaObj.getMonth() === mes && fechaObj.getFullYear() === anio) {
        procesarFilaVentasParaTop(fila, stats);
      }
    }
    ordenarYGuardarTop(stats, 10, etiquetaReporte, nombresSedes[sede], filasParaGuardar);
  });

  if (filasParaGuardar.length > 0) {
    sheetDestino.getRange(sheetDestino.getLastRow() + 1, 1, filasParaGuardar.length, 5).setValues(filasParaGuardar);
  }
}

function procesarFilaVentasParaTop(fila, stats) {
  let cantidadVenta = parseInt(fila[3]) || 0;
  let nombreProducto = fila[4] ? fila[4].toString().trim() : "";
  if (nombreProducto === "") return; 

  let totalDinero = parseFloat(fila[6]) || 0;
  
  if (nombreProducto.startsWith("PACK:")) {
    let contenidoClean = nombreProducto.replace("PACK: ", "");
    let partes = contenidoClean.split(" + ");
    let totalUnidadesEnUnPack = 0;
    let componentesPack = [];

    partes.forEach(parte => {
       let division = parte.split("x ");
       if(division.length === 2) {
           let cantEnPack = parseInt(division[0]);
           let nombreReal = division[1].trim();
           totalUnidadesEnUnPack += cantEnPack;
           componentesPack.push({ nombre: nombreReal, q: cantEnPack });
       }
    });

    componentesPack.forEach(comp => {
       let cantidadRealTotal = cantidadVenta * comp.q;
       let dineroProporcional = 0;
       if(totalUnidadesEnUnPack > 0) {
           dineroProporcional = (comp.q / totalUnidadesEnUnPack) * totalDinero;
       }
       if (!stats[comp.nombre]) stats[comp.nombre] = { cant: 0, total: 0 };
       stats[comp.nombre].cant += cantidadRealTotal;
       stats[comp.nombre].total += dineroProporcional;
    });
  } else {
    if (!stats[nombreProducto]) stats[nombreProducto] = { cant: 0, total: 0 };
    stats[nombreProducto].cant += cantidadVenta;
    stats[nombreProducto].total += totalDinero;
  }
}

function generarTopClientesMensual(mes, anio, categoria, nombreHojaDestino) {
  const ss = SpreadsheetApp.openById(ID_HOJA);
  let sheetDestino = ss.getSheetByName(nombreHojaDestino);
  
  if (!sheetDestino) {
    sheetDestino = ss.insertSheet(nombreHojaDestino);
    sheetDestino.appendRow(["FECHA", "CLIENTE", "SEDE", "NRO DE COMPRAS", "TOTAL GASTADO (S/.)"]);
    sheetDestino.getRange(1, 1, 1, 5).setFontWeight("bold").setBackground("#C9DAF8");
  }

  const sedes = ["M", "S", "SMP"];
  const nombresSedes = { "M": "MAGDALENA", "S": "SURCO", "SMP": "SAN MARTIN DE PORRES" };
  const etiquetaReporte = `${obtenerNombreMes(mes)} ${anio}`;
  let filasParaGuardar = [];

  sedes.forEach(sede => {
    const hojaClientes = ss.getSheetByName("CLIENTES_" + sede);
    if (!hojaClientes) return;

    const dataClientes = hojaClientes.getDataRange().getValues();
    let stats = {};

    for (let i = 1; i < dataClientes.length; i++) {
      let fila = dataClientes[i];
      let nombreCliente = fila[0] ? fila[0].toString().trim().toUpperCase() : "";
      let categoriaFila = fila[6] ? fila[6].toString().trim() : "";
      let totalVenta = parseFloat(fila[7]) || 0;
      let fechaObj = parseFecha(fila[8]);

      if (fechaObj && fechaObj.getMonth() === mes && fechaObj.getFullYear() === anio) {
        if (nombreCliente === "CLIENTES VARIOS" || nombreCliente === "") continue;
        if (categoria !== "TODOS" && categoriaFila !== categoria) continue;

        if (!stats[nombreCliente]) stats[nombreCliente] = { cant: 0, total: 0 };
        stats[nombreCliente].cant += 1; 
        stats[nombreCliente].total += totalVenta;
      }
    }

    ordenarYGuardarTop(stats, 10, etiquetaReporte, nombresSedes[sede], filasParaGuardar);
  });

  if (filasParaGuardar.length > 0) {
    sheetDestino.getRange(sheetDestino.getLastRow() + 1, 1, filasParaGuardar.length, 5).setValues(filasParaGuardar);
  }
}

function ordenarYGuardarTop(stats, limite, etiqueta, nombreSede, filasParaGuardar) {
  let arrayStats = [];
  for (let key in stats) {
    arrayStats.push({ nombre: key, cantidad: stats[key].cant, total: stats[key].total });
  }

  arrayStats.sort((a, b) => b.cantidad - a.cantidad);
  let topN = arrayStats.slice(0, limite);

  topN.forEach(item => {
    filasParaGuardar.push([etiqueta, item.nombre, nombreSede, item.cantidad, item.total]);
  });
}

function limpiarReportesAntiguos(nombreHoja) {
  const ss = SpreadsheetApp.openById(ID_HOJA);
  const sheet = ss.getSheetByName(nombreHoja);
  if (!sheet || sheet.getLastRow() < 2) return;
  
  let data = sheet.getRange(2, 1, sheet.getLastRow()-1, sheet.getLastColumn()).getValues();
  let tags = [];
  
  data.forEach(r => {
    if (!tags.includes(r[0])) tags.push(r[0]);
  });

  if (tags.length >= 3) {
    let tagReciente = tags[tags.length - 1]; 
    let newData = data.filter(r => r[0] === tagReciente);
    
    sheet.getRange(2, 1, sheet.getLastRow()-1, sheet.getLastColumn()).clearContent();
    if(newData.length > 0) {
      sheet.getRange(2, 1, newData.length, newData[0].length).setValues(newData);
    }
  }
}

function parseFecha(fechaString) {
  if (!fechaString) return null;
  if (fechaString instanceof Date) return fechaString;
  let partes = fechaString.toString().split("/");
  if (partes.length !== 3) return null;
  return new Date(partes[2], partes[1] - 1, partes[0]);
}

function obtenerNombreMes(numMes) {
  const meses = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"];
  return meses[numMes] || "Mes Desconocido";
}
