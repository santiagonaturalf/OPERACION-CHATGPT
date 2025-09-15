/**
 * @OnlyCurrentDoc
 * Script para el flujo de trabajo completo de operaciones: Envasado, Adquisiciones y Dashboard.
 * Versión Final.
 */

// --- LÓGICA DE MENÚS Y DISPARADORES ---

function onOpen() {
  setupProjectSheets();
  const ui = SpreadsheetApp.getUi();

  const operationsMenu = ui.createMenu('Gestión de Operaciones')
    .addItem('🚀 Abrir Dashboard de Operaciones', 'showDashboard')
    .addSeparator()
    .addItem('📝 Generar Adquisiciones', 'showAcquisitionEditor')
    .addItem('🚚 Comanda Rutas', 'showComandaRutasDialog')
    .addItem('💬 Panel de Notificaciones (nuevo)', 'openNotificationPanel')
    .addSeparator()
    .addSeparator();

  operationsMenu.addToUi();

  ui.createMenu('Módulo de Finanzas')
    .addItem('💰 Importar Movimientos', 'showImportMovementsDialog')
    .addItem('📦 Importar Pedidos (Pegar)', 'showPasteImportDialog')
    .addItem('📊 Conciliar Ingresos (Ventas)', 'showConciliationDialog')
    .addToUi();
}

function onEdit(e) {
  const range = e.range;
  const sheet = range.getSheet();
  const sheetName = sheet.getName();
  const row = range.getRow();
  const col = range.getColumn();
  if (sheetName === "Lista de Adquisiciones" && row > 1 && (col === 2 || col === 3)) {
    recalculateRowInventory(sheet, row);
  }
}

// --- SETUP & CONFIGURACIÓN ---

/**
 * Crea todas las hojas necesarias para la aplicación si no existen y notifica al usuario.
 */
function setupProjectSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const createdSheets = [];

  // Helper function to create a sheet with headers if it doesn't exist
  const ensureSheetExists = (sheetName, headers, index) => {
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName, index);
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
      sheet.setFrozenRows(1);
      createdSheets.push(sheetName);

    }
    return sheet;
  };

  // Define all required sheets and their headers
  const sheetsToEnsure = [
    { name: "Orders", headers: ["Order #", "Nombre y apellido", "Email", "Phone", "Shipping Address", "Shipping City", "Shipping Region", "Shipping Postcode", "Item Name", "Item SKU", "Item Quantity", "Item Price", "Line Total", "Tax Rate", "Tax Amount", "Importe total del pedido", "Payment Method", "Transaction ID", "Estado del pago", "Furgon"], index: 0 },
    { name: "SKU", headers: ["Nombre Producto", "Producto Base", "Formato Compra", "Cantidad Compra", "Unidad Compra", "Categoría", "Cantidad Venta", "Unidad Venta", "Proveedor", "Teléfono"], index: 1 },
    { name: "Proveedores", headers: ["Nombre", "Teléfono"], index: 2 },
    { name: "MovimientosBancarios", headers: ["MONTO", "DESCRIPCIÓN MOVIMIENTO", "FECHA", "SALDO", "N° DOCUMENTO", "SUCURSAL", "CARGO/ABONO", "Asignado a Pedido"], index: 3 },
    { name: "AsignacionesHistoricas", headers: ["ID_Pago", "ID_Pedido", "Nombre_Banco", "Nombre_Pedido", "Monto", "Fecha_Asignacion"], index: 4 },
    { name: "Lista de Envasado", headers: ["Cantidad", "Inventario", "Nombre Producto"], index: 5 },
    { name: "Lista de Adquisiciones", headers: ["Producto Base", "Cantidad a Comprar", "Formato de Compra", "Inventario Actual", "Unidad Inventario Actual", "Necesidad de Venta", "Unidad Venta", "Inventario al Finalizar", "Unidad Inventario Final", "Precio Adq. Anterior", "Precio Adq. HOY", "Proveedor"], index: 6 },
    { name: "ClientBankData", headers: ["PaymentIdentifier", "CustomerRUT", "CustomerName", "LastSeen"], index: 7 }
  ];

  sheetsToEnsure.forEach(sheetInfo => {
    ensureSheetExists(sheetInfo.name, sheetInfo.headers, sheetInfo.index);
  });

  // Special check for 'Asignado a Pedido' column in case the sheet already existed
  const movementsSheet = ss.getSheetByName("MovimientosBancarios");
  const currentMovementsHeaders = movementsSheet.getRange(1, 1, 1, movementsSheet.getLastColumn()).getValues()[0];
  if (currentMovementsHeaders.indexOf("Asignado a Pedido") === -1) {
    movementsSheet.getRange(1, currentMovementsHeaders.length + 1).setValue("Asignado a Pedido").setFontWeight("bold");
  }

  // Special check for 'Furgon' column in Orders sheet
  const ordersSheet = ss.getSheetByName("Orders");
  const currentOrdersHeaders = ordersSheet.getRange(1, 1, 1, ordersSheet.getLastColumn()).getValues()[0];
  if (currentOrdersHeaders.indexOf("Furgon") === -1) {
    ordersSheet.getRange(1, currentOrdersHeaders.length + 1).setValue("Furgon").setFontWeight("bold");
  }

  if (createdSheets.length > 0) {
    SpreadsheetApp.getUi().alert(`Se han creado las siguientes hojas que faltaban para el correcto funcionamiento: ${createdSheets.join(', ')}.`);
  }
}

function approveMatch(paymentId, orderId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const movementsSheet = ss.getSheetByName("MovimientosBancarios");
    const ordersSheet = ss.getSheetByName("Orders");
    const assignmentsSheet = ss.getSheetByName("AsignacionesHistoricas");

    // --- Update MovimientosBancarios ---
    const paymentRowIndex = parseInt(paymentId.split('|')[1]);
    const assignedCol = movementsSheet.getRange(1, 1, 1, movementsSheet.getLastColumn()).getValues()[0].indexOf("Asignado a Pedido") + 1;
    if (assignedCol === 0) throw new Error("No se encontró la columna 'Asignado a Pedido'.");

    const existingVal = movementsSheet.getRange(paymentRowIndex, assignedCol).getValue();
    if(existingVal) {
      return { status: "error", message: `Este pago ya ha sido asignado al pedido #${existingVal}.` };
    }
    movementsSheet.getRange(paymentRowIndex, assignedCol).setValue(orderId);

    const paymentData = movementsSheet.getRange(paymentRowIndex, 1, 1, assignedCol).getValues()[0];
    const paymentAmount = paymentData[movementsSheet.getRange(1, 1, 1, movementsSheet.getLastColumn()).getValues()[0].indexOf("MONTO")];
    const paymentDesc = paymentData[movementsSheet.getRange(1, 1, 1, movementsSheet.getLastColumn()).getValues()[0].indexOf("DESCRIPCIÓN MOVIMIENTO")];


    // --- Update Orders ---
    const ordersData = ordersSheet.getDataRange().getValues();
    const headers = ordersData.shift();
    const orderIdCol = 0; // Column A
    const statusCol = 7; // Column H

    let orderCustomerName = '';
    let rowsUpdated = 0;
    ordersData.forEach((row, index) => {
      if (String(row[orderIdCol]) === String(orderId)) {
        ordersSheet.getRange(index + 2, statusCol + 1).setValue("Procesando Conciliacion Aprobada");
        if (!orderCustomerName) {
            orderCustomerName = row[1]; // Column B
        }
        rowsUpdated++;
      }
    });

    if(rowsUpdated === 0) throw new Error(`No se encontraron filas para el pedido #${orderId} para actualizar.`);

    // --- Log to AsignacionesHistoricas ---
    if(assignmentsSheet) {
      assignmentsSheet.appendRow([paymentId, orderId, paymentDesc, orderCustomerName, paymentAmount, new Date()]);
    }

    // --- (NEW) Update ClientBankData ---
    const clientBankSheet = ss.getSheetByName("ClientBankData");
    if (clientBankSheet) {
      const paymentIdentifier = extractNameFromDescription(paymentDesc);
      const customerRUT = ordersData.find(r => String(r[orderIdCol]) === String(orderId))[16];

      if (paymentIdentifier && customerRUT) {
        const bankData = clientBankSheet.getDataRange().getValues();
        const identifierCol = 0;
        let existingRow = -1;

        for (let i = 1; i < bankData.length; i++) {
          if (bankData[i][identifierCol] === paymentIdentifier) {
            existingRow = i + 1;
            break;
          }
        }

        if (existingRow !== -1) {
          clientBankSheet.getRange(existingRow, 4).setValue(new Date());
        } else {
          clientBankSheet.appendRow([paymentIdentifier, customerRUT, orderCustomerName, new Date()]);
        }
      }
    }

    SpreadsheetApp.flush();
    return { status: "success", message: `Pedido #${orderId} asignado correctamente.` };
  } catch (e) {
    Logger.log(e);
    return { status: "error", message: e.toString() };
  }
}

function approveOrderForManagement(orderId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ordersSheet = ss.getSheetByName("Orders");

    const ordersData = ordersSheet.getDataRange().getValues();
    ordersData.shift(); // remove headers
    const orderIdCol = 0; // Column A
    const statusCol = 7; // Column H

    let updatedRows = 0;
    ordersData.forEach((row, index) => {
      if (String(row[orderIdCol]) === String(orderId)) {
        ordersSheet.getRange(index + 2, statusCol + 1).setValue("APROBADO POR GERENCIA");
        updatedRows++;
      }
    });

    if (updatedRows > 0) {
      SpreadsheetApp.flush();
      return { status: "success", message: `Pedido #${orderId} aprobado por gerencia.` };
    } else {
      return { status: "error", message: `No se encontró el pedido #${orderId}.` };
    }
  } catch (e) {
    Logger.log(e);
    return { status: "error", message: e.toString() };
  }
}

function cancelOrder(orderId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ordersSheet = ss.getSheetByName("Orders");

    const ordersData = ordersSheet.getDataRange().getValues();
    ordersData.shift(); // remove headers
    const orderIdCol = 0; // Column A
    const statusCol = 7; // Column H

    let updatedRows = 0;
    ordersData.forEach((row, index) => {
      if (String(row[orderIdCol]) === String(orderId)) {
        ordersSheet.getRange(index + 2, statusCol + 1).setValue("Cancelado");
        updatedRows++;
      }
    });

    if (updatedRows > 0) {
      SpreadsheetApp.flush();
      return { status: "success", message: `Pedido #${orderId} cancelado.` };
    } else {
      return { status: "error", message: `No se encontró el pedido #${orderId}.` };
    }
  } catch (e) {
    Logger.log(e);
    return { status: "error", message: e.toString() };
  }
}

function deleteOrder(orderId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = getSheet_('Orders'); // Using the helper
    const idx = getHeaderIndexes_(sheet, H_ORDENES);

    if (idx.pedido < 0 || idx.cantidad < 0) {
      return { status: 'error', message: 'No se encontraron las columnas "N° Pedido" o "Cantidad" en la hoja "Orders".' };
    }

    const data = sheet.getDataRange().getValues();
    let rowsUpdated = 0;

    data.forEach((row, i) => {
      if (i === 0) return; // Skip header row

      const currentOrderId = String(row[idx.pedido]).trim();
      if (currentOrderId === String(orderId).trim()) {
        const rowRange = sheet.getRange(i + 1, 1, 1, sheet.getLastColumn());
        rowRange.setBackground('#ff0000'); // Red background

        const quantityCell = sheet.getRange(i + 1, idx.cantidad + 1);
        const currentQuantity = quantityCell.getValue();
        if (!String(currentQuantity).startsWith('E')) {
           quantityCell.setValue('E' + currentQuantity);
        }
        rowsUpdated++;
      }
    });

    if (rowsUpdated > 0) {
      SpreadsheetApp.flush();
      return { status: 'success', message: `Pedido #${orderId} (${rowsUpdated} filas) marcado como eliminado.` };
    } else {
      return { status: 'error', message: `No se encontró el pedido #${orderId}.` };
    }
  } catch (e) {
    Logger.log(`Error en deleteOrder: ${e.stack}`);
    return { status: 'error', message: `Ocurrió un error al eliminar el pedido: ${e.message}` };
  }
}


/**
 * Marca filas específicas en la hoja "Orders" como eliminadas.
 * Acepta un array de números de fila.
 */
function deleteSelectedRows(rowNumbers) {
  if (!rowNumbers || !Array.isArray(rowNumbers) || rowNumbers.length === 0) {
    return { status: 'error', message: 'No se proporcionaron filas para eliminar.' };
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = getSheet_('Orders');
    const idx = getHeaderIndexes_(sheet, H_ORDENES);

    if (idx.cantidad < 0) {
      return { status: 'error', message: 'No se encontró la columna "Cantidad" en la hoja "Orders".' };
    }

    let rowsUpdated = 0;
    rowNumbers.forEach(rowNum => {
      // Validar que el número de fila es un número válido y mayor que 1 (para no afectar el encabezado)
      const n = parseInt(rowNum, 10);
      if (isNaN(n) || n <= 1) return;

      const rowRange = sheet.getRange(n, 1, 1, sheet.getLastColumn());
      rowRange.setBackground('#ff0000'); // Fondo rojo

      const quantityCell = sheet.getRange(n, idx.cantidad + 1);
      const currentQuantity = quantityCell.getValue();
      if (!String(currentQuantity).startsWith('E')) {
        quantityCell.setValue('E' + currentQuantity);
      }
      rowsUpdated++;
    });

    if (rowsUpdated > 0) {
      SpreadsheetApp.flush();
      return { status: 'success', message: `${rowsUpdated} fila(s) marcada(s) como eliminada(s).` };
    } else {
      return { status: 'error', message: 'No se actualizó ninguna fila. Verifica los números de fila proporcionados.' };
    }
  } catch (e) {
    Logger.log(`Error en deleteSelectedRows: ${e.stack}`);
    return { status: 'error', message: `Ocurrió un error al eliminar las filas: ${e.message}` };
  }
}

/**
 * Reincorpora filas específicas en la hoja "Orders" que fueron marcadas como eliminadas.
 * Acepta un array de números de fila.
 */
function reincorporateSelectedRows(rowNumbers) {
  if (!rowNumbers || !Array.isArray(rowNumbers) || rowNumbers.length === 0) {
    return { status: 'error', message: 'No se proporcionaron filas para reincorporar.' };
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = getSheet_('Orders');
    const idx = getHeaderIndexes_(sheet, H_ORDENES);

    if (idx.cantidad < 0) {
      return { status: 'error', message: 'No se encontró la columna "Cantidad" en la hoja "Orders".' };
    }

    let rowsUpdated = 0;
    rowNumbers.forEach(rowNum => {
      const n = parseInt(rowNum, 10);
      if (isNaN(n) || n <= 1) return;

      const rowRange = sheet.getRange(n, 1, 1, sheet.getLastColumn());
      rowRange.setBackground('#d9ead3'); // Fondo verde para indicar reincorporación

      const quantityCell = sheet.getRange(n, idx.cantidad + 1);
      const currentQuantity = String(quantityCell.getValue());
      if (currentQuantity.startsWith('E')) {
        quantityCell.setValue(currentQuantity.substring(1));
      }
      rowsUpdated++;
    });

    if (rowsUpdated > 0) {
      SpreadsheetApp.flush();
      return { status: 'success', message: `${rowsUpdated} fila(s) reincorporada(s) exitosamente.` };
    } else {
      return { status: 'error', message: 'No se actualizó ninguna fila.' };
    }
  } catch (e) {
    Logger.log(`Error en reincorporateSelectedRows: ${e.stack}`);
    return { status: 'error', message: `Ocurrió un error al reincorporar las filas: ${e.message}` };
  }
}


// --- LÓGICA DE COMANDA RUTAS ---

function showComandaRutasDialog() {
  const html = HtmlService.createHtmlOutputFromFile('ComandaRutasDialog')
    .setWidth(1200)
    .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, 'Comanda Rutas');
}

/**
 * Returns unique orders (one row per order number) for the routing step.
 */
function getOrdersForRouting() {
  const sheet = SpreadsheetApp.getActive().getSheetByName('Orders');
  if (!sheet) throw new Error('No se encontró la hoja "Orders".');
  const data = sheet.getDataRange().getValues();
  const headers = data.shift();
  const idx = indexer(headers);
  const unique = new Map();
  data.forEach(row => {
    const id = String(row[idx.numPedido] || '').trim();
    if (!id || unique.has(id)) return;
    unique.set(id, {
      orderNumber: id,
      customerName: row[idx.nombre] || '',
      phone: row[idx.telefono] || '',
      address: row[idx.direccion] || '',
      department: row[idx.depto] || '',
      commune: row[idx.comuna] || '',
      status: row[idx.estado] || '',
      van: row[idx.furgon] || ''
    });
  });
  return [...unique.values()];
}

/**
 * Saves changes for a single order number across all matching rows.
 */
function saveSingleOrderChange(order) {
  const sheet = SpreadsheetApp.getActive().getSheetByName('Orders');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idx = indexer(headers);
  const ops = [];
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idx.numPedido]) === String(order.orderNumber)) {
      const r = i + 1;
      if (idx.direccion >= 0) ops.push(() => sheet.getRange(r, idx.direccion + 1).setValue(order.address || ''));
      if (idx.depto     >= 0) ops.push(() => sheet.getRange(r, idx.depto + 1).setValue(order.department || ''));
      if (idx.comuna    >= 0) ops.push(() => sheet.getRange(r, idx.comuna + 1).setValue(order.commune || ''));
      if (idx.furgon    >= 0) ops.push(() => sheet.getRange(r, idx.furgon + 1).setValue(order.van || ''));
      if (idx.telefono  >= 0) ops.push(() => sheet.getRange(r, idx.telefono + 1).setValue(order.phone || ''));
    }
  }
  ops.forEach(fn => fn());
  SpreadsheetApp.flush();
  return { status: 'success', message: `Guardado #${order.orderNumber}` };
}

/**
 * Saves multiple orders in a batch, updating all rows for each order number.
 */
function saveRouteChanges(orders) {
  const sheet = SpreadsheetApp.getActive().getSheetByName('Orders');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idx = indexer(headers);
  const rowsById = new Map();
  for (let i = 1; i < data.length; i++) {
    const id = String(data[i][idx.numPedido] || '');
    if (!id) continue;
    if (!rowsById.has(id)) rowsById.set(id, []);
    rowsById.get(id).push(i + 1);
  }
  const ops = [];
  orders.forEach(o => {
    (rowsById.get(String(o.orderNumber)) || []).forEach(r => {
      if (idx.direccion >= 0) ops.push(() => sheet.getRange(r, idx.direccion + 1).setValue(o.address || ''));
      if (idx.depto     >= 0) ops.push(() => sheet.getRange(r, idx.depto + 1).setValue(o.department || ''));
      if (idx.comuna    >= 0) ops.push(() => sheet.getRange(r, idx.comuna + 1).setValue(o.commune || ''));
      if (idx.furgon    >= 0) ops.push(() => sheet.getRange(r, idx.furgon + 1).setValue(o.van || ''));
      if (idx.telefono  >= 0) ops.push(() => sheet.getRange(r, idx.telefono + 1).setValue(o.phone || ''));
    });
  });
  ops.forEach(fn => fn());
  SpreadsheetApp.flush();
  return { status: 'success', message: `Se guardaron ${orders.length} pedido(s).` };
}

/**
 * Column index helper (supports multiple synonyms for column headings). Update synonyms as needed.
 */
function indexer(headers) {
  const norm = s => String(s || '').toLowerCase().trim();
  const idxOf = (...names) => headers.findIndex(h => names.includes(norm(h)));
  return {
    numPedido: idxOf('order #', 'número de pedido','numero de pedido','nº pedido','n° pedido','n pedido','número de ped'),
    nombre:    idxOf('nombre y apellido', 'nombre completo','cliente','nombre'),
    telefono:  idxOf('teléfono','telefono','phone','tel'),
    direccion: idxOf('dirección','direccion','shipping address','dirección líneas 1','direccion lineas 1'),
    depto:     idxOf('depto.','depto','departamento','depto/condominio','depto/condomi','dirección líneas 2','direccion lineas 2'),
    comuna:    idxOf('comuna','shipping region','ciudad'),
    estado:    idxOf('estado','estado del pago'),
    furgon:    idxOf('furgón','furgon','van','furgón asignado'),
    cantidad:  idxOf('item quantity', 'cantidad')
  };
}

function processRouteXLData(pastedText, vanName) {
  if (!vanName) {
    throw new Error("Se requiere un nombre de furgón para procesar la ruta.");
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const lines = pastedText.split('\n');
  const orderedOrderNumbers = lines.map(line => {
    const match = line.match(/#\d+/);
    return match ? match[0] : null;
  }).filter(Boolean);

  if (orderedOrderNumbers.length === 0) {
    throw new Error("No se pudieron encontrar números de pedido válidos (ej: #1234) en el texto pegado.");
  }

  const ordersSheet = ss.getSheetByName('Orders');
  if (!ordersSheet) throw new Error('No se encontró la hoja "Orders".');
  const ordersData = ordersSheet.getDataRange().getValues();
  const headers = ordersData.shift();

  const ordersMap = {};
  ordersData.forEach(row => {
    const orderNumber = String(row[0]);
    if (!ordersMap[orderNumber]) ordersMap[orderNumber] = [];
    ordersMap[orderNumber].push(row);
  });

  const sortedData = [];
  orderedOrderNumbers.forEach(orderNumberWithHash => {
    const cleanOrderNumber = orderNumberWithHash.replace('#', '');
    if (ordersMap[cleanOrderNumber]) {
      sortedData.push(...ordersMap[cleanOrderNumber]);
    }
  });

  const routeSheetName = `Ruta Optimizada - ${vanName}`;
  let routeSheet = ss.getSheetByName(routeSheetName);
  if (routeSheet) {
    routeSheet.clear();
  } else {
    routeSheet = ss.insertSheet(routeSheetName);
  }

  routeSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (sortedData.length > 0) {
    routeSheet.getRange(2, 1, sortedData.length, sortedData[0].length).setValues(sortedData);
  }
  routeSheet.autoResizeColumns(1, headers.length);

  return generatePrintableRouteSheets(vanName);
}

function generatePrintableRouteSheets(vanName) {
  if (!vanName) {
    throw new Error("Se requiere un nombre de furgón para generar las hojas de ruta.");
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const routeSheetName = `Ruta Optimizada - ${vanName}`;
  const routeSheet = ss.getSheetByName(routeSheetName);
  if (!routeSheet) {
    throw new Error(`Primero debe procesar los datos de RouteXL para el furgón "${vanName}".`);
  }

  const routeData = routeSheet.getDataRange().getValues();
  const headers = routeData.shift();
  const vanColumnIndex = headers.indexOf('Furgón');

  const orderSequence = [];
  const seenOrders = new Set();
  routeData.forEach(row => {
    const orderNumber = row[0];
    if (orderNumber && !seenOrders.has(orderNumber)) {
        orderSequence.push(orderNumber);
        seenOrders.add(orderNumber);
    }
  });

  const packagingOrderSequence = [...orderSequence].reverse(); // Crear una copia invertida
  const finalPackagingData = [];
  packagingOrderSequence.forEach((orderNumber, index) => {
      finalPackagingData.push([
          packagingOrderSequence.length - index, // Orden descendente (4, 3, 2, 1)
          orderNumber,      // Nº Pedido (en orden inverso)
          "\n\n\n",         // Numero de Bultos (con saltos de línea)
          "\n\n\n"          // Nombre Envasador (con saltos de línea)
      ]);
  });

  const packagingSheetName = `Orden de Envasado - ${vanName}`;
  let packagingSheet = ss.getSheetByName(packagingSheetName);
  if (packagingSheet) {
    packagingSheet.clear();
  } else {
    packagingSheet = ss.insertSheet(packagingSheetName);
  }

  // Añadir título principal a la hoja de envasado
  packagingSheet.getRange("A1").setValue(vanName).setFontSize(18).setFontWeight('bold').setHorizontalAlignment('center');
  packagingSheet.getRange("A1:D1").merge();

  const packagingHeaders = ["Orden Ruta", "Nº Pedido", "Numero de Bultos", "Nombre Envasador"];
  packagingSheet.getRange("A2:D2").setValues([packagingHeaders]).setFontWeight('bold');

  if (finalPackagingData.length > 0) {
    packagingSheet.getRange(3, 1, finalPackagingData.length, 4).setValues(finalPackagingData);

    // Aplicar formato a la tabla (incluyendo encabezados)
    const tableRange = packagingSheet.getRange(2, 1, finalPackagingData.length + 1, 4);
    tableRange.setHorizontalAlignment("center")
              .setVerticalAlignment("middle")
              .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  }

  // Ajustar anchos de columna
  packagingSheet.autoResizeColumns(1, 2); // Auto-ajustar Orden y Pedido
  packagingSheet.setColumnWidth(3, 200);  // Ancho para Numero de Bultos
  packagingSheet.setColumnWidth(4, 200);  // Ancho para Nombre Envasador

  const loadingData = [];
  orderSequence.forEach((orderNumber, index) => {
      const orderRow = routeData.find(row => row[0] === orderNumber);
      if(orderRow) {
        const address = orderRow[4] || '';
        const dept = orderRow[5] || '';
        const fullAddress = [address, dept].filter(Boolean).join(', ');

        loadingData.push([
          index + 1,          // Orden Carga
          orderNumber,        // Nº Pedido
          orderRow[1],        // Cliente
          "\n\n\n",           // BULTOS con saltos de línea para altura
          fullAddress,        // Dirección Completa
          orderRow[6],        // Comuna
          orderRow[3] || ''   // TELEFONO
        ]);
      }
  });

  const loadingSheetName = `Orden de Carga - ${vanName}`;
  let loadingSheet = ss.getSheetByName(loadingSheetName);
  if (loadingSheet) {
    loadingSheet.clear();
  } else {
    loadingSheet = ss.insertSheet(loadingSheetName);
  }

  // Añadir título principal
  loadingSheet.getRange("A1").setValue(vanName).setFontSize(18).setFontWeight('bold').setHorizontalAlignment('center');
  loadingSheet.getRange("A1:G1").merge();

  // Encabezados de la tabla
  const loadingHeaders = ["Orden Carga", "Nº Pedido", "Cliente", "BULTOS", "Dirección Completa", "Comuna", "TELEFONO"];
  loadingSheet.getRange("A2:G2").setValues([loadingHeaders]).setFontWeight('bold');

  // Escribir datos si existen
  if (loadingData.length > 0) {
    loadingSheet.getRange(3, 1, loadingData.length, 7).setValues(loadingData);
  }

  // Aplicar formato a toda la tabla
  if (loadingData.length > 0) {
    const tableRange = loadingSheet.getRange(2, 1, loadingData.length + 1, 7);
    tableRange.setHorizontalAlignment("center")
              .setVerticalAlignment("middle")
              .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  }

  // Ajustar anchos de columna
  loadingSheet.setColumnWidth(1, 50);     // Ancho fijo y pequeño para Orden Carga (Col A)
  loadingSheet.autoResizeColumns(2, 2);   // Auto-ajustar Nº Pedido y Cliente (Col B, C)
  loadingSheet.setColumnWidth(4, 400);    // Ancho fijo para BULTOS (Col D), aumentado
  loadingSheet.autoResizeColumns(5, 3);   // Auto-ajustar Dirección, Comuna y Teléfono (Col E, F, G)


  const spreadsheetId = ss.getId();
  const packagingPdfUrl = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export?format=pdf&gid=${packagingSheet.getSheetId()}&portrait=true&fitw=true&gridlines=true&printtitle=false`;
  const routePdfUrl = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export?format=pdf&gid=${loadingSheet.getSheetId()}&portrait=false&fitw=true&gridlines=true&printtitle=false&sheetnames=false`;

  return {
    status: 'success',
    message: `Se han generado las hojas de ruta y envasado para ${vanName}.`,
    packagingPdfUrl: packagingPdfUrl,
    routePdfUrl: routePdfUrl
  };
}

function cleanupRouteSheets() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const allSheets = ss.getSheets();
    let deletedCount = 0;

    allSheets.forEach(sheet => {
      const sheetName = sheet.getName();
      if (sheetName.startsWith('Ruta Optimizada -') || sheetName.startsWith('Orden de Envasado -') || sheetName.startsWith('Orden de Carga -')) {
        ss.deleteSheet(sheet);
        deletedCount++;
      }
    });

    return { status: 'success', message: `Limpieza completada. Se eliminaron ${deletedCount} hojas.` };
  } catch (e) {
    Logger.log(e);
    return { status: 'error', message: e.toString() };
  }
}


// --- MÓDULO DE FINANZAS ---

function showImportMovementsDialog() {
  const html = HtmlService.createHtmlOutputFromFile('ImportMovementsDialog')
    .setWidth(600)
    .setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(html, 'Importar Movimientos Bancarios');
}

function importBankMovements(data) {
  if (!data || typeof data !== 'string') {
    throw new Error("No se proporcionaron datos válidos para importar.");
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("MovimientosBancarios");
  if (!sheet) {
    throw new Error("No se encontró la hoja 'MovimientosBancarios'. Por favor, vuelve a abrir el documento para que se cree automáticamente.");
  }

  const newRows = data.trim().split('\n').map(line => line.split('\t'));
  if (newRows.length === 0) {
    return "No se encontraron filas para importar.";
  }

  // 1. Read existing data using a more robust method
  const allData = sheet.getDataRange().getValues();
  const headers = allData.shift(); // Remove header row
  const existingData = allData;   // The rest is existing data

  Logger.log(`Total historical rows read: ${existingData.length}`);

  // Key with Amount, Description, and Date for robust duplicate detection.
  const existingKeys = new Set(existingData.map(row =>
    // Using MONTO (col 0), DESCRIPCION (col 1), y FECHA (col 2)
    `${String(row[0]).trim()}|${String(row[1]).trim()}|${String(row[2]).trim()}`
  ));

  if (existingKeys.size > 0) {
    Logger.log(`Sample historical key (Amount + Desc + Date): ${existingKeys.values().next().value}`);
  }

  // 3. Filter out duplicates from the new rows
  const rowsToInsert = [];
  let duplicateCount = 0;

  newRows.forEach((row, index) => {
    // Key with Amount, Description, and Date for robust duplicate detection.
    const key = `${String(row[0]).trim()}|${String(row[1]).trim()}|${String(row[2]).trim()}`;
    if (index === 0) {
      Logger.log(`Sample new key (Amount + Desc + Date): ${key}`);
      Logger.log(`Does historical set have this new key? ${existingKeys.has(key)}`);
    }
    if (!existingKeys.has(key)) {
      rowsToInsert.push(row);
      existingKeys.add(key); // Add new key to set to avoid duplicate imports in the same batch
    } else {
      duplicateCount++;
    }
  });

  // 4. Insert only the new, non-duplicate rows
  if (rowsToInsert.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rowsToInsert.length, rowsToInsert[0].length).setValues(rowsToInsert);
  }

  // 5. Update the return message
  let message = `Importación completada.`;
  if (rowsToInsert.length > 0) {
    message += ` Se añadieron ${rowsToInsert.length} nuevos movimientos.`;
  }
  if (duplicateCount > 0) {
    message += ` Se omitieron ${duplicateCount} movimientos duplicados.`;
  }
  if (rowsToInsert.length === 0 && duplicateCount === 0) {
    message = "No se importó nada. Revisa los datos pegados.";
  }

  return message;
}

function showFinanceDashboard() {
  const html = HtmlService.createHtmlOutputFromFile('FinanceDashboardDialog')
    .setWidth(500)
    .setHeight(350);
  SpreadsheetApp.getUi().showModalDialog(html, 'Módulo de Finanzas');
}

function showConciliationDialog() {
  const html = HtmlService.createHtmlOutputFromFile('SalesReconciliationDialog')
    .setWidth(1000)
    .setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, 'Conciliar Ingresos de Ventas');
}

function getReconciliationData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const movementsSheet = ss.getSheetByName("MovimientosBancarios");
  const ordersSheet = ss.getSheetByName("Orders");
  const clientBankSheet = ss.getSheetByName("ClientBankData");

  if (!movementsSheet || !ordersSheet || !clientBankSheet) {
    throw new Error("Una o más hojas requeridas no se encontraron: MovimientosBancarios, Orders, ClientBankData.");
  }

  // --- 1. Fetch all necessary data ---
  const movementsData = movementsSheet.getDataRange().getValues();
  const ordersData = ordersSheet.getDataRange().getValues();
  const clientBankData = clientBankSheet.getDataRange().getValues();

  // --- 2. Prepare initial lists ---
  const movementsHeaders = movementsData.shift();
  const assignedColIdx = movementsHeaders.indexOf("Asignado a Pedido");
  const chargeColIdx = movementsHeaders.indexOf("CARGO/ABONO");
  const amountColIdx = movementsHeaders.indexOf("MONTO");
  const descColIdx = movementsHeaders.indexOf("DESCRIPCIÓN MOVIMIENTO");
  const dateColIdx = movementsHeaders.indexOf("FECHA");

  let unassignedPayments = [];
  movementsData.forEach((row, index) => {
    if (row[chargeColIdx] === 'A' && !row[assignedColIdx]) {
      const amount = parseFloat(String(row[amountColIdx]).replace(/[^0-9,-]+/g,"").replace(",", "."));
      if (isNaN(amount) || amount <= 0) return;
      let paymentDate;
      const dateCell = row[dateColIdx];
      if (dateCell instanceof Date) paymentDate = dateCell;
      else if (typeof dateCell === 'string' && dateCell) paymentDate = parseDDMMYYYY(dateCell);
      if (!paymentDate || isNaN(paymentDate.getTime())) return;
      unassignedPayments.push({ amount, desc: row[descColIdx], date: paymentDate, extractedName: extractNameFromDescription(row[descColIdx]), paymentId: `row|${index + 2}` });
    }
  });

  ordersData.shift();
  const REAL_ORDER_ID_COL = 0, REAL_CUSTOMER_NAME_COL = 1, REAL_STATUS_COL = 7, REAL_ORDER_DATE_COL = 8, REAL_TOTAL_AMOUNT_COL = 15, REAL_PAYMENT_METHOD_COL = 18, REAL_PHONE_COL = 3, REAL_RUT_COL = 16;
  const pendingOrdersMap = {};
  ordersData.forEach((row, index) => {
    const orderId = row[REAL_ORDER_ID_COL];
    if (!orderId) return;
    const status = String(row[REAL_STATUS_COL]).trim();
    const method = row[REAL_PAYMENT_METHOD_COL];
    const orderDate = new Date(row[REAL_ORDER_DATE_COL]);
    const isEligible = (method === 'bacs' && (status === 'En Espera de Pago' || status === 'Procesando') && orderDate instanceof Date && !isNaN(orderDate));
    if (isEligible && !pendingOrdersMap[orderId]) {
       const totalAmount = parseFloat(String(row[REAL_TOTAL_AMOUNT_COL]).replace(/[^0-9,-]+/g,"").replace(",","."));
       if (isNaN(totalAmount) || totalAmount <= 0) return;
       pendingOrdersMap[orderId] = { orderId, customerName: row[REAL_CUSTOMER_NAME_COL], phone: row[REAL_PHONE_COL], normalizedPhone: normalizePhoneNumber(row[REAL_PHONE_COL]), totalAmount, date: orderDate, status, rowNumber: index + 2, customerRUT: row[REAL_RUT_COL] };
    }
  });
  let pendingOrders = Object.values(pendingOrdersMap);

  // --- 3. Matching Logic ---
  const historicalSuggestions = [], highConfidenceSuggestions = [], lowConfidenceSuggestions = [];
  const matchedPaymentIds = new Set(), matchedOrderIds = new Set();

  const clientBankMap = new Map(clientBankData.slice(1).map(row => [row[0], row[1]]));

  // Tier 1: Historical Matching
  unassignedPayments.forEach(payment => {
    const paymentIdentifier = payment.extractedName;
    const customerRUT = clientBankMap.get(paymentIdentifier);
    if (customerRUT) {
      const order = pendingOrders.find(o => o.customerRUT === customerRUT && !matchedOrderIds.has(o.orderId));
      if (order) {
        historicalSuggestions.push({ payment, order });
        matchedPaymentIds.add(payment.paymentId);
        matchedOrderIds.add(order.orderId);
      }
    }
  });

  // Tiers 2 & 3: Score-Based Matching
  unassignedPayments.filter(p => !matchedPaymentIds.has(p.paymentId)).forEach(payment => {
    let bestMatch = { order: null, score: 0, amountScore: 0, nameScore: 0, dateScore: 0 };
    pendingOrders.filter(o => !matchedOrderIds.has(o.orderId)).forEach(order => {
      if (payment.date < new Date(order.date.getTime() - 24*60*60*1000)) return;
      const amountDiff = Math.abs(payment.amount - order.totalAmount);
      let amountScore = 0;
      if (amountDiff === 0) amountScore = 100;
      else if (amountDiff < 5000) amountScore = 100 - (amountDiff / 50);
      else return;
      const msPerDay = 1000 * 60 * 60 * 24;
      const dayDifference = Math.floor((new Date(payment.date.getFullYear(), payment.date.getMonth(), payment.date.getDate()) - new Date(order.date.getFullYear(), order.date.getMonth(), order.date.getDate())) / msPerDay);
      if (dayDifference < 0) return;
      const dateScore = Math.max(0, 100 - (dayDifference * 10));
      if (dateScore <= 0 && dayDifference > 0) return;
      const nameScore = calculateNameSimilarity(payment.extractedName, order.customerName);
      if (nameScore < 20) return;
      const totalScore = (amountScore * 0.5) + (nameScore * 0.3) + (dateScore * 0.2);
      if (totalScore > bestMatch.score) bestMatch = { order, score: totalScore, amountScore, nameScore, dateScore };
    });

    if (bestMatch.order && bestMatch.score > 65) {
      const suggestion = { payment, order: bestMatch.order, score: Math.round(bestMatch.score), amountScore: Math.round(bestMatch.amountScore), nameScore: Math.round(bestMatch.nameScore), dateScore: Math.round(bestMatch.dateScore) };
      if (bestMatch.amountScore === 100) highConfidenceSuggestions.push(suggestion);
      else lowConfidenceSuggestions.push(suggestion);
      matchedPaymentIds.add(payment.paymentId);
      matchedOrderIds.add(bestMatch.order.orderId);
    }
  });

  // --- 4. Prepare return data ---
  const formatDate = (date) => (date instanceof Date && !isNaN(date)) ? Utilities.formatDate(date, Session.getScriptTimeZone(), "dd/MM/yyyy") : "Fecha Inválida";
  const formatSuggestion = s => ({ ...s, payment: { ...s.payment, date: formatDate(s.payment.date) }, order: { ...s.order, date: formatDate(s.order.date) } });

  const manualListOrders = pendingOrders.filter(o => o.status === 'En Espera de Pago');

  const summaryState = loadSummaryState(); // Load persisted summary state

  return {
    historicalSuggestions: historicalSuggestions.map(formatSuggestion),
    highConfidenceSuggestions: highConfidenceSuggestions.map(formatSuggestion),
    lowConfidenceSuggestions: lowConfidenceSuggestions.map(formatSuggestion),
    unmatchedPayments: unassignedPayments.map(p => ({ ...p, date: formatDate(p.date) })),
    manualListOrders: manualListOrders.map(o => ({ ...o, date: formatDate(o.date) })),
    approvedSummary: summaryState // Add summary state to the return object
  };
}

/**
 * Saves the client-side summary state to user properties.
 * @param {Array} summaryData The array of summary objects to save.
 */
function saveSummaryState(summaryData) {
  try {
    const properties = PropertiesService.getUserProperties();
    properties.setProperty('reconciliationSummary', JSON.stringify(summaryData));
  } catch (e) {
    Logger.log(`Error saving summary state: ${e.toString()}`);
    // No need to throw an error to the client for this, as it's a background save.
  }
}

/**
 * Loads the summary state from user properties.
 * @returns {Array} The parsed summary array or an empty array if not found or error.
 */
function loadSummaryState() {
  try {
    const properties = PropertiesService.getUserProperties();
    const jsonState = properties.getProperty('reconciliationSummary');
    return jsonState ? JSON.parse(jsonState) : [];
  } catch (e) {
    Logger.log(`Error loading summary state: ${e.toString()}`);
    return [];
  }
}

/**
 * Clears the summary state from user properties.
 */
function clearSummaryState() {
  try {
    const properties = PropertiesService.getUserProperties();
    properties.deleteProperty('reconciliationSummary');
    return { status: 'success', message: 'Resumen limpiado.' };
  } catch (e) {
    Logger.log(`Error clearing summary state: ${e.toString()}`);
    return { status: 'error', message: `Error al limpiar: ${e.toString()}` };
  }
}

// --- GESTIÓN DE PEDIDOS (AGREGAR POR LOTE) ---

/**
 * Muestra un diálogo para agregar nuevos pedidos pegando texto.
 */
function showAppendOrdersDialog() {
  const html = HtmlService.createHtmlOutputFromFile('AppendOrdersDialog')
    .setWidth(650)
    .setHeight(450);
  SpreadsheetApp.getUi().showModalDialog(html, 'Agregar Pedidos por Lote');
}

/**
 * Analiza texto separado por tabulaciones y lo anexa a la hoja "Orders".
 * Asume que la primera fila del texto pegado es un encabezado y la ignora.
 */
function appendOrdersFromPastedText(textData) {
  try {
    if (!textData || typeof textData !== 'string' || textData.trim() === '') {
      throw new Error("No se proporcionaron datos para importar.");
    }

    const sheet = getSheet_('Orders');

    // Dividir el texto en filas y luego en celdas por tabulación.
    let rows = textData.trim().split('\n').map(row => row.split('\t'));

    // Se asume que el usuario incluye encabezados, por lo que se elimina la primera fila.
    rows.shift();

    if (rows.length === 0) {
      return { status: 'success', message: "No se encontraron filas de datos para agregar (se omitió el encabezado)." };
    }

    // Anexar las nuevas filas a la hoja.
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);

    SpreadsheetApp.flush();
    return { status: 'success', message: `Se agregaron ${rows.length} nuevos pedidos exitosamente.` };

  } catch (e) {
    Logger.log(`Error en appendOrdersFromPastedText: ${e.stack}`);
    return { status: 'error', message: `Ocurrió un error: ${e.message}` };
  }
}


// --- GESTIÓN DE PEDIDOS (ELIMINAR/VER) ---

/**
 * Muestra un reporte de los pedidos marcados como eliminados.
 */
function showDeletedOrders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  try {
    const sheet = getSheet_('Orders');
    // Extend H_ORDENES for this function
    const H_ORDENES_EXT = {
      ...H_ORDENES,
      nombre: ['nombre completo', 'cliente', 'nombre', 'nombre y apellido']
    };
    const idx = getHeaderIndexes_(sheet, H_ORDENES_EXT);

    if (idx.pedido < 0 || idx.cantidad < 0 || idx.producto < 0 || idx.nombre < 0) {
      throw new Error('Faltan una o más columnas requeridas en "Orders" (Pedido, Cantidad, Producto, Nombre).');
    }

    const data = sheet.getDataRange().getValues();
    const deletedOrders = [];

    data.forEach((row, i) => {
      if (i === 0) return; // Skip header

      const quantity = String(row[idx.cantidad]);
      if (quantity.startsWith('E')) {
        deletedOrders.push({
          orderNumber: row[idx.pedido],
          customerName: row[idx.nombre],
          productName: row[idx.producto],
          quantity: quantity
        });
      }
    });

    if (deletedOrders.length === 0) {
      ui.alert('No se encontraron pedidos eliminados.');
      return;
    }

    let report = `Reporte de Pedidos Eliminados (${deletedOrders.length} items):\n\n`;
    const groupedByOrder = {};
    deletedOrders.forEach(item => {
        if (!groupedByOrder[item.orderNumber]) {
            groupedByOrder[item.orderNumber] = {
                customerName: item.customerName,
                items: []
            };
        }
        groupedByOrder[item.orderNumber].items.push(`- ${item.productName} (Cantidad: ${item.quantity})`);
    });

    for (const orderNumber in groupedByOrder) {
        report += `Pedido: #${orderNumber}\n`;
        report += `Cliente: ${groupedByOrder[orderNumber].customerName}\n`;
        report += groupedByOrder[orderNumber].items.join('\n');
        report += `\n\n`;
    }

    // Use a preformatted block for better readability in the alert
    const output = HtmlService.createHtmlOutput(`<pre>${report}</pre>`).setWidth(500).setHeight(400);
    ui.showModalDialog(output, 'Pedidos Eliminados');

  } catch (e) {
    Logger.log(e);
    ui.alert('Error', `Ocurrió un error al generar el reporte: ${e.message}`, ui.ButtonSet.OK);
  }
}

// --- DASHBOARD V2 (IMPLEMENTACIÓN DEL USUARIO) ---

/**
 * Obtiene los datos de los pedidos para el nuevo panel de eliminación.
 * Agrupa los artículos por número de pedido e incluye el número de fila de cada artículo.
 * Omite los artículos que ya han sido marcados como eliminados (con 'E' en la cantidad).
 */
function getOrdersForDeletion() {
  try {
    const sheet = getSheet_('Orders');
    const data = sheet.getDataRange().getValues();
    const headers = data.shift(); // Saca los encabezados

    // Usar el indexer para encontrar columnas de forma robusta
    const idx = indexer(headers);
    // Añadir índices para columnas que no están en el indexer estándar
    const norm = s => String(s || '').toLowerCase().trim();
    idx.producto = headers.findIndex(h => ['item name', 'nombre producto', 'producto'].includes(norm(h)));
    idx.cantidad = headers.findIndex(h => ['item quantity', 'cantidad'].includes(norm(h)));

    // Validar que las columnas esenciales existen
    if (idx.numPedido < 0 || idx.producto < 0 || idx.cantidad < 0) {
      throw new Error("No se encontraron columnas críticas como 'Número de pedido', 'Item Name' o 'Item Quantity'.");
    }

    const orders = {};

    data.forEach((row, i) => {
      const orderId = String(row[idx.numPedido] || '').trim();
      const quantity = String(row[idx.cantidad] || '');

      // Omitir filas sin ID de pedido o ya marcadas como eliminadas
      if (!orderId || quantity.startsWith('E')) {
        return;
      }

      // Si es la primera vez que vemos este ID de pedido, creamos la entrada principal
      if (!orders[orderId]) {
        orders[orderId] = {
          orderNumber: orderId,
          customerName: row[idx.nombre] || 'N/A',
          status: row[idx.estado] || 'N/A',
          commune: row[idx.comuna] || 'N/A',
          van: row[idx.furgon] || 'N/A',
          items: []
        };
      }

      // Añadir el artículo (producto) a la lista de ese pedido
      orders[orderId].items.push({
        productName: row[idx.producto] || 'Producto sin nombre',
        quantity: quantity,
        rowNumber: i + 2 // i es 0-indexed y la fila 1 era el encabezado, así que +2
      });
    });

    return { ok: true, orders: orders };
  } catch (e) {
    Logger.log(`Error en getOrdersForDeletion: ${e.stack}`);
    return { ok: false, error: e.toString() };
  }
}

/**
 * Obtiene los datos de los pedidos MARCADOS COMO ELIMINADOS para el panel de reincorporación.
 * Solo incluye artículos cuya cantidad comienza con 'E'.
 */
function getDeletedOrders() {
  try {
    const sheet = getSheet_('Orders');
    const data = sheet.getDataRange().getValues();
    const headers = data.shift();

    const idx = indexer(headers);
    const norm = s => String(s || '').toLowerCase().trim();
    idx.producto = headers.findIndex(h => ['item name', 'nombre producto', 'producto'].includes(norm(h)));
    idx.cantidad = headers.findIndex(h => ['item quantity', 'cantidad'].includes(norm(h)));

    if (idx.numPedido < 0 || idx.producto < 0 || idx.cantidad < 0) {
      throw new Error("No se encontraron columnas críticas como 'Número de pedido', 'Item Name' o 'Item Quantity'.");
    }

    const orders = {};

    data.forEach((row, i) => {
      const orderId = String(row[idx.numPedido] || '').trim();
      const quantity = String(row[idx.cantidad] || '');

      // Solo incluir filas marcadas como eliminadas
      if (!orderId || !quantity.startsWith('E')) {
        return;
      }

      if (!orders[orderId]) {
        orders[orderId] = {
          orderNumber: orderId,
          customerName: row[idx.nombre] || 'N/A',
          status: row[idx.estado] || 'N/A',
          commune: row[idx.comuna] || 'N/A',
          van: row[idx.furgon] || 'N/A',
          items: []
        };
      }

      orders[orderId].items.push({
        productName: row[idx.producto] || 'Producto sin nombre',
        quantity: quantity.substring(1), // Quita la 'E' para mostrar el número original
        rowNumber: i + 2
      });
    });

    return { ok: true, orders: orders };
  } catch (e) {
    Logger.log(`Error en getDeletedOrders: ${e.stack}`);
    return { ok: false, error: e.toString() };
  }
}


function showDashboard() {
  const html = HtmlService.createTemplateFromFile('DashboardDialog').evaluate()
    .setWidth(1200)
    .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, 'Dashboard Operaciones SNF');
}

/*************************** CONFIG ***************************/
// Nombre de hojas esperadas
const SH_ORDENES            = 'Orders';          // Debe contener cabeza con: N° Pedido, Nombre Producto, Cantidad, Comuna (o similar)
const SH_LISTA_ADQ          = 'Lista de Adquisiciones';// Debe contener las columnas del ejemplo entregado

// Claves de encabezados (mapeo robusto por nombre)
const H_ORDENES = {
  pedido:  ['N° Pedido','Nº Pedido','Numero Pedido','Número de Pedido','Pedido'],
  producto:['Nombre Producto','Producto','Item','Ítem'],
  cantidad:['Cantidad','Qty','Cantidad Venta','Cant'],
  comuna:  ['Comuna','Ciudad','Sector']
};

const H_ADQ = {
  productoBase:        ['Producto Base','Producto','Nombre Producto','Base'],
  cantComprar:         ['Cantidad a Comprar','Cantidad','Cant Comprar'],
  formatoCompra:       ['Formato de Compra','Formato','Presentación'],
  invActual:           ['Inventario Actual','Stock Actual','Inventario'],
  unidadInvActual:     ['Unidad Inventario Actual','Unidad Inv Actual','Unidad Inventario'],
  necesidadVenta:      ['Necesidad de Venta','Necesidad','Venta Necesaria'],
  unidadVenta:         ['Unidad Venta','Unidad Venta (Nombre)','Unidad Vta'],
  invFinalizar:        ['Inventario al Finalizar','Inventario Final','Stock Final'],
  unidadInvFinal:      ['Unidad Inventario Final','Unidad Inv Final','Unidad Final'],
  precioAdqAnterior:   ['Precio Adq. Anterior','Precio Anterior'],
  precioAdqHoy:        ['Precio Adq. HOY','Precio Hoy','Precio Actual'],
  proveedor:           ['Proveedor','Vendor']
};

/*************************** UTILIDADES ***************************/
function getSheet_(name) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(name);
  if (!sh) throw new Error('No existe la hoja: ' + name);
  return sh;
}

function mapHeaders_(row) {
  const map = {};
  row.forEach((h, i) => {
    map[(h||'').toString().trim()] = i;
  });
  return map;
}

function pickIdx_(headerIndexMap, aliases){
  for (const alias of aliases){
    const k = Object.keys(headerIndexMap).find(x => x.toLowerCase() === alias.toLowerCase());
    if (k) return headerIndexMap[k];
  }
  // También aceptar contiene (más laxo)
  for (const alias of aliases){
    const k = Object.keys(headerIndexMap).find(x => x.toLowerCase().includes(alias.toLowerCase()));
    if (k) return headerIndexMap[k];
  }
  return -1;
}

function getHeaderIndexes_(sh, headerAliases){
  const lastCol = sh.getLastColumn();
  if (lastCol === 0) throw new Error('Hoja vacía: ' + sh.getName());
  const headers = sh.getRange(1,1,1,lastCol).getValues()[0];
  const imap = mapHeaders_(headers);
  const out = {};
  Object.keys(headerAliases).forEach(key => {
    out[key] = pickIdx_(imap, headerAliases[key]);
  });
  return out;
}

/*************************** DISTRIBUCIÓN POR COMUNAS ***************************/
/**
 * Devuelve [{comuna, cantidadPedidos}]
 */
function getDashboardSummary() {
  try {
    const sh = getSheet_(SH_ORDENES);
    const idx = getHeaderIndexes_(sh, H_ORDENES);

    if (idx.pedido < 0 || idx.cantidad < 0) {
      return { ok: false, error: 'No se encontraron las columnas de "N° Pedido" o "Cantidad" en "' + SH_ORDENES + '".' };
    }

    const data = sh.getRange(2, 1, Math.max(0, sh.getLastRow() - 1), sh.getLastColumn()).getValues();

    const pedidosUnicos = new Set();
    let totalPaquetes = 0;

    for (const row of data) {
      const pedido = (row[idx.pedido] || '').toString().trim();
      if (pedido) {
        pedidosUnicos.add(pedido);
      }

      const cantidad = parseFloat(row[idx.cantidad]);
      if (!isNaN(cantidad)) {
        totalPaquetes += cantidad;
      }
    }

    return {
      ok: true,
      data: {
        pedidosHoy: pedidosUnicos.size,
        paquetesHoy: totalPaquetes
      }
    };
  } catch (e) {
    Logger.log(`Error en getDashboardSummary: ${e.stack}`);
    return { ok: false, error: e.toString() };
  }
}

function getDistribucionComunas(){
  const sh = getSheet_(SH_ORDENES);
  const idx = getHeaderIndexes_(sh, H_ORDENES);
  if (idx.comuna < 0 || idx.pedido < 0) {
    return { ok:false, error:'No se ubicaron columnas de Comuna y/o Pedido en "' + SH_ORDENES + '".' };
  }
  const data = sh.getRange(2,1,Math.max(0, sh.getLastRow()-1), sh.getLastColumn()).getValues();
  const map = new Map();
  for (const row of data){
    const comuna = (row[idx.comuna]||'SIN COMUNA').toString().trim();
    if (!comuna) continue;
    map.set(comuna, (map.get(comuna)||0)+1);
  }
  const arr = Array.from(map, ([comuna, cantidadPedidos]) => ({comuna, cantidadPedidos}));
  arr.sort((a,b)=> b.cantidadPedidos - a.cantidadPedidos);
  return { ok:true, items: arr };
}

/*************************** BUSCADOR ***************************/
/**
 * Busca por nombre de producto (parcial, case-insensitive) y/o Nº de Pedido.
 * Responde filas agregadas por producto con: { producto, cantidadVendida, pedidos:[...], inventarioActual, adquisicionesHoy:{cantidad, formato}, inventarioFinal }
 */
function buscarProductosYPedidos(filtro){
  filtro = filtro || {}; // {productoText, numeroPedido}
  const productoText = (filtro.productoText||'').toString().trim().toLowerCase();
  const numeroPedido = (filtro.numeroPedido||'').toString().trim();

  // 1) Ordenes (para cantidades y pedidos asociados)
  const shOrd = getSheet_(SH_ORDENES);
  const idxOrd = getHeaderIndexes_(shOrd, H_ORDENES);
  if (idxOrd.producto < 0 || idxOrd.cantidad < 0) {
    return { ok:false, error:'Faltan columnas en "' + SH_ORDENES + '" (Producto/Cantidad).' };
  }
  const dataOrd = shOrd.getRange(2,1,Math.max(0, shOrd.getLastRow()-1), shOrd.getLastColumn()).getValues();

  // 2) Lista de Adquisiciones (para inventarios y adquisiciones de hoy)
  const shAdq = getSheet_(SH_LISTA_ADQ);
  const idxAdq = getHeaderIndexes_(shAdq, H_ADQ);
  const dataAdq = shAdq.getRange(2,1,Math.max(0, shAdq.getLastRow()-1), shAdq.getLastColumn()).getValues();

  // Mapa rápido productoBase -> info de adquisiciones/inventarios
  const infoAdq = new Map();
  for (const r of dataAdq){
    const p = (idxAdq.productoBase>=0 ? r[idxAdq.productoBase] : '').toString().trim();
    if (!p) continue;
    infoAdq.set(p.toLowerCase(), {
      inventarioActual: r[idxAdq.invActual] ?? '',
      unidadInvActual:  r[idxAdq.unidadInvActual] ?? '',
      adquisicionesHoy: {
        cantidad: r[idxAdq.cantComprar] ?? '',
        formato:  r[idxAdq.formatoCompra] ?? ''
      },
      inventarioFinal:  r[idxAdq.invFinalizar] ?? '',
      unidadInvFinal:   r[idxAdq.unidadInvFinal] ?? ''
    });
  }

  // Agregación por producto
  const agg = new Map(); // key: producto (tal cual aparece en Ordenes)
  for (const r of dataOrd){
    const prod = (r[idxOrd.producto]||'').toString().trim();
    if (!prod) continue;
    if (productoText && !prod.toLowerCase().includes(productoText)) continue;
    if (numeroPedido){
      const pedidoVal = idxOrd.pedido>=0 ? (r[idxOrd.pedido]||'').toString().trim() : '';
      if (pedidoVal !== numeroPedido) continue;
    }
    const qty = parseFloat(r[idxOrd.cantidad]) || 0;
    const pedido = idxOrd.pedido>=0 ? (r[idxOrd.pedido]||'').toString().trim() : '';

    if (!agg.has(prod)) agg.set(prod, { cantidadVendida:0, pedidos:new Set() });
    const obj = agg.get(prod);
    obj.cantidadVendida += qty;
    if (pedido) obj.pedidos.add(pedido);
  }

  // Empaquetar respuesta + merge con adquisiciones/inventarios (por producto base, intentando normalizar)
  const items = [];
  for (const [producto, val] of agg.entries()){
    // Heurística simple para mapear a Producto Base: usar primera palabra o el nombre completo; probar variantes
    const keyCandidates = [producto, producto.split(' ')[0]]
      .map(s => s.toLowerCase());
    let info = null;
    for (const k of keyCandidates){
      if (infoAdq.has(k)) { info = infoAdq.get(k); break; }
    }

    items.push({
      producto,
      cantidadVendida: val.cantidadVendida,
      pedidos: Array.from(val.pedidos),
      inventarioActual: info ? info.inventarioActual + (info.unidadInvActual?(' ' + info.unidadInvActual):'') : '',
      adquisicionesHoy: info ? info.adquisicionesHoy : {cantidad:'', formato:''},
      inventarioFinal:  info ? info.inventarioFinal + (info.unidadInvFinal?(' ' + info.unidadInvFinal):'') : ''
    });
  }

  // Ordenar por cantidad vendida desc
  items.sort((a,b)=> (b.cantidadVendida||0) - (a.cantidadVendida||0));
  return { ok:true, items };
}

/*************************** BOOTSTRAP ***************************/
function doGet(){
  return HtmlService.createTemplateFromFile('DashboardDialog').evaluate()
    .setTitle('Dashboard Operaciones SNF')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename){
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// --- FLUJO DE ENVASADO ---

function startPackagingProcess() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ordersSheet = ss.getSheetByName('Orders');
  const skuSheet = ss.getSheetByName('SKU');
  if (!ordersSheet || !skuSheet) {
    SpreadsheetApp.getUi().alert('Error: Faltan las hojas "Orders" o "SKU".');
    return;
  }
  const newProducts = getNewProducts(ordersSheet, skuSheet);
  if (newProducts.length > 0) {
    showBatchUpdateDialog(newProducts);
  } else {
    showCategorySelectionDialog();
  }
}

function showBatchUpdateDialog(productList) {
  const template = HtmlService.createTemplateFromFile('Dialog');
  template.productList = JSON.stringify(productList);
  template.baseProducts = JSON.stringify(getExistingBaseProducts()); // Pass the suggestions
  const html = template.evaluate().setWidth(1200).setHeight(700); // Increased dialog size
  SpreadsheetApp.getUi().showModalDialog(html, 'Paso 1: Añadir Nuevos Productos a SKU');
}

function saveSkuData(data) {
  const skuSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('SKU');
  skuSheet.appendRow([data.nombreProducto, data.productoBase, data.formato, data.cantidad, data.unidad, data.categoria, data.cantidadVenta, data.unidadVenta, '']);
  return { status: 'success' };
}

function triggerCategoryDialog() {
  showCategorySelectionDialog();
}

/** PASO 2 · Panel interno para Envasado (Modal) **/
function showCategorySelectionDialog() {
  // Abre un diálogo modal central en lugar de un panel lateral.
  const html = HtmlService.createHtmlOutputFromFile('CategoryPanel')
    .setWidth(900)
    .setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, 'Paso 2: Categorías para Envasado');
}

function getPackagingData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ordersSheet = ss.getSheetByName('Orders');
  const skuSheet = ss.getSheetByName('SKU');
  const skuMap = getSkuMap(skuSheet);
  const orderData = ordersSheet.getRange("J2:K" + ordersSheet.getLastRow()).getValues();
  const productTotals = {};
  orderData.forEach(([name, qty]) => {
    if (name && qty) {
      if (!productTotals[name]) { productTotals[name] = 0; }
      productTotals[name] += parseInt(qty, 10) || 0;
    }
  });

  const stockMap = getStockFromOrders();
  const categorySummary = {};

  for (const productName in productTotals) {
    const category = skuMap[productName] ? skuMap[productName].category : 'Sin Categoría';
    if (!categorySummary[category]) {
      categorySummary[category] = { count: 0, products: {} };
    }
    categorySummary[category].count++;

    const inventoryInfo = stockMap[productName];
    let inventoryValue = 'No encontrado';
    if (inventoryInfo) {
      const stock = String(inventoryInfo.quantity);
      const unit = String(inventoryInfo.unit);
      if (stock || unit) {
        inventoryValue = `${stock} ${unit}`.trim();
      }
    }

    categorySummary[category].products[productName] = {
      total: productTotals[productName],
      stock: inventoryValue
    };
  }
  return categorySummary;
}

/**
 * Crea un mapa de stock desde la hoja "Orders".
 * @returns {Object<string, {quantity: any, unit: any}>} Un mapa donde las claves son
 *   nombres de productos y los valores son objetos con stock y unidad.
 */
function getStockFromOrders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ordersSheet = ss.getSheetByName('Orders');
  if (!ordersSheet) {
    Logger.log("Error: La hoja 'Orders' no fue encontrada.");
    return {};
  }
  const stockMap = {};
  const lastRow = ordersSheet.getLastRow();
  if (lastRow < 2) {
    return stockMap;
  }

  // Columnas: J (Nombre Producto, index 10), AA (Stock Real, index 27), AB (Unidad Venta, index 28)
  const data = ordersSheet.getRange(2, 10, lastRow - 1, 19).getValues(); // Lee desde la columna J hasta la AB

  data.forEach(row => {
    const productName = row[0];  // Columna J (índice 0 en el rango J:AB)
    const stockReal = row[17]; // Columna AA (índice 17 en el rango J:AB)
    const unit = row[18];      // Columna AB (índice 18 en el rango J:AB)

    // Solo agregar si el producto existe y no ha sido mapeado aún.
    // Esto asume que el stock es el mismo para todas las entradas del mismo producto.
    if (productName && stockMap[productName] === undefined) {
      stockMap[productName] = {
        quantity: stockReal,
        unit: unit
      };
    }
  });

  return stockMap;
}

function generatePackagingSheet(selectedCategories) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const data = getPackagingData();

  // Obtener el mapa de stock directamente desde la hoja "Orders".
  const stockMap = getStockFromOrders();

  // Crear una hoja con un nombre único basado en la fecha
  const date = new Date();
  const formattedDate = Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd");
  const sheetName = `Lista de Envasado - ${formattedDate}`;

  let sheet = ss.getSheetByName(sheetName);
  if (sheet) {
    sheet.clear(); // Limpiar la hoja si ya existe para este día
  } else {
    sheet = ss.insertSheet(sheetName); // Crear una nueva si no existe
  }

  sheet.activate(); // Activar la hoja para que el usuario la vea

  let currentRow = 1;

  // Título principal
  sheet.getRange(currentRow, 1, 1, 3).merge().setValue("Lista de Envasado").setFontWeight("bold").setFontSize(14).setHorizontalAlignment("center");
  currentRow += 2; // Espacio después del título

  // Encabezados de la tabla
  const headers = ["Cantidad", "Inventario Actual", "Nombre Producto"];
  const headerRange = sheet.getRange(currentRow, 1, 1, 3);
  headerRange.setValues([headers]).setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle");
  sheet.setFrozenRows(currentRow);
  currentRow++;

  // Llenar datos por categoría
  selectedCategories.sort().forEach(category => {
    sheet.getRange(currentRow, 1, 1, 3).merge().setValue(category.toUpperCase()).setFontWeight("bold").setHorizontalAlignment("center").setBackground("#f2f2f2");
    currentRow++;

    const products = data[category].products;
    const sortedProductNames = Object.keys(products).sort();

    const productRows = [];
    sortedProductNames.forEach(productName => {
      const inventoryInfo = stockMap[productName];

      // Formatear el valor del inventario para incluir la unidad
      let inventoryValue = 'No encontrado';
      if (inventoryInfo) {
        const stock = String(inventoryInfo.quantity);
        const unit = String(inventoryInfo.unit);
        // Mostrar valor solo si hay cantidad o unidad.
        if (stock || unit) {
          inventoryValue = `${stock} ${unit}`.trim();
        }
      }

      productRows.push([products[productName].total, inventoryValue, productName]);
    });

    if (productRows.length > 0) {
      const dataRange = sheet.getRange(currentRow, 1, productRows.length, 3);
      dataRange.setValues(productRows);
      dataRange.setHorizontalAlignment("center").setVerticalAlignment("middle");
      currentRow += productRows.length;
    }
    currentRow++; // Añadir una fila en blanco entre categorías para mayor claridad
  });

  // Ajustar anchos de columna
  sheet.setColumnWidth(1, 100); // Ancho para "Cantidad"
  sheet.setColumnWidth(2, 150); // Ancho para "Inventario Actual"
  sheet.setColumnWidth(3, 350); // Ancho para "Nombre Producto"

  // Construir y devolver la URL del PDF para impresión inmediata
  const printUrl = `https://docs.google.com/spreadsheets/d/${ss.getId()}/export?format=pdf&gid=${sheet.getSheetId()}&portrait=true&fitw=true&gridlines=true&printtitle=false`;
  return printUrl;
}

/**
 * Elimina todas las hojas de "Lista de Envasado" antiguas.
 */
function deletePackagingSheets() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const allSheets = ss.getSheets();
    let deletedCount = 0;

    allSheets.forEach(sheet => {
      const sheetName = sheet.getName();
      if (sheetName.startsWith('Lista de Envasado -')) {
        ss.deleteSheet(sheet);
        deletedCount++;
      }
    });

    if (deletedCount > 0) {
      return { status: 'success', message: `Limpieza completada. Se eliminaron ${deletedCount} hoja(s) de envasado.` };
    } else {
      return { status: 'info', message: 'No se encontraron hojas de envasado antiguas para eliminar.' };
    }
  } catch (e) {
    Logger.log(e);
    return { status: 'error', message: `Ocurrió un error: ${e.toString()}` };
  }
}

// --- FLUJO DE ADQUISICIONES ---

/**
 * Genera y guarda automáticamente la lista de adquisiciones.
 * Calcula las necesidades basadas en los pedidos y SKU, y luego guarda el plan.
 */
function updateAcquisitionListAutomated() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ordersSheet = ss.getSheetByName('Orders');
    const skuSheet = ss.getSheetByName('SKU');

    if (!ordersSheet || !skuSheet) {
      Logger.log('Omitiendo actualización automática de adquisiciones: Faltan las hojas "Orders" o "SKU".');
      return;
    }

    // 0. Get current inventory
    const inventoryMap = getCurrentInventory();
    
    // 1. Generar el plan de adquisiciones
    const { productToSkuMap, baseProductPurchaseOptions } = getPurchaseDataMaps(skuSheet);
    const baseProductNeeds = calculateBaseProductNeeds(ordersSheet, productToSkuMap);
    const acquisitionPlan = createAcquisitionPlan(baseProductNeeds, baseProductPurchaseOptions, inventoryMap);

    // 2. Transformar el plan al formato que espera `saveAcquisitions`
    const finalPlan = Object.values(acquisitionPlan).map(p => {
      const suggestedFormatString = `${p.suggestedFormat.name} (${p.suggestedFormat.size} ${p.suggestedFormat.unit})`;
      const allFormatStrings = p.availableFormats.map(f => `${f.name} (${f.size} ${f.unit})`);

      return {
        productName: p.productName,
        quantity: p.suggestedQty,
        selectedFormatString: suggestedFormatString,
        supplier: p.supplier,
        totalNeed: p.totalNeed,
        unit: p.unit,
        allFormatStrings: allFormatStrings,
        allFormatObjects: p.availableFormats.map(f => ({...f}))
      };
    });

    // 3. Guardar el plan utilizando la función existente
    // Esta función ya se encarga de limpiar la hoja, escribir encabezados y obtener el inventario actual.
    saveAcquisitions(finalPlan);
    Logger.log("La lista de adquisiciones se ha actualizado automáticamente.");

  } catch (e) {
    Logger.log(`Error durante la actualización automática de adquisiciones: ${e.toString()}`);
    // No mostramos una alerta al usuario para no ser intrusivos, pero lo registramos.
  }
}

function getAcquisitionDataForEditor(mode = 'wholesale') {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ordersSheet = ss.getSheetByName('Orders');
  const skuSheet = ss.getSheetByName('SKU');
  const proveedoresSheet = ss.getSheetByName('Proveedores');

  if (!ordersSheet || !skuSheet || !proveedoresSheet) {
    throw new Error('Faltan una o más hojas requeridas: "Orders", "SKU", o "Proveedores".');
  }

  // --- NEW: Cargar categorías desde SKU ---
  const baseToCategory = new Map();
  const allCategoriesSet = new Set();
  if (skuSheet.getLastRow() > 1) {
    const skuCategoryData = skuSheet.getRange("B2:F" + skuSheet.getLastRow()).getValues();
    skuCategoryData.forEach(row => {
      const baseProduct = row[0]; // Col B
      const category = row[4];    // Col F
      if (baseProduct && category) {
        baseToCategory.set(normalizeKey(baseProduct), category);
        allCategoriesSet.add(category);
      }
    });
  }

  // --- START MODIFICATION: Check if "Lista de Adquisiciones" is populated ---
  const acquisitionsSheet = ss.getSheetByName("Lista de Adquisiciones");
  const populatedAcquisitionsMap = new Map();
  let isAcquisitionsSheetPopulated = false;

  if (acquisitionsSheet && acquisitionsSheet.getLastRow() > 1) {
      isAcquisitionsSheetPopulated = true;
      const acquisitionsData = acquisitionsSheet.getRange(2, 1, acquisitionsSheet.getLastRow() - 1, acquisitionsSheet.getLastColumn()).getValues();
      const headers = acquisitionsSheet.getRange(1, 1, 1, acquisitionsSheet.getLastColumn()).getValues()[0];
      const productCol = headers.indexOf("Producto Base");
      const qtyCol = headers.indexOf("Cantidad a Comprar");

      if (productCol !== -1 && qtyCol !== -1) {
          acquisitionsData.forEach(row => {
              const productName = row[productCol];
              const qty = parseFloat(String(row[qtyCol]).replace(",", ".")) || 0;
              if (productName) {
                  populatedAcquisitionsMap.set(normalizeKey(productName), qty);
              }
          });
      }
  }
  // --- END MODIFICATION ---

  const inventoryMap = getCurrentInventory();
  const { productToSkuMap, baseProductPurchaseOptions } = getPurchaseDataMaps(skuSheet);
  const baseProductNeeds = calculateBaseProductNeeds(ordersSheet, productToSkuMap);
  const acquisitionPlan = createAcquisitionPlan(baseProductNeeds, baseProductPurchaseOptions, inventoryMap, mode);

  // --- START MODIFICATION: Override suggested quantities if "Lista de Adquisiciones" is populated ---
  if (isAcquisitionsSheetPopulated) {
      acquisitionPlan.forEach(item => {
          const key = normalizeKey(item.productName);
          if (populatedAcquisitionsMap.has(key)) {
              item.suggestedQty = populatedAcquisitionsMap.get(key);
          } else {
              item.suggestedQty = 0; // If not in the populated list, default to 0
          }
      });
  }
  // --- END MODIFICATION ---

  // Enriquecer plan con categorías
  let hasUncategorized = false;
  acquisitionPlan.forEach(item => {
    const key = normalizeKey(item.productName);
    if (baseToCategory.has(key)) {
      item.category = baseToCategory.get(key);
    } else {
      item.category = 'Sin Categoría';
      hasUncategorized = true;
    }
  });
  if (hasUncategorized) {
      allCategoriesSet.add('Sin Categoría');
  }

  const supplierData = proveedoresSheet.getRange("A2:A" + proveedoresSheet.getLastRow()).getValues().flat().filter(String);
  const supplierSet = new Set(supplierData);
  supplierSet.add("Patio Mayorista");

  return {
    acquisitionPlan: acquisitionPlan,
    allSuppliers: Array.from(supplierSet).sort(),
    allCategories: Array.from(allCategoriesSet).sort()
  };
}

function showAcquisitionEditor() {
  const dataForEditor = getAcquisitionDataForEditor();
  const template = HtmlService.createTemplateFromFile('AcquisitionEditorDialog');
  // Pasar el objeto de datos directamente al template. La serialización se hará en el lado del cliente.
  template.data = dataForEditor;
  const html = template.evaluate().setWidth(1100).setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, 'Editar Borrador de Adquisiciones');
}

function saveAcquisitions(finalPlan) {
  // finalPlan es un array de objetos desde el cliente.
  // Cada objeto: { productName, quantity, selectedFormatString, supplier, totalNeed, unit, allFormatStrings, allFormatObjects }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Lista de Adquisiciones");
  if (sheet) {
    sheet.clear();
    sheet.clearConditionalFormatRules();
  } else {
    sheet = ss.insertSheet("Lista de Adquisiciones");
  }

  // Escribir datos en un formato plano para mayor robustez, con una columna de proveedor.
  const headers = ["Producto Base", "Cantidad a Comprar", "Formato de Compra", "Inventario Actual", "Unidad Inventario Actual", "Necesidad de Venta", "Unidad Venta", "Inventario al Finalizar", "Unidad Inventario Final", "Precio Adq. Anterior", "Precio Adq. HOY", "Proveedor"];
  sheet.getRange("A1:L1").setValues([headers]).setFontWeight("bold");
  sheet.getRange("A1:C1").setBackground("#d9ead3");
  sheet.getRange("D1:E1").setBackground("#fff2cc");
  sheet.getRange("F1:K1").setBackground("#f4cccc");
  sheet.getRange("L1").setBackground("#d9d9d9");
  sheet.setFrozenRows(1);

  const inventoryMap = getCurrentInventory(); // Get current inventory
  const priceMap = getHistoricalPrices(); // Get historical prices
  const dataToWrite = [];

  finalPlan.forEach(p => {
    const selectedFormatObject = p.allFormatObjects.find(f => `${f.name} (${f.size} ${f.unit})` === p.selectedFormatString);
    const formatSize = selectedFormatObject ? selectedFormatObject.size : 0;

    const currentInventory = inventoryMap[p.productName] || { quantity: 0, unit: p.unit };

    const purchasedAmount = (parseFloat(p.quantity) || 0) * formatSize;
    const finalInventory = currentInventory.quantity + purchasedAmount - (parseFloat(p.totalNeed) || 0);

    const history = priceMap[p.productName] || [];
    const precioHoy = history.length > 0 ? history[0].price : "";
    const precioAnterior = history.length > 1 ? history[1].price : "";

    const rowData = [
      p.productName,
      p.quantity,
      p.selectedFormatString,
      currentInventory.quantity, // Use actual inventory
      currentInventory.unit,     // Use actual inventory unit
      p.totalNeed,
      p.unit,
      finalInventory,
      p.unit,
      precioAnterior, // Columna J
      precioHoy,      // Columna K
      p.supplier || "Sin Proveedor"
    ];
    dataToWrite.push(rowData);
  });

  if (dataToWrite.length > 0) {
    sheet.getRange(2, 1, dataToWrite.length, headers.length).setValues(dataToWrite);

    // Aplicar la validación de datos a toda la columna de formato de una vez
    const formatColumnRange = sheet.getRange("C2:C" + (dataToWrite.length + 1));
    // Nota: Esta validación será la misma para todas las celdas (la del último producto).
    // Una validación por celda es necesaria si los formatos varían mucho.
    finalPlan.forEach((p, index) => {
      const rule = SpreadsheetApp.newDataValidation().requireValueInList(p.allFormatStrings).build();
      sheet.getRange(index + 2, 3).setDataValidation(rule);
    });
  }

  sheet.autoResizeColumns(1, headers.length);

  return { status: "success", message: "Lista de adquisiciones guardada con éxito." };
}

function clearAcquisitionList() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("Lista de Adquisiciones");
    if (sheet) {
      sheet.clear();
      sheet.clearConditionalFormatRules();
    } else {
      sheet = ss.insertSheet("Lista de Adquisiciones");
    }

    const headers = ["Producto Base", "Cantidad a Comprar", "Formato de Compra", "Inventario Actual", "Unidad Inventario Actual", "Necesidad de Venta", "Unidad Venta", "Inventario al Finalizar", "Unidad Inventario Final", "Precio Adq. Anterior", "Precio Adq. HOY", "Proveedor"];
    sheet.getRange("A1:L1").setValues([headers]).setFontWeight("bold");
    sheet.getRange("A1:C1").setBackground("#d9ead3");
    sheet.getRange("D1:E1").setBackground("#fff2cc");
    sheet.getRange("F1:K1").setBackground("#f4cccc");
    sheet.getRange("L1").setBackground("#d9d9d9");
    sheet.setFrozenRows(1);

    SpreadsheetApp.flush(); // Ensure changes are saved
    return { status: 'success', message: 'La lista de adquisiciones ha sido limpiada con éxito.' };
  } catch (e) {
    Logger.log(`Error en clearAcquisitionList: ${e.stack}`);
    return { status: 'error', message: `Ocurrió un error al limpiar la lista: ${e.message}` };
  }
}

function generateAcquisitionDRAFT() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ordersSheet = ss.getSheetByName('Orders');
  const skuSheet = ss.getSheetByName('SKU');
  if (!ordersSheet || !skuSheet) {
    SpreadsheetApp.getUi().alert('Faltan las hojas "Orders" o "SKU".');
    return;
  }
  const { productToSkuMap, baseProductPurchaseOptions } = getPurchaseDataMaps(skuSheet);
  const baseProductNeeds = calculateBaseProductNeeds(ordersSheet, productToSkuMap);
  const acquisitionPlan = createAcquisitionPlan(baseProductNeeds, baseProductPurchaseOptions);
  let sheet = ss.getSheetByName("Lista de Adquisiciones");
  if (sheet) {
    sheet.clear();
    sheet.clearConditionalFormatRules();
  } else {
    sheet = ss.insertSheet("Lista de Adquisiciones");
  }
  const headers = ["Producto Base", "Cantidad a Comprar", "Formato de Compra", "Inventario Actual", "Unidad Inventario Actual", "Necesidad de Venta", "Unidad Venta", "Inventario al Finalizar", "Unidad Inventario Final", "Precio Adq. Anterior", "Precio Adq. HOY"];
  sheet.getRange("A1:K1").setValues([headers]).setFontWeight("bold");
  sheet.getRange("A1:C1").setBackground("#d9ead3");
  sheet.getRange("D1:E1").setBackground("#fff2cc");
  sheet.getRange("F1:K1").setBackground("#f4cccc");
  sheet.setFrozenRows(1);
  const dataBySupplier = groupPlanBySupplier(acquisitionPlan);
  let currentRow = 2;
  const sortedSuppliers = Object.keys(dataBySupplier).sort();
  sortedSuppliers.forEach(supplier => {
    sheet.getRange(currentRow, 1, 1, headers.length).merge().setValue(supplier).setFontWeight("bold").setHorizontalAlignment("center").setBackground("#d9d9d9");
    currentRow++;
    const products = dataBySupplier[supplier];
    products.forEach(p => {
      const suggestedFormatString = `${p.suggestedFormat.name} (${p.suggestedFormat.size} ${p.suggestedFormat.unit})`;
      const totalComprado = p.suggestedQty * p.suggestedFormat.size;
      const inventarioFinal = 0 + totalComprado - p.totalNeed;
      sheet.getRange(currentRow, 1, 1, headers.length).setValues([[p.productName, p.suggestedQty, suggestedFormatString, 0, p.unit, p.totalNeed, p.saleUnit, inventarioFinal, p.unit, "", ""]]);
      const formatOptions = p.availableFormats.map(f => `${f.name} (${f.size} ${f.unit})`);
      const rule = SpreadsheetApp.newDataValidation().requireValueInList(formatOptions).build();
      sheet.getRange(currentRow, 3).setDataValidation(rule);
      currentRow++;
    });
  });
  sheet.autoResizeColumns(1, headers.length);
  SpreadsheetApp.getUi().alert("Borrador de 'Lista de Adquisiciones' generado con éxito.");
}

/**
 * Clears all data from the "Lista de Adquisiciones" sheet, preserving the header row.
 */
function clearAcquisitionsList() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Lista de Adquisiciones");
    if (sheet) {
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
      }
      SpreadsheetApp.flush();
      return { status: "success", message: "La lista de adquisiciones ha sido limpiada." };
    } else {
      // This case should ideally not happen if setupProjectSheets() runs correctly.
      return { status: "error", message: "La hoja 'Lista de Adquisiciones' no fue encontrada." };
    }
  } catch (e) {
    Logger.log(`Error en clearAcquisitionsList: ${e.stack}`);
    return { status: "error", message: `Ocurrió un error al limpiar la lista: ${e.message}` };
  }
}

function recalculateRowInventory(sheet, row) {
  const dataRange = sheet.getRange(`A${row}:H${row}`);
  const values = dataRange.getValues()[0];
  const [productoBase, cantidadAComprar, formatoDeCompra, inventarioActual, unidadInvActual, necesidadDeVenta, unidadVenta] = values;
  const multiplierMatch = String(formatoDeCompra).match(/\((\d+(\.\d+)?)/);
  const multiplier = multiplierMatch ? parseFloat(multiplierMatch[1]) : 0;
  const totalComprado = (parseFloat(String(cantidadAComprar).replace(",", ".")) || 0) * multiplier;
  const inventarioFinal = (parseFloat(String(inventarioActual).replace(",", ".")) || 0) + totalComprado - (parseFloat(String(necesidadDeVenta).replace(",", ".")) || 0);
  sheet.getRange(row, 8).setValue(inventarioFinal);
}



/**
 * Lee la hoja "Historico Adquisiciones" y devuelve un mapa de precios históricos por producto.
 * @returns {Object<string, Array<{date: Date, price: number}>>} Un mapa donde las claves son
 *   nombres de productos y los valores son arrays de objetos de precio, ordenados por fecha descendente.
 */
function getHistoricalPrices() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const historicoSheet = ss.getSheetByName("Historico Adquisiciones");
  const priceMap = {};

  if (!historicoSheet) {
    Logger.log("Advertencia: No se encontró la hoja 'Historico Adquisiciones'. No se mostrarán precios.");
    return priceMap;
  }

  const lastRow = historicoSheet.getLastRow();
  if (lastRow < 2) {
    return priceMap; // Hoja vacía o solo con encabezados
  }

  // Columnas: B (Fecha), C (Producto Base), F (Precio Compra)
  const data = historicoSheet.getRange("B2:F" + lastRow).getValues();

  data.forEach(row => {
    const date = row[0];        // de la columna B
    const productName = row[1]; // de la columna C
    const price = row[4];       // de la columna F

    if (productName && date && price) {
      const normalizedProductName = normalizeKey(productName);
      if (!priceMap[normalizedProductName]) {
        priceMap[normalizedProductName] = [];
      }
      priceMap[normalizedProductName].push({
        date: new Date(date),
        price: parseFloat(String(price).replace(",", ".")) || 0
      });
    }
  });

  // Ordenar los precios de cada producto por fecha, de más reciente a más antiguo
  for (const product in priceMap) {
    priceMap[product].sort((a, b) => b.date - a.date);
  }

  return priceMap;
}

function getCurrentInventory() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inventorySheet = ss.getSheetByName("Inventario Actual");
  const inventoryMap = {};

  if (!inventorySheet) {
    Logger.log("Warning: La hoja 'Inventario Actual' no fue encontrada. El inventario actual será 0.");
    return inventoryMap;
  }

  const lastRow = inventorySheet.getLastRow();
  if (lastRow < 2) {
    return inventoryMap; // Sheet is empty or has only headers
  }

  // Read data from columns B (Producto Base), C (Cantidad Stock Real), D (Unidad Venta)
  const data = inventorySheet.getRange(2, 2, lastRow - 1, 3).getValues();

  data.forEach(row => {
    const productName = row[0]; // from column B
    const quantity = row[1];    // from column C
    const unit = row[2];        // from column D
    if (productName) {
      inventoryMap[normalizeKey(productName)] = {
        quantity: parseFloat(String(quantity).replace(",", ".")) || 0,
        unit: unit || ''
      };
    }
  });

  return inventoryMap;
}

/**
 * Lee la hoja "Historico Adquisiciones", la ordena por fecha y crea un mapa con el proveedor más reciente para cada producto base.
 * @returns {Object<string, string>} Un mapa donde las claves son nombres de "Producto Base" y los valores son el nombre del proveedor más reciente.
 */
function getLatestSuppliersFromHistory() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const historicoSheet = ss.getSheetByName("Historico Adquisiciones");
  const latestSuppliers = {};

  if (!historicoSheet || historicoSheet.getLastRow() < 2) {
    Logger.log("Advertencia: No se encontró la hoja 'Historico Adquisiciones' o está vacía. No se pudo obtener el historial de proveedores.");
    return latestSuppliers;
  }

  // 1. Get all data and headers
  const range = historicoSheet.getDataRange();
  const values = range.getValues();
  const headers = values.shift(); // Remove headers from data

  // 2. Find column indices dynamically
  const dateCol = headers.indexOf("Fecha de Guardado");
  const productCol = headers.indexOf("Producto Base");
  const supplierCol = headers.indexOf("Nombre de Proveedor");

  if (dateCol === -1 || productCol === -1 || supplierCol === -1) {
    Logger.log("Error: No se encontraron las columnas requeridas ('Fecha de Guardado', 'Producto Base', 'Nombre de Proveedor') en 'Historico Adquisiciones'.");
    return latestSuppliers;
  }

  // 3. Sort data by date descending (most recent first)
  values.sort((a, b) => {
    const dateA = new Date(a[dateCol]);
    const dateB = new Date(b[dateCol]);
    return dateB - dateA; // Sort descending
  });

  // 4. Iterate through the sorted data to find the latest supplier for each product
  for (const row of values) {
    const productoBase = row[productCol];
    const proveedor = row[supplierCol];

    // If we find a product and a supplier, and we haven't already saved one for this product, add it.
    // Since the data is sorted by date, the first one we find will be the most recent.
    if (productoBase && proveedor) {
      const normalizedProductoBase = normalizeKey(productoBase);
      if (!latestSuppliers[normalizedProductoBase]) {
        latestSuppliers[normalizedProductoBase] = String(proveedor).trim();
      }
    }
  }

  Logger.log("Proveedores más recientes obtenidos del historial (corregido): " + JSON.stringify(latestSuppliers));
  return latestSuppliers;
}



// --- FUNCIONES AUXILIARES ---


function parseDDMMYYYY(dateString) {
  if (!dateString || typeof dateString !== 'string') return null;
  const parts = dateString.split('/');
  if (parts.length !== 3) return null;
  // new Date(year, monthIndex, day)
  return new Date(parts[2], parts[1] - 1, parts[0]);
}

function extractNameFromDescription(description) {
  if (!description || typeof description !== 'string') return '';
  const match = description.match(/(?:transf de|de)\s(.+)/i);
  if (match && match[1]) {
    return match[1].replace(/[0-9]/g, '').trim();
  }
  let cleaned = description.replace(/transf/i, '')
                           .replace(/pago/i, '')
                           .replace(/[0-9]/g, '')
                           .trim();
  return cleaned;
}

function calculateNameSimilarity(nameFromPayment, nameFromOrder) {
  if (!nameFromPayment || !nameFromOrder) return 0;

  const normalize = (str) => str.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").split(/\s+/);

  const wordsFromPayment = normalize(nameFromPayment);
  const wordsFromOrder = normalize(nameFromOrder);

  if (wordsFromPayment.length === 0 || wordsFromOrder.length === 0) return 0;

  let matches = 0;
  for (const pWord of wordsFromPayment) {
    for (const oWord of wordsFromOrder) {
      if (oWord.startsWith(pWord) || pWord.startsWith(oWord)) {
        matches++;
        break;
      }
    }
  }

  return (matches / wordsFromPayment.length) * 100;
}

function normalizePhoneNumber(phone) {
  if (!phone) return '';
  const originalPhoneStr = String(phone);
  let phoneStr = originalPhoneStr.trim();

  // Clean up common prefixes like '=' or '+'
  if (phoneStr.startsWith('=') || phoneStr.startsWith('+')) {
    phoneStr = phoneStr.substring(1);
  }
  if (phoneStr.startsWith('+')) { // In case of '=+'
    phoneStr = phoneStr.substring(1);
  }

  // Handle the `...123` suffix
  if (phoneStr.endsWith('123')) {
    let coreNumber = phoneStr.slice(0, -3);
    if (coreNumber.length === 9 && coreNumber.startsWith('9')) {
      return `56${coreNumber}`;
    }
  }

  // Handle standard Chilean formats if the special suffix format didn't match
  if (phoneStr.startsWith('569') && phoneStr.length === 11) {
    return phoneStr;
  }
  if (phoneStr.length === 9 && phoneStr.startsWith('9')) {
    return `56${phoneStr}`;
  }
  if (phoneStr.length === 8) {
    return `569${phoneStr}`;
  }

  // Final Fallback: If no specific format was matched, strip all non-numeric characters.
  return originalPhoneStr.replace(/\D/g, '');
}

function getNewProducts(ordersSheet, skuSheet) {
  const ordersData = ordersSheet.getRange('J2:J' + ordersSheet.getLastRow()).getValues();
  const skuData = skuSheet.getRange('A2:A' + skuSheet.getLastRow()).getValues();
  const orderProducts = ordersData.map(row => row[0]).filter(String);
  const skuProducts = new Set(skuData.map(row => row[0]).filter(String));
  return [...new Set(orderProducts)].filter(product => !skuProducts.has(product));
}

/**
 * Gets a sorted, unique list of "Producto Base" names from the SKU sheet for autocomplete suggestions.
 * @returns {string[]} A sorted array of unique product base names.
 */
function getExistingBaseProducts() {
  try {
    const skuSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("SKU");
    // Return empty array if the sheet doesn't exist or is empty
    if (!skuSheet || skuSheet.getLastRow() < 2) {
      return [];
    }
    // Column B is "Producto Base". Read from row 2 to the last row.
    const range = skuSheet.getRange(2, 2, skuSheet.getLastRow() - 1, 1);
    const values = range.getValues().flat().filter(String); // Get all values, flatten array, and remove empty strings
    const uniqueValues = [...new Set(values)]; // Get unique values using a Set
    return uniqueValues.sort(); // Return sorted unique values
  } catch (e) {
    Logger.log(`Error in getExistingBaseProducts: ${e.message}`);
    return []; // Return empty array on error
  }
}

function getSkuMap(skuSheet) {
  const skuData = skuSheet.getRange("A2:I" + skuSheet.getLastRow()).getValues();
  const skuMap = {};
  skuData.forEach(row => {
    let [name, base, format, qty, unit, category, saleQty, saleUnit, supplier] = row;
    if (name) {
      category = normalizeString(category);
      unit = normalizeUnit(unit);
      saleUnit = normalizeUnit(saleUnit);
      skuMap[name] = { base, format, qty, unit, category, saleQty, saleUnit, supplier };
    }
  });
  return skuMap;
}

function getPurchaseDataMaps(skuSheet) {
  const skuData = skuSheet.getRange("A2:I" + skuSheet.getLastRow()).getValues();
  const productToSkuMap = {};
  const baseProductPurchaseOptions = {};
  const baseProductSaleUnits = {}; // To store the sale unit for each base product

  skuData.forEach(row => {
    const [nombreProducto, productoBase, formatoCompra, cantidadCompra, unidadCompra, cat, cantVenta, unidadVenta, proveedor] = row;

    if (!productoBase) return; // Skip rows without a base product

    const normalizedProductoBase = normalizeKey(productoBase);

    if (nombreProducto) {
      productToSkuMap[nombreProducto] = {
        productoBase: normalizedProductoBase, // Use the normalized key
        cantidadVenta: parseFloat(String(cantVenta).replace(',', '.')) || 0,
        unidadVenta: normalizeUnit(unidadVenta)
      };
    }

    if (unidadVenta) {
      baseProductSaleUnits[normalizedProductoBase] = normalizeUnit(unidadVenta);
    }

    if (formatoCompra) {
      if (!baseProductPurchaseOptions[normalizedProductoBase]) {
        baseProductPurchaseOptions[normalizedProductoBase] = { options: [], suppliers: new Set() };
      }

      const options = baseProductPurchaseOptions[normalizedProductoBase].options;
      const parsedCantidad = parseFloat(String(cantidadCompra).replace(',', '.')) || 0;
      const normalizedUnidadCompra = normalizeUnit(unidadCompra);
      const normalizedFormatoCompra = normalizeString(formatoCompra);

      const exists = options.some(o =>
        o.name === normalizedFormatoCompra &&
        o.size === parsedCantidad &&
        o.unit === normalizedUnidadCompra);

      if (!exists) {
        options.push({ name: normalizedFormatoCompra, size: parsedCantidad, unit: normalizedUnidadCompra });
      }

      if (proveedor) baseProductPurchaseOptions[normalizedProductoBase].suppliers.add(proveedor);
    }
  });

  // Add minimum purchase formats (1kg / 1 unit) if they don't exist
  for (const productoBase in baseProductPurchaseOptions) {
      const saleUnit = baseProductSaleUnits[productoBase];
      if (saleUnit === 'Kg' || saleUnit === 'Unidad') {
          const options = baseProductPurchaseOptions[productoBase].options;
          const hasMinOption = options.some(o => o.size === 1 && o.unit === saleUnit);
          if (!hasMinOption) {
              options.push({
                  name: saleUnit === 'Kg' ? 'Bolsa' : 'Unidad', // Generic name for the minimum unit
                  size: 1,
                  unit: saleUnit
              });
          }
      }
  }

  return { productToSkuMap, baseProductPurchaseOptions };
}

function calculateBaseProductNeeds(ordersSheet, productToSkuMap) {
  const orderData = ordersSheet.getRange("J2:K" + ordersSheet.getLastRow()).getValues();
  const baseProductNeeds = {};
  orderData.forEach(([name, qty]) => {
    if (name && qty && productToSkuMap[name]) {
      const skuInfo = productToSkuMap[name];
      const baseProduct = skuInfo.productoBase;
      const saleUnit = normalizeUnit(skuInfo.unidadVenta);
      const totalSaleAmount = (parseInt(qty, 10) || 0) * skuInfo.cantidadVenta;
      if (!baseProductNeeds[baseProduct]) baseProductNeeds[baseProduct] = {};
      if (!baseProductNeeds[baseProduct][saleUnit]) baseProductNeeds[baseProduct][saleUnit] = 0;
      baseProductNeeds[baseProduct][saleUnit] += totalSaleAmount;
    }
  });
  return baseProductNeeds;
}

function createAcquisitionPlan(baseProductNeeds, baseProductPurchaseOptions, inventoryMap, mode = 'wholesale') {
  const acquisitionPlan = [];
  const latestSuppliers = getLatestSuppliersFromHistory();

  // If mode is 'all', we start with all products. Otherwise, only products with needs.
  const productSource = (mode === 'all')
    ? Object.keys(baseProductPurchaseOptions)
    : Object.keys(baseProductNeeds);

  const sortedBaseProducts = productSource.sort((a, b) => a.localeCompare(b));

  for (const baseProduct of sortedBaseProducts) {
    if (baseProductPurchaseOptions[baseProduct]) {
      const needs = baseProductNeeds[baseProduct] || {};
      const purchaseInfo = baseProductPurchaseOptions[baseProduct];
      const purchaseOptions = purchaseInfo.options.sort((a, b) => b.size - a.size);

      // Determine the unit of measurement. Fallback logic is important.
      const needUnit = Object.keys(needs)[0] || purchaseOptions[0]?.unit || 'Unidad';
      const totalNeed = needs[needUnit] || 0;

      const inventoryInfo = (inventoryMap && inventoryMap[baseProduct])
        ? inventoryMap[baseProduct]
        : { quantity: 0, unit: needUnit };

      let netNeed = Math.max(0, totalNeed - inventoryInfo.quantity);

      // The core logic change: ONLY skip if mode is NOT 'all' AND there's no need.
      if (mode !== 'all' && netNeed <= 0) {
        continue;
      }

      const supplier = getBestSupplier(purchaseInfo, latestSuppliers[baseProduct]);

      // If there is no net need, we add the item with 0 quantity and continue.
      if (netNeed <= 0) {
        acquisitionPlan.push(createAcquisitionItem(
          baseProduct,
          totalNeed,
          needUnit,
          supplier,
          purchaseOptions,
          purchaseOptions[0] || { name: 'Sin formato', size: 0, unit: needUnit }, // Default suggested format
          0, // Suggested quantity is 0
          inventoryInfo
        ));
        continue; // Go to the next product
      }

      // --- Existing logic for netNeed > 0 ---
      if (mode === 'just-in-time') {
        let remainingNeed = netNeed;

        // Iterar de la opción más grande a la más pequeña
        purchaseOptions.forEach(option => {
          if (option.unit === needUnit && option.size > 0 && remainingNeed > 0) {
            const numToBuy = Math.floor(remainingNeed / option.size);
            if (numToBuy > 0) {
              acquisitionPlan.push(createAcquisitionItem(baseProduct, totalNeed, needUnit, supplier, purchaseOptions, option, numToBuy, inventoryInfo));
              remainingNeed -= numToBuy * option.size;
            }
          }
        });

        // Si queda un remanente, comprar una unidad del formato más pequeño disponible
        if (remainingNeed > 0) {
          const smallestOption = purchaseOptions.slice().reverse().find(o => o.unit === needUnit && o.size > 0);
          if (smallestOption) {
            acquisitionPlan.push(createAcquisitionItem(baseProduct, totalNeed, needUnit, supplier, purchaseOptions, smallestOption, 1, inventoryInfo));
          }
        }
      } else { // modo 'wholesale' (Compra Normal/Minimo Mayorista)
        // Buscar el formato más pequeño que sea IGUAL O MAYOR a la necesidad neta.
        // Se busca en reverso (del más pequeño al más grande) porque las opciones vienen ordenadas de mayor a menor.
        const idealOption = purchaseOptions.slice().reverse().find(o => o.size >= netNeed && o.unit === needUnit);

        if (idealOption) {
          // Si se encuentra un formato ideal, comprar solo 1 de ese.
          acquisitionPlan.push(createAcquisitionItem(baseProduct, totalNeed, needUnit, supplier, purchaseOptions, idealOption, 1, inventoryInfo));
        } else {
          // Si NINGÚN formato individualmente cubre la necesidad (ej: necesito 15kg, pero el formato más grande es de 12kg),
          // usar el formato más grande disponible y calcular cuántos se necesitan.
          const biggestOption = purchaseOptions[0]; // La opción más grande
          if (biggestOption && biggestOption.unit === needUnit) {
            const numToBuy = netNeed > 0 ? Math.ceil(netNeed / biggestOption.size) : 0;
            if (numToBuy > 0) {
              acquisitionPlan.push(createAcquisitionItem(baseProduct, totalNeed, needUnit, supplier, purchaseOptions, biggestOption, numToBuy, inventoryInfo));
            }
          }
        }
      }
    }
  }
  return acquisitionPlan;
}

/** Helper para crear un item del plan de adquisición y evitar repetición de código. */
function createAcquisitionItem(productName, totalNeed, unit, supplier, availableFormats, suggestedFormat, suggestedQty, inventoryInfo) {
  return {
    productName,
    totalNeed,
    unit,
    saleUnit: unit, // Asumimos que la unidad de venta es la misma que la de necesidad
    supplier,
    availableFormats,
    suggestedFormat,
    suggestedQty,
    currentInventory: inventoryInfo.quantity,
    currentInventoryUnit: inventoryInfo.unit
  };
}

/** Helper para determinar el mejor proveedor. */
function getBestSupplier(purchaseInfo, historicalSupplier) {
    const skuSuppliers = Array.from(purchaseInfo.suppliers);
    return historicalSupplier || (skuSuppliers.length > 0 ? skuSuppliers[0] : "Patio Mayorista");
}

function groupPlanBySupplier(acquisitionPlan) {
  const dataBySupplier = {};
  for (const productName in acquisitionPlan) {
    const productData = acquisitionPlan[productName];
    const supplier = productData.supplier || "Sin Proveedor";
    if (!dataBySupplier[supplier]) dataBySupplier[supplier] = [];
    dataBySupplier[supplier].push(productData);
  }
  return dataBySupplier;
}

/**
 * Normaliza una cadena para ser usada como clave (lowercase, trim).
 * @param {string} str La cadena a normalizar.
 * @returns {string} La cadena normalizada.
 */
function normalizeKey(str) {
  if (!str || typeof str !== 'string') return '';
  return str.trim().toLowerCase();
}

function normalizeString(str) {
  if (!str || typeof str !== 'string') return '';
  return str.trim().toLowerCase().replace(/\w\S*/g, (w) => (w.replace(/^\w/, (c) => c.toUpperCase())));
}

function normalizeUnit(str) {
  if (!str || typeof str !== 'string') return '';
  const s = str.trim().toLowerCase();
  if (s.startsWith('kilo')) { return 'Kg';}
  if (s.startsWith('gr')) { return 'Gr';}
  if (s.startsWith('unidad')) { return 'Unidad';}
  if (s.startsWith('bandeja')) { return 'Bandeja';}
  return normalizeString(s);
}

function showPasteImportDialog() {
  const html = HtmlService.createHtmlOutputFromFile('PasteImportDialog')
    .setWidth(600)
    .setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, 'Importar Pedidos por Copiado y Pegado');
}

function getOAuthToken() {
  DriveApp.getFolderById('root'); // Force Drive scope.
  return ScriptApp.getOAuthToken();
}

function importOrdersFromPastedText(textData) {
  try {
    if (!textData || typeof textData !== 'string') {
        throw new Error("No se proporcionaron datos de texto para importar.");
    }

    // 1. Parsear el texto
    const rows = textData.trim().split('\n').map(row => row.split('\t'));
    if (rows.length < 2) {
      throw new Error("Los datos pegados deben incluir al menos una fila de encabezado y una fila de datos.");
    }

    const sourceHeaders = rows.shift().map(h => normalizeHeader(h));
    const sourceData = rows;

    // 2. Obtener encabezados de destino y crear mapa
    const mainSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const ordersSheet = mainSpreadsheet.getSheetByName("Orders");
    if (!ordersSheet) {
      throw new Error("No se encontró la hoja 'Orders' en el libro principal.");
    }
    const targetHeaders = ordersSheet.getRange(1, 1, 1, ordersSheet.getLastColumn()).getValues()[0];
    const normalizedTargetHeaders = targetHeaders.map(h => normalizeHeader(h));

    const columnIndexMap = normalizedTargetHeaders.map(targetHeader => sourceHeaders.indexOf(targetHeader));

    // 3. Reordenar los datos
    const reorderedData = sourceData.map(sourceRow => {
        const newRow = [];
        columnIndexMap.forEach((sourceIndex, targetIndex) => {
            newRow[targetIndex] = (sourceIndex !== -1) ? sourceRow[sourceIndex] : "";
        });
        return newRow;
    });

    // 4. Escribir los datos en la hoja
    ordersSheet.getRange(2, 1, ordersSheet.getMaxRows() - 1, ordersSheet.getMaxColumns()).clearContent();
    if (reorderedData.length > 0) {
        ordersSheet.getRange(2, 1, reorderedData.length, reorderedData[0].length).setValues(reorderedData);
    }

    // --- LÓGICA PARA AGREGAR PRODUCTO BASE ---
    const skuSheet = mainSpreadsheet.getSheetByName("SKU");
    if (!skuSheet) {
      Logger.log("Advertencia: No se encontró la hoja 'SKU'. No se pudo poblar la columna 'Producto Base'.");
    } else {
      const skuData = skuSheet.getRange("A2:B" + skuSheet.getLastRow()).getValues();
      // Normalize keys by trimming whitespace to make matching more robust.
      const skuMap = new Map(skuData.map(row => [(row[0] || '').toString().trim(), row[1]]));

      const headers = ordersSheet.getRange(1, 1, 1, ordersSheet.getLastColumn()).getValues()[0];
      let productNameColIndex = headers.indexOf("Item Name");
      if (productNameColIndex === -1) {
        productNameColIndex = headers.indexOf("Nombre Producto");
      }
      const productoBaseCol = 26; // Columna Z

      if (productNameColIndex !== -1) {
        ordersSheet.getRange(1, productoBaseCol).setValue("Producto Base").setFontWeight("bold");

        if (reorderedData.length > 0) {
          const importedData = ordersSheet.getRange(2, 1, reorderedData.length, ordersSheet.getLastColumn()).getValues();
          const valuesForZ = [];

          for (let i = 0; i < importedData.length; i++) {
            // Normalize product name from order by trimming whitespace before lookup.
            const productName = (importedData[i][productNameColIndex] || '').toString().trim();
            const baseProduct = skuMap.get(productName) || "";
            valuesForZ.push([baseProduct]);
          }

          if (valuesForZ.length > 0) {
            ordersSheet.getRange(2, productoBaseCol, valuesForZ.length, 1).setValues(valuesForZ);
          }
        }
      } else {
        Logger.log("Advertencia: No se encontró la columna 'Item Name' en la hoja 'Orders'. No se pudo poblar la columna 'Producto Base'.");
      }
    }
    // --- FIN LÓGICA ---

    // Forzar la actualización de la hoja para que los valores de "Producto Base" estén disponibles para la siguiente sección.
    SpreadsheetApp.flush();

    // --- INICIO: LÓGICA PARA AGREGAR STOCK REAL Y UNIDAD DE VENTA (v2) ---
    const inventarioSheet = mainSpreadsheet.getSheetByName("Inventario Actual");
    if (!inventarioSheet) {
      Logger.log("ADVERTENCIA: No se encontró la hoja 'Inventario Actual'. No se pudo poblar 'Stock Real' ni 'Unidad Venta'.");
    } else {
      Logger.log("Paso 1: Leyendo y procesando 'Inventario Actual'.");
      const inventarioData = inventarioSheet.getDataRange().getValues();
      const inventarioHeaders = inventarioData.shift();
      const productCol = inventarioHeaders.indexOf("Producto Base");
      const stockCol = inventarioHeaders.indexOf("Stock Real");
      const unitCol = inventarioHeaders.indexOf("Unidad Venta");
      const timestampCol = inventarioHeaders.indexOf("Timestamp");

      const latestInventory = new Map();

      if (productCol !== -1 && stockCol !== -1 && unitCol !== -1 && timestampCol !== -1) {
        inventarioData.forEach(row => {
          const productName = row[productCol];
          if (!productName) return; // Si no hay nombre de producto, saltar fila

          // Lógica de parseo de fecha robusta
          let timestamp;
          const rawDate = row[timestampCol];

          if (rawDate instanceof Date && !isNaN(rawDate)) {
            // Si ya es un objeto Date válido, lo usamos directamente.
            timestamp = rawDate;
          } else {
            // Si es un string o número, intentamos convertirlo.
            const dateParts = String(rawDate).split(/[\s/:]+/);
            if (dateParts.length >= 6) { // Formato D/M/YYYY HH:MM:SS
              timestamp = new Date(dateParts[2], dateParts[1] - 1, dateParts[0], dateParts[3], dateParts[4], dateParts[5]);
            } else if (dateParts.length === 3) { // Formato D/M/YYYY
              timestamp = new Date(dateParts[2], dateParts[1] - 1, dateParts[0]);
            } else {
              // Fallback para otros formatos que new Date() pueda entender (ej. ISO) o si es inválido.
              timestamp = new Date(rawDate);
            }
          }

          const normalizedProduct = normalizeKey(productName); // Normalizar la clave
          if (normalizedProduct && !isNaN(timestamp.getTime())) {
            const existing = latestInventory.get(normalizedProduct);
            if (!existing || timestamp > existing.timestamp) {
              latestInventory.set(normalizedProduct, {
                stock: row[stockCol],
                unit: row[unitCol],
                timestamp: timestamp
              });
            }
          }
        });
        Logger.log(`Paso 1 completado. Se procesaron ${latestInventory.size} productos únicos en el inventario.`);
      } else {
        Logger.log("ADVERTENCIA: No se encontraron todas las columnas necesarias ('Producto Base', 'Stock Real', 'Unidad Venta', 'Timestamp') en 'Inventario Actual'.");
      }

      Logger.log("Paso 2: Agregando encabezados a la hoja 'Orders'.");
      const stockRealCol = 27; // Columna AA
      const unidadVentaCol = 28; // Columna AB
      ordersSheet.getRange(1, stockRealCol).setValue("Stock Real").setFontWeight("bold");
      ordersSheet.getRange(1, unidadVentaCol).setValue("Unidad Venta").setFontWeight("bold");

      Logger.log("Paso 3: Poblando las nuevas columnas.");
      if (reorderedData.length > 0) {
        const ordersDataRange = ordersSheet.getRange(2, 1, reorderedData.length, ordersSheet.getLastColumn());
        const ordersData = ordersDataRange.getValues();
        const productoBaseColIndex = 25; // Columna Z tiene el "Producto Base"

        const valuesForAA = [];
        const valuesForAB = [];
        let foundCount = 0;

        for (let i = 0; i < ordersData.length; i++) {
          const baseProduct = ordersData[i][productoBaseColIndex];
          const normalizedBaseProduct = normalizeKey(baseProduct); // Normalizar para la búsqueda
          const inventoryInfo = latestInventory.get(normalizedBaseProduct);

          if (inventoryInfo) {
            valuesForAA.push([inventoryInfo.stock]);
            valuesForAB.push([inventoryInfo.unit]);
            foundCount++;
          } else {
            valuesForAA.push([""]);
            valuesForAB.push([""]);
          }
        }

        Logger.log(`Paso 3 completado. Se encontraron ${foundCount} de ${ordersData.length} productos en el inventario.`);

        if (valuesForAA.length > 0) {
          ordersSheet.getRange(2, stockRealCol, valuesForAA.length, 1).setValues(valuesForAA);
          ordersSheet.getRange(2, unidadVentaCol, valuesForAB.length, 1).setValues(valuesForAB);
          Logger.log("Datos escritos en las columnas AA y AB.");
        }
      }
    }
    // --- FIN: LÓGICA PARA AGREGAR STOCK REAL Y UNIDAD DE VENTA (v2) ---

    return `¡Éxito! Se han importado ${reorderedData.length} filas de pedidos.`;

  } catch (e) {
    Logger.log(`Error en importOrdersFromPastedText: ${e.stack}`);
    throw new Error(`Ocurrió un error durante la importación: ${e.message}`);
  }
}

function normalizeHeader(header) {
    if (typeof header !== 'string') return '';
    const normalized = header.toString().toLowerCase().trim().replace(/:/g, '');

    const mappings = {
        'número de pedido': 'order #',
        'nombre completo': 'nombre y apellido',
        'cantidad': 'item quantity',
        'total de la línea del pedido': 'line total',
        'nombre producto': 'item name',
        'rut cliente': 'rut cliente',
        'metodo de pago': 'payment method',
        'importe total del pedido': 'importe total del pedido',
        'depto/condominio': 'shipping city', // Asumiendo que Depto/Condominio puede mapear a ciudad de envío si es necesario
        'comuna': 'shipping region' // Asumiendo que Comuna mapea a región de envío
    };

    return mappings[normalized] || normalized;
}

function importOrdersFromXLSX(fileId) {
  let tempSheetId = null;
  try {
    const resource = {
      title: `[Temp] Importación de Pedidos - ${new Date().toISOString()}`,
      mimeType: MimeType.GOOGLE_SHEETS
    };
    const tempFile = Drive.Files.copy(resource, fileId);
    tempSheetId = tempFile.id;
    const tempSpreadsheet = SpreadsheetApp.openById(tempSheetId);
    const tempSheet = tempSpreadsheet.getSheets()[0];
    const sourceDataWithHeaders = tempSheet.getDataRange().getValues();
    if (!sourceDataWithHeaders || sourceDataWithHeaders.length < 2) {
      throw new Error("El archivo seleccionado está vacío o no tiene datos.");
    }

    const sourceHeaders = sourceDataWithHeaders.shift().map(h => normalizeHeader(h));

    const mainSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const ordersSheet = mainSpreadsheet.getSheetByName("Orders");
    if (!ordersSheet) {
      throw new Error("No se encontró la hoja 'Orders' en el libro principal.");
    }
    const targetHeaders = ordersSheet.getRange(1, 1, 1, ordersSheet.getLastColumn()).getValues()[0];
    const normalizedTargetHeaders = targetHeaders.map(h => normalizeHeader(h));

    const columnIndexMap = normalizedTargetHeaders.map(targetHeader => sourceHeaders.indexOf(targetHeader));

    const reorderedData = sourceDataWithHeaders.map(sourceRow => {
        const newRow = [];
        columnIndexMap.forEach((sourceIndex, targetIndex) => {
            newRow[targetIndex] = (sourceIndex !== -1) ? sourceRow[sourceIndex] : "";
        });
        return newRow;
    });

    ordersSheet.getRange(2, 1, ordersSheet.getMaxRows() - 1, ordersSheet.getMaxColumns()).clearContent();
    if (reorderedData.length > 0) {
        ordersSheet.getRange(2, 1, reorderedData.length, reorderedData[0].length).setValues(reorderedData);
    }

    return `¡Éxito! Se han importado ${reorderedData.length} filas de pedidos.`;

  } catch (e) {
    Logger.log(`Error en importOrdersFromXLSX: ${e.toString()}\n${e.stack}`);
    if (e.message.includes("You do not have permission to call Drive.Files.copy")) {
        throw new Error("Error de Permisos: La API de Google Drive no está activada. Por favor, actívala en el editor de Apps Script (Servicios > +) y vuelve a intentarlo.");
    }
    throw new Error(`Ocurrió un error durante la importación: ${e.message}`);
  } finally {
    if (tempSheetId) {
      Drive.Files.remove(tempSheetId);
      Logger.log(`Archivo temporal eliminado: ${tempSheetId}`);
    }
  }
}


/**********************
 * PANEL DE NOTIFICACIONES
 **********************/

function openNotificationPanel() {
  const html = HtmlService.createHtmlOutputFromFile('NotificationPanel')
    .setWidth(1000)
    .setHeight(720);
  SpreadsheetApp.getUi().showModalDialog(html, 'Panel de Notificación a Proveedores');
}

/**
 * Lee "Lista de Adquisiciones" y arma:
 *  providers: [{ name, phone, items:[{name, presentation, qty}] }]
 * Donde phone sale de "Proveedores" (A: Nombre, B: Teléfono) con fallback a "SKU" (I: Proveedor, J: Teléfono).
 */
function api_getPanelData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const adquis = ss.getSheetByName('Lista de Adquisiciones');
  const proveedoresSheet = ss.getSheetByName('Proveedores');
  const skuSheet = ss.getSheetByName('SKU');

  // Mapa de todos los proveedores
  const allSuppliers = api_getAllSuppliers();

  if (!adquis || adquis.getLastRow() < 2) return { providers: [], allSuppliers: allSuppliers };

  // --- (MODIFIED) Create a map for Product Base -> Category from SKU sheet ---
  const categoryByProduct = new Map();
  if (skuSheet && skuSheet.getLastRow() > 1) {
    // B: Producto Base, F: Categoría
    skuSheet.getRange(2, 2, skuSheet.getLastRow() - 1, 5).getValues()
      .forEach(([baseProduct, , , , category]) => {
        const bp = String(baseProduct || '').trim();
        const cat = String(category || '').trim();
        if (bp && cat) {
          categoryByProduct.set(bp, cat);
        }
      });
  }

  // Mapas de teléfonos
  const phoneBySupplier = new Map();
  if (proveedoresSheet && proveedoresSheet.getLastRow() > 1) {
    proveedoresSheet.getRange(2, 1, proveedoresSheet.getLastRow() - 1, 2).getValues()
      .forEach(([name, phone]) => {
        if (name) phoneBySupplier.set(String(name).trim(), String(phone || '').trim());
      });
  }
  if (skuSheet && skuSheet.getLastRow() > 1) {
    // I: Proveedor, J: Teléfono
    skuSheet.getRange(2, 9, skuSheet.getLastRow() - 1, 2).getValues()
      .forEach(([supplier, phone]) => {
        const s = String(supplier || '').trim();
        if (s && !phoneBySupplier.has(s) && phone) {
          phoneBySupplier.set(s, String(phone).trim());
        }
      });
  }

  // Leemos adquisiciones: A: Producto Base, B: Cantidad a Comprar, C: Formato, D: Inv. Actual, E: Unidad Inv., F: Necesidad Venta, L: Proveedor
  const data = adquis.getRange(2, 1, adquis.getLastRow() - 1, 12).getValues();
  const bySupplier = new Map();

  data.forEach(row => {
    const productBase = String(row[0] || '').trim();
    const qty         = parseFloat(String(row[1] || '0').replace(',', '.')) || 0;
    const formatStr   = String(row[2] || '').trim();
    const invActual   = parseFloat(String(row[3] || '0').replace(',', '.')) || 0;
    const invUnit     = String(row[4] || 'un.').trim();
    const salesNeed   = parseFloat(String(row[5] || '0').replace(',', '.')) || 0;
    const supplier    = String(row[11] || '').trim();

    if (!supplier || !productBase || qty === 0) return;

    if (!bySupplier.has(supplier)) {
      bySupplier.set(supplier, {
        name: supplier,
        phone: phoneBySupplier.get(supplier) || '',
        items: []
      });
    }
    bySupplier.get(supplier).items.push({
      name: productBase,
      presentation: formatStr,
      qty: qty,
      currentInventory: invActual,
      salesNeed: salesNeed,
      unit: invUnit,
      category: categoryByProduct.get(productBase) || 'Sin Categoría' // --- (ADDED) ---
    });
  });

  // Ordenamos alfabéticamente y devolvemos
  const providers = Array.from(bySupplier.values())
    .sort((a,b)=> a.name.localeCompare(b.name));
  return { providers, allSuppliers };
}

/**
 * Lee la hoja "SKU" y agrupa los productos por proveedor para la vista de "Favoritos".
 *  providers: [{ name, phone, items:[{name, presentation, qty}] }]
 * Donde phone sale de "Proveedores" (A: Nombre, B: Teléfono) con fallback a "SKU" (I: Proveedor, J: Teléfono).
 */
function api_getFavoritesData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const skuSheet = ss.getSheetByName('SKU');
  const proveedoresSheet = ss.getSheetByName('Proveedores');

  // Mapa de todos los proveedores
  const allSuppliers = api_getAllSuppliers();

  if (!skuSheet || skuSheet.getLastRow() < 2) return { providers: [], allSuppliers: allSuppliers };

  // Mapas de teléfonos (misma lógica que en api_getPanelData)
  const phoneBySupplier = new Map();
  if (proveedoresSheet && proveedoresSheet.getLastRow() > 1) {
    proveedoresSheet.getRange(2, 1, proveedoresSheet.getLastRow() - 1, 2).getValues()
      .forEach(([name, phone]) => {
        if (name) phoneBySupplier.set(String(name).trim(), String(phone || '').trim());
      });
  }
  // Fallback a la hoja SKU para teléfonos
  // I: Proveedor (9), J: Teléfono (10)
  const skuPhoneData = skuSheet.getRange(2, 9, skuSheet.getLastRow() - 1, 2).getValues();
  skuPhoneData.forEach(([supplier, phone]) => {
      const s = String(supplier || '').trim();
      if (s && !phoneBySupplier.has(s) && phone) {
        phoneBySupplier.set(s, String(phone).trim());
      }
    });


  // Leemos SKU: B: Producto Base, C: Formato Compra, E: Unidad Compra, F: Categoría, I: Proveedor
  const data = skuSheet.getRange(2, 1, skuSheet.getLastRow() - 1, 9).getValues();
  const bySupplier = new Map();

  data.forEach(row => {
    const productBase = String(row[1] || '').trim(); // Col B: Producto Base
    const formatStr   = String(row[2] || '').trim(); // Col C: Formato Compra
    const unit        = String(row[4] || 'un.').trim(); // Col E: Unidad Compra
    const category    = String(row[5] || 'Sin Categoría').trim(); // Col F: Categoría
    const supplier    = String(row[8] || '').trim(); // Col I: Proveedor

    if (!supplier || !productBase) return; // Un producto base y un proveedor son necesarios

    if (!bySupplier.has(supplier)) {
      bySupplier.set(supplier, {
        name: supplier,
        phone: phoneBySupplier.get(supplier) || '',
        items: []
      });
    }

    // Evitar duplicados de productos base para el mismo proveedor
    const existingItems = bySupplier.get(supplier).items;
    const isDuplicate = existingItems.some(item => item.name === productBase);

    if (!isDuplicate) {
        existingItems.push({
          name: productBase,
          presentation: formatStr,
          qty: 1, // Default a 1, el usuario lo puede cambiar
          currentInventory: 0, // No aplica en esta vista
          salesNeed: 0, // No aplica en esta vista
          unit: unit,
          category: category // --- (ADDED) ---
        });
    }
  });

  // Ordenamos alfabéticamente y devolvemos
  const providers = Array.from(bySupplier.values())
    .sort((a,b)=> a.name.localeCompare(b.name));
  return { providers, allSuppliers };
}

/**
 * Crea/actualiza el teléfono del proveedor en la hoja "Proveedores".
 */
function api_updateProviderPhone(supplierName, rawPhone) {
  if (!supplierName) throw new Error('Falta nombre de proveedor');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('Proveedores');
  if (!sh) {
    sh = ss.insertSheet('Proveedores');
    sh.getRange(1,1,1,2).setValues([['Nombre','Teléfono']]).setFontWeight('bold');
  }

  const phone = normalizePhoneNumber(rawPhone); // ya existe en tu código
  const last = sh.getLastRow();
  if (last < 2) {
    sh.appendRow([supplierName, phone]);
    return 'OK';
  }

  const range = sh.getRange(2,1,last-1,2).getValues();
  for (let i=0;i<range.length;i++){
    if (String(range[i][0]).trim() === supplierName) {
      sh.getRange(i+2, 2).setValue(phone);
      return 'OK';
    }
  }
  sh.appendRow([supplierName, phone]);
  return 'OK';
}

/**
 * Construye el link de WhatsApp con el formato de mensaje solicitado.
 * No abre ventanas; solo devuelve la URL para que el cliente la copie/abra.
 */
function api_updatePurchaseQuantity(productName, newQuantity) {
  if (!productName) throw new Error('Falta el nombre del producto.');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('Lista de Adquisiciones');
  if (!sh || sh.getLastRow() < 2) {
    throw new Error("No se encontró la hoja 'Lista de Adquisiciones' o está vacía.");
  }

  const data = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues(); // A: Producto Base, B: Cantidad
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]).trim() === productName) {
      const rowIdx = i + 2;
      sh.getRange(rowIdx, 2).setValue(newQuantity);
      // Disparar el recálculo del inventario final en la misma fila
      recalculateRowInventory(sh, rowIdx);
      return { status: 'success', message: `Cantidad de '${productName}' actualizada.` };
    }
  }

  throw new Error(`No se encontró el producto '${productName}' en la lista.`);
}

function api_buildWhatsappLink(rawPhone, supplierName, items) {
  if (!items || !Array.isArray(items) || items.length === 0) {
    throw new Error('No hay ítems seleccionados');
  }
  const phone = normalizePhoneNumber(rawPhone); // reutiliza tu helper
  const intro = '¡Hola! Te envío nuestro pedido para hoy:';
  const lines = items.map(i=>{
    const qty = Math.max(1, parseInt(i.qty,10)||1);
    const pres = i.presentation ? `, ${i.presentation}` : '';
    return `- *${qty}* ${i.name}${pres}`;
  });

  const text = [intro, ...lines, '', '¡Gracias!'].join('\n');
  const url  = `https://api.whatsapp.com/send/?phone=${encodeURIComponent(phone)}&text=${encodeURIComponent(text)}`;
  return url;
}

/**
 * Actualiza la Categoría en la hoja SKU para un producto específico.
 * Busca por "Nombre Producto". Si encuentra varias filas con el mismo nombre, actualiza todas.
 * Devuelve { ok:boolean, updated:number }.
 */
function api_updateProductCategory(productName, newCategory) {
  if (!productName || !newCategory) throw new Error('Datos insuficientes');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('SKU');
  if (!sheet) throw new Error("No se encontró la hoja 'SKU'.");

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { ok: false, msg: 'SKU vacío' };

  const range = sheet.getRange(1, 1, 1, sheet.getLastColumn());
  const headers = range.getValues()[0];
  const nameCol = headers.indexOf('Nombre Producto');
  const catCol  = headers.indexOf('Categoría');
  if (nameCol === -1 || catCol === -1) {
    throw new Error("Faltan columnas 'Nombre Producto' y/o 'Categoría' en SKU");
  }

  const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  let updated = 0;
  for (let i = 0; i < data.length; i++) {
    const rowIndex = i + 2;
    const name = String(data[i][nameCol]).trim();
    if (name && name === String(productName).trim()) {
      sheet.getRange(rowIndex, catCol + 1).setValue(newCategory);
      updated++;
    }
  }
  SpreadsheetApp.flush();
  return { ok: updated > 0, updated };
}

/**
 * (Opcional) Lista de todas las categorías existentes en SKU
 * para poblar el selector del panel con opciones reales.
 */
function getAllCategories() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sku = ss.getSheetByName('SKU');
  if (!sku) return [];
  const last = sku.getLastRow();
  if (last < 2) return [];
  const headers = sku.getRange(1, 1, 1, sku.getLastColumn()).getValues()[0];
  const catCol = headers.indexOf('Categoría');
  if (catCol === -1) return [];
  const values = sku.getRange(2, catCol + 1, last - 1, 1).getValues().flat();
  return [...new Set(values.filter(v => v && String(v).trim() !== ''))].sort();
}

function api_getAllSuppliers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('Proveedores');
  if (!sh || sh.getLastRow() < 2) return [];
  const suppliers = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().flat().filter(String);
  return [...new Set(suppliers)].sort();
}

function api_reassignProductSupplier(productName, newSupplierName) {
  if (!productName || !newSupplierName) {
    throw new Error('Faltan el nombre del producto o el nuevo proveedor.');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('Lista de Adquisiciones');
  if (!sh || sh.getLastRow() < 2) {
    throw new Error("No se encontró la hoja 'Lista de Adquisiciones' o está vacía.");
  }

  const data = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const productCol = headers.indexOf("Producto Base");
  const supplierCol = headers.indexOf("Proveedor");

  if (productCol === -1 || supplierCol === -1) {
    throw new Error("No se encontraron las columnas 'Producto Base' o 'Proveedor' en la hoja 'Lista de Adquisiciones'.");
  }

  let updated = false;
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][productCol]).trim() === productName) {
      const rowIdx = i + 2;
      sh.getRange(rowIdx, supplierCol + 1).setValue(newSupplierName);
      updated = true;
      break; // Asumimos que el producto es único en la lista
    }
  }

  if (!updated) {
    throw new Error(`No se encontró el producto '${productName}' en la lista.`);
  }

  // --- NEW LOGIC TO UPDATE SKU SHEET ---
  const skuSheet = ss.getSheetByName('SKU');
  if (skuSheet && skuSheet.getLastRow() > 1) {
    const skuHeaders = skuSheet.getRange(1, 1, 1, skuSheet.getLastColumn()).getValues()[0];
    const skuProductBaseCol = skuHeaders.indexOf("Producto Base");
    const skuSupplierCol = skuHeaders.indexOf("Proveedor");

    if (skuProductBaseCol !== -1 && skuSupplierCol !== -1) {
      const skuData = skuSheet.getRange(2, 1, skuSheet.getLastRow() - 1, skuSheet.getLastColumn()).getValues();
      for (let i = 0; i < skuData.length; i++) {
        if (String(skuData[i][skuProductBaseCol]).trim() === productName) {
          const rowIdx = i + 2;
          skuSheet.getRange(rowIdx, skuSupplierCol + 1).setValue(newSupplierName);
        }
      }
    }
  }
  // --- END NEW LOGIC ---

  return { status: 'success', message: `Proveedor de '${productName}' actualizado a '${newSupplierName}'.` };
}

/**********************
 * DATA CLEANING FLOW (PASO 2)
 **********************/

/**
 * Orchestrates the data cleaning process.
 * This is the entry point for the "Paso 2: Limpiar Datos" button.
 * It sequentially checks for duplicates and new suppliers, showing a dialog for the first issue found.
 */
function startDashboardRefresh() {
  // First, try to find and show the duplicate orders dialog.
  const duplicatesFoundAndShown = findAndShowDuplicateDialog();

  // If no duplicates were found, try to find and show the new suppliers dialog.
  if (!duplicatesFoundAndShown) {
    const newSuppliersFoundAndShown = findAndShowNewSupplierDialog();

    // If no new suppliers were found either, inform the user.
    if (!newSuppliersFoundAndShown) {
      SpreadsheetApp.getUi().alert('No hay datos para limpiar.');
    }
  }
}

/**
 * Finds customers with multiple order numbers and shows a dialog to manage them.
 * @returns {boolean} True if the dialog was shown, false otherwise.
 */
function findAndShowDuplicateDialog() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ordersSheet = ss.getSheetByName('Orders');
  if (!ordersSheet || ordersSheet.getLastRow() < 2) {
    return false; // No data to check
  }

  const headers = ordersSheet.getRange(1, 1, 1, ordersSheet.getLastColumn()).getValues()[0];
  const idx = indexer(headers);
  const orderIdCol = idx.numPedido;
  const customerNameCol = idx.nombre;

  if (orderIdCol < 0 || customerNameCol < 0) {
    Logger.log("Could not find 'Order #' or 'Nombre y apellido' columns.");
    return false;
  }

  const data = ordersSheet.getRange(2, 1, ordersSheet.getLastRow() - 1, Math.max(orderIdCol, customerNameCol) + 1).getValues();

  const ordersByCustomer = {};
  data.forEach(row => {
    const orderId = String(row[orderIdCol]).trim();
    const customerName = String(row[customerNameCol]).trim();

    if (orderId && customerName) {
      if (!ordersByCustomer[customerName]) {
        ordersByCustomer[customerName] = new Set();
      }
      ordersByCustomer[customerName].add(orderId);
    }
  });

  const duplicates = {};
  for (const customer in ordersByCustomer) {
    if (ordersByCustomer[customer].size > 1) {
      duplicates[customer] = Array.from(ordersByCustomer[customer]);
    }
  }

  if (Object.keys(duplicates).length > 0) {
    const template = HtmlService.createTemplateFromFile('DuplicateDialog');
    template.duplicates = JSON.stringify(duplicates);
    const html = template.evaluate().setWidth(600).setHeight(400);
    SpreadsheetApp.getUi().showModalDialog(html, 'Limpiar Pedidos Duplicados');
    return true;
  }

  return false;
}

/**
 * Deletes all rows associated with the given order numbers.
 * This is called from the DuplicateDialog.
 * @param {string[]} orderNumbers - An array of order numbers to delete.
 * @returns {string} A result message.
 */
function deleteOrdersByNumber(orderNumbers) {
  if (!orderNumbers || !Array.isArray(orderNumbers) || orderNumbers.length === 0) {
    throw new Error('No se proporcionaron números de pedido para eliminar.');
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Orders');
    if (!sheet) throw new Error('No se encontró la hoja "Orders".');

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const idx = indexer(headers);
    const orderIdCol = idx.numPedido;
    const quantityCol = idx.cantidad;

    if (orderIdCol < 0 || quantityCol < 0) {
      throw new Error("No se encontraron las columnas 'Order #' o 'Item Quantity' en la hoja 'Orders'.");
    }

    const data = sheet.getDataRange().getValues();
    let rowsUpdated = 0;
    const ordersToDelete = new Set(orderNumbers);

    for (let i = 1; i < data.length; i++) { // Start from 1 to skip header
      const currentOrderId = String(data[i][orderIdCol]).trim();
      if (ordersToDelete.has(currentOrderId)) {
        const rowNum = i + 1;
        const rowRange = sheet.getRange(rowNum, 1, 1, sheet.getLastColumn());
        rowRange.setBackground('#ff0000'); // Red background

        const quantityCell = sheet.getRange(rowNum, quantityCol + 1);
        const currentQuantity = quantityCell.getValue();
        if (!String(currentQuantity).startsWith('E')) {
          quantityCell.setValue('E' + currentQuantity);
        }
        rowsUpdated++;
      }
    }

    if (rowsUpdated > 0) {
      SpreadsheetApp.flush();
      return `Se marcaron ${rowsUpdated} fila(s) de ${orderNumbers.length} pedido(s) como eliminadas.`;
    } else {
      return 'No se encontraron filas que coincidieran con los pedidos seleccionados.';
    }
  } catch (e) {
    Logger.log(`Error en deleteOrdersByNumber: ${e.stack}`);
    throw new Error(`Ocurrió un error: ${e.message}`);
  }
}

/**
 * Finds suppliers from the SKU sheet that are not in the Proveedores sheet
 * and shows a dialog to add them.
 * @returns {boolean} True if the dialog was shown, false otherwise.
 */
function findAndShowNewSupplierDialog() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const skuSheet = ss.getSheetByName('SKU');
  const proveedoresSheet = ss.getSheetByName('Proveedores');

  if (!skuSheet) {
    Logger.log("SKU sheet not found. Cannot check for new suppliers.");
    return false;
  }

  // Get all unique suppliers from SKU sheet (Column I)
  const skuSuppliers = new Set();
  if (skuSheet.getLastRow() > 1) {
    const skuSupplierData = skuSheet.getRange(2, 9, skuSheet.getLastRow() - 1, 1).getValues();
    skuSupplierData.forEach(row => {
      const supplier = String(row[0]).trim();
      if (supplier) {
        skuSuppliers.add(supplier);
      }
    });
  }

  // Get all existing suppliers from Proveedores sheet (Column A)
  const existingSuppliers = new Set();
  if (proveedoresSheet && proveedoresSheet.getLastRow() > 1) {
    const existingSupplierData = proveedoresSheet.getRange(2, 1, proveedoresSheet.getLastRow() - 1, 1).getValues();
    existingSupplierData.forEach(row => {
      const supplier = String(row[0]).trim();
      if (supplier) {
        existingSuppliers.add(supplier);
      }
    });
  }

  // Find suppliers that are in SKU but not in Proveedores
  const newSuppliers = [];
  for (const supplier of skuSuppliers) {
    if (!existingSuppliers.has(supplier)) {
      newSuppliers.push(supplier);
    }
  }

  if (newSuppliers.length > 0) {
    const template = HtmlService.createTemplateFromFile('NewSupplierDialog');
    template.newSuppliers = JSON.stringify(newSuppliers);
    template.existingSuppliers = JSON.stringify(Array.from(existingSuppliers).sort()); // Pass existing suppliers
    const html = template.evaluate().setWidth(800).setHeight(500); // Increased width and height
    SpreadsheetApp.getUi().showModalDialog(html, 'Añadir Nuevos Proveedores');
    return true;
  }

  return false;
}

/**
 * Saves new suppliers and their phone numbers to the 'Proveedores' sheet.
 * @param {Object} supplierData - An object where keys are supplier names and values are phone numbers.
 * @returns {string} A result message.
 */
function saveOrAssignSuppliers(data) {
  if (!data || (!data.newEntries && !data.assignments)) {
    throw new Error('No se proporcionaron datos válidos para guardar.');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const proveedoresSheet = ss.getSheetByName('Proveedores');
  const skuSheet = ss.getSheetByName('SKU');
  let newEntriesCount = 0;
  let assignmentsCount = 0;
  let updatedRowsCount = 0;

  // 1. Procesa nuevas entradas
  if (data.newEntries && data.newEntries.length > 0) {
    const rowsToAppend = data.newEntries.map(entry => [entry.name, entry.phone]);
    if(rowsToAppend.length > 0){
        proveedoresSheet.getRange(proveedoresSheet.getLastRow() + 1, 1, rowsToAppend.length, 2).setValues(rowsToAppend);
        newEntriesCount = rowsToAppend.length;
    }
  }

  // 2. Procesa asignaciones
  if (data.assignments && data.assignments.length > 0) {
    const skuData = skuSheet.getDataRange().getValues();
    const headers = skuData.shift(); // remove headers
    const supplierColIndex = headers.indexOf('Proveedor');

    if (supplierColIndex === -1) {
      throw new Error("No se encontró la columna 'Proveedor' en la hoja 'SKU'.");
    }

    const assignmentsMap = new Map(data.assignments.map(a => [a.from, a.to]));

    skuData.forEach((row, i) => {
      const currentSupplier = row[supplierColIndex];
      if (assignmentsMap.has(currentSupplier)) {
        const newSupplier = assignmentsMap.get(currentSupplier);
        skuSheet.getRange(i + 2, supplierColIndex + 1).setValue(newSupplier);
        updatedRowsCount++;
      }
    });
    assignmentsCount = assignmentsMap.size;
  }

  SpreadsheetApp.flush();

  let message = "Proceso completado.\n";
  if (newEntriesCount > 0) {
    message += `- Se crearon ${newEntriesCount} nuevos proveedores.\n`;
  }
  if (updatedRowsCount > 0) {
    message += `- Se actualizaron ${updatedRowsCount} registros en la hoja SKU para ${assignmentsCount} asignaciones.`;
  }
  if (newEntriesCount === 0 && updatedRowsCount === 0) {
    message = "No se realizaron cambios."
  }

  return message;
}
