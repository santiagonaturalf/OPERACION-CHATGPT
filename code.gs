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
    { name: "Orders", headers: ["Order #", "Nombre y apellido", "Email", "Phone", "Shipping Address", "Shipping City", "Shipping Region", "Shipping Postcode", "Item Name", "Item SKU", "Item Quantity", "Item Price", "Line Total", "Tax Rate", "Tax Amount", "Importe total del pedido", "Payment Method", "Transaction ID", "Estado del pago"], index: 0 },
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
    numPedido: idxOf('número de pedido','numero de pedido','nº pedido','n° pedido','n pedido','número de ped'),
    nombre:    idxOf('nombre completo','cliente','nombre','nombre y apellido'),
    telefono:  idxOf('teléfono','telefono','phone','tel'),
    direccion: idxOf('dirección','direccion','shipping address','dirección líneas 1','direccion lineas 1'),
    depto:     idxOf('depto.','depto','departamento','depto/condominio','depto/condomi','dirección líneas 2','direccion lineas 2'),
    comuna:    idxOf('comuna','shipping region','ciudad'),
    estado:    idxOf('estado','estado del pago'),
    furgon:    idxOf('furgón','furgon','van','furgón asignado')
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

  const packagingHeaders = ["Orden Ruta", "Nº Pedido", "Numero de Bultos", "Nombre Envasador"];
  packagingSheet.getRange("A1:D1").setValues([packagingHeaders]).setFontWeight('bold');

  if (finalPackagingData.length > 0) {
    packagingSheet.getRange(2, 1, finalPackagingData.length, 4).setValues(finalPackagingData);

    // Aplicar formato a la tabla
    const tableRange = packagingSheet.getRange(1, 1, finalPackagingData.length + 1, 4);
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

  return {
    historicalSuggestions: historicalSuggestions.map(formatSuggestion),
    highConfidenceSuggestions: highConfidenceSuggestions.map(formatSuggestion),
    lowConfidenceSuggestions: lowConfidenceSuggestions.map(formatSuggestion),
    unmatchedPayments: unassignedPayments.map(p => ({ ...p, date: formatDate(p.date) })),
    manualListOrders: manualListOrders.map(o => ({ ...o, date: formatDate(o.date) }))
  };
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
  const categorySummary = {};
  for (const productName in productTotals) {
    const category = skuMap[productName] ? skuMap[productName].category : 'Sin Categoría';
    if (!categorySummary[category]) { categorySummary[category] = { count: 0, products: {} }; }
    categorySummary[category].count++;
    categorySummary[category].products[productName] = productTotals[productName];
  }
  return categorySummary;
}

function generatePackagingSheet(selectedCategories) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const data = getPackagingData();

  // Obtener los mapas de datos necesarios para el inventario
  const inventoryMap = getCurrentInventory();
  const skuSheet = ss.getSheetByName('SKU');
  if (!skuSheet) throw new Error("La hoja 'SKU' no fue encontrada.");
  const skuMap = getSkuMap(skuSheet);

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
      const skuInfo = skuMap[productName];
      const baseProduct = skuInfo ? skuInfo.base : null;
      const inventoryInfo = baseProduct ? inventoryMap[baseProduct] : null;
      // Formatear el valor del inventario para incluir la unidad
      const inventoryValue = inventoryInfo ? `${inventoryInfo.quantity} ${inventoryInfo.unit}` : 'No encontrado';

      productRows.push([products[productName], inventoryValue, productName]);
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

function getAcquisitionDataForEditor() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ordersSheet = ss.getSheetByName('Orders');
  const skuSheet = ss.getSheetByName('SKU');
  const proveedoresSheet = ss.getSheetByName('Proveedores');

  if (!ordersSheet || !skuSheet || !proveedoresSheet) {
    throw new Error('Faltan una o más hojas requeridas: "Orders", "SKU", o "Proveedores".');
  }

  // 0. Get current inventory first
  const inventoryMap = getCurrentInventory();

  // 1. Generar el plan de adquisiciones (lógica reutilizada)
  const { productToSkuMap, baseProductPurchaseOptions } = getPurchaseDataMaps(skuSheet);
  const baseProductNeeds = calculateBaseProductNeeds(ordersSheet, productToSkuMap);
  const acquisitionPlan = createAcquisitionPlan(baseProductNeeds, baseProductPurchaseOptions, inventoryMap);

  // 2. Obtener la lista de proveedores
  const supplierData = proveedoresSheet.getRange("A2:A" + proveedoresSheet.getLastRow()).getValues().flat().filter(String);
  const supplierSet = new Set(supplierData);
  supplierSet.add("Patio Mayorista"); // Asegurarse de que "Patio Mayorista" esté disponible

  // Convertir el plan de un objeto a un array para que sea más fácil de manejar en el lado del cliente
  const planAsArray = Object.values(acquisitionPlan);

  return {
    acquisitionPlan: planAsArray,
    allSuppliers: Array.from(supplierSet).sort()
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
      if (!priceMap[productName]) {
        priceMap[productName] = [];
      }
      priceMap[productName].push({
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
      inventoryMap[productName] = {
        quantity: parseFloat(String(quantity).replace(",", ".")) || 0,
        unit: unit || ''
      };
    }
  });

  return inventoryMap;
}

/**
 * Lee la hoja "Historico Adquisiciones" y crea un mapa con el proveedor más reciente para cada producto base.
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

  // Columnas: C (Producto Base), H (Proveedor).
  const data = historicoSheet.getRange("C2:H" + historicoSheet.getLastRow()).getValues();

  // Iterar hacia atrás para encontrar la entrada más reciente primero.
  for (let i = data.length - 1; i >= 0; i--) {
    const row = data[i];
    const productoBase = row[0]; // Índice 0 en el rango C:H corresponde a la columna C
    const proveedor = row[5];    // Índice 5 en el rango C:H corresponde a la columna H

    // Si encontramos un producto y un proveedor, y aún no lo hemos guardado, lo añadimos al mapa.
    if (productoBase && proveedor && !latestSuppliers[productoBase]) {
      latestSuppliers[productoBase] = String(proveedor).trim();
    }
  }

  Logger.log("Proveedores más recientes obtenidos del historial: " + JSON.stringify(latestSuppliers));
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
  skuData.forEach(row => {
    const [nombreProducto, productoBase, formatoCompra, cantidadCompra, unidadCompra, cat, cantVenta, unidadVenta, proveedor] = row;
    if (nombreProducto) {
      productToSkuMap[nombreProducto] = {
        productoBase,
        cantidadVenta: parseFloat(String(cantVenta).replace(',', '.')) || 0,
        unidadVenta: normalizeUnit(unidadVenta)
      };
    }
    if (productoBase && formatoCompra) {
      if (!baseProductPurchaseOptions[productoBase]) {
        baseProductPurchaseOptions[productoBase] = { options: [], suppliers: new Set() };
      }
      baseProductPurchaseOptions[productoBase].options.push({
        name: formatoCompra,
        size: parseFloat(String(cantidadCompra).replace(',', '.')) || 0,
        unit: normalizeUnit(unidadCompra)
      });
      if (proveedor) baseProductPurchaseOptions[productoBase].suppliers.add(proveedor);
    }
  });
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

function createAcquisitionPlan(baseProductNeeds, baseProductPurchaseOptions, inventoryMap) {
  const acquisitionPlan = {};
  const latestSuppliers = getLatestSuppliersFromHistory(); // Llama a la nueva función

  for (const baseProduct in baseProductNeeds) {
    if (baseProductPurchaseOptions[baseProduct]) {
      const needs = baseProductNeeds[baseProduct];
      const purchaseInfo = baseProductPurchaseOptions[baseProduct];
      const purchaseOptions = purchaseInfo.options;
      const needUnit = Object.keys(needs)[0];
      const totalNeed = needs[needUnit];
      let bestOption = null;
      let minWaste = Infinity;

      // Get current inventory for this product, defaulting to 0
      const inventoryInfo = (inventoryMap && inventoryMap[baseProduct]) ? inventoryMap[baseProduct] : { quantity: 0, unit: needUnit };
      const netNeed = Math.max(0, totalNeed - inventoryInfo.quantity);

      purchaseOptions.forEach((option) => {
        if (option.unit === needUnit && option.size > 0) {
          const numToBuy = netNeed > 0 ? Math.ceil(netNeed / option.size) : 0;
          const waste = (numToBuy * option.size) - netNeed;
          if (waste < minWaste) {
            minWaste = waste;
            bestOption = { ...option, suggestedQty: numToBuy };
          }
        }
      });

      if (bestOption) {
        // --- NUEVA LÓGICA PARA PROVEEDOR ---
        const historicalSupplier = latestSuppliers[baseProduct];
        const skuSuppliers = Array.from(purchaseInfo.suppliers);
        const defaultSkuSupplier = skuSuppliers.length > 0 ? skuSuppliers[0] : "Patio Mayorista";

        acquisitionPlan[baseProduct] = {
          productName: baseProduct,
          totalNeed,
          unit: needUnit,
          saleUnit: needUnit,
          supplier: historicalSupplier || defaultSkuSupplier, // Usa el proveedor histórico con fallback
          availableFormats: purchaseOptions,
          suggestedFormat: bestOption,
          suggestedQty: bestOption.suggestedQty,
          currentInventory: inventoryInfo.quantity,
          currentInventoryUnit: inventoryInfo.unit
        };
      }
    }
  }
  return acquisitionPlan;
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
  if (!adquis || adquis.getLastRow() < 2) return { providers: [] };

  const proveedoresSheet = ss.getSheetByName('Proveedores');
  const skuSheet = ss.getSheetByName('SKU');

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
      unit: invUnit
    });
  });

  // Ordenamos alfabéticamente y devolvemos
  const providers = Array.from(bySupplier.values())
    .sort((a,b)=> a.name.localeCompare(b.name));
  return { providers };
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
