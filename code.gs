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
    .addItem('💬 Notificar a Proveedores', 'startNotificationProcess')
    .addSeparator()
    .addItem('📈 Analizar Adquisiciones', 'runPriceAnalysis')
    .addSeparator();

  const maintenanceSubMenu = ui.createMenu('Mantenimiento y Reportes')
    .addItem('⚙️ Calcular Costo de Adquisición en Orders', 'calculateOrderLineCost')
    .addItem('🔍 Revisar Formatos Desconocidos', 'runCostResolutionTool')
    .addItem('⚠️ Detectar Anomalías de Costo', 'runAnomalyReviewTool')
    .addSeparator()
    .addItem('✨ Normalizar Productos Base', 'showNormalizationDialog')
    .addItem('📂 Normalizar Categorías', 'showCategoryNormalizerDialog')
    .addItem('🔍 Verificador de Coherencia', 'showConsistencyCheckerDialog')
    .addSeparator()
    .addItem('🧾 Reporte de Inconsistencias', 'showInconsistencyReportDialog');

  operationsMenu.addSubMenu(maintenanceSubMenu);
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

/**
 * Calculates the acquisition cost for each line item in the 'Orders' sheet.
 * It adds a 'Costo Adquisicion' column if it doesn't exist.
 * If a product's cost is not found in 'CostosVenta', it uses the order line's total as a fallback.
 * This function is intended to be triggered manually from a menu.
 */
function calculateOrderLineCost(showAlert = true) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const ordersSheet = ss.getSheetByName('Orders');
  const costosSheet = ss.getSheetByName('CostosVenta');

  if (!ordersSheet || !costosSheet) {
    if (showAlert) ui.alert("Faltan las hojas 'Orders' o 'CostosVenta' para realizar el cálculo.");
    throw new Error("Missing required sheets.");
  }

  try {
    const startTime = new Date();
    if (showAlert) {
      ui.alert("Iniciando el cálculo de Costo de Adquisición. Esto puede tardar unos momentos. Por favor, espera el mensaje de confirmación.");
    }

    // 1. Create a map of product costs from 'CostosVenta' for efficient lookup.
    const productCostMap = new Map();
    if (costosSheet.getLastRow() > 1) {
      const costosData = costosSheet.getRange("B2:C" + costosSheet.getLastRow()).getValues();
      for (let i = costosData.length - 1; i >= 0; i--) {
        const row = costosData[i];
        const productName = row[0];
        const cost = parseFloat(String(row[1]).replace(',', '.'));
        if (productName && !productCostMap.has(productName) && !isNaN(cost)) {
          productCostMap.set(productName, cost);
        }
      }
    }

    // 2. Find or create the "Costo Adquisicion" column.
    const headerRange = ordersSheet.getRange(1, 1, 1, ordersSheet.getMaxColumns());
    const headers = headerRange.getValues()[0];
    const targetHeader = "Costo Adquisicion";
    let targetColPosition = headers.indexOf(targetHeader) + 1;

    if (targetColPosition === 0) { // Header doesn't exist, create it.
      const requestedCol = 27; // Column AA
      const lastCol = ordersSheet.getLastColumn();

      if (requestedCol <= lastCol && ordersSheet.getRange(1, requestedCol).getValue() !== "") {
        targetColPosition = lastCol + 1; // Fallback to the end if AA is taken
      } else {
        targetColPosition = requestedCol; // Use AA
      }
      ordersSheet.getRange(1, targetColPosition).setValue(targetHeader).setFontWeight("bold");
    }

    // 3. Iterate through orders and calculate costs.
    const lastRow = ordersSheet.getLastRow();
    if (lastRow < 2) {
      if (showAlert) ui.alert("No hay pedidos para procesar.");
      return "No hay pedidos para procesar.";
    }

    const dataRange = ordersSheet.getRange(2, 1, lastRow - 1, Math.max(13, targetColPosition));
    const ordersData = dataRange.getValues();
    const costsToUpdate = [];

    const PRODUCT_NAME_COL = 9;  // Col J
    const QUANTITY_COL = 10;     // Col K
    const LINE_TOTAL_COL = 12;   // Col M

    for (let i = 0; i < ordersData.length; i++) {
      const row = ordersData[i];
      const productName = row[PRODUCT_NAME_COL];

      if (!productName || productName.toString().trim() === "") {
        costsToUpdate.push([""]);
        continue;
      }

      const quantity = parseInt(row[QUANTITY_COL], 10) || 0;
      const lineTotal = parseFloat(String(row[LINE_TOTAL_COL]).replace(',', '.')) || 0;

      let acquisitionCost;
      if (productCostMap.has(productName)) {
        acquisitionCost = productCostMap.get(productName) * quantity;
      } else {
        acquisitionCost = lineTotal; // Fallback: Use line total from column M
      }
      costsToUpdate.push([acquisitionCost]);
    }

    // 4. Write new costs to the sheet in a single batch.
    if (costsToUpdate.length > 0) {
      ordersSheet.getRange(2, targetColPosition, costsToUpdate.length, 1).setValues(costsToUpdate);
    }

    const endTime = new Date();
    const duration = (endTime - startTime) / 1000;

    const message = `Cálculo completado en ${duration.toFixed(1)} segundos. Se ha actualizado la columna '${targetHeader}' para ${costsToUpdate.length} filas.`;
    if (showAlert) {
      ui.alert(message);
    }
    return message;

  } catch (e) {
    Logger.log(`Error en calculateOrderLineCost: ${e.stack}`);
    if (showAlert) {
      ui.alert(`Ocurrió un error durante el cálculo: ${e.message}`);
    }
    throw e;
  }
}

/**
 * Finds products that have been sold but whose purchase format size cannot be determined.
 * This happens when a format (e.g., "bins") is not defined with a size in the SKU sheet.
 * @returns {Array<Object>} An array of objects, each representing an unresolved item.
 */
function findUnresolvedCostFormats() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const historicoSheet = ss.getSheetByName("Historico Adquisiciones");
  const skuSheet = ss.getSheetByName("SKU");
  const ordersSheet = ss.getSheetByName("Orders");

  // Build the necessary maps
  const skuMap = new Map();
  const purchaseFormatMap = new Map();
  skuSheet.getDataRange().getValues().slice(1).forEach(row => {
    const nombreProducto = row[0];
    const productoBase = row[1];
    const formatoCompra = row[2];
    const cantidadCompra = parseFloat(String(row[3]).replace(',', '.')) || 0;
    const unidadVenta = normalizeUnit(row[7]);
    if (nombreProducto && productoBase) {
      skuMap.set(nombreProducto, { productoBase, unidadVenta });
    }
    if (productoBase && formatoCompra) {
      // Use toLowerCase() for case-insensitive matching
      purchaseFormatMap.set(`${productoBase.toLowerCase()}|${formatoCompra.toLowerCase()}`, cantidadCompra);
    }
  });

  const priceMap = new Map();
  historicoSheet.getDataRange().getValues().slice(1).forEach(row => {
    const productoBase = row[2];
    const formato = row[3];
    // This is the corrected logic: always get the latest format from the bottom of the sheet
    if (productoBase) {
      priceMap.set(productoBase.toLowerCase(), { formato });
    }
  });

  const productsSoldToday = new Set();
  ordersSheet.getDataRange().getValues().slice(1).forEach(row => {
    const productName = row[9];
    if (productName) productsSoldToday.add(productName);
  });

  const unresolvedItems = [];
  const unresolvedKeys = new Set(); // To avoid duplicates

  productsSoldToday.forEach(productName => {
    const skuInfo = skuMap.get(productName);
    if (!skuInfo) return;

    const priceInfo = priceMap.get(skuInfo.productoBase.toLowerCase());
    if (!priceInfo || !priceInfo.formato) return;

    const match = priceInfo.formato.toString().match(/\(([\d.,]+)/);
    if (!match) { // If it doesn't have a "(size)" part
      const formatKey = `${skuInfo.productoBase.toLowerCase()}|${priceInfo.formato.toLowerCase()}`;
      if (!purchaseFormatMap.has(formatKey)) {
        const uniqueKey = `${skuInfo.productoBase}|${priceInfo.formato}`;
        if (!unresolvedKeys.has(uniqueKey)) {
          unresolvedItems.push({
            productoBase: skuInfo.productoBase,
            formatoCompra: priceInfo.formato,
            unidadVenta: skuInfo.unidadVenta
          });
          unresolvedKeys.add(uniqueKey);
        }
      }
    }
  });

  return unresolvedItems;
}

/**
 * Runs the cost resolution tool. Finds unresolved costs and opens a dialog for the user to fix them.
 */
function runCostResolutionTool() {
  const unresolvedItems = findUnresolvedCostFormats();

  if (unresolvedItems.length === 0) {
    SpreadsheetApp.getUi().alert("¡Buenas noticias! No se encontraron formatos de compra desconocidos. Todos los costos pueden ser calculados correctamente.");
    return;
  }

  const template = HtmlService.createTemplateFromFile('ResolveCostsDialog');
  template.unresolvedItemsJSON = JSON.stringify(unresolvedItems);

  const html = template.evaluate()
    .setWidth(600)
    .setHeight(400);

  SpreadsheetApp.getUi().showModalDialog(html, 'Resolver Formatos Desconocidos');
}

/**
 * Saves the resolved format information from the dialog to the SKU sheet.
 * @param {Array<Object>} data An array of objects, each representing a format to be defined.
 * @returns {string} A confirmation message.
 */
function saveResolvedCosts(data) {
  if (!data || !Array.isArray(data) || data.length === 0) {
    throw new Error("No se proporcionaron datos válidos para guardar.");
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const skuSheet = ss.getSheetByName("SKU");
  if (!skuSheet) {
    throw new Error("No se encontró la hoja 'SKU'.");
  }

  const rowsToAppend = [];
  data.forEach(item => {
    // Create a new row for the SKU sheet.
    // We create a placeholder "Nombre Producto" as the main purpose is to define the format.
    const newRow = [
      `FORMATO - ${item.formatoCompra} para ${item.productoBase}`, // Nombre Producto (placeholder)
      item.productoBase,      // Producto Base
      item.formatoCompra,     // Formato Compra
      item.cantidadCompra,    // Cantidad Compra
      item.unidadCompra,      // Unidad Compra
      '', // Categoría
      '', // Cantidad Venta
      '', // Unidad Venta
      '', // Proveedor
      ''  // Teléfono
    ];
    rowsToAppend.push(newRow);
  });

  if (rowsToAppend.length > 0) {
    skuSheet.getRange(skuSheet.getLastRow() + 1, 1, rowsToAppend.length, rowsToAppend[0].length).setValues(rowsToAppend);
  }

  return `${rowsToAppend.length} nuevo(s) formato(s) guardado(s) en la hoja SKU. Por favor, vuelve a ejecutar la herramienta de cálculo de costos o el análisis que estabas realizando.`;
}

/**
 * Scans the 'Orders' sheet to find anomalies where the acquisition cost is higher than the sale price.
 * @returns {Array<Object>} An array of objects, each representing an anomalous row with its details.
 */
function findPriceCostAnomalies() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ordersSheet = ss.getSheetByName('Orders');
  if (!ordersSheet) {
    throw new Error("No se encontró la hoja 'Orders'.");
  }

  const headers = ordersSheet.getRange(1, 1, 1, ordersSheet.getLastColumn()).getValues()[0];
  const productNameCol = findColumnIndexRobust(headers, ["item name", "nombre producto"]);
  const priceCol = findColumnIndexRobust(headers, ["item price", "precio actual del producto"]);
  const costCol = findColumnIndexRobust(headers, ["costo adquisicion"]);

  if (productNameCol === -1 || priceCol === -1 || costCol === -1) {
    throw new Error("Una o más columnas requeridas no se encontraron: 'Nombre Producto' (o 'Item Name'), 'Precio actual del producto' (o 'Item Price'), 'Costo Adquisicion'.");
  }

  const data = ordersSheet.getDataRange().getValues();
  const anomalies = [];

  for (let i = 1; i < data.length; i++) { // Start from row 2 (index 1)
    const row = data[i];
    const price = parseFloat(String(row[priceCol]).replace(/\./g, '').replace(',', '.')) || 0;
    const cost = parseFloat(String(row[costCol]).replace(/\./g, '').replace(',', '.')) || 0;

    // Anomaly condition: cost is greater than price, and price is not zero
    if (price > 0 && cost > price) {
      anomalies.push({
        rowNumber: i + 1, // Add 1 to convert 0-based index to 1-based row number
        productName: row[productNameCol],
        salePrice: price,
        incorrectCost: cost
      });
    }
  }
  return anomalies;
}

/**
 * Runs the anomaly review tool. Finds price-cost anomalies and shows a dialog for user correction.
 */
function runAnomalyReviewTool() {
  const anomalies = findPriceCostAnomalies();

  if (anomalies.length === 0) {
    SpreadsheetApp.getUi().alert("¡Buenas noticias! No se encontraron anomalías de costo vs. precio en la hoja 'Orders'.");
    return;
  }

  const template = HtmlService.createTemplateFromFile('ReviewAnomaliesDialog');
  template.anomaliesJSON = JSON.stringify(anomalies);

  const html = template.evaluate()
    .setWidth(700)
    .setHeight(500);

  SpreadsheetApp.getUi().showModalDialog(html, 'Revisar Anomalías de Costo');
}

/**
 * Saves the manually corrected costs from the anomaly review dialog back to the 'Orders' sheet.
 * @param {Array<Object>} corrections An array of objects, each with a rowNumber and a newCost.
 * @returns {string} A confirmation message.
 */
function saveCorrectedCosts(corrections) {
  if (!corrections || !Array.isArray(corrections) || corrections.length === 0) {
    throw new Error("No se proporcionaron datos válidos para guardar.");
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ordersSheet = ss.getSheetByName('Orders');
  if (!ordersSheet) {
    throw new Error("No se encontró la hoja 'Orders'.");
  }

  const headers = ordersSheet.getRange(1, 1, 1, ordersSheet.getLastColumn()).getValues()[0];
  const costCol = headers.indexOf("Costo Adquisicion");

  if (costCol === -1) {
    throw new Error("No se encontró la columna 'Costo Adquisicion' para actualizar.");
  }

  const costColPosition = costCol + 1;

  let updatedCount = 0;
  corrections.forEach(correction => {
    if (correction.rowNumber && correction.newCost !== null && !isNaN(correction.newCost)) {
      ordersSheet.getRange(correction.rowNumber, costColPosition).setValue(correction.newCost);
      updatedCount++;
    }
  });

  return `${updatedCount} costo(s) ha(n) sido corregido(s) exitosamente.`;
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

      // Si la hoja creada es 'CostosVenta', poblarla inmediatamente.
      if (sheetName === "CostosVenta") {
        Logger.log("La hoja 'CostosVenta' no existía. Se procederá a poblarla con los datos de hoy.");
        actualizarCostosDeVentaDiarios();
      }
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
    { name: "ClientBankData", headers: ["PaymentIdentifier", "CustomerRUT", "CustomerName", "LastSeen"], index: 7 },
    { name: "Historico Adquisiciones", headers: ["ID", "Fecha de Registro", "Producto Base", "Formato de Compra", "Cantidad Comprada", "Precio Compra", "Costo Total Compra", "Proveedor"], index: 8 },
    { name: "CostosVenta", headers: ["Fecha", "Nombre Producto", "Costo Adquisicion"], index: 9 },
    { name: "Anomalías de Precios", headers: ["Fecha", "Nombre Producto", "Costo de Hoy", "Costo Promedio Histórico", "Desviación Estándar", "Nivel de Desviación (StdDevs)", "Mensaje"], index: 10 }
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
        const sheet = ss.getSheetByName('Orders');
        const data = sheet.getDataRange().getValues();
        const rowsToDelete = [];

        // Find all rows matching the orderId, starting from the end to avoid shifting issues
        for (let i = data.length - 1; i >= 1; i--) {
            if (String(data[i][0]) === String(orderId)) {
                rowsToDelete.push(i + 1);
            }
        }

        if (rowsToDelete.length > 0) {
            rowsToDelete.forEach(rowNum => sheet.deleteRow(rowNum));
            SpreadsheetApp.flush();
            return { status: 'success', message: `Pedido #${orderId} (${rowsToDelete.length} filas) eliminado exitosamente.` };
        } else {
            return { status: 'error', message: `No se encontró el pedido #${orderId} para eliminar.` };
        }
    } catch (e) {
        Logger.log(`Error en deleteOrder: ${e.stack}`);
        return { status: 'error', message: `Ocurrió un error al eliminar el pedido: ${e.message}` };
    }
}


// --- LÓGICA DE COMANDA RUTAS ---

function showComandaRutasDialog() {
  const html = HtmlService.createHtmlOutputFromFile('ComandaRutasDialog')
    .setWidth(1200)
    .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, 'Comanda Rutas');
}

function getOrdersForRouting() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ordersSheet = ss.getSheetByName('Orders');
  if (!ordersSheet) {
    throw new Error('No se encontró la hoja "Orders".');
  }

  const lastRow = ordersSheet.getLastRow();
  if (lastRow < 2) return [];

  // Expand range to read up to the last column to dynamically find 'Furgón'
  const dataRange = ordersSheet.getRange(2, 1, lastRow - 1, ordersSheet.getLastColumn());
  const values = dataRange.getValues();
  const headers = ordersSheet.getRange(1, 1, 1, ordersSheet.getLastColumn()).getValues()[0];
  const vanColumnIndex = headers.indexOf('Furgón');

  const uniqueOrders = {};
  values.forEach((row) => {
    const orderId = row[0];
    if (orderId && !uniqueOrders[orderId]) {
      uniqueOrders[orderId] = {
        orderNumber: orderId,
        customerName: row[1] || '',
        phone: row[3] || '',
        address: row[4] || '',
        department: row[5] || '',
        commune: row[6] || '',
        status: row[7] || '',
        van: vanColumnIndex !== -1 ? (row[vanColumnIndex] || '') : ''
      };
    }
  });

  return Object.values(uniqueOrders);
}

function saveSingleOrderChange(orderData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ordersSheet = ss.getSheetByName('Orders');
    if (!ordersSheet) {
      throw new Error('No se encontró la hoja "Orders".');
    }

    const headers = ordersSheet.getRange(1, 1, 1, ordersSheet.getLastColumn()).getValues()[0];
    const vanHeader = 'Furgón';
    let vanColumn = headers.indexOf(vanHeader) + 1;
    if (vanColumn === 0) {
      vanColumn = ordersSheet.getLastColumn() + 1;
      ordersSheet.getRange(1, vanColumn).setValue(vanHeader).setFontWeight('bold');
    }

    const orderNumbers = ordersSheet.getRange("A2:A" + ordersSheet.getLastRow()).getValues().flat().map(String);
    const rowIndex = orderNumbers.indexOf(String(orderData.orderNumber));

    if (rowIndex === -1) {
      Logger.log(`No se encontró el pedido #${orderData.orderNumber} para el auto-guardado.`);
      return { status: 'warning', message: `No se encontró el pedido ${orderData.orderNumber}.` };
    }

    const row = rowIndex + 2;
    ordersSheet.getRange(row, 5).setValue(orderData.address);
    ordersSheet.getRange(row, 6).setValue(orderData.department);
    ordersSheet.getRange(row, 7).setValue(orderData.commune);
    ordersSheet.getRange(row, vanColumn).setValue(orderData.van);

    return { status: 'success', message: `Pedido #${orderData.orderNumber} guardado.` };
  } catch (e) {
    Logger.log(e);
    return { status: 'error', message: e.toString() };
  }
}

function saveRouteChanges(updatedOrders) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ordersSheet = ss.getSheetByName('Orders');
  if (!ordersSheet) {
    throw new Error('No se encontró la hoja "Orders".');
  }

  const headers = ordersSheet.getRange(1, 1, 1, ordersSheet.getLastColumn()).getValues()[0];
  const vanHeader = 'Furgón';
  let vanColumn = headers.indexOf(vanHeader) + 1;
  if (vanColumn === 0) {
    vanColumn = ordersSheet.getLastColumn() + 1;
    ordersSheet.getRange(1, vanColumn).setValue(vanHeader).setFontWeight('bold');
  }

  const allOrderNumbers = ordersSheet.getRange("A2:A" + ordersSheet.getLastRow()).getValues().flat().map(String);

  updatedOrders.forEach(order => {
    const rowIndex = allOrderNumbers.indexOf(String(order.orderNumber));
    if (rowIndex !== -1) {
      const row = rowIndex + 2;
      ordersSheet.getRange(row, 5).setValue(order.address);
      ordersSheet.getRange(row, 6).setValue(order.department);
      ordersSheet.getRange(row, 7).setValue(order.commune);
      ordersSheet.getRange(row, vanColumn).setValue(order.van);
    }
  });

  return { status: 'success', message: 'Cambios guardados con éxito.' };
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


// --- LÓGICA DEL DASHBOARD ---

function showDashboard() {
  updateAcquisitionListAutomated(); // Actualiza la lista de adquisiciones automáticamente
  const html = HtmlService.createHtmlOutputFromFile('LauncherDialog')
    .setWidth(400)
    .setHeight(250);
  SpreadsheetApp.getUi().showModalDialog(html, 'Abrir Dashboard');
}

// --- FUNCIONES DE LA APLICACIÓN WEB ---

/**
 * Punto de entrada principal para la aplicación web. Sirve el HTML del dashboard.
 * @param {Object} e - El objeto de evento de la solicitud GET.
 * @returns {HtmlOutput} El contenido HTML para ser renderizado.
 */
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('DashboardDialog')
    .setTitle('Dashboard de Operaciones')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Devuelve la URL de la aplicación web implementada.
 * Esta función es llamada por el diálogo lanzador para saber qué URL abrir.
 * @returns {string} La URL de la aplicación web.
 */
function getWebAppUrl() {
  // Para que esto funcione, el script debe estar implementado como una aplicación web.
  // Ir a "Implementar" > "Nueva implementación", seleccionar "Aplicación web"
  // y asegurarse de que el acceso esté configurado como "Cualquier usuario" o según sea necesario.
  return ScriptApp.getService().getUrl();
}

function startDashboardRefresh() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ordersSheet = ss.getSheetByName('Orders');
  if (!ordersSheet) {
    throw new Error('No se encontró la hoja "Orders".');
  }
  const orderData = ordersSheet.getRange("A2:B" + ordersSheet.getLastRow()).getValues();
  const customerOrders = {};
  orderData.forEach(([orderNumber, customerName]) => {
    if (customerName) {
      if (!customerOrders[customerName]) customerOrders[customerName] = new Set();
      customerOrders[customerName].add(orderNumber);
    }
  });
  const duplicates = {};
  for (const customer in customerOrders) {
    if (customerOrders[customer].size > 1) {
      duplicates[customer] = Array.from(customerOrders[customer]);
    }
  }
  if (Object.keys(duplicates).length > 0) {
    showDuplicateDialog(duplicates);
  } else {
    checkForNewSuppliers();
  }
}

function showDuplicateDialog(duplicateData) {
  const template = HtmlService.createTemplateFromFile('DuplicateDialog');
  template.duplicates = JSON.stringify(duplicateData);
  const html = template.evaluate().setWidth(700).setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, 'Resolver Pedidos Duplicados');
}

function deleteOrdersByNumber(orderNumbersToDelete) {
  if (!orderNumbersToDelete || orderNumbersToDelete.length === 0) return "No se seleccionó ningún pedido para eliminar.";
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Orders');
  const data = sheet.getDataRange().getValues();
  const rowsToDelete = [];
  for (let i = data.length - 1; i >= 1; i--) {
    if (orderNumbersToDelete.includes(String(data[i][0]))) {
      rowsToDelete.push(i + 1);
    }
  }
  if (rowsToDelete.length > 0) {
    rowsToDelete.forEach(rowNum => sheet.deleteRow(rowNum));
    checkForNewSuppliers();
    return `Se eliminaron ${rowsToDelete.length} filas. Continuando con el chequeo de proveedores...`;
  } else {
    return "No se encontraron los pedidos seleccionados.";
  }
}

function checkForNewSuppliers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const skuSheet = ss.getSheetByName("SKU");
  const proveedoresSheet = ss.getSheetByName("Proveedores");
  if (!skuSheet || !proveedoresSheet) {
    SpreadsheetApp.getUi().alert("Faltan las hojas 'SKU' o 'Proveedores'.");
    return;
  }
  const skuSuppliers = new Set(skuSheet.getRange("I2:I" + skuSheet.getLastRow()).getValues().flat().filter(String));
  const existingSuppliers = new Set(proveedoresSheet.getRange("A2:A" + proveedoresSheet.getLastRow()).getValues().flat().filter(String));
  const newSuppliers = [...skuSuppliers].filter(s => !existingSuppliers.has(s));
  if (newSuppliers.length > 0) {
    showNewSupplierDialog(newSuppliers);
  } else {
    SpreadsheetApp.getUi().alert("Todos los datos están limpios y consistentes. Ahora puedes cargar las métricas en el dashboard.");
  }
}

function showNewSupplierDialog(newSuppliers) {
  const template = HtmlService.createTemplateFromFile('NewSupplierDialog');
  template.newSuppliers = JSON.stringify(newSuppliers);
  const html = template.evaluate().setWidth(600).setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(html, 'Añadir Teléfonos de Proveedores Nuevos');
}

function saveNewSuppliers(supplierData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const proveedoresSheet = ss.getSheetByName("Proveedores");
  const dataToAppend = Object.entries(supplierData);
  if (dataToAppend.length > 0) {
    proveedoresSheet.getRange(proveedoresSheet.getLastRow() + 1, 1, dataToAppend.length, 2).setValues(dataToAppend);
  }
  return "Proveedores guardados. Ya puedes cargar las métricas en el dashboard.";
}

// --- FUNCIONES PÚBLICAS PARA EL DASHBOARD ---

function getDashboardSummaryMetrics() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ordersSheet = ss.getSheetByName('Orders');

    // --- Lógica de costos duplicada para independencia ---
    const costosSheet = ss.getSheetByName('CostosVenta');
    const productCostMap = new Map();
    if (costosSheet && costosSheet.getLastRow() > 1) {
      const costosData = costosSheet.getRange("B2:C" + costosSheet.getLastRow()).getValues();
      for (let i = costosData.length - 1; i >= 0; i--) {
        const row = costosData[i];
        const productName = row[0];
        const cost = parseFloat(String(row[1]).replace(',', '.'));
        if (productName && !productCostMap.has(productName) && !isNaN(cost)) {
          productCostMap.set(productName, cost);
        }
      }
    }
    // --- Fin de lógica duplicada ---

    let totalSales = 0;
    let totalCosts = 0;
    const orderIds = new Set();

    if (ordersSheet && ordersSheet.getLastRow() > 1) {
        const ordersData = ordersSheet.getRange("A2:M" + ordersSheet.getLastRow()).getValues();
        ordersData.forEach(row => {
            const orderId = row[0];
            const productName = row[9]; // Columna J: Nombre Producto
            const quantity = row[10];
            const lineTotal = row[12];

            if (orderId) orderIds.add(orderId);
            if (lineTotal) totalSales += parseFloat(lineTotal) || 0;

            if (productName && quantity) {
                if (productCostMap.has(productName)) {
                    totalCosts += productCostMap.get(productName) * (parseInt(quantity, 10) || 0);
                } else {
                    // Fallback: use the line total for this item as its cost for margin calculation
                    totalCosts += parseFloat(String(row[12]).replace(',', '.')) || 0;
                }
            }
        });
    }

    const grossMargin = totalSales - totalCosts;
    const marginPercentage = totalSales > 0 ? (grossMargin / totalSales) * 100 : 0;

    return {
      totalSales: totalSales,
      orderCount: orderIds.size,
      grossMargin: grossMargin,
      marginPercentage: marginPercentage
    };
  } catch (e) {
    Logger.log(`ERROR en getDashboardSummaryMetrics: ${e.stack}`);
    return { error: `Error en Métricas de Resumen: ${e.message}` };
  }
}

function getDashboardCostMetrics() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ordersSheet = ss.getSheetByName('Orders');
    const costosSheet = ss.getSheetByName('CostosVenta');

    const productCostMap = new Map();
    if (costosSheet && costosSheet.getLastRow() > 1) {
      const costosData = costosSheet.getRange("B2:C" + costosSheet.getLastRow()).getValues();
      for (let i = costosData.length - 1; i >= 0; i--) {
        const row = costosData[i];
        const productName = row[0];
        const cost = parseFloat(String(row[1]).replace(',', '.'));
        if (productName && !productCostMap.has(productName) && !isNaN(cost)) {
          productCostMap.set(productName, cost);
        }
      }
    }

    let totalCosts = 0;
    if (ordersSheet && ordersSheet.getLastRow() > 1) {
        // Read up to column M (13) to include Line Total
        const ordersData = ordersSheet.getRange("A2:M" + ordersSheet.getLastRow()).getValues();
        const lineTotalCol = 12; // Column M
        
        ordersData.forEach(row => {
            const productName = row[9]; // Columna J: Nombre Producto
            const quantity = row[10];   // Columna K
            
            if (productName && quantity) {
                if (productCostMap.has(productName)) {
                    totalCosts += productCostMap.get(productName) * (parseInt(quantity, 10) || 0);
                } else {
                    // Fallback: use the line total for this item as its cost
                    const lineTotal = parseFloat(String(row[lineTotalCol]).replace(',', '.')) || 0;
                    totalCosts += lineTotal;
                }
            }
        });
    }

    return {
      totalCosts: totalCosts,
      productsWithoutCostCount: 0, // This is now 0 as we have a fallback
      productsWithoutCostNames: [] // This is now empty
    };
  } catch (e) {
    Logger.log(`ERROR en getDashboardCostMetrics: ${e.stack}`);
    return { error: `Error en Métricas de Costos: ${e.message}` };
  }
}

function getDashboardTopProducts() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ordersSheet = ss.getSheetByName('Orders');
    const productQuantities = {};

    if (ordersSheet && ordersSheet.getLastRow() > 1) {
        const ordersData = ordersSheet.getRange("J2:K" + ordersSheet.getLastRow()).getValues(); // Leer desde la columna J
        ordersData.forEach(row => {
            const productName = row[0]; // J es index 0 en este rango
            const quantity = row[1];    // K es index 1 en este rango
            if (productName && quantity) {
                productQuantities[productName] = (productQuantities[productName] || 0) + (parseInt(quantity, 10) || 0);
            }
        });
    }

    const topSoldProducts = Object.entries(productQuantities).sort(([, a], [, b]) => b - a).slice(0, 5);
    return topSoldProducts;
  } catch (e) {
    Logger.log(`ERROR en getDashboardTopProducts: ${e.stack}`);
    return { error: `Error en Top Productos: ${e.message}` };
  }
}

function getDashboardCommuneDistribution() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ordersSheet = ss.getSheetByName('Orders');
    const communeCounts = {};

    if (ordersSheet && ordersSheet.getLastRow() > 1) {
        const ordersData = ordersSheet.getRange("A2:G" + ordersSheet.getLastRow()).getValues();
        ordersData.forEach(row => {
            const orderId = row[0];     // A es index 0
            const commune = row[6];     // G es index 6
            if (commune) {
                const orderKey = `${orderId}-${commune}`;
                if (!communeCounts[orderKey]) {
                    communeCounts[orderKey] = commune;
                }
            }
        });
    }

    const communeTally = {};
    Object.values(communeCounts).forEach(c => communeTally[c] = (communeTally[c] || 0) + 1);
    const communeDistribution = Object.entries(communeTally).sort(([, a], [, b]) => b - a);
    return communeDistribution;
  } catch (e) {
    Logger.log(`ERROR en getDashboardCommuneDistribution: ${e.stack}`);
    return { error: `Error en Distribución por Comuna: ${e.message}` };
  }
}


function futureModulePlaceholder() {
  SpreadsheetApp.getUi().alert("Este módulo será implementado en una futura actualización.");
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

function showCategorySelectionDialog() {
  const html = HtmlService.createHtmlOutputFromFile('CategoryDialog').setWidth(500).setHeight(450);
  SpreadsheetApp.getUi().showModalDialog(html, 'Paso 2: Seleccionar Categorías para Envasado');
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
  let sheet = ss.getSheetByName("Lista de Envasado");
  if (sheet) { sheet.clear(); } else { sheet = ss.insertSheet("Lista de Envasado"); }
  let currentRow = 1;
  sheet.getRange("A1:C1").setValues([["Cantidad", "Inventario", "Nombre Producto"]]).setFontWeight("bold");
  selectedCategories.sort().forEach(category => {
    currentRow++;
    sheet.getRange(currentRow, 1, 1, 3).merge().setValue(category.toUpperCase()).setFontWeight("bold").setHorizontalAlignment("center").setBackground("#f2f2f2");
    currentRow++;
    const products = data[category].products;
    const sortedProductNames = Object.keys(products).sort();
    sortedProductNames.forEach(productName => {
      sheet.getRange(currentRow, 1).setValue(products[productName]);
      sheet.getRange(currentRow, 3).setValue(productName);
      currentRow++;
    });
  });
  sheet.autoResizeColumns(1, 3);
  const printUrl = `https://docs.google.com/spreadsheets/d/${ss.getId()}/export?format=pdf&gid=${sheet.getSheetId()}&portrait=true&fitw=true&gridlines=false&printtitle=false`;
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

function startNotificationProcess() {
  const html = HtmlService.createHtmlOutputFromFile('NotificationDialog').setWidth(600).setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(html, 'Panel de Notificación a Proveedores');
}

function getSupplierList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const acquisitionsSheet = ss.getSheetByName("Lista de Adquisiciones");
  if (!acquisitionsSheet || acquisitionsSheet.getLastRow() < 2) {
    return [];
  }
  const supplierData = acquisitionsSheet.getRange("L2:L" + acquisitionsSheet.getLastRow()).getValues();
  const suppliers = new Set();
  supplierData.forEach(row => {
    if (row[0]) {
      suppliers.add(String(row[0]).trim());
    }
  });
  return Array.from(suppliers).sort();
}

function getOrdersForSupplier(supplierName) {
  if (!supplierName) {
    throw new Error("Se requiere un nombre de proveedor.");
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const acquisitionsSheet = ss.getSheetByName("Lista de Adquisiciones");
  const proveedoresSheet = ss.getSheetByName("Proveedores");
  const skuSheet = ss.getSheetByName("SKU");

  if (!acquisitionsSheet) throw new Error("No se encuentra la hoja 'Lista de Adquisiciones'.");
  if (!proveedoresSheet) throw new Error("No se encuentra la hoja 'Proveedores'.");
  if (!skuSheet) throw new Error("No se encuentra la hoja 'SKU'.");

  const phoneMapProveedores = new Map();
  if (proveedoresSheet.getLastRow() > 1) {
    const phoneData = proveedoresSheet.getRange("A2:B" + proveedoresSheet.getLastRow()).getValues();
    phoneData.forEach(([name, phone]) => {
      if (name && phone) phoneMapProveedores.set(String(name).trim(), String(phone).trim());
    });
  }
  const phoneMapSku = new Map();
  if (skuSheet.getLastRow() > 1) {
    const skuSupplierData = skuSheet.getRange("I2:J" + skuSheet.getLastRow()).getValues();
    skuSupplierData.forEach(([supplier, phone]) => {
      if (supplier && phone) {
        const supName = String(supplier).trim();
        if (!phoneMapSku.has(supName)) phoneMapSku.set(supName, String(phone).trim());
      }
    });
  }
  const phone = phoneMapProveedores.get(supplierName) || phoneMapSku.get(supplierName) || 'No encontrado';

  const orders = [];
  if (acquisitionsSheet.getLastRow() > 1) {
    const allData = acquisitionsSheet.getRange("A2:L" + acquisitionsSheet.getLastRow()).getValues();
    allData.forEach(row => {
      const [product, quantity, format, , , , , , , , , supplier] = row;
      if (supplier && String(supplier).trim() === supplierName) {
        if (product && quantity && parseFloat(String(quantity).replace(',', '.')) !== 0) {
          orders.push({
            product: String(product).trim(),
            quantity: quantity,
            format: String(format).trim()
          });
        }
      }
    });
  }

  return {
    phone: phone,
    orders: orders
  };
}

// --- MÓDULO DE ANÁLISIS DE PRECIOS ---

/**
 * Procesa la hoja "Reporte Adquisiciones" para registrar las compras diarias
 * en la hoja "Historico Adquisiciones", evitando duplicados.
 * Aplica la lógica de corrección de compras si los datos correspondientes están presentes.
 */
function procesarReporteAdquisiciones() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const reporteSheet = ss.getSheetByName("Reporte Adquisiciones");
  if (!reporteSheet) {
    ui.alert('Error: No se encontró la hoja "Reporte Adquisiciones".');
    return;
  }

  const historicoSheet = ss.getSheetByName("Historico Adquisiciones");
  if (!historicoSheet) {
    ui.alert('Error: No se encontró la hoja "Historico Adquisiciones".');
    return;
  }

  try {
    const reporteData = reporteSheet.getDataRange().getValues();
    reporteData.shift(); // Quitar encabezados

    const historicoData = historicoSheet.getDataRange().getValues();
    historicoData.shift(); // Quitar encabezados

    const productoBaseRowMap = new Map();
    historicoData.forEach((row, index) => {
      const productoBase = row[2]; // Columna C
      if (productoBase) {
        productoBaseRowMap.set(productoBase.toString().trim(), index + 2);
      }
    });

    let updatedCount = 0;
    let appendedCount = 0;
    const today = new Date();

    reporteData.forEach(row => {
      const ID_COL = 12;
      const ID = row[ID_COL] ? row[ID_COL].toString().trim() : null;
      if (!ID) return;

      const productoBase = row[1].toString().trim(); // Columna B en Reporte

      let cantidadReal, formatoReal;
      const CORRECCION_CANTIDAD_COL = 8;
      const CORRECCION_FORMATO_COL = 9;
      if (row[CORRECCION_CANTIDAD_COL] && row[CORRECCION_CANTIDAD_COL] !== "") {
        cantidadReal = row[CORRECCION_CANTIDAD_COL];
        formatoReal = row[CORRECCION_FORMATO_COL] || '';
      } else {
        cantidadReal = row[3]; // CANTIDAD_COMPRA_COL
        formatoReal = row[2]; // FORMATO_COMPRA_COL
      }

      const newRowData = [
        ID,                         // Col A: ID
        today,                      // Col B: Fecha de Registro
        productoBase,               // Col C: Producto Base
        formatoReal,                // Col D: Formato de Compra
        cantidadReal,               // Col E: Cantidad Comprada
        row[4],                     // Col F: Precio Compra
        row[5],                     // Col G: Costo Total Compra
        row[6]                      // Col H: Proveedor
      ];

      const rowIndex = productoBaseRowMap.get(productoBase);
      if (rowIndex) {
        historicoSheet.getRange(rowIndex, 1, 1, newRowData.length).setValues([newRowData]);
        updatedCount++;
      } else {
        historicoSheet.appendRow(newRowData);
        appendedCount++;
        productoBaseRowMap.set(productoBase, historicoSheet.getLastRow());
      }
    });

    if (updatedCount > 0 || appendedCount > 0) {
      ui.alert(`Proceso completado. Se actualizaron ${updatedCount} productos y se añadieron ${appendedCount} productos nuevos al historial de precios.`);
    } else {
      ui.alert('No se encontraron nuevas adquisiciones para procesar en el "Reporte Adquisiciones".');
    }

  } catch (e) {
    Logger.log(e);
    ui.alert(`Ha ocurrido un error durante el procesamiento: ${e.message}`);
  }
}

/**
 * Limpia y recalcula la hoja 'CostosVenta' para todos los productos vendidos en el día actual.
 * Se basa en los últimos precios de la hoja 'Historico Adquisiciones'.
 * Diseñada para ser ejecutada diariamente por un trigger.
 */
function actualizarCostosDeVentaDiarios() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    // 1. Obtener todas las hojas y datos necesarios
    const historicoSheet = ss.getSheetByName("Historico Adquisiciones");
    const skuSheet = ss.getSheetByName("SKU");
    const ordersSheet = ss.getSheetByName("Orders");
    const costosSheet = ss.getSheetByName("CostosVenta");

    if (!historicoSheet || !skuSheet || !ordersSheet || !costosSheet) {
      Logger.log("Error: Falta una o más hojas requeridas para la actualización diaria de costos.");
      return;
    }

    // 2. Crear mapas de búsqueda para un acceso eficiente
    // Mapa de SKU: Nombre Producto -> { productoBase, cantidadVenta }
    const skuMap = new Map();
    skuSheet.getDataRange().getValues().slice(1).forEach(row => {
      const nombreProducto = row[0];
      const productoBase = row[1];
      const cantidadVenta = parseFloat(String(row[6]).replace(',', '.')) || 0;
      if (nombreProducto && productoBase) {
        skuMap.set(nombreProducto, { productoBase, cantidadVenta });
      }
    });

    // Mapa de Precios: Producto Base -> { precioUnitario, formato }
    const priceMap = new Map();
    historicoSheet.getDataRange().getValues().slice(1).forEach(row => {
      const productoBase = row[2]; // Col C
      const formato = row[3];      // Col D
      const precioUnitario = parseFloat(String(row[5]).replace(',', '.')) || 0; // Col F - 'Precio Compra'
      if (productoBase) {
        priceMap.set(productoBase, { precioUnitario, formato });
      }
    });

    // 3. Obtener todos los productos únicos de la hoja Orders (se asume que todos son para procesar hoy)
    const productsSoldToday = new Set();
    ordersSheet.getDataRange().getValues().slice(1).forEach(row => {
      const productName = row[9]; // Columna J: Nombre Producto
      if (productName) {
        productsSoldToday.add(productName);
      }
    });

    if (productsSoldToday.size === 0) {
      Logger.log("No se vendieron productos hoy. La hoja de CostosVenta no se actualizará.");
      costosSheet.getRange("A2:C" + Math.max(costosSheet.getMaxRows(), 2)).clearContent(); // Limpiar la hoja si no hay ventas
      return;
    }

    // 4. Calcular los costos para los productos vendidos hoy
    const newCostosData = [];
    productsSoldToday.forEach(productName => {
      const skuInfo = skuMap.get(productName);
      if (!skuInfo) return;

      const priceInfo = priceMap.get(skuInfo.productoBase);
      if (!priceInfo) return;

      // Extraer el tamaño del formato (lógica robusta)
      const match = priceInfo.formato.toString().match(/\(([\d.,]+)/);
      let formatoSize = 1;
      if (match && match[1]) {
        const cleanedString = match[1].replace(/\./g, '').replace(',', '.');
        const parsedSize = parseFloat(cleanedString);
        if (!isNaN(parsedSize) && parsedSize > 0) {
          formatoSize = parsedSize;
        }
      }

      const costoPorUnidadBase = priceInfo.precioUnitario / formatoSize;
      const costoFinal = costoPorUnidadBase * skuInfo.cantidadVenta;

      if (!isNaN(costoFinal)) {
        newCostosData.push([new Date(), productName, costoFinal]);
      }
    });

    // 5. Limpiar la hoja y escribir los nuevos datos
    costosSheet.getRange("A2:C" + Math.max(costosSheet.getMaxRows(), 2)).clearContent();
    if (newCostosData.length > 0) {
      costosSheet.getRange(2, 1, newCostosData.length, 3).setValues(newCostosData);
    }

    Logger.log(`Hoja 'CostosVenta' actualizada con ${newCostosData.length} productos.`);

  } catch (e) {
    Logger.log(`Error en la actualización diaria de costos: ${e.stack}`);
  }
}

/**
 * Orquesta el proceso completo de análisis de precios,
 * llamando primero al procesamiento de adquisiciones y luego al cálculo de costos.
 */
function runPriceAnalysis() {
  try {
    procesarReporteAdquisiciones();
    const analysisData = getAnalysisData();
    showPriceApprovalDashboard(analysisData);
  } catch (e) {
    Logger.log(e);
    SpreadsheetApp.getUi().alert("Ocurrió un error en el proceso de análisis: " + e.message);
  }
}

/**
 * Analiza las adquisiciones del día, calcula los costos de venta y detecta anomalías.
 * @returns {{allCosts: Array<Array<any>>, anomalies: Array<object>}} Un objeto que contiene todos los costos calculados y una lista de anomalías detectadas.
 */
function getAnalysisData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const DEVIATION_THRESHOLD = 2.5;

  const historicoSheet = ss.getSheetByName("Historico Adquisiciones");
  const skuSheet = ss.getSheetByName("SKU");
  const costosSheet = ss.getSheetByName("CostosVenta");
  const ordersSheet = ss.getSheetByName("Orders");

  if (!historicoSheet || !skuSheet || !costosSheet || !ordersSheet) {
    throw new Error("Faltan una o más hojas requeridas. Ejecute el setup inicial.");
  }

  const historicoData = historicoSheet.getDataRange().getValues();
  const skuData = skuSheet.getDataRange().getValues();
  const costosData = costosSheet.getDataRange().getValues();
  const ordersData = ordersSheet.getDataRange().getValues();

  historicoData.shift();
  skuData.shift();
  costosData.shift();
  ordersData.shift();

  // Crear mapa de precios de venta (el más reciente por producto)
  const salePrices = {};
  const ORDERS_ITEM_NAME_COL = 8;
  const ORDERS_ITEM_PRICE_COL = 11;
  for (let i = ordersData.length - 1; i >= 0; i--) {
      const row = ordersData[i];
      const productName = row[ORDERS_ITEM_NAME_COL];
      const price = parseFloat(row[ORDERS_ITEM_PRICE_COL]);
      if (productName && !salePrices[productName] && price > 0) {
          salePrices[productName] = price;
      }
  }

  const skuMap = {};
  skuData.forEach(row => {
    const productoBase = row[1];
    if (!productoBase) return;
    if (!skuMap[productoBase]) skuMap[productoBase] = [];
    skuMap[productoBase].push({
      nombreProducto: row[0],
      cantidadVenta: parseFloat(String(row[6]).replace(',', '.')) || 0,
      unidadVenta: normalizeUnit(row[7])
    });
  });

  const historicalCosts = {};
  costosData.forEach(row => {
    const productName = row[1];
    const cost = parseFloat(row[2]);
    if (!productName || isNaN(cost)) return;
    if (!historicalCosts[productName]) historicalCosts[productName] = [];
    historicalCosts[productName].push(cost);
  });

  const analysisResults = [];
  const today = new Date();
  const todayString = today.toDateString();
  const processedProducts = new Set();

  const todayAcquisitions = historicoData.filter(acq => new Date(acq[1]).toDateString() === todayString);

  todayAcquisitions.forEach(acq => {
    const productoBase = acq[2];
    const formato = acq[3];
    const precioUnitario = parseFloat(String(acq[6]).replace(',', '.')) || 0;

    const match = formato.toString().match(/\(([\d.,]+)/);
    let formatoSize = 1;
    if (match && match[1]) {
      // Elimina los separadores de miles (puntos) y luego reemplaza la coma decimal por un punto.
      const cleanedString = match[1].replace(/\./g, '').replace(',', '.');
      const parsedSize = parseFloat(cleanedString);
      if (!isNaN(parsedSize) && parsedSize > 0) {
        formatoSize = parsedSize;
      }
    }
    const costoPorUnidadBase = precioUnitario / formatoSize;

    if (skuMap[productoBase]) {
      skuMap[productoBase].forEach(sku => {
        if (processedProducts.has(sku.nombreProducto)) return;

        const costoFinal = costoPorUnidadBase * sku.cantidadVenta;
        const history = historicalCosts[sku.nombreProducto] || [];
        const stats = calculateStats(history);

        let status = 'ok';
        let deviationLevel = 0;
        if (history.length < 2) {
            status = 'new';
        } else if (stats.stdDev > 0) {
            deviationLevel = (costoFinal - stats.mean) / stats.stdDev;
            if (Math.abs(deviationLevel) > DEVIATION_THRESHOLD) {
                status = 'anomaly';
            }
        }

        analysisResults.push({
          nombreProducto: sku.nombreProducto,
          costoHoy: costoFinal,
          status: status,
          productoBase: productoBase,
          formatoCompra: formato,
          proveedor: acq[8],
          precioUnitarioCompra: precioUnitario,
          costoPromedio: stats.mean,
          desviacionEstandar: stats.stdDev,
          nivelDesviacion: deviationLevel,
          precioVenta: salePrices[sku.nombreProducto] || 0
        });
        processedProducts.add(sku.nombreProducto);
      });
    }
  });

  return analysisResults;
}

/**
 * Muestra un diálogo modal con las anomalías de precios para su aprobación.
 * @param {object} analysisData - El objeto que contiene allCosts y anomalies.
 */
function showPriceApprovalDashboard(analysisData) {
  const template = HtmlService.createTemplateFromFile('PriceApprovalDialog');
  template.analysisResults = analysisData;

  const html = template.evaluate()
      .setWidth(900)
      .setHeight(600);

  SpreadsheetApp.getUi().showModalDialog(html, 'Revisar y Aprobar Precios de Adquisición');
}

/**
 * Guarda los datos de costos y anomalías aprobados en sus respectivas hojas.
 * Esta función es llamada desde el dashboard de aprobación.
 * @param {object} data - Un objeto que contiene las listas 'costs' y 'anomalies'.
 * @returns {string} Un mensaje de confirmación.
 */
function commitPriceData(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const costosSheet = ss.getSheetByName("CostosVenta");
  const anomaliasSheet = ss.getSheetByName("Anomalías de Precios");
  const today = new Date();
  const todayString = today.toDateString();

  try {
    // 1. Preparar los datos para escribir
    const costsToWrite = data.map(item => [today, item.nombreProducto, item.costoHoy]);

    const anomaliesToWrite = data
      .filter(item => item.status === 'anomaly')
      .map(a => [
        today, a.nombreProducto, a.costoHoy, a.costoPromedio, a.desviacionEstandar, a.nivelDesviacion,
        `Costo de hoy (${a.costoHoy.toFixed(2)}) se desvía en ${a.nivelDesviacion.toFixed(2)} stddevs.`
      ]);

    // 2. Guardar los costos de venta
    if (costsToWrite.length > 0) {
      const allCostosData = costosSheet.getDataRange().getValues();
      const rowsToDeleteCosts = [];
      allCostosData.forEach((row, index) => {
        if (index > 0 && new Date(row[0]).toDateString() === todayString) {
          rowsToDeleteCosts.push(index + 1);
        }
      });
      for (let i = rowsToDeleteCosts.length - 1; i >= 0; i--) {
        costosSheet.deleteRow(rowsToDeleteCosts[i]);
      }
      costosSheet.getRange(costosSheet.getLastRow() + 1, 1, costsToWrite.length, 3).setValues(costsToWrite);
    }

    // 3. Guardar las anomalías
    const allAnomaliasData = anomaliasSheet.getDataRange().getValues();
    const rowsToDeleteAnomalies = [];
    allAnomaliasData.forEach((row, index) => {
      if (index > 0 && new Date(row[0]).toDateString() === todayString) {
        rowsToDeleteAnomalies.push(index + 1);
      }
    });
    for (let i = rowsToDeleteAnomalies.length - 1; i >= 0; i--) {
      anomaliasSheet.deleteRow(rowsToDeleteAnomalies[i]);
    }

    if (anomaliesToWrite.length > 0) {
      anomaliasSheet.getRange(anomaliasSheet.getLastRow() + 1, 1, anomaliesToWrite.length, 7).setValues(anomaliesToWrite);
    }

    return "Los precios han sido aprobados y guardados correctamente.";

  } catch (e) {
    Logger.log(e);
    throw new Error("Ocurrió un error al guardar los datos: " + e.message);
  }
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

// --- NORMALIZACIÓN DE PRODUCTOS ---

/**
 * Shows the dialog for the automated inconsistency report.
 */
function showInconsistencyReportDialog() {
  try {
    const html = HtmlService.createHtmlOutputFromFile('InconsistencyReportDialog')
      .setWidth(600)
      .setHeight(400);
    SpreadsheetApp.getUi().showModalDialog(html, 'Reporte de Inconsistencias');
  } catch (e) {
    Logger.log(e);
    SpreadsheetApp.getUi().alert(`Error al abrir el reporte: ${e.message}`);
  }
}

/**
 * Runs a consistency check on all base products and returns a list of those with inconsistencies.
 * @returns {string[]} An array of product base names that have inconsistent data.
 */
function runFullInconsistencyCheck() {
  const allProducts = getNormalizedBaseProducts();
  if (!allProducts || allProducts.length === 0) {
    return [];
  }

  const inconsistentProducts = [];
  allProducts.forEach(productName => {
    try {
      const inconsistencies = getInconsistenciesForProduct(productName);
      if (Object.keys(inconsistencies).length > 0) {
        inconsistentProducts.push(productName);
      }
    } catch (e) {
      // Log error for a specific product but continue the check for others
      Logger.log(`Error checking inconsistencies for product "${productName}": ${e.message}`);
    }
  });

  return inconsistentProducts;
}

/**
 * Shows the dialog for normalizing product categories.
 */
function showCategoryNormalizerDialog() {
  try {
    const html = HtmlService.createHtmlOutputFromFile('CategoryNormalizerDialog')
      .setWidth(800)
      .setHeight(600);
    SpreadsheetApp.getUi().showModalDialog(html, 'Normalizador de Categorías');
  } catch (e) {
    Logger.log(e);
    SpreadsheetApp.getUi().alert(`Error al abrir el normalizador de categorías: ${e.message}`);
  }
}

/**
 * Applies the category normalization mapping to the 'Categoría' column in the 'SKU' sheet.
 * @param {Object.<string, string>} fixes - An object mapping old category names to new normalized names.
 * @returns {string} A success message.
 */
function applyCategoryFixes(fixes) {
  if (!fixes || Object.keys(fixes).length === 0) {
    return "No se proporcionaron cambios para aplicar.";
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const skuSheet = ss.getSheetByName("SKU");
  if (!skuSheet) {
    throw new Error("No se encontró la hoja 'SKU'.");
  }

  const range = skuSheet.getDataRange();
  const values = range.getValues();
  const headers = values[0];

  const categoriaColIndex = headers.indexOf("Categoría");
  if (categoriaColIndex === -1) {
    throw new Error("No se encontró la columna 'Categoría' en la hoja 'SKU'.");
  }

  let updatedCells = 0;
  // Start from row 1 to skip headers
  for (let i = 1; i < values.length; i++) {
    const currentCategory = values[i][categoriaColIndex];
    if (fixes.hasOwnProperty(currentCategory)) {
      values[i][categoriaColIndex] = fixes[currentCategory];
      updatedCells++;
    }
  }

  if (updatedCells > 0) {
    range.setValues(values);
    return `¡Éxito! Se actualizaron ${updatedCells} celdas de categoría.`;
  } else {
    return "No se encontraron categorías que necesitaran actualización.";
  }
}

/**
 * Analyzes the 'SKU' sheet to find category variations and products without a category.
 * @returns {{variations: Object.<string, string[]>, productsWithoutCategory: string[]}}
 *   An object containing grouped category variations and a list of products missing a category.
 */
function getCategoryVariations() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const skuSheet = ss.getSheetByName("SKU");
  if (!skuSheet) {
    throw new Error("No se encontró la hoja 'SKU'.");
  }

  const lastRow = skuSheet.getLastRow();
  if (lastRow < 2) {
    return { variations: {}, productsWithoutCategory: [] };
  }

  // Read "Nombre Producto" (col A) and "Categoría" (col F)
  const dataRange = skuSheet.getRange("A2:F" + lastRow);
  const values = dataRange.getValues();

  const productsWithoutCategory = [];
  const categoryGroups = {};

  values.forEach(row => {
    const productName = row[0];
    const category = row[5]; // Column F

    if (productName && (!category || String(category).trim() === '')) {
      productsWithoutCategory.push(productName);
    } else if (category && typeof category === 'string' && category.trim() !== '') {
      const trimmedCategory = category.trim();
      // Normalize by making it lowercase. This will group "Frutas" and "frutas".
      const normalizedKey = trimmedCategory.toLowerCase();

      if (!categoryGroups[normalizedKey]) {
        categoryGroups[normalizedKey] = [];
      }
      // Add the original variation to the group if it's not already there
      if (categoryGroups[normalizedKey].indexOf(trimmedCategory) === -1) {
        categoryGroups[normalizedKey].push(trimmedCategory);
      }
    }
  });

  // Filter out groups that only have one variation, as they don't need normalization.
  const finalVariations = {};
  for (const key in categoryGroups) {
    if (categoryGroups[key].length > 1) {
      // Suggest the first variation as the standard name
      const suggestion = categoryGroups[key][0];
      finalVariations[suggestion] = categoryGroups[key];
    }
  }

  return {
    variations: finalVariations,
    productsWithoutCategory: productsWithoutCategory
  };
}

/**
 * Bridge function to open the consistency checker for a specific product.
 * Called from the inconsistency report dialog.
 * @param {string} productName - The name of the product to check.
 */
function openCheckerForProduct(productName) {
  showConsistencyCheckerDialog(productName);
}

/**
 * Shows the dialog for checking data consistency for a base product.
 * @param {string} [productNameToSelect] - Optional. A product name to pre-select in the dialog.
 */
function showConsistencyCheckerDialog(productNameToSelect) {
  try {
    const template = HtmlService.createTemplateFromFile('ConsistencyCheckerDialog');
    template.productNameToSelect = productNameToSelect || null;
    const html = template.evaluate()
      .setWidth(700)
      .setHeight(600);
    SpreadsheetApp.getUi().showModalDialog(html, 'Verificador de Coherencia');
  } catch (e) {
    Logger.log(e);
    SpreadsheetApp.getUi().alert(`Error al abrir el verificador: ${e.message}`);
  }
}

/**
 * Shows the dialog for normalizing "Producto Base" names.
 */
function showNormalizationDialog() {
  try {
    const variations = getProductoBaseVariations();
    const template = HtmlService.createTemplateFromFile('NormalizationDialog');
    template.variations = JSON.stringify(variations || {});
    const html = template.evaluate().setWidth(700).setHeight(500);
    SpreadsheetApp.getUi().showModalDialog(html, 'Normalizar Productos Base');
  } catch (e) {
    Logger.log(e);
    SpreadsheetApp.getUi().alert(`Error al abrir el normalizador: ${e.message}`);
  }
}

/**
 * Gets all unique "Producto Base" variations and groups them by a normalized name.
 * @returns {Object.<string, string[]>} An object where keys are the suggested normalized names
 *   and values are arrays of the variations found.
 */
function getProductoBaseVariations() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const skuSheet = ss.getSheetByName("SKU");
  if (!skuSheet) {
    throw new Error("No se encontró la hoja 'SKU'.");
  }

  const lastRow = skuSheet.getLastRow();
  if (lastRow < 2) {
    return {}; // No data to process
  }

  const productoBaseData = skuSheet.getRange(2, 2, lastRow - 1, 1).getValues().flat();

  const groupedVariations = {};

  productoBaseData.forEach(name => {
    if (!name || typeof name !== 'string' || name.trim() === '') {
      return; // Skip empty or invalid names
    }
    const trimmedName = name.trim();
    // Simple normalization: lowercase and remove plural 's' at the end.
    // This can be improved with more complex heuristics if needed.
    const normalizedKey = trimmedName.toLowerCase().replace(/s$/, "").replace(/es$/, "");

    if (!groupedVariations[normalizedKey]) {
      groupedVariations[normalizedKey] = [];
    }
    // Add the original name to the group if it's not already there
    if (groupedVariations[normalizedKey].indexOf(trimmedName) === -1) {
      groupedVariations[normalizedKey].push(trimmedName);
    }
  });

  // Filter out groups that only have one variation, as they don't need normalization.
  const result = {};
  for (const key in groupedVariations) {
    if (groupedVariations[key].length > 1) {
      // We can make the key (suggestion) more presentable, e.g., capitalize it.
      const suggestion = key.charAt(0).toUpperCase() + key.slice(1);
      result[suggestion] = groupedVariations[key];
    }
  }

  return result;
}

/**
 * Applies the normalization mapping to the 'Producto Base' column in the 'SKU' sheet.
 * @param {Object.<string, string>} normalizationMap - An object mapping old names to new normalized names.
 * @returns {string} A success message.
 */
function applyNormalization(normalizationMap) {
  if (!normalizationMap || Object.keys(normalizationMap).length === 0) {
    return "No se proporcionaron cambios para aplicar.";
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const skuSheet = ss.getSheetByName("SKU");
  if (!skuSheet) {
    throw new Error("No se encontró la hoja 'SKU'.");
  }

  const range = skuSheet.getDataRange();
  const values = range.getValues();
  const headers = values[0];

  // Find the 'Producto Base' column index dynamically
  const productoBaseColIndex = headers.indexOf("Producto Base");
  if (productoBaseColIndex === -1) {
    throw new Error("No se encontró la columna 'Producto Base' en la hoja 'SKU'.");
  }

  let updatedRows = 0;
  // Start from row 1 to skip headers
  for (let i = 1; i < values.length; i++) {
    const currentName = values[i][productoBaseColIndex];
    if (normalizationMap.hasOwnProperty(currentName)) {
      values[i][productoBaseColIndex] = normalizationMap[currentName];
      updatedRows++;
    }
  }

  if (updatedRows > 0) {
    range.setValues(values);
    return `¡Éxito! Se actualizaron ${updatedRows} filas en la hoja SKU.`;
  } else {
    return "No se encontraron filas que necesitaran actualización.";
  }
}

/**
 * Gets a sorted list of unique "Producto Base" names from the SKU sheet.
 * @returns {string[]} A sorted array of unique product base names.
 */
function getNormalizedBaseProducts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const skuSheet = ss.getSheetByName("SKU");
  if (!skuSheet) {
    throw new Error("No se encontró la hoja 'SKU'.");
  }
  const lastRow = skuSheet.getLastRow();
  if (lastRow < 2) {
    return [];
  }
  const productoBaseData = skuSheet.getRange(2, 2, lastRow - 1, 1).getValues().flat();
  const uniqueNames = [...new Set(productoBaseData)];
  return uniqueNames.filter(name => name && typeof name === 'string' && name.trim() !== '').sort();
}

/**
 * Finds inconsistencies in specified columns for a given "Producto Base".
 * @param {string} productBaseName - The name of the base product to check.
 * @returns {Object.<string, string[]>} An object where keys are column names with inconsistencies
 *   and values are the array of different values found.
 */
function getInconsistenciesForProduct(productBaseName) {
  if (!productBaseName) {
    throw new Error("Se requiere un nombre de producto base.");
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const skuSheet = ss.getSheetByName("SKU");
  if (!skuSheet) {
    throw new Error("No se encontró la hoja 'SKU'.");
  }

  const data = skuSheet.getDataRange().getValues();
  const headers = data.shift(); // Get and remove headers

  const productoBaseColIndex = headers.indexOf("Producto Base");
  if (productoBaseColIndex === -1) {
    throw new Error("No se encontró la columna 'Producto Base'.");
  }

  // Define columns to check for consistency.
  const columnsToCheck = ["Formato Compra", "Unidad Compra", "Categoría", "Unidad Venta", "Proveedor"];
  const columnIndices = columnsToCheck.map(colName => headers.indexOf(colName));

  // Filter rows that match the selected productBaseName
  const productRows = data.filter(row => row[productoBaseColIndex] === productBaseName);

  if (productRows.length <= 1) {
    return {}; // No inconsistencies possible with 0 or 1 row
  }

  const inconsistencies = {};

  columnIndices.forEach((colIndex, i) => {
    if (colIndex === -1) return; // Skip if column doesn't exist

    const columnName = columnsToCheck[i];
    const values = new Set(productRows.map(row => String(row[colIndex]).trim()));

    if (values.size > 1) {
      inconsistencies[columnName] = Array.from(values);
    }
  });

  return inconsistencies;
}

/**
 * Applies consistency fixes to all rows of a given base product in the SKU sheet.
 * @param {string} productBaseName - The base product to update.
 * @param {Object.<string, string>} fixes - An object mapping column names to their new, unified value.
 * @returns {string} A success message.
 */
function applyConsistencyFixes(productBaseName, fixes) {
  if (!productBaseName || !fixes || Object.keys(fixes).length === 0) {
    throw new Error("No se proporcionaron suficientes datos para aplicar las correcciones.");
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const skuSheet = ss.getSheetByName("SKU");
  if (!skuSheet) {
    throw new Error("No se encontró la hoja 'SKU'.");
  }

  const range = skuSheet.getDataRange();
  const values = range.getValues();
  const headers = values[0];

  const productoBaseColIndex = headers.indexOf("Producto Base");
  if (productoBaseColIndex === -1) {
    throw new Error("No se encontró la columna 'Producto Base'.");
  }

  const fixColumns = {};
  for (const colName in fixes) {
    const colIndex = headers.indexOf(colName);
    if (colIndex === -1) {
      throw new Error(`La columna a corregir '${colName}' no fue encontrada.`);
    }
    fixColumns[colName] = colIndex;
  }

  let updatedRows = 0;
  // Start from 1 to skip header row
  for (let i = 1; i < values.length; i++) {
    if (values[i][productoBaseColIndex] === productBaseName) {
      for (const colName in fixes) {
        const colIndex = fixColumns[colName];
        values[i][colIndex] = fixes[colName];
      }
      updatedRows++;
    }
  }

  if (updatedRows > 0) {
    range.setValues(values);
    return `¡Éxito! Se actualizaron ${updatedRows} filas para el producto '${productBaseName}'.`;
  } else {
    return "No se encontraron filas para actualizar para el producto especificado.";
  }
}


// --- FUNCIONES AUXILIARES ---

/**
 * Calcula la media y la desviación estándar de una población de un array de números.
 * @param {number[]} data - Un array de números.
 * @returns {{mean: number, stdDev: number}} Un objeto con la media y la desviación estándar.
 */
function calculateStats(data) {
  if (!data || data.length === 0) {
    return { mean: 0, stdDev: 0 };
  }

  const n = data.length;
  const mean = data.reduce((a, b) => a + b, 0) / n;

  if (n < 2) {
    return { mean: mean, stdDev: 0 };
  }

  const variance = data.reduce((a, b) => a + Math.pow(b - mean, 2), 0) / n;
  const stdDev = Math.sqrt(variance);

  return { mean: mean, stdDev: stdDev };
}

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

/**
 * Crea o actualiza los disparadores de tiempo necesarios para el proyecto.
 */
function setupTriggers() {
  const functionName = "actualizarCostosDeVentaDiarios";

  // Eliminar triggers antiguos para la misma función para evitar duplicados.
  const allTriggers = ScriptApp.getProjectTriggers();
  let triggerExists = false;
  for (const trigger of allTriggers) {
    if (trigger.getHandlerFunction() === functionName) {
      triggerExists = true;
      // Opcional: podrías borrar y recrear si necesitas cambiar la configuración.
      // Por ahora, si ya existe, asumimos que está bien configurado.
    }
  }

  // Si no existe un disparador para esta función, crear uno nuevo.
  if (!triggerExists) {
    ScriptApp.newTrigger(functionName)
      .timeBased()
      .everyDays(1)
      .atHour(2) // Se ejecuta todos los días entre las 2 y 3 AM.
      .create();
    Logger.log(`Disparador diario para '${functionName}' creado.`);
  } else {
    Logger.log(`El disparador para '${functionName}' ya existe.`);
  }
}
