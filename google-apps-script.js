/**
 * ==============================================================================
 * CL'S BROWNIES — Google Sheets Backend Apps Script
 * ==============================================================================
 * 
 * INSTRUCTIONS TO DEPLOY:
 * 1. Open Google Sheets (https://sheets.new)
 * 2. Rename your spreadsheet: "CL'S BROWNIES — Orders Database"
 * 3. In the top menu, click: Extensions > Apps Script
 * 4. Erase any code inside Code.gs, and PASTE THIS ENTIRE FILE.
 * 5. Click the "Save" icon (Cmd+S or Ctrl+S).
 * 6. Click the blue "Deploy" button (top right) > "New deployment".
 * 7. Click the gear icon next to "Select type" > choose "Web app".
 * 8. Set the configuration:
 *      - Description: "CL Brownies Orders API"
 *      - Execute as: "Me" (your email)
 *      - Who has access: "Anyone"   <-- CRITICAL! Must be "Anyone" so the site can write orders!
 * 9. Click "Deploy". Grant permissions if prompted (Advanced > Go to Untitled project).
 * 10. COPY the "Web app URL" (it will end in /exec).
 * 11. Paste that URL into the Admin Dashboard Settings or configure it in the app!
 * ==============================================================================
 */

const SHEET_NAME = "Orders";

/**
 * Handle HTTP GET Requests
 * Used to fetch orders and test connectivity
 */
function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) ? e.parameter.action : "getOrders";

  if (action === "ping") {
    return createJsonResponse({
      status: "success",
      message: "CL's Brownies Google Sheets API is online and healthy!",
      timestamp: new Date().toISOString()
    });
  }

  if (action === "getOrders") {
    const orders = fetchAllOrders();
    return createJsonResponse({
      status: "success",
      count: orders.length,
      orders: orders
    });
  }

  if (action === "createOrder") {
    let order = {};
    if (e.parameter && e.parameter.data) {
      try {
        order = JSON.parse(e.parameter.data);
      } catch (err) {
        order = e.parameter;
      }
    } else if (e && e.parameter) {
      order = e.parameter;
    }
    const result = appendNewOrder(order);
    return createJsonResponse({
      status: "success",
      message: "Order successfully recorded in Google Sheets!",
      orderId: result.orderId,
      row: result.row
    });
  }

  if (action === "updateOrder") {
    const orderId = e && e.parameter ? e.parameter.orderId : null;
    let updates = {};
    if (e.parameter && e.parameter.updates) {
      try { updates = JSON.parse(e.parameter.updates); } catch (err) { updates = {}; }
    } else if (e && e.parameter) {
      if (e.parameter.status) updates.status = e.parameter.status;
      if (e.parameter.paymentStatus) updates.paymentStatus = e.parameter.paymentStatus;
    }
    const success = updateOrderRow(orderId, updates);
    return createJsonResponse({
      status: success ? "success" : "error",
      message: success ? "Order status updated." : "Order ID not found."
    });
  }

  if (action === "deleteOrder") {
    const orderId = e && e.parameter ? e.parameter.orderId : null;
    const success = deleteOrderRow(orderId);
    return createJsonResponse({
      status: success ? "success" : "error",
      message: success ? "Order deleted." : "Order ID not found."
    });
  }

  return createJsonResponse({
    status: "error",
    message: "Unknown action parameter: " + action
  });
}

/**
 * Handle HTTP POST Requests
 * Used to submit new orders and update status from website/dashboard
 */
function doPost(e) {
  try {
    let payload = {};
    if (e && e.postData && e.postData.contents) {
      try {
        payload = JSON.parse(e.postData.contents);
      } catch (jsonErr) {
        // Fallback if sent as urlencoded or plain text
        payload = e.parameter || {};
      }
    } else if (e && e.parameter) {
      payload = e.parameter;
    }

    const action = payload.action || "createOrder";

    if (action === "createOrder") {
      const order = payload.order || payload;
      const result = appendNewOrder(order);
      return createJsonResponse({
        status: "success",
        message: "Order successfully recorded in Google Sheets!",
        orderId: result.orderId,
        row: result.row
      });
    }

    if (action === "updateOrder") {
      const orderId = payload.orderId;
      const updates = payload.updates || {};
      const success = updateOrderRow(orderId, updates);
      return createJsonResponse({
        status: success ? "success" : "error",
        message: success ? "Order status updated." : "Order ID not found."
      });
    }

    if (action === "deleteOrder") {
      const orderId = payload.orderId;
      const success = deleteOrderRow(orderId);
      return createJsonResponse({
        status: success ? "success" : "error",
        message: success ? "Order deleted." : "Order ID not found."
      });
    }

    return createJsonResponse({
      status: "error",
      message: "Unsupported POST action: " + action
    });

  } catch (error) {
    return createJsonResponse({
      status: "error",
      message: "Error processing request: " + error.toString()
    });
  }
}

/**
 * Helper to produce standardized JSON response with CORS headers
 */
function createJsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Get or automatically initialize the "Orders" sheet with styled headers
 */
function getOrdersSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }

  // If sheet is fresh or has no headers, set up layout
  if (sheet.getLastRow() === 0) {
    setupSheetLayout(sheet);
  }

  return sheet;
}

/**
 * Configure professional columns, styling, and frozen header
 */
function setupSheetLayout(sheet) {
  sheet.clear();

  const headers = [
    "Order ID",
    "Created At",
    "Customer Name",
    "Instagram",
    "Phone",
    "Box Size",
    "Quantity",
    "Extra Toppings",
    "Fulfilment",
    "Preferred Date",
    "Total (€)",
    "Order Status",
    "Payment Status",
    "Notes"
  ];

  // Set header text
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Style header row (Dark Brown / Cocoa #2B1810 with bold white text)
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground("#24140E");
  headerRange.setFontColor("#FAF9F6");
  headerRange.setFontWeight("bold");
  headerRange.setFontFamily("Arial");
  headerRange.setFontSize(10);
  headerRange.setHorizontalAlignment("center");
  headerRange.setVerticalAlignment("middle");
  sheet.setRowHeight(1, 38);

  // Freeze top row
  sheet.setFrozenRows(1);

  // Set default column widths for readability
  sheet.setColumnWidth(1, 110); // Order ID
  sheet.setColumnWidth(2, 160); // Created At
  sheet.setColumnWidth(3, 150); // Customer Name
  sheet.setColumnWidth(4, 140); // Instagram
  sheet.setColumnWidth(5, 130); // Phone
  sheet.setColumnWidth(6, 120); // Box Size
  sheet.setColumnWidth(7, 80);  // Quantity
  sheet.setColumnWidth(8, 220); // Toppings
  sheet.setColumnWidth(9, 110); // Fulfilment
  sheet.setColumnWidth(10, 120); // Preferred Date
  sheet.setColumnWidth(11, 90);  // Total (€)
  sheet.setColumnWidth(12, 120); // Order Status
  sheet.setColumnWidth(13, 130); // Payment Status
  sheet.setColumnWidth(14, 260); // Notes
}

/**
 * Append a newly submitted order to the sheet
 */
function appendNewOrder(order) {
  const sheet = getOrdersSheet();

  const toppingsStr = Array.isArray(order.toppings) 
    ? order.toppings.join(", ") 
    : (order.toppings || "");

  const orderId = order.id || ("CL-" + Math.floor(1000 + Math.random() * 9000));
  const createdAt = order.createdAt || new Date().toISOString();

  const rowData = [
    orderId,
    createdAt,
    order.name || "Customer",
    order.instagram || "@unspecified",
    order.phone || "",
    order.boxSize || "Box of 6",
    Number(order.quantity) || 1,
    toppingsStr,
    order.fulfilment || "Collection",
    order.preferredDate || "",
    Number(order.total) || 14,
    order.status || "New",
    order.paymentStatus || "Unpaid",
    order.notes || ""
  ];

  sheet.appendRow(rowData);
  const newRowIdx = sheet.getLastRow();

  // Basic styling for the new row
  const rowRange = sheet.getRange(newRowIdx, 1, 1, rowData.length);
  rowRange.setVerticalAlignment("middle");
  rowRange.setFontFamily("Arial");
  rowRange.setFontSize(10);
  sheet.setRowHeight(newRowIdx, 28);

  // Center align columns: ID, Qty, Fulfilment, Date, Total, Statuses
  sheet.getRange(newRowIdx, 1).setHorizontalAlignment("center");
  sheet.getRange(newRowIdx, 7).setHorizontalAlignment("center");
  sheet.getRange(newRowIdx, 9).setHorizontalAlignment("center");
  sheet.getRange(newRowIdx, 10).setHorizontalAlignment("center");
  sheet.getRange(newRowIdx, 11).setHorizontalAlignment("center");
  sheet.getRange(newRowIdx, 12).setHorizontalAlignment("center");
  sheet.getRange(newRowIdx, 13).setHorizontalAlignment("center");

  return { orderId: orderId, row: newRowIdx };
}

/**
 * Fetch all orders from the sheet
 */
function fetchAllOrders() {
  const sheet = getOrdersSheet();
  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) {
    return [];
  }

  const values = sheet.getRange(2, 1, lastRow - 1, 14).getValues();
  const orders = [];

  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    if (!row[0]) continue; // Skip blank rows

    const toppings = row[7] ? String(row[7]).split(",").map(function(t) { return t.trim(); }) : [];

    orders.push({
      id: String(row[0]),
      createdAt: row[1] ? (row[1] instanceof Date ? row[1].toISOString() : String(row[1])) : new Date().toISOString(),
      name: String(row[2] || ""),
      instagram: String(row[3] || ""),
      phone: String(row[4] || ""),
      boxSize: String(row[5] || "Box of 6"),
      boxCount: parseInt(String(row[5]).replace(/\D/g, '')) || 6,
      quantity: Number(row[6]) || 1,
      toppings: toppings,
      fulfilment: String(row[8] || "Collection"),
      preferredDate: row[9] ? (row[9] instanceof Date ? Utilities.formatDate(row[9], "GMT", "yyyy-MM-dd") : String(row[9])) : "",
      total: Number(row[10]) || 0,
      status: String(row[11] || "New"),
      paymentStatus: String(row[12] || "Unpaid"),
      notes: String(row[13] || "")
    });
  }

  // Return sorted with newest first
  orders.reverse();
  return orders;
}

/**
 * Update order row by Order ID
 */
function updateOrderRow(orderId, updates) {
  const sheet = getOrdersSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return false;

  const idColValues = sheet.getRange(2, 1, lastRow - 1, 1).getValues();

  for (let i = 0; i < idColValues.length; i++) {
    if (String(idColValues[i][0]).trim() === String(orderId).trim()) {
      const rowIndex = i + 2;

      if (updates.status !== undefined) {
        sheet.getRange(rowIndex, 12).setValue(updates.status);
      }
      if (updates.paymentStatus !== undefined) {
        sheet.getRange(rowIndex, 13).setValue(updates.paymentStatus);
      }
      if (updates.notes !== undefined) {
        sheet.getRange(rowIndex, 14).setValue(updates.notes);
      }
      return true;
    }
  }
  return false;
}

/**
 * Delete order row by Order ID
 */
function deleteOrderRow(orderId) {
  const sheet = getOrdersSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return false;

  const idColValues = sheet.getRange(2, 1, lastRow - 1, 1).getValues();

  for (let i = 0; i < idColValues.length; i++) {
    if (String(idColValues[i][0]).trim() === String(orderId).trim()) {
      sheet.deleteRow(i + 2);
      return true;
    }
  }
  return false;
}

/**
 * Custom Menu in Google Sheets UI for easy maintenance
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🧁 CL's Brownies")
    .addItem("Format Sheet Headers", "menuSetupLayout")
    .addItem("Add Sample Order", "menuAddSampleOrder")
    .addToUi();
}

function menuSetupLayout() {
  const sheet = getOrdersSheet();
  setupSheetLayout(sheet);
  SpreadsheetApp.getUi().alert("Sheet headers and column widths initialized successfully!");
}

function menuAddSampleOrder() {
  appendNewOrder({
    id: "CL-DEMO",
    name: "Alex Sample",
    instagram: "@alex_dublin",
    phone: "087 000 1234",
    boxSize: "Box of 6",
    quantity: 1,
    toppings: ["Kinder Bueno", "Milk Chocolate Drizzle"],
    total: 16,
    fulfilment: "Collection",
    preferredDate: Utilities.formatDate(new Date(), "GMT", "yyyy-MM-dd"),
    status: "New",
    paymentStatus: "Unpaid",
    notes: "Sample test order generated from Google Sheets menu."
  });
  SpreadsheetApp.getUi().alert("Sample order added to the table!");
}
