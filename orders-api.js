/**
 * CL'S BROWNIES — Shared Backend API Client
 * Seamless dual-layer store: Google Sheets Web App + LocalStorage fallback.
 * Works immediately offline / demo, and connects instantly to Google Sheets.
 */
(function(window) {
  'use strict';

  const STORAGE_KEY = 'cls_brownies_orders_v1';
  const CONFIG_KEY = 'cls_brownies_sheets_url';

  // Realistic sample orders for Dublin bakery
  const INITIAL_DEMO_ORDERS = [
    {
      id: "CL-1048",
      createdAt: new Date(Date.now() - 25 * 60 * 1000).toISOString(), // 25 mins ago
      name: "Sarah O'Connor",
      instagram: "@sarah.oc",
      phone: "087 123 4567",
      boxSize: "Box of 6",
      boxCount: 6,
      quantity: 1,
      toppings: ["Lotus Biscoff Biscuit", "Salted Caramel Drizzle"],
      total: 16,
      fulfilment: "Collection",
      preferredDate: new Date(Date.now() + 86400000).toISOString().split('T')[0], // Tomorrow
      status: "New",
      paymentStatus: "Unpaid",
      notes: "Can I collect around 4:30pm? It's for my sister's birthday box!"
    },
    {
      id: "CL-1047",
      createdAt: new Date(Date.now() - 2 * 3600 * 1000).toISOString(), // 2 hours ago
      name: "Liam Byrne",
      instagram: "@liambyrne_",
      phone: "085 987 6543",
      boxSize: "Box of 12",
      boxCount: 12,
      quantity: 1,
      toppings: ["Kinder Bueno", "Milk Chocolate Drizzle", "Mini Oreos"],
      total: 23,
      fulfilment: "Delivery",
      preferredDate: new Date(Date.now() + 2 * 86400000).toISOString().split('T')[0],
      status: "Confirmed",
      paymentStatus: "Paid (Revolut)",
      notes: "Delivery to Rathmines please if feasible. Revolut sent!"
    },
    {
      id: "CL-1046",
      createdAt: new Date(Date.now() - 4 * 3600 * 1000).toISOString(),
      name: "Aoife Murphy",
      instagram: "@aoifemurph_dublin",
      phone: "083 456 7890",
      boxSize: "Box of 6",
      boxCount: 6,
      quantity: 1,
      toppings: ["Pistachio Cream", "White Chocolate Drizzle"],
      total: 16,
      fulfilment: "Collection",
      preferredDate: new Date().toISOString().split('T')[0], // Today
      status: "Baking",
      paymentStatus: "Paid (Revolut)",
      notes: "Pistachio dream box. Agreed pickup 5pm."
    },
    {
      id: "CL-1045",
      createdAt: new Date(Date.now() - 6 * 3600 * 1000).toISOString(),
      name: "Conor Doyle",
      instagram: "@conordoyle",
      phone: "",
      boxSize: "Box of 4",
      boxCount: 4,
      quantity: 1,
      toppings: ["Fresh Raspberries", "Dark Chocolate Drizzle"],
      total: 12,
      fulfilment: "Collection",
      preferredDate: new Date().toISOString().split('T')[0],
      status: "Ready",
      paymentStatus: "Paid (Revolut)",
      notes: "Will swing by right after college lecture at 3pm."
    },
    {
      id: "CL-1042",
      createdAt: new Date(Date.now() - 28 * 3600 * 1000).toISOString(),
      name: "Emma Walsh",
      instagram: "@emma_w_dublin",
      phone: "086 234 5678",
      boxSize: "Box of 12",
      boxCount: 12,
      quantity: 2,
      toppings: ["Crushed Oreos", "Salted Caramel Drizzle"],
      total: 44,
      fulfilment: "Collection",
      preferredDate: new Date(Date.now() - 86400000).toISOString().split('T')[0],
      status: "Fulfilled",
      paymentStatus: "Paid (Revolut)",
      notes: "Huge hit at our office party! Will definitely reorder."
    }
  ];

  // Initialize Local Storage if blank
  function initLocalStorage() {
    try {
      const existing = localStorage.getItem(STORAGE_KEY);
      if (!existing) {
        localStorage.setItem(STORAGE_KEY, JSON.stringify(INITIAL_DEMO_ORDERS));
      }
    } catch (e) {
      console.warn("LocalStorage access failed:", e);
    }
  }

  function getLocalOrders() {
    try {
      const data = localStorage.getItem(STORAGE_KEY);
      return data ? JSON.parse(data) : INITIAL_DEMO_ORDERS;
    } catch (e) {
      return INITIAL_DEMO_ORDERS;
    }
  }

  function saveLocalOrders(orders) {
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(orders));
      window.dispatchEvent(new CustomEvent('cls_orders_updated', { detail: { orders } }));
    } catch (e) {
      console.error("Failed to save orders locally:", e);
    }
  }

  // Generate unique order ID
  function generateOrderId() {
    const randomNum = Math.floor(1000 + Math.random() * 9000);
    return `CL-${randomNum}`;
  }

  const DEFAULT_SHEETS_URL = 'https://script.google.com/macros/s/AKfycbwL1p_VacEPYCOhuWmlBhBJYYCFS70JRfA_L51Zbg9X2psxrpLq9MJr2HEK4i1R5uGf/exec';

  // POST a write action to the Apps Script (handled by doPost).
  // text/plain avoids a CORS preflight; no-cors means the response is opaque.
  function postToSheets(url, payload) {
    return fetch(url, {
      method: 'POST',
      mode: 'no-cors',
      redirect: 'follow',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify(payload)
    });
  }

  const BrowniesAPI = {
    // Get configured Google Apps Script Web App URL
    getSheetsUrl() {
      try {
        const stored = localStorage.getItem(CONFIG_KEY);
        if (stored !== null && stored !== '') return stored;
        return DEFAULT_SHEETS_URL;
      } catch (e) {
        return DEFAULT_SHEETS_URL;
      }
    },

    setSheetsUrl(url) {
      try {
        const cleanUrl = (url || '').trim();
        localStorage.setItem(CONFIG_KEY, cleanUrl);
        window.dispatchEvent(new CustomEvent('cls_config_updated', { detail: { url: cleanUrl } }));
        return true;
      } catch (e) {
        return false;
      }
    },

    isConnected() {
      const url = this.getSheetsUrl();
      return Boolean(url && url.startsWith('https://script.google.com/'));
    },

    // Test Google Sheets Connection
    async testConnection(url) {
      const targetUrl = (url || this.getSheetsUrl()).trim();
      if (!targetUrl) {
        return { success: false, message: "No Google Apps Script URL provided." };
      }
      try {
        const testEndpoint = targetUrl + (targetUrl.includes('?') ? '&' : '?') + 'action=ping&_=' + Date.now();
        const res = await fetch(testEndpoint, {
          method: 'GET',
          redirect: 'follow',
          mode: 'cors'
        });
        if (!res.ok) {
          throw new Error(`HTTP ${res.status}: ${res.statusText}`);
        }
        const data = await res.json();
        if (data && data.status === 'success') {
          return { success: true, message: data.message || "Connected to Google Sheets successfully!" };
        }
        return { success: false, message: data.message || "Invalid response format from Apps Script." };
      } catch (err) {
        // Many Apps Scripts redirect cross-origin. If CORS blocks reading JSON, check fallback
        return {
          success: false,
          message: `Connection error: ${err.message}. Please verify the Web App deployment is set to 'Anyone' access.`
        };
      }
    },

    // Get all orders (from Google Sheets if configured, with local fallback)
    async getOrders() {
      initLocalStorage();
      const sheetsUrl = this.getSheetsUrl();

      if (!sheetsUrl) {
        return { orders: getLocalOrders(), source: 'local' };
      }

      try {
        const fetchUrl = sheetsUrl + (sheetsUrl.includes('?') ? '&' : '?') + 'action=getOrders&_=' + Date.now();
        const response = await fetch(fetchUrl, {
          method: 'GET',
          redirect: 'follow',
          mode: 'cors'
        });

        if (response.ok) {
          const data = await response.json();
          if (data && data.status === 'success' && Array.isArray(data.orders)) {
            // Merge remote orders with local state to ensure no loss
            saveLocalOrders(data.orders);
            return { orders: data.orders, source: 'sheets' };
          }
        }
      } catch (err) {
        console.warn("Could not fetch remote Google Sheet orders, using local storage cache:", err);
      }

      return { orders: getLocalOrders(), source: 'local_fallback' };
    },

    // Submit new customer order
    async createOrder(orderInput) {
      initLocalStorage();
      const orderId = orderInput.id || generateOrderId();
      const newOrder = {
        id: orderId,
        createdAt: new Date().toISOString(),
        name: orderInput.name || "Customer",
        instagram: orderInput.instagram ? (orderInput.instagram.startsWith('@') ? orderInput.instagram : '@' + orderInput.instagram) : "@unspecified",
        phone: orderInput.phone || "",
        boxSize: orderInput.boxSize || "Box of 6",
        boxCount: orderInput.boxCount || 6,
        quantity: Number(orderInput.quantity) || 1,
        toppings: Array.isArray(orderInput.toppings) ? orderInput.toppings : [],
        total: Number(orderInput.total) || 14,
        fulfilment: orderInput.fulfilment === 'delivery' ? 'Delivery' : 'Collection',
        preferredDate: orderInput.preferredDate || "",
        status: "New",
        paymentStatus: "Unpaid",
        notes: orderInput.notes || ""
      };

      // 1. Immediately save to local storage
      const currentOrders = getLocalOrders();
      currentOrders.unshift(newOrder);
      saveLocalOrders(currentOrders);

      // 2. If Google Sheets URL configured, post asynchronously
      const sheetsUrl = this.getSheetsUrl();
      let remoteSynced = false;

      if (sheetsUrl) {
        try {
          await postToSheets(sheetsUrl, { action: 'createOrder', order: newOrder });
          remoteSynced = true;
        } catch (postErr) {
          console.warn("Failed to push order to Google Sheets:", postErr);
        }
      }

      return {
        success: true,
        order: newOrder,
        orderId: newOrder.id,
        remoteSynced: remoteSynced
      };
    },

    // Update status or payment status of an order
    async updateOrder(orderId, updates) {
      const orders = getLocalOrders();
      const index = orders.findIndex(o => o.id === orderId);
      if (index === -1) {
        return { success: false, message: "Order not found" };
      }

      orders[index] = { ...orders[index], ...updates };
      saveLocalOrders(orders);

      const sheetsUrl = this.getSheetsUrl();
      if (sheetsUrl) {
        try {
          await postToSheets(sheetsUrl, { action: 'updateOrder', orderId, updates });
        } catch (e) {
          console.warn("Failed to sync update to Google Sheets:", e);
        }
      }

      return { success: true, order: orders[index] };
    },

    // Delete or cancel an order
    async deleteOrder(orderId) {
      const orders = getLocalOrders().filter(o => o.id !== orderId);
      saveLocalOrders(orders);

      const sheetsUrl = this.getSheetsUrl();
      if (sheetsUrl) {
        try {
          await postToSheets(sheetsUrl, { action: 'deleteOrder', orderId });
        } catch (e) {
          console.warn("Failed to sync deletion to Google Sheets:", e);
        }
      }

      return { success: true };
    },

    // Reset back to sample orders
    resetDemoData() {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(INITIAL_DEMO_ORDERS));
      window.dispatchEvent(new CustomEvent('cls_orders_updated', { detail: { orders: INITIAL_DEMO_ORDERS } }));
      return INITIAL_DEMO_ORDERS;
    },

    // Export orders as downloadable CSV
    exportToCSV() {
      const orders = getLocalOrders();
      if (!orders.length) return false;

      const headers = [
        "Order ID",
        "Created At",
        "Customer Name",
        "Instagram Handle",
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

      const csvRows = [headers.join(',')];

      orders.forEach(o => {
        const row = [
          `"${o.id}"`,
          `"${o.createdAt}"`,
          `"${(o.name || '').replace(/"/g, '""')}"`,
          `"${(o.instagram || '').replace(/"/g, '""')}"`,
          `"${(o.phone || '').replace(/"/g, '""')}"`,
          `"${(o.boxSize || '').replace(/"/g, '""')}"`,
          o.quantity || 1,
          `"${(Array.isArray(o.toppings) ? o.toppings.join('; ') : '').replace(/"/g, '""')}"`,
          `"${o.fulfilment || 'Collection'}"`,
          `"${o.preferredDate || ''}"`,
          o.total || 0,
          `"${o.status || 'New'}"`,
          `"${o.paymentStatus || 'Unpaid'}"`,
          `"${(o.notes || '').replace(/"/g, '""')}"`
        ];
        csvRows.push(row.join(','));
      });

      const csvContent = "data:text/csv;charset=utf-8," + encodeURIComponent(csvRows.join('\n'));
      const link = document.createElement("a");
      link.setAttribute("href", csvContent);
      link.setAttribute("download", `cls_brownies_orders_${new Date().toISOString().slice(0, 10)}.csv`);
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      return true;
    }
  };

  // Run initial setup
  initLocalStorage();

  window.BrowniesAPI = BrowniesAPI;
})(window);
