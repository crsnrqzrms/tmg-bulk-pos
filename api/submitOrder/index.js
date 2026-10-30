const { Client } = require("@microsoft/microsoft-graph-client");
require("isomorphic-fetch");

module.exports = async function (context, req) {
  const payload = req.body;

  if (!payload || !payload.cart || payload.cart.length === 0) {
    context.res = { status: 400, body: "Missing order items in payload." };
    return;
  }

  const timestamp = new Date().toISOString();
  const orderId = "POS-" + Date.now();
  const userEmail = req.headers["x-ms-client-principal-name"] || "Store User";

  // Build the 29-column row layout matching the POS schema
  const rowsToInsert = payload.cart.map(item => [
    timestamp,
    userEmail,
    orderId,
    payload.orNumber,
    payload.brand,
    payload.shop,
    payload.paymentMethod,
    payload.paymentRef || "",
    payload.customerName || "",
    payload.customerTin || "",
    payload.discountScheme || "NONE",
    payload.discountId || "",
    payload.subtotal,
    payload.totalDiscount,
    payload.finalTotal,
    payload.vatableSales,
    payload.vatExemptSales,
    payload.vatAmount,
    item.itemCode,
    item.itemDesc,
    item.size,
    item.subBrand,
    item.mainCat + " - " + item.subCat,
    item.price,
    item.qty,
    item.price * item.qty,
    item.lineDiscount,
    (item.price * item.qty) - item.lineDiscount,
    payload.notes || ""
  ]);

  const driveItemId = process.env.EXCEL_FILE_DRIVE_ITEM_ID;

  // If local / test mode without credentials, simulate success
  if (!driveItemId || !process.env.MICROSOFT_GRAPH_CLIENT_SECRET) {
    context.res = {
      status: 200,
      headers: { "Content-Type": "application/json" },
      body: {
        status: "Success (Local Mock)",
        orderId: orderId,
        orNumber: payload.orNumber,
        shop: payload.shop,
        paymentMethod: payload.paymentMethod,
        finalTotal: payload.finalTotal,
        vatableSales: payload.vatableSales,
        vatExemptSales: payload.vatExemptSales,
        vatAmount: payload.vatAmount
      }
    };
    return;
  }

  try {
    const graphClient = Client.init({
      authProvider: async (done) => {
        const token = await getAccessToken();
        done(null, token);
      }
    });

    // Append all rows to POS_Transactions table in Excel Online
    await graphClient
      .api(`/me/drive/items/${driveItemId}/workbook/tables/POS_Transactions/rows`)
      .post({ values: rowsToInsert });

    context.res = {
      status: 200,
      headers: { "Content-Type": "application/json" },
      body: {
        status: "Success",
        orderId: orderId,
        orNumber: payload.orNumber,
        shop: payload.shop,
        paymentMethod: payload.paymentMethod,
        finalTotal: payload.finalTotal,
        vatableSales: payload.vatableSales,
        vatExemptSales: payload.vatExemptSales,
        vatAmount: payload.vatAmount
      }
    };
  } catch (err) {
    context.res = { status: 500, body: "Error saving transaction to Excel: " + err.message };
  }
};

async function getAccessToken() {
  const params = new URLSearchParams({
    client_id: process.env.MICROSOFT_GRAPH_CLIENT_ID,
    client_secret: process.env.MICROSOFT_GRAPH_CLIENT_SECRET,
    grant_type: "client_credentials",
    scope: "https://graph.microsoft.com/.default"
  });

  const res = await fetch(`https://login.microsoftonline.com/${process.env.MICROSOFT_GRAPH_TENANT_ID}/oauth2/v2.0/token`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: params.toString()
  });

  const data = await res.json();
  return data.access_token;
}