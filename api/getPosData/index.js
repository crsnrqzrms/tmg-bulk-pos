const { Client } = require("@microsoft/microsoft-graph-client");
require("isomorphic-fetch");

module.exports = async function (context, req) {
  try {
    const driveItemId = process.env.EXCEL_FILE_DRIVE_ITEM_ID;

    // Check if cloud credentials exist; otherwise fallback to test defaults
    if (!driveItemId || !process.env.MICROSOFT_GRAPH_CLIENT_SECRET) {
      context.res = {
        status: 200,
        headers: { "Content-Type": "application/json" },
        body: {
          brands: ["Manam"],
          catalog: [],
          shopMap: { "Manam": ["MANAM GREENHILLS"] },
          thresholdMap: { "MANAM|MANAM GREENHILLS": 10000 },
          paymentMethods: [
            { name: "Cash", requireRef: false },
            { name: "Credit / Debit Card", requireRef: true },
            { name: "GCash", requireRef: true }
          ],
          discountList: [
            { name: "Senior Citizen (RA 9994)", rule: "STATUTORY_HEADCOUNT", val: 20, requireId: true },
            { name: "PWD (RA 10754)", rule: "STATUTORY_HEADCOUNT", val: 20, requireId: true }
          ],
          userEmail: req.headers["x-ms-client-principal-name"] || "Local Terminal"
        }
      };
      return;
    }

    // Authenticate with Microsoft Graph
    const graphClient = Client.init({
      authProvider: async (done) => {
        const token = await getAccessToken();
        done(null, token);
      }
    });

    // 1. Fetch Mapping tab (Catalog)
    const mappingRes = await graphClient
      .api(`/me/drive/items/${driveItemId}/workbook/tables/Mapping/rows`)
      .get();
    const catalogRows = mappingRes.value.map((r, idx) => ({
      id: idx,
      brand: r.values[0][0],
      subBrand: r.values[0][1],
      mainCat: r.values[0][2],
      subCat: r.values[0][3],
      itemDesc: r.values[0][4],
      size: r.values[0][5],
      itemCode: r.values[0][6],
      price: parseFloat(String(r.values[0][7]).replace(/[^0-9.-]+/g, "")) || 0
    }));

    const brands = [...new Set(catalogRows.map(r => r.brand).filter(Boolean))].sort();

    // 2. Fetch Shop List tab
    const shopRes = await graphClient
      .api(`/me/drive/items/${driveItemId}/workbook/tables/ShopList/rows`)
      .get();
    const shopMap = {};
    const thresholdMap = {};
    shopRes.value.forEach(r => {
      const b = String(r.values[0][0] || "").trim();
      const s = String(r.values[0][1] || "").trim();
      const t = parseFloat(String(r.values[0][2] || "").replace(/[^0-9.-]+/g, "")) || 0;
      if (b && s) {
        if (!shopMap[b]) shopMap[b] = [];
        shopMap[b].push(s);
        thresholdMap[(b + "|" + s).toUpperCase()] = t;
      }
    });

    // 3. Fetch Payment Methods
    const payRes = await graphClient
      .api(`/me/drive/items/${driveItemId}/workbook/tables/PaymentMethods/rows`)
      .get();
    const paymentMethods = payRes.value
      .filter(r => r.values[0][2] === true || String(r.values[0][2]).toUpperCase() === "TRUE")
      .map(r => ({
        name: String(r.values[0][0]).trim(),
        requireRef: r.values[0][1] === true || String(r.values[0][1]).toUpperCase() === "TRUE"
      }));

    // 4. Fetch Discount Mapping
    const discRes = await graphClient
      .api(`/me/drive/items/${driveItemId}/workbook/tables/DiscountMapping/rows`)
      .get();
    const discountList = discRes.value
      .filter(r => r.values[0][4] === true || String(r.values[0][4]).toUpperCase() === "TRUE")
      .map(r => ({
        name: String(r.values[0][0]).trim(),
        rule: String(r.values[0][1] || "STATUTORY_HEADCOUNT").toUpperCase(),
        val: parseFloat(r.values[0][2]) || 0,
        requireId: r.values[0][3] === true || String(r.values[0][3]).toUpperCase() === "TRUE"
      }));

    context.res = {
      status: 200,
      headers: { "Content-Type": "application/json" },
      body: {
        brands,
        catalog: catalogRows,
        shopMap,
        thresholdMap,
        paymentMethods,
        discountList,
        userEmail: req.headers["x-ms-client-principal-name"] || "Store User"
      }
    };
  } catch (err) {
    context.res = { status: 500, body: "Error fetching POS catalog: " + err.message };
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