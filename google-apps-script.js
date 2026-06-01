// ============================================================
//  ATHLETE GROCERY PWA — Google Apps Script Backend
//  Paste this in: Extensions > Apps Script > Code.gs
//  Then: Deploy > New Deployment > Web App
//    - Execute as: Me
//    - Who has access: Anyone
//  Copy the Web App URL into the frontend CONFIG object.
// ============================================================

const SHEET_NAME_PRODUCTS  = "Base Produits";
const SHEET_NAME_LIST      = "Liste de Course";

// ── Helpers ──────────────────────────────────────────────────

function getOrCreateSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  return sheet;
}

function seedProductsIfEmpty(sheet) {
  if (sheet.getLastRow() > 1) return; // already has data

  // Headers
  sheet.getRange(1, 1, 1, 11).setValues([[
    "ID", "Categorie", "Nom", "Prix (€)", "Unite", "Quantite",
    "Calories/100g", "Proteines/100g", "Glucides/100g", "Lipides/100g", "Saison"
  ]]);
  sheet.getRange(1, 1, 1, 11).setFontWeight("bold").setBackground("#1e3a5f").setFontColor("#ffffff");

  // Sample summer-season products for a 120kg athlete
  const products = [
    // [ID, Categorie, Nom, Prix, Unite, Quantite, Kcal, Prot, Gluc, Lip, Saison]
    // ── Viandes & Poissons ────────────────────────────────────
    ["P001","Viande & Poisson","Blanc de poulet (1kg)","6.99","kg","1","110","23.1","0","1.2","Toute année"],
    ["P002","Viande & Poisson","Steak haché 5% MG (x4)","5.49","barquette","1","121","20.7","0","4.3","Toute année"],
    ["P003","Viande & Poisson","Saumon frais (500g)","7.99","paquet","1","208","20.1","0","13.6","Toute année"],
    ["P004","Viande & Poisson","Thon en boite naturel (x5)","4.99","lot","1","116","25.5","0","0.6","Toute année"],
    ["P005","Viande & Poisson","Dinde escalope (600g)","5.29","barquette","1","107","22.6","0","1.4","Toute année"],
    ["P006","Viande & Poisson","Sardines à l'huile (x3)","3.49","lot","1","208","17.8","0","14.8","Toute année"],
    // ── Œufs & Produits laitiers ──────────────────────────────
    ["P007","Œufs & Laitiers","Œufs x12 plein air","3.59","boite","1","155","13.0","1.1","11.0","Toute année"],
    ["P008","Œufs & Laitiers","Skyr nature 0% (x4 500g)","5.89","lot","1","63","11.0","4.0","0.2","Toute année"],
    ["P009","Œufs & Laitiers","Fromage blanc 0% (1kg)","2.99","pot","1","46","8.0","3.4","0.1","Toute année"],
    ["P010","Œufs & Laitiers","Lait demi-écrémé (2L)","2.09","bouteille","1","46","3.2","4.8","1.5","Toute année"],
    ["P011","Œufs & Laitiers","Parmesan râpé (100g)","2.49","sachet","1","392","35.7","0","28.1","Toute année"],
    // ── Féculents & Céréales ──────────────────────────────────
    ["P012","Féculents","Flocons d'avoine (1kg)","2.19","paquet","1","372","13.5","59.8","7.1","Toute année"],
    ["P013","Féculents","Riz basmati complet (1kg)","2.49","paquet","1","350","8.0","71.3","2.8","Toute année"],
    ["P014","Féculents","Pâtes complètes (500g)","1.49","paquet","1","356","13.4","65.0","2.5","Toute année"],
    ["P015","Féculents","Pommes de terre (2kg)","2.29","filet","1","77","2.0","17.0","0.1","Toute année"],
    ["P016","Féculents","Pain complet (500g)","1.99","paquet","1","247","9.5","41.0","3.3","Toute année"],
    ["P017","Féculents","Quinoa (500g)","3.99","paquet","1","368","14.1","57.2","6.1","Toute année"],
    // ── Légumes (Été) ─────────────────────────────────────────
    ["P018","Légumes Été","Tomates vrac (1kg)","2.49","kg","1","18","0.9","3.5","0.2","Été"],
    ["P019","Légumes Été","Courgettes (1kg)","1.99","kg","1","17","1.2","2.6","0.3","Été"],
    ["P020","Légumes Été","Poivrons (3 pièces)","2.29","lot","1","31","1.0","6.0","0.3","Été"],
    ["P021","Légumes Été","Épinards frais (400g)","2.49","sachet","1","23","2.9","0.4","0.4","Été"],
    ["P022","Légumes Été","Brocoli (500g)","1.79","pièce","1","34","2.8","4.0","0.4","Été"],
    ["P023","Légumes Été","Concombre (pièce)","0.89","pièce","1","15","0.6","2.5","0.1","Été"],
    ["P024","Légumes Été","Aubergine (pièce)","1.29","pièce","1","25","1.0","5.0","0.2","Été"],
    ["P025","Légumes Été","Haricots verts (500g)","2.19","sachet","1","31","1.8","5.7","0.2","Été"],
    ["P026","Légumes Été","Salade mélange (sachet)","1.59","sachet","1","15","1.2","1.7","0.3","Été"],
    // ── Fruits (Été) ──────────────────────────────────────────
    ["P027","Fruits Été","Pastèque (pièce ~3kg)","4.99","pièce","1","30","0.6","7.0","0.2","Été"],
    ["P028","Fruits Été","Cerises (500g)","3.99","barquette","1","63","1.1","14.3","0.2","Été"],
    ["P029","Fruits Été","Abricots (500g)","2.99","barquette","1","48","1.4","10.5","0.4","Été"],
    ["P030","Fruits Été","Pêches / Nectarines (1kg)","3.49","kg","1","42","1.0","9.5","0.1","Été"],
    ["P031","Fruits Été","Fraises (500g)","3.29","barquette","1","32","0.7","7.0","0.3","Été"],
    ["P032","Fruits Été","Bananes (1kg)","1.79","kg","1","89","1.1","23.0","0.3","Toute année"],
    ["P033","Fruits Été","Myrtilles (250g)","3.49","barquette","1","57","0.7","14.5","0.3","Été"],
    // ── Légumineuses ─────────────────────────────────────────
    ["P034","Légumineuses","Pois chiches cuits (800g)","1.19","boite","1","119","7.2","17.4","2.6","Toute année"],
    ["P035","Légumineuses","Lentilles vertes (500g)","1.69","paquet","1","353","24.0","60.1","1.8","Toute année"],
    ["P036","Légumineuses","Haricots rouges cuits (800g)","1.09","boite","1","127","8.7","20.2","0.5","Toute année"],
    // ── Graisses & Huiles ─────────────────────────────────────
    ["P037","Matières Grasses","Huile d'olive (1L)","7.99","bouteille","1","884","0","0","100","Toute année"],
    ["P038","Matières Grasses","Beurre de cacahuète (500g)","4.29","pot","1","588","25.1","20.1","50.4","Toute année"],
    ["P039","Matières Grasses","Amandes (200g)","3.99","sachet","1","579","21.2","4.9","50.6","Toute année"],
    ["P040","Matières Grasses","Graines de chia (200g)","3.49","sachet","1","486","17.0","42.1","31.0","Toute année"],
    ["P041","Matières Grasses","Avocat x3","2.99","filet","1","160","2.0","8.5","14.7","Toute année"],
    // ── Condiments & Divers ───────────────────────────────────
    ["P042","Condiments","Sauce soja (150ml)","1.99","bouteille","1","60","5.6","6.7","0.1","Toute année"],
    ["P043","Condiments","Miel (250g)","3.49","pot","1","304","0.3","82.4","0","Toute année"],
    ["P044","Condiments","Café soluble (200g)","4.49","pot","1","2","0.1","0.4","0.2","Toute année"],
    ["P045","Condiments","Levure bière (100g)","3.29","sachet","1","116","13.8","15.6","1.4","Toute année"],
    // ── Boissons ──────────────────────────────────────────────
    ["P046","Boissons","Eau minérale 6x1.5L","3.99","pack","1","0","0","0","0","Toute année"],
    ["P047","Boissons","Jus d'orange pur jus (1L)","2.49","bouteille","1","44","0.7","10.4","0.2","Toute année"],
  ];

  if (products.length > 0) {
    sheet.getRange(2, 1, products.length, 11).setValues(products);
  }
  sheet.autoResizeColumns(1, 11);
}

function seedListIfEmpty(sheet) {
  if (sheet.getLastRow() > 0) return;
  sheet.getRange(1, 1, 1, 3).setValues([["ID Produit", "Nom", "Quantite"]]);
  sheet.getRange(1, 1, 1, 3).setFontWeight("bold").setBackground("#1e3a5f").setFontColor("#ffffff");
}

// ── GET handler ───────────────────────────────────────────────

function doGet(e) {
  try {
    const productsSheet = getOrCreateSheet(SHEET_NAME_PRODUCTS);
    const listSheet     = getOrCreateSheet(SHEET_NAME_LIST);

    seedProductsIfEmpty(productsSheet);
    seedListIfEmpty(listSheet);

    // Read products (skip header row)
    const pData    = productsSheet.getDataRange().getValues();
    const pHeaders = pData[0];
    const products = pData.slice(1).map(row => {
      const obj = {};
      pHeaders.forEach((h, i) => obj[h] = row[i]);
      return obj;
    });

    // Read saved list (skip header row)
    const lData    = listSheet.getDataRange().getValues();
    const lHeaders = lData[0];
    const list     = lData.slice(1)
      .filter(r => r[0] !== "")
      .map(row => {
        const obj = {};
        lHeaders.forEach((h, i) => obj[h] = row[i]);
        return obj;
      });

    const payload = JSON.stringify({ products, list });
    return ContentService.createTextOutput(payload)
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── POST handler ──────────────────────────────────────────────

function doPost(e) {
  try {
    const body    = JSON.parse(e.postData.contents);
    const newList = body.list || [];

    const listSheet = getOrCreateSheet(SHEET_NAME_LIST);

    // Clear existing data (keep header)
    const lastRow = listSheet.getLastRow();
    if (lastRow > 1) {
      listSheet.getRange(2, 1, lastRow - 1, 3).clearContent();
    }

    // Write new list
    if (newList.length > 0) {
      const rows = newList.map(item => [
        item["ID Produit"] || item.id || "",
        item["Nom"]        || item.name || "",
        item["Quantite"]   || item.qty  || 1
      ]);
      listSheet.getRange(2, 1, rows.length, 3).setValues(rows);
    }

    return ContentService.createTextOutput(JSON.stringify({ success: true, saved: newList.length }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
