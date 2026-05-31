function checkStockAlert() {

  var token = "8244721253:AAGVKI3ikWaScpBQtjCFamPEuCYRbAUZysM";   // tukar token baru
  var chat_id = "8186247289";

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Stock");
  var data = sheet.getDataRange().getValues();

  var message = "⚠️ STOK RENDAH:\n\n";
  var alert = false;

  for (var i = 1; i < data.length; i++) {

    var bahan = data[i][0];        // Column A
    var stok_awal = data[i][1];    // Column B
    var digunakan = data[i][2];    // Column C
    var baki = data[i][3];         // Column D
    var minimum = data[i][4];      // Column E
    // auto kira baki kalau kosong
    if (!baki) {
      baki = stok_awal - digunakan;
      sheet.getRange(i + 1, 4).setValue(baki);
    }

    // check stok rendah
    if (baki <= minimum) {
      message += "🔴 " + bahan + "\n";
      message += "Baki: " + baki + " | Min: " + minimum + "\n\n";

      // highlight merah
      sheet.getRange(i + 1, 4).setBackground("#ff4d4d");

      alert = true;
    } else {
      // reset warna kalau ok
      sheet.getRange(i + 1, 4).setBackground("#ffffff");
    }
  }

  // hantar telegram kalau ada alert
  if (alert) {
    var url = "https://api.telegram.org/bot" + token + "/sendMessage?chat_id=" 
              + chat_id + "&text=" + encodeURIComponent(message);

    UrlFetchApp.fetch(url);
  }
}
// ======================
// doGet
// ======================
function doGet(e) {
  if (e.parameter.action === "getMaterials") {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Stock");  // Nama sheet anda
    const data = sheet.getRange("A2:A1000").getValues()   // Ambil sehingga row 1000
                     .flat()
                     .filter(String)                       // Buang kosong
                     .map(item => item.toString().trim());
    
    return ContentService.createTextOutput(JSON.stringify(data))
             .setMimeType(ContentService.MimeType.JSON);
  }
  if (e.parameter.action === "getBalanceByMaterial") {
  var material = (e.parameter.material || "").trim();

  if (!material) {
    return ContentService.createTextOutput(JSON.stringify({ error: "Material tidak dinyatakan" }))
             .setMimeType(ContentService.MimeType.JSON);
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Stock");
  const data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === material.toLowerCase()) {
      return ContentService.createTextOutput(JSON.stringify({
        material: data[i][0],
        baki: data[i][3]
      })).setMimeType(ContentService.MimeType.JSON);
    }
  }

  return ContentService.createTextOutput(JSON.stringify({ error: "Material tidak dijumpai" }))
           .setMimeType(ContentService.MimeType.JSON);
}
if (e.parameter.action === "getUsageHistory") {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("MaterialUsage");
  const data = sheet.getDataRange().getValues();

  if (data.length <= 1) {
    return ContentService.createTextOutput(JSON.stringify([]))
             .setMimeType(ContentService.MimeType.JSON);
  }

  var history = [];

  for (var i = 1; i < data.length; i++) {
    history.push({
      timestamp : data[i][0] ? new Date(data[i][0]).toLocaleString('ms-MY') : "",
      nama      : data[i][1],
      material  : data[i][2],
      kuantiti  : data[i][3],
      unit      : data[i][4],
      tujuan    : data[i][5]
    });
  }

  // Terbalik supaya paling baru di atas
  history.reverse();

  return ContentService.createTextOutput(JSON.stringify(history))
           .setMimeType(ContentService.MimeType.JSON);
}
  // ... kod POST anda yang lain
}

// ======================
// doPost - VERSI DIPERBAIKI
// ======================
function doPost(e) {
  try {
    var type = e.parameter.type || "";

    // ==================== RESTOCK ====================
    if (type === "restock") {
      var material = (e.parameter.material || "").trim();
      var kuantiti = Number(e.parameter.kuantiti) || 0;

      if (!material || kuantiti <= 0) {
        return ContentService.createTextOutput("Error: Data restock tidak lengkap");
      }

      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var stockSheet = ss.getSheetByName("Stock");
      var restockSheet = ss.getSheetByName("Restock");

      var data = stockSheet.getDataRange().getValues();

      for (var i = 1; i < data.length; i++) {
        if (String(data[i][0]).trim() === material) {
          var stokAwalBaru = Number(data[i][1]) + kuantiti;
          var bakiBaru     = Number(data[i][3]) + kuantiti;

          stockSheet.getRange(i + 1, 2).setValue(stokAwalBaru);
          stockSheet.getRange(i + 1, 4).setValue(bakiBaru);

          restockSheet.appendRow([new Date(), material, kuantiti]);
          sendTelegramRestock(material, kuantiti);

          return ContentService.createTextOutput("Restock Success");
        }
      }
      return ContentService.createTextOutput("Error: Material tidak dijumpai");
    }

    // ==================== MATERIAL USAGE ====================
    var nama    = (e.parameter.nama || "").trim();
    var tujuan  = (e.parameter.tujuan || "").trim();
    var itemsJson = e.parameter.items;

    if (!nama) return ContentService.createTextOutput("Error: Nama peminjam diperlukan");
    if (!itemsJson) return ContentService.createTextOutput("Error: Tiada barang dipilih");

    var items = JSON.parse(itemsJson);

    if (!Array.isArray(items) || items.length === 0) {
      return ContentService.createTextOutput("Error: Tiada barang");
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var stockSheet = ss.getSheetByName("Stock");
    var usageSheet = ss.getSheetByName("MaterialUsage");

    if (!stockSheet || !usageSheet) {
      return ContentService.createTextOutput("Error: Sheet Stock atau MaterialUsage tidak dijumpai");
    }

    var stockData = stockSheet.getDataRange().getValues();
    var timestamp = new Date();

    for (var j = 0; j < items.length; j++) {
      var item = items[j];
      var material = (item.material || "").trim();
      var kuantiti = Number(item.kuantiti) || 0;
      var unit     = (item.unit || "").trim();

      if (!material || kuantiti <= 0) {
        return ContentService.createTextOutput("Error: Data tidak lengkap untuk " + material);
      }

      var found = false;

      for (var i = 1; i < stockData.length; i++) {
        if (String(stockData[i][0]).trim() === material) {

          var stokAwal     = Number(stockData[i][1]) || 0;
          var stokDigunakan = Number(stockData[i][2]) || 0;
          var baki         = Number(stockData[i][3]) || 0;

          if (baki <= 0) {
            sendTelegramOutOfStock(material);
            return ContentService.createTextOutput(`❌ STOK TELAH HABIS!\nMaterial: ${material}`);
          }

          if (kuantiti > baki) {
            sendTelegramAlert(material, baki);
            return ContentService.createTextOutput(
              `❌ STOK TIDAK CUKUP!\nMaterial: ${material}\nBaki: ${baki}\nDiminta: ${kuantiti}`
            );
          }

          // Update stok
          stokDigunakan += kuantiti;
          baki = stokAwal - stokDigunakan;

          stockSheet.getRange(i + 1, 3).setValue(stokDigunakan);
          stockSheet.getRange(i + 1, 4).setValue(baki);

          // Rekod ke sheet
          usageSheet.appendRow([timestamp, nama, material, kuantiti, unit, tujuan]);

          // Alert stok rendah selepas pinjam
          if (baki <= 10 && baki > 0) {
            sendTelegramAlert(material, baki);
          }

          found = true;
          break;
        }
      }

      if (!found) {
        return ContentService.createTextOutput(`Error: Material "${material}" tidak dijumpai`);
      }
    }

    // Hantar SATU notifikasi Telegram untuk semua barang
    sendTelegramUsageMulti(nama, tujuan, items);

    return ContentService.createTextOutput(`Berjaya! ${items.length} barang telah direkodkan.`);
    
  } catch (err) {
    console.log("doPost Error: " + err.message);   // Untuk debug
    return ContentService.createTextOutput("Error: " + err.message);
  }
}

// ======================
// TELEGRAM FUNCTIONS - WAJIB ADA
// ======================
function sendTelegramUsage(nama, material, kuantiti, unit, tujuan) {
  var token = "8244721253:AAGVKI3ikWaScpBQtjCFamPEuCYRbAUZysM";
  var chatId = "8186247289";

  var msg = "📦 Material Dipinjam\n" +
            "👤 Nama: " + nama + "\n" +
            "📦 Material: " + material + "\n" +
            "🔢 Kuantiti: " + kuantiti + 
            (unit ? "unit " + unit : "") + "\n" +   // Tambah unit di sini
            "📍 Tujuan: " + tujuan;

  try {
    UrlFetchApp.fetch("https://api.telegram.org/bot" + token + 
      "/sendMessage?chat_id=" + chatId + "&text=" + encodeURIComponent(msg));
  } catch(err) {
    Logger.log("Telegram usage error: " + err);
  }
}

function sendTelegramAlert(material, baki) {
  var token = "8244721253:AAGVKI3ikWaScpBQtjCFamPEuCYRbAUZysM";
  var chatId = "8186247289";

  var msg = "⚠️ STOCK RENDAH\nMaterial: " + material + "\nBaki Tinggal: " + baki;

  try {
    UrlFetchApp.fetch("https://api.telegram.org/bot" + token + 
      "/sendMessage?chat_id=" + chatId + "&text=" + encodeURIComponent(msg));
  } catch(err) {
    Logger.log("Telegram alert error: " + err);
  }
}

function sendTelegramOutOfStock(material) {
  var token = "8244721253:AAGVKI3ikWaScpBQtjCFamPEuCYRbAUZysM";
  var chatId = "8186247289";

  var msg = "❌ STOCK HABIS\nMaterial: " + material;

  try {
    UrlFetchApp.fetch("https://api.telegram.org/bot" + token + 
      "/sendMessage?chat_id=" + chatId + "&text=" + encodeURIComponent(msg));
  } catch(err) {
    Logger.log("Telegram out-of-stock error: " + err);
  }
}

function sendTelegramRestock(material, kuantiti) {
  var token = "8244721253:AAGVKI3ikWaScpBQtjCFamPEuCYRbAUZysM";
  var chatId = "8186247289";

  var msg = "🔄 RESTOCK BARANG\nMaterial: " + material + "\nKuantiti Tambah: " + kuantiti;

  try {
    UrlFetchApp.fetch("https://api.telegram.org/bot" + token + 
      "/sendMessage?chat_id=" + chatId + "&text=" + encodeURIComponent(msg));
  } catch(err) {
    Logger.log("Telegram restock error: " + err);
  }
}
// ======================
// Dapatkan senarai material untuk autocomplete
// ======================
function getMaterialList() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var stockSheet = ss.getSheetByName("Stock");
    
    if (!stockSheet) return ["Error: Sheet Stock tidak dijumpai"];
    
    var data = stockSheet.getDataRange().getValues();
    var materials = [];
    
    for (var i = 1; i < data.length; i++) {
      var mat = String(data[i][0]).trim();
      if (mat) materials.push(mat);
    }
    
    return materials.sort();
  } catch (err) {
    console.log("Error getMaterialList: " + err.message);
    return ["Ralat memuat senarai"];
  }
}

// ======================
// TELEGRAM USAGE MULTI (dengan Unit)
// ======================
function sendTelegramUsageMulti(nama, tujuan, items) {
  var token = "8244721253:AAGVKI3ikWaScpBQtjCFamPEuCYRbAUZysM";
  var chatId = "8186247289";

  var msg = "📦 **Material Dipinjam (Multi)**\n" +
            "👤 Nama: " + nama + "\n" +
            "📍 Tujuan: " + tujuan + "\n\n" +
            "📋 Senarai Barang:\n";

  items.forEach(function(item) {
    var unitText = item.unit ? " " + item.unit : "";
    msg += "• " + item.material + " → " + item.kuantiti + unitText + "\n";
  });

  msg += "\nJumlah barang: " + items.length + "\n⏰ " + new Date().toLocaleString('ms-MY');

  try {
    UrlFetchApp.fetch("https://api.telegram.org/bot" + token + 
      "/sendMessage?chat_id=" + chatId + "&text=" + encodeURIComponent(msg));
    console.log("Telegram Multi Usage berjaya");
  } catch(err) {
    console.log("Telegram multi usage error: " + err.message);
  }
}
