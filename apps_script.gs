// Google Apps Script — extensão simples para guardar ofertas numa Google Sheet.
// Crie uma planilha com uma aba chamada Ofertas e cole este código em Extensões > Apps Script.
const SHEET_NAME = "Ofertas";
function doPost(e) {
  const data = JSON.parse(e.postData.contents);
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (sh.getLastRow() === 0) sh.appendRow(["nome","foto","preco","preco_anterior","categoria","link","data"]);
  sh.appendRow([data.name,data.image,data.price,data.oldPrice,data.category,data.link,new Date()]);
  return ContentService.createTextOutput(JSON.stringify({ok:true})).setMimeType(ContentService.MimeType.JSON);
}
function doGet() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const rows = sh.getDataRange().getValues();
  const data = rows.slice(1).map(r=>({name:r[0],image:r[1],price:r[2],oldPrice:r[3],category:r[4],link:r[5],date:new Date(r[6]).getTime()}));
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}