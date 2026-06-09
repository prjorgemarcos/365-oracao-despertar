// ============================================================
// 365 DIAS DE ORAÇÃO — PROJETO DESPERTAR
// Google Apps Script — Cole este código no Apps Script
// ============================================================

const SHEET_ID = '1NnkvQUMYfOn-o5mrcQoJyGH4GF1YTdYe-k-Po0hZ3sY';
const SHEET_NAME = 'Igrejas';

const COLS = ['id','name','city','state','juba','contact','whatsapp','day','status'];

function getSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    // Cabeçalho
    sheet.getRange(1, 1, 1, COLS.length).setValues([
      ['ID','Nome da Igreja','Município','UF','JUBA','Responsável','WhatsApp','Dia','Status']
    ]);
    sheet.getRange(1, 1, 1, COLS.length)
      .setBackground('#6b21a8')
      .setFontColor('#ffffff')
      .setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function doPost(e) {
  const cors = ContentService.createTextOutput();
  cors.setMimeType(ContentService.MimeType.JSON);

  try {
    const payload = JSON.parse(e.postData.contents);
    let result = {};

    if (payload.action === 'list') {
      result = listChurches();
    } else if (payload.action === 'save') {
      result = saveChurch(payload.church);
    } else if (payload.action === 'delete') {
      result = deleteChurch(payload.id);
    } else {
      result = { error: 'Ação desconhecida' };
    }

    cors.setContent(JSON.stringify(result));
  } catch(err) {
    cors.setContent(JSON.stringify({ error: err.message }));
  }

  return cors;
}

function doGet(e) {
  // Permite testar a URL no navegador
  return ContentService.createTextOutput(
    JSON.stringify(listChurches())
  ).setMimeType(ContentService.MimeType.JSON);
}

function listChurches() {
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { churches: [] };

  const churches = data.slice(1).map(row => {
    const obj = {};
    COLS.forEach((key, i) => { obj[key] = row[i] !== undefined ? String(row[i]) : ''; });
    return obj;
  }).filter(c => c.id); // ignora linhas vazias

  return { churches };
}

function saveChurch(church) {
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();

  // Procura linha existente pelo id
  let foundRow = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(church.id)) {
      foundRow = i + 1; // 1-indexed
      break;
    }
  }

  const row = COLS.map(k => church[k] || '');

  if (foundRow > 0) {
    // Atualiza linha existente
    sheet.getRange(foundRow, 1, 1, COLS.length).setValues([row]);
  } else {
    // Insere nova linha
    sheet.appendRow(row);
  }

  return { ok: true };
}

function deleteChurch(id) {
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      return { ok: true };
    }
  }

  return { error: 'Igreja não encontrada' };
}
