// ============================================================
// GAS API Backend — 留言板範例（前後端分離測試）
// ============================================================
var SHEET_NAME = '留言';

function doGet(e) {
  return handleRequest(e, 'get');
}

function doPost(e) {
  return handleRequest(e, 'post');
}

function handleRequest(e, method) {
  try {
    if (method === 'get') {
      var action = e.parameter.action;
      if (action === 'list') return toJson(listMessages());
      if (action === 'setup') return toJson(setupSheet());
      if (action === 'add') return toJson(addMessage(e.parameter.name, e.parameter.message));
      return toJson({ success: false, message: '未知 action（可用：list, setup, add）' });
    }

    if (method === 'post') {
      var data = JSON.parse(e.postData.contents);
      if (data.action === 'add') return toJson(addMessage(data.name, data.message));
      return toJson({ success: false, message: '未知 action（可用：add）' });
    }

    return toJson({ success: false, message: '未知請求' });
  } catch (err) {
    return toJson({ success: false, message: err.message });
  }
}

function toJson(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function setupSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['時間', '姓名', '留言']);
    sheet.setFrozenRows(1);
  }
  return { success: true, message: '試算表已初始化' };
}

function listMessages() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) return { success: false, message: '未初始化，請先執行 setup' };
  var data = sheet.getDataRange().getValues();
  var messages = [];
  for (var i = 1; i < data.length; i++) {
    messages.push({ time: data[i][0], name: data[i][1], message: data[i][2] });
  }
  messages.reverse();
  return { success: true, messages: messages };
}

function authorizeWebApp() {
  var url = ScriptApp.getService().getUrl();
  Logger.log('Web App URL: ' + url);
  return '授權完成！Web App URL: ' + url;
}

function addMessage(name, message) {
  if (!name || !message) return { success: false, message: '請填寫姓名和留言' };
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) return { success: false, message: '未初始化，請先執行 setup' };
  var now = Utilities.formatDate(new Date(), 'Asia/Hong_Kong', 'yyyy-MM-dd HH:mm:ss');
  sheet.appendRow([now, name, message]);
  return { success: true, message: '留言已送出！' };
}
