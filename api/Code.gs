// ============================================================
// Colour Your Life — API Backend（for GitHub Pages）
// ============================================================
var SPREADSHEET_ID      = '19r9JIEFHKsh_FF5Qw9WV4GXVsgQw2TBMxFM6h7i0E3Q';
var SHEET_NAME_GROUPS   = '組別';
var SHEET_NAME_STATIONS = '攤位';
var SHEET_NAME_RECORDS  = '記錄';
var SHEET_NAME_CONFIG   = '設定';
var ADMIN_PASSWORD_PROP = 'ADMIN_PASSWORD';

var GROUP_NAMES = [];
['A','B','C','D','E'].forEach(function(p) {
  var max = (p === 'A') ? 5 : 6;
  for (var i = 1; i <= max; i++) GROUP_NAMES.push(p + i);
});
var DEFAULT_PASSWORD = '123';

var RANKS = [
  { threshold: 0,    title: '特工見習生' },
  { threshold: 150,  title: '特工助理' },
  { threshold: 300,  title: '助理特工' },
  { threshold: 450,  title: '特工' },
  { threshold: 600,  title: '特工隊長' },
  { threshold: 750,  title: '特工首席隊長' },
  { threshold: 900,  title: '特工署助理署長' },
  { threshold: 1050, title: '特工署副署長' },
  { threshold: 1200, title: '特工署署長' },
];

function doGet(e) {
  try {
    var p = e.parameter;
    var action = p.action;
    var result;
    if (action === 'login')         result = loginGroup(p.groupId, p.password);
    else if (action === 'profile')  result = getGroupProfile(p.groupId);
    else if (action === 'leaderboard') result = getLeaderboard();
    else if (action === 'loginAdmin')  result = loginAdmin(p.password);
    else if (action === 'loginStation') result = loginStationAdmin(p.stationId, p.password);
    else if (action === 'pendingGroups') result = getStationPendingGroups(p.stationId);
    else if (action === 'allStations')   result = getAllStations();
    else if (action === 'allGroups')     result = getAllGroups();
    else if (action === 'allRecords')    result = getAllRecords();
    else if (action === 'addExp')        result = addExp(p.groupId, p.stationId, p.admin || 'admin');
    else if (action === 'addBonus')      result = addBonusExp(p.groupId, Number(p.exp), p.reason, p.admin || 'admin');
    else if (action === 'resetGroup')    result = resetGroup(p.groupId);
    else if (action === 'resetAll')      result = resetAllGroups();
    else if (action === 'setup')         result = setupSheets();
    else result = { success: false, message: '未知 action' };

    var json = JSON.stringify(result);
    if (p.callback) return ContentService.createTextOutput(p.callback + '(' + json + ')').setMimeType(ContentService.MimeType.JAVASCRIPT);
    return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    var json = JSON.stringify({ success: false, message: err.message });
    var cb = e && e.parameter && e.parameter.callback;
    if (cb) return ContentService.createTextOutput(cb + '(' + json + ')').setMimeType(ContentService.MimeType.JAVASCRIPT);
    return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
  }
}

function setupSheets() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var gs = ss.getSheetByName(SHEET_NAME_GROUPS);
  if (!gs) {
    gs = ss.insertSheet(SHEET_NAME_GROUPS);
    gs.appendRow(['groupId', '密碼', '總經驗值']);
    GROUP_NAMES.forEach(function(g) { gs.appendRow([g, DEFAULT_PASSWORD, 0]); });
    gs.setFrozenRows(1);
  }
  var st = ss.getSheetByName(SHEET_NAME_STATIONS);
  if (!st) {
    st = ss.insertSheet(SHEET_NAME_STATIONS);
    st.appendRow(['stationId', '地點', '名稱', '經驗值']);
    var STATIONS_DATA = [
      { id: 'S01', location: 'N301室',  name: 'VASK 天賦識別',                 exp: 120 },
      { id: 'S02', location: 'N301A室', name: '與熊隊長合照：懷疑與自信',      exp: 130 },
      { id: 'S03', location: 'N301室',  name: 'STEM 結構力學測試',              exp: 100 },
      { id: 'S04', location: 'N301室',  name: '錦鋰想點：生涯規劃桌上遊戲',    exp: 150 },
      { id: 'S05', location: 'N301A室', name: '能量補給：「小食部」站點',       exp: 110 },
      { id: 'S06', location: 'N301室',  name: '空間追蹤：功能房間定位',         exp: 150 },
      { id: 'S07', location: '禮堂',     name: '拔河',                            exp: 130 },
      { id: 'S08', location: '',         name: '主線任務：飛索',                  exp: 200 },
      { id: 'S09', location: '',         name: '主線任務：遊繩下降',              exp: 200 },
      { id: 'S99', location: '全校',     name: '🤝 隱藏任務：導師脈絡對接',      exp: 50 },
    ];
    STATIONS_DATA.forEach(function(s) { st.appendRow([s.id, s.location, s.name, s.exp]); });
    st.setFrozenRows(1);
  }
  var rs = ss.getSheetByName(SHEET_NAME_RECORDS);
  if (!rs) {
    rs = ss.insertSheet(SHEET_NAME_RECORDS);
    rs.appendRow(['recordId', 'groupId', 'stationId', '經驗值', '時間', '管理員']);
    rs.setFrozenRows(1);
  }
  rs = ss.getSheetByName(SHEET_NAME_RECORDS);
  if (rs && rs.getFrozenRows() === 0) rs.setFrozenRows(1);
  var cs = ss.getSheetByName(SHEET_NAME_CONFIG);
  if (!cs) {
    cs = ss.insertSheet(SHEET_NAME_CONFIG);
    cs.appendRow(['key', 'value']);
    cs.appendRow([ADMIN_PASSWORD_PROP, 'admin123']);
    cs.setFrozenRows(1);
  }
  return { success: true, message: '試算表已初始化' };
}

function getSheet(name) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(name);
  if (!sheet) setupSheets();
  return ss.getSheetByName(name);
}

function getSetting(key) {
  var sheet = getSheet(SHEET_NAME_CONFIG);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === key) return data[i][1];
  }
  return null;
}

function getRank(exp) {
  for (var i = RANKS.length - 1; i >= 0; i--) {
    if (exp >= RANKS[i].threshold) return { title: RANKS[i].title, threshold: RANKS[i].threshold };
  }
  return { title: RANKS[0].title, threshold: RANKS[0].threshold };
}

function getNextRank(exp) {
  for (var i = 0; i < RANKS.length - 1; i++) {
    if (exp < RANKS[i + 1].threshold) return RANKS[i + 1];
  }
  return null;
}

function loginGroup(groupId, password) {
  try {
    var sheet = getSheet(SHEET_NAME_GROUPS);
    if (!sheet) return { success: false, message: '系統錯誤：無法讀取組別資料表。' };
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).toUpperCase() === String(groupId).toUpperCase()) {
        if (String(data[i][1]) !== String(password)) return { success: false, message: '密碼錯誤。' };
        return { success: true, groupId: data[i][0], totalExp: data[i][2] || 0 };
      }
    }
    return { success: false, message: '找不到此組別編號。' };
  } catch (e) { return { success: false, message: '系統錯誤：' + e.message }; }
}

function getGroupProfile(groupId) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var gs = ss.getSheetByName(SHEET_NAME_GROUPS);
    var rs = ss.getSheetByName(SHEET_NAME_RECORDS);
    var stationsSheet = ss.getSheetByName(SHEET_NAME_STATIONS);
    if (!gs || !rs || !stationsSheet) { setupSheets(); gs = ss.getSheetByName(SHEET_NAME_GROUPS); rs = ss.getSheetByName(SHEET_NAME_RECORDS); stationsSheet = ss.getSheetByName(SHEET_NAME_STATIONS); if (!gs || !rs || !stationsSheet) return { success: false, message: '系統錯誤：無法讀取資料表。' }; }
    var gData = gs.getDataRange().getValues();
    var rData = rs.getDataRange().getValues();
    var sData = stationsSheet.getDataRange().getValues();
    var found = null;
    for (var i = 1; i < gData.length; i++) {
      if (String(gData[i][0]).toUpperCase() === String(groupId).toUpperCase()) {
        found = { groupId: gData[i][0], totalExp: gData[i][2] || 0 }; break;
      }
    }
    if (!found) return { success: false, message: '找不到組別。' };
    var validStationIds = {};
    for (var ks = 1; ks < sData.length; ks++) validStationIds[String(sData[ks][0]).toUpperCase()] = true;
    var completed = {}; var groupRecords = [];
    for (var j = 1; j < rData.length; j++) {
      if (String(rData[j][1]).toUpperCase() === String(found.groupId).toUpperCase()) {
        if (validStationIds[String(rData[j][2]).toUpperCase()]) completed[String(rData[j][2])] = true;
        groupRecords.push({ recordId: rData[j][0], stationId: rData[j][2], exp: rData[j][3], time: rData[j][4], admin: rData[j][5] });
      }
    }
    var stations = []; var allDone = true;
    for (var k = 1; k < sData.length; k++) {
      var sid = String(sData[k][0]); var isDone = !!completed[sid];
      stations.push({ stationId: sid, location: sData[k][1], name: sData[k][2], exp: sData[k][3], completed: isDone });
      if (!isDone) allDone = false;
    }
    var exp = found.totalExp; var rank = getRank(exp); var nextRank = getNextRank(exp);
    return { success: true, groupId: found.groupId, totalExp: exp, rank: rank.title, nextRank: nextRank ? nextRank.title : null, nextRankThreshold: nextRank ? nextRank.threshold : null, stations: stations, allCompleted: allDone, records: groupRecords };
  } catch (e) { return { success: false, message: '系統錯誤：' + e.message }; }
}

function getAllGroups() {
  try {
    var sheet = getSheet(SHEET_NAME_GROUPS); if (!sheet) return [];
    var data = sheet.getDataRange().getValues(); var result = [];
    for (var i = 1; i < data.length; i++) result.push({ groupId: data[i][0], totalExp: data[i][2] || 0 });
    return result;
  } catch (e) { return []; }
}

function getLeaderboard() {
  try {
    var sheet = getSheet(SHEET_NAME_GROUPS); if (!sheet) return [];
    var data = sheet.getDataRange().getValues(); var result = [];
    for (var i = 1; i < data.length; i++) { var exp = data[i][2] || 0; result.push({ groupId: data[i][0], totalExp: exp, rank: getRank(exp).title }); }
    result.sort(function(a, b) { return b.totalExp - a.totalExp; });
    return result;
  } catch (e) { return []; }
}

function loginAdmin(password) {
  try { var correct = getSetting(ADMIN_PASSWORD_PROP); return { success: password === correct }; } catch (e) { return { success: false, message: '系統錯誤：' + e.message }; }
}

function loginStationAdmin(stationId, password) {
  try {
    var correct = getSetting(ADMIN_PASSWORD_PROP);
    if (password !== correct) return { success: false, message: '密碼錯誤' };
    var sheet = getSheet(SHEET_NAME_STATIONS); if (!sheet) return { success: false, message: '系統錯誤：無法讀取攤位表。' };
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).toUpperCase() === String(stationId).toUpperCase()) {
        var exp = Number(data[i][3]);
        if (exp === 0) return { success: false, message: '此攤位不能直接加分，請用額外任務。' };
        return { success: true, station: { stationId: data[i][0], location: data[i][1], name: data[i][2], exp: exp } };
      }
    }
    return { success: false, message: '找不到攤位編號' };
  } catch (e) { return { success: false, message: '系統錯誤：' + e.message }; }
}

function getStationPendingGroups(stationId) {
  try {
    var gs = getSheet(SHEET_NAME_GROUPS); var rs = getSheet(SHEET_NAME_RECORDS); var ss = getSheet(SHEET_NAME_STATIONS);
    if (!gs || !rs) return [];
    var gData = gs.getDataRange().getValues(); var rData = rs.getDataRange().getValues();
    var validIds = {};
    if (ss) { var sd = ss.getDataRange().getValues(); for (var x = 1; x < sd.length; x++) validIds[String(sd[x][0]).toUpperCase()] = true; }
    var done = {};
    for (var i = 1; i < rData.length; i++) { var rsid = String(rData[i][2]).toUpperCase(); if (validIds[rsid] && rsid === String(stationId).toUpperCase()) done[String(rData[i][1]).toUpperCase()] = true; }
    var result = [];
    for (var j = 1; j < gData.length; j++) { var gid = String(gData[j][0]); if (!done[gid.toUpperCase()]) result.push({ groupId: gid, totalExp: gData[j][2] || 0 }); }
    result.sort(function(a, b) { return a.groupId.localeCompare(b.groupId); });
    return result;
  } catch (e) { return []; }
}

function getAllStations() {
  try {
    var sheet = getSheet(SHEET_NAME_STATIONS); if (!sheet) return [];
    var data = sheet.getDataRange().getValues(); var result = [];
    for (var i = 1; i < data.length; i++) result.push({ stationId: data[i][0], location: data[i][1], name: data[i][2], exp: data[i][3] });
    return result;
  } catch (e) { return []; }
}

function getAllRecords() {
  try {
    var sheet = getSheet(SHEET_NAME_RECORDS); if (!sheet) return [];
    var data = sheet.getDataRange().getValues(); var result = [];
    for (var i = 1; i < data.length; i++) result.push({ recordId: data[i][0], groupId: data[i][1], stationId: data[i][2], exp: data[i][3], time: data[i][4], admin: data[i][5] });
    return result;
  } catch (e) { return []; }
}

function addExp(groupId, stationId, adminName) {
  try {
    var stationsSheet = getSheet(SHEET_NAME_STATIONS); if (!stationsSheet) return { success: false, message: '系統錯誤：無法讀取攤位表。' };
    var sData = stationsSheet.getDataRange().getValues(); var exp = 0; var stationName = ''; var validIds = {};
    for (var i = 1; i < sData.length; i++) { validIds[String(sData[i][0]).toUpperCase()] = true; if (String(sData[i][0]) === String(stationId)) { exp = Number(sData[i][3]); stationName = sData[i][2]; } }
    if (exp === 0) return { success: false, message: '找不到攤位或經驗值為 0。' };
    var recordsSheet = getSheet(SHEET_NAME_RECORDS); if (!recordsSheet) return { success: false, message: '系統錯誤：無法讀取記錄表。' };
    var rData = recordsSheet.getDataRange().getValues();
    for (var j = 1; j < rData.length; j++) {
      // Only match against known station IDs (skip bonus/other records)
      var rid = String(rData[j][2]);
      if (String(rData[j][1]).toUpperCase() === String(groupId).toUpperCase() && rid === String(stationId) && validIds[rid.toUpperCase()]) {
        return { success: false, message: '此組別已完成此攤位，不能重複加分。' };
      }
    }
    var timestamp = Utilities.formatDate(new Date(), 'Asia/Hong_Kong', 'yyyy-MM-dd HH:mm:ss');
    var recordId = 'R' + new Date().getTime();
    recordsSheet.appendRow([recordId, groupId, stationId, exp, timestamp, adminName]);
    var gs = getSheet(SHEET_NAME_GROUPS); if (!gs) return { success: false, message: '系統錯誤：無法讀取組別表。' };
    var gData = gs.getDataRange().getValues();
    for (var k = 1; k < gData.length; k++) { if (String(gData[k][0]).toUpperCase() === String(groupId).toUpperCase()) { var newExp = (Number(gData[k][2]) || 0) + exp; gs.getRange(k + 1, 3).setValue(newExp); break; } }
    return { success: true, message: '已為 ' + groupId + ' 加入 ' + exp + ' EXP（' + stationName + '）' };
  } catch (e) { return { success: false, message: '系統錯誤：' + e.message }; }
}

function addBonusExp(groupId, exp, reason, adminName) {
  try {
    if (!exp || exp <= 0) return { success: false, message: '請輸入有效的經驗值。' };
    if (!reason) return { success: false, message: '請輸入加分原因。' };
    var recordsSheet = getSheet(SHEET_NAME_RECORDS); if (!recordsSheet) return { success: false, message: '系統錯誤：無法讀取記錄表。' };
    var timestamp = Utilities.formatDate(new Date(), 'Asia/Hong_Kong', 'yyyy-MM-dd HH:mm:ss');
    var recordId = 'R' + new Date().getTime();
    recordsSheet.appendRow([recordId, groupId, reason, exp, timestamp, adminName]);
    var gs = getSheet(SHEET_NAME_GROUPS); if (!gs) return { success: false, message: '系統錯誤：無法讀取組別表。' };
    var gData = gs.getDataRange().getValues();
    for (var k = 1; k < gData.length; k++) { if (String(gData[k][0]).toUpperCase() === String(groupId).toUpperCase()) { var newExp = (Number(gData[k][2]) || 0) + exp; gs.getRange(k + 1, 3).setValue(newExp); break; } }
    return { success: true, message: '已為 ' + groupId + ' 額外加入 ' + exp + ' EXP（' + reason + '）' };
  } catch (e) { return { success: false, message: '系統錯誤：' + e.message }; }
}

function resetGroup(groupId) {
  try {
    var rs = getSheet(SHEET_NAME_RECORDS); if (!rs) return { success: false, message: '系統錯誤：無法讀取記錄表。' };
    var rData = rs.getDataRange().getValues(); var rowsToDelete = [];
    for (var i = 1; i < rData.length; i++) { if (String(rData[i][1]).toUpperCase() === String(groupId).toUpperCase()) rowsToDelete.push(i + 1); }
    for (var m = rowsToDelete.length - 1; m >= 0; m--) { rs.deleteRow(rowsToDelete[m]); }
    var gs = getSheet(SHEET_NAME_GROUPS); if (!gs) return { success: false, message: '系統錯誤：無法讀取組別表。' };
    var gData = gs.getDataRange().getValues();
    for (var n = 1; n < gData.length; n++) { if (String(gData[n][0]).toUpperCase() === String(groupId).toUpperCase()) { gs.getRange(n + 1, 3).setValue(0); break; } }
    return { success: true, message: '已重置 ' + groupId + ' 的所有記錄。' };
  } catch (e) { return { success: false, message: '系統錯誤：' + e.message }; }
}

function resetAllGroups() {
  try {
    var rs = getSheet(SHEET_NAME_RECORDS); if (!rs) return { success: false, message: '系統錯誤：無法讀取記錄表。' };
    var rData = rs.getDataRange().getValues();
    for (var i = rData.length - 1; i >= 1; i--) { rs.deleteRow(i + 1); }
    var gs = getSheet(SHEET_NAME_GROUPS); if (!gs) return { success: false, message: '系統錯誤：無法讀取組別表。' };
    var gData = gs.getDataRange().getValues();
    for (var j = 1; j < gData.length; j++) { gs.getRange(j + 1, 3).setValue(0); }
    return { success: true, message: '已重置所有組別記錄。' };
  } catch (e) { return { success: false, message: '系統錯誤：' + e.message }; }
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Colour Your Life').addItem('初始化試算表', 'setupSheets').addToUi();
}
