/* 無人機小組・時間協調系統 — 後端
   部署方式：開一份新的 Google 試算表 → 上方選單「擴充功能 → Apps Script」→
   把這個檔案的內容整個貼進去（取代預設內容）→ 存檔 → 執行一次 init()（會跳出授權視窗，
   照著點下去即可，因為是你自己的程式）→ 部署 → 新增部署作業 → 類型選「網頁應用程式」→
   執行身份「我」→ 具有存取權的使用者「所有人」→ 部署 → 複製網址貼到 apps-script-config.js。
   詳細步驟見 README.md。 */

var MEMBER_COUNT = 9;
var TASK_COLS = ['id','group','title','owner','date','slot','status','pct','day','p0','p1','span','logs'];

/* 手動執行一次即可（也可以不執行，第一次有人打開網頁時會自動建立） */
function init(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSheets_(ss);
  Logger.log('資料表已就緒：' + ss.getUrl());
}

function doGet(e){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ensureSheets_(ss);
  return json_({
    members: readMembers_(sh.members),
    groups: readGroups_(sh.groups),
    tasks: readTasks_(sh.tasks)
  });
}

function doPost(e){
  var lock = LockService.getScriptLock();
  var result = {};
  try{
    lock.waitLock(10000);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var body = JSON.parse(e.postData.contents);
    var action = body.action, payload = body.payload || {};
    if(action === 'writeMember') writeMember_(ss, payload);
    else if(action === 'writeGroups') writeGroups_(ss, payload);
    else if(action === 'createTask') result.id = createTask_(ss, payload);
    else if(action === 'updateTask') updateTask_(ss, payload.id, payload.patch || {});
    else if(action === 'deleteTask') deleteTask_(ss, payload.id);
    else if(action === 'resetAll') resetAll_(ss);
    else if(action === 'uploadPhoto') result.url = uploadPhoto_(payload.dataUrl);
    else throw new Error('未知的操作：' + action);
    return json_({ ok: true, result: result });
  }catch(err){
    return json_({ ok: false, error: String((err && err.message) || err) });
  }finally{
    lock.releaseLock();
  }
}

/* ─────────── 資料表結構 ─────────── */
function ensureSheets_(ss){
  var members = ss.getSheetByName('Members');
  if(!members){
    members = ss.insertSheet('Members');
    members.appendRow(['id','name','grid','reason','imported','method','subj']);
    for(var i=0;i<MEMBER_COUNT;i++){
      members.appendRow([i, '組員 '+(i+1), JSON.stringify(blankGrid_(true)), JSON.stringify(blankGrid_('')), false, '', '']);
    }
  }
  var groups = ss.getSheetByName('Groups');
  if(!groups){
    groups = ss.insertSheet('Groups');
    groups.appendRow(['fuselage','wing','tail','locked']);
    groups.appendRow(['[]','[]','[]', false]);
  }
  var tasks = ss.getSheetByName('Tasks');
  if(!tasks){
    tasks = ss.insertSheet('Tasks');
    tasks.appendRow(TASK_COLS);
  }
  var def = ss.getSheetByName('Sheet1');
  if(def && ss.getSheets().length > 3) ss.deleteSheet(def);
  return { members: members, groups: groups, tasks: tasks };
}
function blankGrid_(v){
  var g = [];
  for (var d=0; d<7; d++){
    var row = [];
    for (var p=0; p<13; p++) row.push(v);
    g.push(row);
  }
  return g;
}
function safeParse_(s, def){ try{ return s ? JSON.parse(s) : def; }catch(e){ return def; } }
function json_(obj){
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ─────────── 讀取 ─────────── */
function readMembers_(sh){
  var rows = sh.getDataRange().getValues(); rows.shift();
  return rows.map(function(r){
    return { id:Number(r[0]), name:r[1], grid:safeParse_(r[2],blankGrid_(true)), reason:safeParse_(r[3],blankGrid_('')),
      imported:!!r[4], method:r[5]||'', subj: r[6] ? safeParse_(r[6],null) : null };
  });
}
function readGroups_(sh){
  var r = sh.getDataRange().getValues()[1] || ['[]','[]','[]',false];
  return { fuselage:safeParse_(r[0],[]), wing:safeParse_(r[1],[]), tail:safeParse_(r[2],[]), locked:!!r[3] };
}
function readTasks_(sh){
  var rows = sh.getDataRange().getValues(); rows.shift();
  return rows.filter(function(r){ return r[0]; }).map(function(r){
    return { id:String(r[0]), group:r[1], title:r[2], owner:r[3], date:r[4], slot:r[5], status:r[6],
      pct:Number(r[7])||0, day:Number(r[8])||0, p0:Number(r[9])||0, p1:Number(r[10])||0,
      span:Number(r[11])||1, logs:safeParse_(r[12],[]) };
  });
}

/* ─────────── 寫入 ─────────── */
function writeMember_(ss, p){
  var sh = ensureSheets_(ss).members;
  var row = Number(p.id) + 2;
  sh.getRange(row,1,1,7).setValues([[p.id, p.name, JSON.stringify(p.grid), JSON.stringify(p.reason),
    !!p.imported, p.method||'', p.subj ? JSON.stringify(p.subj) : '']]);
}
function writeGroups_(ss, p){
  var sh = ensureSheets_(ss).groups;
  sh.getRange(2,1,1,4).setValues([[JSON.stringify(p.fuselage||[]), JSON.stringify(p.wing||[]),
    JSON.stringify(p.tail||[]), !!p.locked]]);
}
function createTask_(ss, p){
  var sh = ensureSheets_(ss).tasks;
  sh.appendRow([p.id, p.group, p.title, p.owner, p.date, p.slot, '待開始', 0, p.day, p.p0, p.p1, p.span||1, '[]']);
  return p.id;
}
function findTaskRow_(sh, id){
  var last = sh.getLastRow();
  if(last < 2) return -1;
  var ids = sh.getRange(2,1,last-1,1).getValues();
  for(var i=0;i<ids.length;i++) if(String(ids[i][0]) === String(id)) return i+2;
  return -1;
}
function updateTask_(ss, id, patch){
  var sh = ensureSheets_(ss).tasks, row = findTaskRow_(sh, id);
  if(row < 0) throw new Error('找不到任務：' + id);
  TASK_COLS.forEach(function(key, i){
    if(Object.prototype.hasOwnProperty.call(patch, key)){
      var v = patch[key];
      if(key === 'logs') v = JSON.stringify(v);
      sh.getRange(row, i+1).setValue(v);
    }
  });
}
function deleteTask_(ss, id){
  var sh = ensureSheets_(ss).tasks, row = findTaskRow_(sh, id);
  if(row > 0) sh.deleteRow(row);
}
function resetAll_(ss){
  var sh = ensureSheets_(ss);
  for(var i=0;i<MEMBER_COUNT;i++){
    sh.members.getRange(i+2,1,1,7).setValues([[i, '組員 '+(i+1), JSON.stringify(blankGrid_(true)),
      JSON.stringify(blankGrid_('')), false, '', '']]);
  }
  sh.groups.getRange(2,1,1,4).setValues([['[]','[]','[]', false]]);
  if(sh.tasks.getLastRow() > 1) sh.tasks.deleteRows(2, sh.tasks.getLastRow()-1);
}

/* ─────────── 進度紀錄照片（存到 Google Drive，回傳可公開檢視的網址） ─────────── */
function uploadPhoto_(dataUrl){
  var m = /^data:(image\/[a-zA-Z0-9.+-]+);base64,(.*)$/.exec(dataUrl || '');
  if(!m) throw new Error('圖片格式錯誤');
  var mime = m[1], bytes = Utilities.base64Decode(m[2]);
  var blob = Utilities.newBlob(bytes, mime, 'log-photo-' + Date.now() + '.jpg');
  var file = getPhotoFolder_().createFile(blob);
  file.setSharingAccess(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return 'https://drive.google.com/uc?export=view&id=' + file.getId();
}
function getPhotoFolder_(){
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('PHOTO_FOLDER_ID'), folder = null;
  if(id){ try{ folder = DriveApp.getFolderById(id); }catch(e){ folder = null; } }
  if(!folder){
    folder = DriveApp.createFolder('無人機小組課表系統照片');
    props.setProperty('PHOTO_FOLDER_ID', folder.getId());
  }
  return folder;
}
