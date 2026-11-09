/**
 * 安心呷｜雲端後台（Google 登入＋照護團隊資料同步）  2026/11/09 版
 *
 * 使用方式：在 Apps Script 編輯器裡「全選 → 刪除」原本的程式碼，整份貼上這個檔案。
 * 部署：「部署 → 管理部署作業 → 編輯（鉛筆）→ 版本：新版本 → 部署」，網址不會變。
 *
 * 試算表的工作表：
 *   帳號     ：帳號、名字、關係、密碼雜湊（舊帳號）、照護團隊、建立時間、Email、同意研究聯絡、Google ID
 *   照護團隊 ：團隊邀請碼、主要管理者、建立時間
 *   紀錄     ：每一筆餐飲／身體、看診、長輩資料、照顧者（同一個團隊的人共用）
 * 照片會存到雲端硬碟的「安心呷 照片」資料夾。
 */

var SS = SpreadsheetApp.getActiveSpreadsheet();
var MAX_CELL = 49000;
// Google Cloud › Google Auth Platform › 用戶端 ›「安心呷網頁版」的用戶端 ID
var GOOGLE_CLIENT_ID = '449399542925-71t2elt3dmbd97mflc623mce93euijk7.apps.googleusercontent.com';

function doPost(e) {
  var d;
  try { d = JSON.parse(e.postData.contents); } catch (err) { return out_({ ok: false, error: 'bad_json' }); }
  if (!d.action) return logLogin_(d);               // 登入紀錄
  try {
    var fn = API[d.action];
    if (!fn) return out_({ ok: false, error: 'unknown_action' });
    return out_(fn(d));
  } catch (err) {
    var msg = String(err && err.message || err);
    return out_({ ok: false, error: /^[a-z_]+$/.test(msg) ? msg : 'server_error', detail: msg.slice(0, 200) });
  }
}

function doGet() {
  return ContentService.createTextOutput('安心呷 雲端後台運作中');
}

/* ---------------- API ---------------- */
var API = {
  // 只允許 Google 登入：帳號密碼的登入、註冊、重設密碼都關閉（已登入的人不受影響）
  register: function () { throw new Error('google_only'); },
  login: function () { throw new Error('google_only'); },
  reset: function () { throw new Error('google_only'); },

  // Google 登入：已註冊 → 直接登入；第一次 → 請 App 顯示「完成註冊」
  google: function (d) {
    var g = verifyGoogle_(d.idToken);
    var row = findGoogleRow_(g);
    if (!row) return { ok: false, error: 'need_profile', email: g.email, name: g.name };
    return googleSession_(String(sheet_('帳號').getRange(row, 1).getValue()), g);
  },

  // 第一次用 Google：建立帳號（帳號名稱就是 Gmail），可輸入照護團隊邀請碼
  googleRegister: function (d) {
    var g = verifyGoogle_(d.idToken);
    var name = clean_(d.name, 20), role = clean_(d.role, 12);
    if (!name) throw new Error('missing_fields');
    return withLock_(function () {
      var row = findGoogleRow_(g);
      if (row) return googleSession_(String(sheet_('帳號').getRange(row, 1).getValue()), g);
      var acc = normAcc_(g.email);
      if (findAccount_(acc)) throw new Error('account_taken');      // 以前用這個 Gmail 當帳號註冊過 → 請他連結舊帳號
      var team = String(d.team || '').trim().toUpperCase();
      if (team) { if (!findTeam_(team)) throw new Error('team_not_found'); }
      else { team = newTeamCode_(); sheet_('照護團隊').appendRow([team, acc, now_()]); }
      sheet_('帳號').appendRow([acc, name, role, '', team, now_(), g.email, d.contact ? '同意' : '', g.sub]);
      return googleSession_(acc, g);
    });
  },

  // 舊的帳號密碼使用者：驗證一次舊密碼，把舊帳號綁到這個 Google 帳號
  googleLink: function (d) {
    var g = verifyGoogle_(d.idToken);
    return withLock_(function () {
      if (findGoogleRow_(g)) throw new Error('already_linked');
      var a = findAccount_(normAcc_(d.account));
      if (!a) throw new Error('account_not_found');
      if (!a.pwd || a.pwd !== pwd_(d.hash)) throw new Error('wrong_password');
      var sh = sheet_('帳號');
      if (String(sh.getRange(a.row, 9).getValue())) throw new Error('already_linked');
      sh.getRange(a.row, 7).setValue(g.email);
      sh.getRange(a.row, 9).setValue(g.sub);
      return googleSession_(a.acc, g);
    });
  },

  // 查詢自己的 Email（舊帳號沒有 Email 時，App 會請他補填）
  me: function (d) {
    var s = check_(d.token);
    var a = findAccount_(s.acc);
    if (!a) throw new Error('account_not_found');
    return { ok: true, email: a.email, contact: a.contact === '同意', google: !!a.sub };
  },

  updateMe: function (d) {
    var s = check_(d.token);
    return withLock_(function () {
      var a = findAccount_(s.acc);
      if (!a) throw new Error('account_not_found');
      var sh = sheet_('帳號');
      if (d.name) sh.getRange(a.row, 2).setValue(clean_(d.name, 20));
      if (d.role) sh.getRange(a.row, 3).setValue(clean_(d.role, 12));
      if ('email' in d && !a.sub) {                               // Google 帳號的 Email 不能改
        var em = cleanEmail_(d.email);
        if (!em) throw new Error('bad_email');
        sh.getRange(a.row, 7).setValue(em);
      }
      if ('contact' in d) sh.getRange(a.row, 8).setValue(d.contact ? '同意' : '');
      return { ok: true, members: members_(s.team) };
    });
  },

  sync: function (d) {
    var s = check_(d.token);
    var since = Number(d.since) || 0;
    var changes = Array.isArray(d.changes) ? d.changes.slice(0, 500) : [];
    return withLock_(function () {
      var sh = sheet_('紀錄');
      var vals = sh.getDataRange().getValues();
      var index = {};
      for (var i = 1; i < vals.length; i++) {
        if (String(vals[i][0]) === s.team) index[vals[i][1] + '|' + String(vals[i][2])] = i;
      }
      var ts = Date.now(), appended = [];
      changes.forEach(function (c) {
        var kind = clean_(c.kind, 10), id = clean_(c.id, 60);
        if (!/^(meal|visit|care|elder)$/.test(kind) || !id) return;
        var json = c.deleted ? '' : JSON.stringify(c.data || {});
        if (json.length > MAX_CELL) return;                      // 太大（照片應先上傳）
        var row = [s.team, kind, id, Number(c.updatedAt) || ts, c.deleted ? 1 : 0, json, s.acc, ts];
        var at = index[kind + '|' + id];
        if (at === -1) return;                                   // 同一批重複
        if (at !== undefined) {
          if (Number(vals[at][3]) > row[3]) return;              // 雲端已有較新的版本
          sh.getRange(at + 1, 1, 1, 8).setValues([row]);
          vals[at] = row;
        } else {
          appended.push(row);
          index[kind + '|' + id] = -1;
        }
      });
      if (appended.length) sh.getRange(sh.getLastRow() + 1, 1, appended.length, 8).setValues(appended);
      var all = vals.slice(1).concat(appended);
      var recs = all.filter(function (r) { return String(r[0]) === s.team && Number(r[7]) > since; })
        .map(function (r) {
          return { kind: r[1], id: String(r[2]), updatedAt: Number(r[3]), deleted: !!Number(r[4]),
                   data: r[5] ? JSON.parse(r[5]) : null, by: r[6] };
        });
      return { ok: true, now: ts, records: recs, members: members_(s.team), team: s.team };
    });
  },

  upload: function (d) {
    var s = check_(d.token);
    var m = String(d.dataUrl || '').match(/^data:(image\/(jpeg|png|webp));base64,(.+)$/);
    if (!m) throw new Error('bad_image');
    var bytes = Utilities.base64Decode(m[3]);
    if (bytes.length > 3 * 1024 * 1024) throw new Error('image_too_large');
    var blob = Utilities.newBlob(bytes, m[1], s.team + '-' + Date.now() + '.' + (m[2] === 'jpeg' ? 'jpg' : m[2]));
    var file = photoFolder_().createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return { ok: true, url: 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w1200' };
  },

  // 下載週報：讀取團隊自己上傳的照片（只能讀自己團隊的）
  photo: function (d) {
    var s = check_(d.token);
    var m = String(d.url || '').match(/[?&]id=([\w-]+)/);
    if (!m) throw new Error('bad_image');
    var file = DriveApp.getFileById(m[1]);
    if (file.getName().indexOf(s.team + '-') !== 0) throw new Error('bad_image');
    var b = file.getBlob();
    return { ok: true, dataUrl: 'data:' + b.getContentType() + ';base64,' + Utilities.base64Encode(b.getBytes()) };
  },

  // 主要管理者：移出成員
  removeMember: function (d) {
    var s = check_(d.token);
    return withLock_(function () {
      requireAdmin_(s);
      var acc = normAcc_(d.account);
      if (acc === s.acc) throw new Error('cannot_remove_self');
      var a = findAccount_(acc);
      if (!a || a.team !== s.team) throw new Error('not_in_team');
      sheet_('帳號').getRange(a.row, 5).setValue('');
      return { ok: true, members: members_(s.team) };
    });
  },

  // 主要管理者：把管理者交給別人
  setAdmin: function (d) {
    var s = check_(d.token);
    return withLock_(function () {
      requireAdmin_(s);
      var a = findAccount_(normAcc_(d.account));
      if (!a || a.team !== s.team) throw new Error('not_in_team');
      sheet_('照護團隊').getRange(teamRow_(s.team).row, 2).setValue(a.acc);
      return { ok: true, members: members_(s.team) };
    });
  }
};

/* ---------------- Google 登入 ---------------- */
// 向 Google 確認登入憑證是真的、是發給安心呷的、還沒過期
function verifyGoogle_(idToken) {
  if (!idToken) throw new Error('google_failed');
  var res = UrlFetchApp.fetch('https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(idToken), { muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) throw new Error('google_failed');
  var t = JSON.parse(res.getContentText());
  if (t.aud !== GOOGLE_CLIENT_ID) throw new Error('google_failed');
  if (t.iss !== 'accounts.google.com' && t.iss !== 'https://accounts.google.com') throw new Error('google_failed');
  if (Number(t.exp) * 1000 < Date.now()) throw new Error('google_failed');
  if (String(t.email_verified) !== 'true') throw new Error('google_failed');
  return { sub: String(t.sub), email: String(t.email || '').toLowerCase(), name: String(t.name || t.given_name || '') };
}
// 在「帳號」工作表 I 欄找這個 Google 帳號（找不到回傳 0）
function findGoogleRow_(g) {
  var sh = sheet_('帳號'), last = sh.getLastRow();
  if (last < 2) return 0;
  var v = sh.getRange(1, 9, last, 1).getValues();
  for (var i = 1; i < v.length; i++) if (String(v[i][0]) === g.sub) return i + 1;
  return 0;
}
function googleSession_(acc, g) {
  var a = findAccount_(acc);
  if (!a) throw new Error('account_not_found');
  if (!a.team) {                                              // 被移出團隊的人重新登入：給他一個新的照護團隊
    a.team = newTeamCode_();
    sheet_('照護團隊').appendRow([a.team, a.acc, now_()]);
    sheet_('帳號').getRange(a.row, 5).setValue(a.team);
  }
  var r = session_(a.acc, a.name, a.role, a.team);
  r.email = g.email;
  r.google = true;
  return r;
}

/* ---------------- 帳號與團隊 ---------------- */
var HEADERS = {
  '帳號': ['帳號', '名字', '與長輩關係', '密碼雜湊', '照護團隊', '建立時間', 'Email', '同意研究聯絡', 'Google ID'],
  '照護團隊': ['團隊邀請碼', '主要管理者', '建立時間'],
  '紀錄': ['照護團隊', '類型', 'ID', '更新時間', '已刪除', '內容', '最後修改者', '伺服器時間']
};
function sheet_(name) {
  var sh = SS.getSheetByName(name);
  if (!sh) {
    sh = SS.insertSheet(name);
    sh.getRange(1, 1, sh.getMaxRows(), HEADERS[name].length).setNumberFormat('@');   // 全部當文字，避免帳號 0912… 或邀請碼被當成數字
    if (name === '紀錄') sh.getRange(1, 4, sh.getMaxRows(), 2).setNumberFormat('0'), sh.getRange(1, 8, sh.getMaxRows(), 1).setNumberFormat('0');
    sh.appendRow(HEADERS[name]);
    sh.setFrozenRows(1);
    if (name === '帳號') sh.hideColumns(4);
  }
  return sh;
}
function findAccount_(acc) {
  var v = sheet_('帳號').getDataRange().getValues();
  for (var i = 1; i < v.length; i++) {
    if (String(v[i][0]) === acc) return { row: i + 1, acc: acc, name: String(v[i][1]), role: String(v[i][2]), pwd: String(v[i][3]),
      team: String(v[i][4]), email: String(v[i][6] || ''), contact: String(v[i][7] || ''), sub: String(v[i][8] || '') };
  }
  return null;
}
function findTeam_(code) {
  var v = sheet_('照護團隊').getDataRange().getValues();
  for (var i = 1; i < v.length; i++) if (String(v[i][0]) === code) return true;
  return false;
}
function teamRow_(team) {
  var v = sheet_('照護團隊').getDataRange().getValues();
  for (var i = 1; i < v.length; i++) if (String(v[i][0]) === team) return { row: i + 1, admin: String(v[i][1]) };
  return null;
}
function requireAdmin_(s) {
  var t = teamRow_(s.team);
  if (!t || t.admin !== s.acc) throw new Error('not_admin');
}
function newTeamCode_() {
  var abc = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789', code;
  do { code = ''; for (var i = 0; i < 6; i++) code += abc.charAt(Math.floor(Math.random() * abc.length)); }
  while (findTeam_(code));
  return code;
}
// 成員名單：多帶 admin（主要管理者）；管理者已離開時，由最早加入的成員接手
function members_(team) {
  var v = sheet_('帳號').getDataRange().getValues();
  var list = v.slice(1).filter(function (r) { return String(r[4]) === team; })
    .map(function (r) { return { account: String(r[0]), name: String(r[1]), role: String(r[2]) }; });
  var t = teamRow_(team), admin = t ? t.admin : '';
  if (list.length && !list.some(function (m) { return m.account === admin; })) {
    admin = list[0].account;
    if (t) sheet_('照護團隊').getRange(t.row, 2).setValue(admin);
  }
  list.forEach(function (m) { if (m.account === admin) m.admin = true; });
  return list;
}
function session_(acc, name, role, team) {
  return { ok: true, token: token_(acc, team), account: acc, name: name, role: role, team: team, members: members_(team) };
}

/* ---------------- 安全 ---------------- */
function secret_() {
  var p = PropertiesService.getScriptProperties(), s = p.getProperty('SECRET');
  if (!s) { s = Utilities.getUuid() + Utilities.getUuid(); p.setProperty('SECRET', s); }
  return s;
}
function hmac_(str) {
  return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(str, secret_()));
}
function pwd_(clientHash) { return hmac_('pwd|' + String(clientHash || '')); }
function token_(acc, team) {
  var body = Utilities.base64EncodeWebSafe(JSON.stringify({ a: acc, t: team }), Utilities.Charset.UTF_8);
  return body + '~' + hmac_('tok|' + body);
}
// 每次使用都確認這個人還在這個照護團隊裡（被移出後就無法再讀寫）
function check_(token) {
  var parts = String(token || '').split('~');
  if (parts.length !== 2 || hmac_('tok|' + parts[0]) !== parts[1]) throw new Error('bad_token');
  var j = JSON.parse(Utilities.newBlob(Utilities.base64DecodeWebSafe(parts[0])).getDataAsString('UTF-8'));
  var a = findAccount_(j.a);
  if (!a) throw new Error('bad_token');
  if (a.team !== j.t) throw new Error('removed');
  return { acc: j.a, team: j.t };
}
function withLock_(fn) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) throw new Error('busy');
  try { return fn(); } finally { lock.releaseLock(); }
}
function photoFolder_() {
  var p = PropertiesService.getScriptProperties(), id = p.getProperty('PHOTO_FOLDER');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  var f = DriveApp.createFolder('安心呷 照片');
  p.setProperty('PHOTO_FOLDER', f.getId());
  return f;
}

/* ---------------- 工具 ---------------- */
function normAcc_(v) { return clean_(v, 60).toLowerCase(); }
function now_() { return Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy/MM/dd HH:mm:ss'); }
// 限制長度，並避免使用者輸入被試算表當成公式執行
function clean_(v, max) {
  var s = String(v == null ? '' : v).trim().slice(0, max || 200);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}
function cleanEmail_(v) {
  v = String(v || '').trim().toLowerCase().slice(0, 100);
  return /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(v) ? v : '';
}
function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------------- 登入紀錄（第一個工作表） ---------------- */
function logLogin_(d) {
  var sheet = SS.getSheets()[0];
  var time = d.time ? new Date(d.time) : new Date();
  sheet.appendRow([
    Utilities.formatDate(time, 'Asia/Taipei', 'yyyy/MM/dd HH:mm:ss'),
    clean_(d.account), clean_(d.method), clean_(d.lang), clean_(d.font),
    clean_(d.platform), clean_(d.version), clean_(d.device)
  ]);
  return ContentService.createTextOutput('ok');
}

/* 貼上新程式碼後，在編輯器上方選「setup」按「執行」一次：補上新欄位標題，並完成授權 */
function setup() {
  sheet_('帳號'); sheet_('照護團隊'); sheet_('紀錄'); photoFolder_(); secret_();
  ['帳號', '照護團隊'].forEach(function (name) {
    var sh = sheet_(name), h = HEADERS[name];
    sh.getRange(1, 1, 1, h.length).setValues([h]);
    sh.getRange(1, 1, sh.getMaxRows(), h.length).setNumberFormat('@');
  });
  UrlFetchApp.fetch('https://oauth2.googleapis.com/tokeninfo?id_token=x', { muteHttpExceptions: true });   // 觸發「連線到外部服務」授權
}
