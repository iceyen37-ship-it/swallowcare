/**
 * 安心呷｜雲端後台（Google 登入＋照護團隊資料同步）  2026/11/10 通知版
 *
 * 使用方式：在 Apps Script 編輯器裡「全選 → 刪除」原本的程式碼，整份貼上這個檔案。
 * 部署：「部署 → 管理部署作業 → 編輯（鉛筆）→ 版本：新版本 → 部署」，網址不會變。
 *
 * 試算表的工作表：
 *   帳號     ：帳號、名字、關係、密碼雜湊（舊帳號）、照護團隊、建立時間、Email、同意研究聯絡、Google ID
 *   照護團隊 ：團隊邀請碼、主要管理者、建立時間
 *   紀錄     ：每一筆餐飲／身體、看診、長輩資料、照顧者（同一個團隊的人共用）
 *   推播     ：每支手機的通知代碼（Token）、帳號、平台、語言
 * 照片會存到雲端硬碟的「安心呷 照片」資料夾。
 *
 * 通知（推播）需要：專案設定 › 指令碼屬性 新增「FCM_SA」＝ Firebase 服務帳戶金鑰 JSON 全文，
 * 再執行一次 setupPush（建立「回診提醒」與「新版本通知」的定時觸發）。
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
    var fresh = [];
    var result = withLock_(function () {
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
          if (kind === 'meal' && !c.deleted && ts - row[3] < 6 * 3600 * 1000) fresh.push(c.data || {});
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
    if (fresh.length) { try { notifyNewRecords_(s, fresh); } catch (e) {} }
    return result;
  },

  // 通知：記下這支手機的通知代碼（同一個代碼換人登入時，改成新的帳號）
  pushRegister: function (d) {
    var s = check_(d.token);
    var tok = String(d.push || '').trim();
    if (!/^[\w:\-]{20,400}$/.test(tok)) throw new Error('bad_push');
    var plat = d.platform === 'android' ? 'android' : 'web';
    var lang = /^(zh|en|id|fil|vi)$/.test(d.lang) ? d.lang : 'zh';
    return withLock_(function () {
      var sh = sheet_('推播'), v = sh.getDataRange().getValues();
      for (var i = 1; i < v.length; i++) {
        if (String(v[i][0]) === tok) { sh.getRange(i + 1, 1, 1, 5).setValues([[tok, s.acc, plat, lang, now_()]]); return { ok: true }; }
      }
      sh.appendRow([tok, s.acc, plat, lang, now_()]);
      return { ok: true };
    });
  },
  // 登出或關閉通知：刪掉這支手機的代碼
  pushUnregister: function (d) {
    var tok = String(d.push || '').trim();
    if (!tok) return { ok: true };
    return withLock_(function () { removePushTokens_([tok]); return { ok: true }; });
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

  // 輸入邀請碼加入其他照護團隊（原本團隊的紀錄留在原團隊）
  joinTeam: function (d) {
    var s = check_(d.token);
    var team = String(d.team || '').trim().toUpperCase();
    if (!/^[A-Z0-9]{6}$/.test(team)) throw new Error('team_not_found');
    return withLock_(function () {
      if (!findTeam_(team)) throw new Error('team_not_found');
      if (team === s.team) throw new Error('same_team');
      var a = findAccount_(s.acc);
      if (!a) throw new Error('account_not_found');
      sheet_('帳號').getRange(a.row, 5).setValue(team);
      var r = session_(a.acc, a.name, a.role, team);
      r.email = a.email; r.google = !!a.sub;
      return r;
    });
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
  '紀錄': ['照護團隊', '類型', 'ID', '更新時間', '已刪除', '內容', '最後修改者', '伺服器時間'],
  '推播': ['Token', '帳號', '平台', '語言', '更新時間']
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


/* ---------------- 通知（推播） ---------------- */
var PUSH_TEXT = {
  rec:   { zh: '{who} 新增了{elder}的紀錄：{item}', en: '{who} added a record for {elder}: {item}', id: '{who} menambahkan catatan untuk {elder}: {item}',
           fil: 'Nagdagdag si {who} ng tala para kay {elder}: {item}', vi: '{who} đã thêm ghi chép cho {elder}: {item}' },
  visit: { zh: '明天是{elder}的回診日{place}，記得帶上要問醫生的問題。', en: 'Tomorrow is {elder}\'s clinic visit{place}.', id: 'Besok jadwal kontrol {elder}{place}.',
           fil: 'Bukas ang check-up ni {elder}{place}.', vi: 'Ngày mai là lịch tái khám của {elder}{place}.' },
  webUp: { zh: '安心呷有新版本 v{v}，打開安心呷，按上方的「立即更新」。', en: 'Ansimjia v{v} is available. Open the app and tap “Update now”.', id: 'Ansimjia v{v} tersedia. Buka aplikasi dan ketuk “Perbarui”.',
           fil: 'May bagong Ansimjia v{v}. Buksan ang app at i-tap ang “I-update”.', vi: 'Đã có Ansimjia v{v}. Mở ứng dụng và chạm “Cập nhật”.' },
  andUp: { zh: '安心呷有新版本 v{v}，請到 Google Play 更新。', en: 'Ansimjia v{v} is available on Google Play.', id: 'Ansimjia v{v} tersedia di Google Play.',
           fil: 'May bagong Ansimjia v{v} sa Google Play.', vi: 'Ansimjia v{v} đã có trên Google Play.' }
};
var WEB_APP_URL = 'https://iceyen37-ship-it.github.io/swallowcare/';
var VERSION_URL = WEB_APP_URL + 'version.json';

function fill_(t, o) { return t.replace(/\{(\w+)\}/g, function (_, k) { return o[k] == null ? '' : o[k]; }); }

// 有人新增飲食／身體紀錄 → 通知同一個照護團隊的其他人（同一個人 10 分鐘內只通知一次）
function notifyNewRecords_(s, items) {
  var cache = CacheService.getScriptCache(), key = 'rec|' + s.team + '|' + s.acc;
  if (cache.get(key)) return;
  cache.put(key, '1', 600);
  var me = findAccount_(s.acc), who = me ? me.name : '';
  var names = items.map(function (x) { return String(x.name || '').trim(); }).filter(String);
  var item = names.slice(0, 3).join('、') + (names.length > 3 ? '…' : '');
  var elder = elderName_(s.team);
  var to = members_(s.team).map(function (m) { return m.account; }).filter(function (a) { return a !== s.acc; });
  sendPush_(to, function (lang) {
    return { title: '安心呷', body: fill_(PUSH_TEXT.rec[lang], { who: who, elder: elder, item: item || '—' }), link: WEB_APP_URL + '?go=meal' };
  });
}
function elderName_(team) {
  var v = sheet_('紀錄').getDataRange().getValues();
  for (var i = v.length - 1; i >= 1; i--) {
    if (String(v[i][0]) === team && v[i][1] === 'elder' && v[i][5]) { try { return JSON.parse(v[i][5]).name || '長輩'; } catch (e) {} }
  }
  return '長輩';
}

// 每天晚上：明天有回診的團隊，通知全部成員（setupPush 會建立這個定時觸發）
function visitReminder() {
  var tomorrow = Utilities.formatDate(new Date(Date.now() + 24 * 3600 * 1000), 'Asia/Taipei', 'yyyy-MM-dd');
  var v = sheet_('紀錄').getDataRange().getValues(), teams = {};
  for (var i = 1; i < v.length; i++) {
    if (v[i][1] !== 'visit' || Number(v[i][4]) || !v[i][5]) continue;
    var d; try { d = JSON.parse(v[i][5]); } catch (e) { continue; }
    if (d.date === tomorrow && !teams[v[i][0]]) teams[v[i][0]] = d;
  }
  Object.keys(teams).forEach(function (team) {
    var d = teams[team], place = [d.clinic, d.doctor].filter(String).join(' ');
    var elder = elderName_(team);
    sendPush_(members_(team).map(function (m) { return m.account; }), function (lang) {
      return { title: '安心呷 回診提醒', body: fill_(PUSH_TEXT.visit[lang], { elder: elder, place: place ? '（' + place + '）' : '' }), link: WEB_APP_URL + '?go=visit' };
    });
  });
}

// 每小時：version.json 的版本變新 → 通知所有人更新（網頁版、Android 分開）
function checkVersionPush() {
  var res = UrlFetchApp.fetch(VERSION_URL + '?t=' + Date.now(), { muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) return;
  var v = JSON.parse(res.getContentText()), p = PropertiesService.getScriptProperties();
  [['web', 'webUp'], ['android', 'andUp']].forEach(function (x) {
    var plat = x[0], cur = String(v[plat] || ''), last = p.getProperty('PUSHED_' + plat);
    if (!cur || cur === last) return;
    p.setProperty('PUSHED_' + plat, cur);
    if (!last) return;                                         // 第一次執行只記下目前版本，不發通知
    sendPush_(null, function (lang) {
      return { title: '安心呷', body: fill_(PUSH_TEXT[x[1]][lang], { v: cur }), link: WEB_APP_URL };
    }, plat);
  });
}

// 測試：只發給自己（執行這個 Apps Script 的 Google 帳號）的手機
function testPush() {
  var me = String(Session.getEffectiveUser().getEmail() || '').toLowerCase();
  var n = sendPush_([me], function () { return { title: '安心呷', body: '通知測試成功！', link: WEB_APP_URL }; });
  Logger.log('已送出 ' + n + ' 則');
}

// 送出通知：accounts = 帳號清單（null 代表全部），build(lang) 回傳 {title, body, link}
function sendPush_(accounts, build, onlyPlatform) {
  var sh = sheet_('推播'), rows = sh.getDataRange().getValues().slice(1);
  var want = accounts ? accounts.reduce(function (o, a) { o[a] = 1; return o; }, {}) : null;
  rows = rows.filter(function (r) { return r[0] && (!want || want[String(r[1])]) && (!onlyPlatform || r[2] === onlyPlatform); });
  if (!rows.length) return 0;
  var sa = JSON.parse(PropertiesService.getScriptProperties().getProperty('FCM_SA') || 'null');
  if (!sa) throw new Error('no_fcm_key');
  var auth = 'Bearer ' + fcmAccessToken_(sa);
  var url = 'https://fcm.googleapis.com/v1/projects/' + sa.project_id + '/messages:send';
  var reqs = rows.map(function (r) {
    var m = build(PUSH_TEXT.rec[r[3]] ? String(r[3]) : 'zh');
    var msg = { token: String(r[0]), data: { link: m.link } };
    if (r[2] === 'android') {
      msg.android = { priority: 'high', notification: { title: m.title, body: m.body, channel_id: 'general', icon: 'ic_stat_notify', color: '#ED6C01' } };
    } else {
      msg.webpush = { headers: { Urgency: 'high' }, notification: { title: m.title, body: m.body, icon: WEB_APP_URL + 'icon-192.png' }, fcm_options: { link: m.link } };
    }
    return { url: url, method: 'post', contentType: 'application/json', headers: { Authorization: auth }, payload: JSON.stringify({ message: msg }), muteHttpExceptions: true };
  });
  var dead = [], sent = 0;
  for (var i = 0; i < reqs.length; i += 50) {
    UrlFetchApp.fetchAll(reqs.slice(i, i + 50)).forEach(function (res, j) {
      var code = res.getResponseCode();
      if (code === 200) sent++;
      else if (code === 404 || /UNREGISTERED|INVALID_ARGUMENT/.test(res.getContentText())) dead.push(String(rows[i + j][0]));
    });
  }
  if (dead.length) withLock_(function () { removePushTokens_(dead); });
  return sent;
}
function removePushTokens_(tokens) {
  var set = tokens.reduce(function (o, t) { o[t] = 1; return o; }, {});
  var sh = sheet_('推播'), v = sh.getDataRange().getValues();
  for (var i = v.length - 1; i >= 1; i--) if (set[String(v[i][0])]) sh.deleteRow(i + 1);
}
// 用服務帳戶金鑰換 Firebase 通知用的存取權杖（快取 50 分鐘）
function fcmAccessToken_(sa) {
  var cache = CacheService.getScriptCache(), hit = cache.get('fcm_at');
  if (hit) return hit;
  var now = Math.floor(Date.now() / 1000);
  var enc = function (o) { return Utilities.base64EncodeWebSafe(JSON.stringify(o)).replace(/=+$/, ''); };
  var input = enc({ alg: 'RS256', typ: 'JWT' }) + '.' + enc({ iss: sa.client_email, scope: 'https://www.googleapis.com/auth/firebase.messaging',
    aud: 'https://oauth2.googleapis.com/token', iat: now, exp: now + 3600 });
  var sig = Utilities.base64EncodeWebSafe(Utilities.computeRsaSha256Signature(input, sa.private_key)).replace(/=+$/, '');
  var res = UrlFetchApp.fetch('https://oauth2.googleapis.com/token', { method: 'post', muteHttpExceptions: true,
    payload: { grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer', assertion: input + '.' + sig } });
  var j = JSON.parse(res.getContentText());
  if (!j.access_token) throw new Error('fcm_auth_failed');
  cache.put('fcm_at', j.access_token, 3000);
  return j.access_token;
}

/* 設定好 FCM_SA 後，在編輯器上方選「setupPush」按「執行」一次 */
function setupPush() {
  sheet_('推播');
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (/^(visitReminder|checkVersionPush)$/.test(t.getHandlerFunction())) ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('visitReminder').timeBased().everyDays(1).atHour(19).inTimezone('Asia/Taipei').create();
  ScriptApp.newTrigger('checkVersionPush').timeBased().everyHours(1).create();
  checkVersionPush();                                          // 記下目前版本
  fcmAccessToken_(JSON.parse(PropertiesService.getScriptProperties().getProperty('FCM_SA')));
  Logger.log('通知設定完成');
}
