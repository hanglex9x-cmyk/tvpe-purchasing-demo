/**
 * TVPE ERP — Web App (Apps Script) v3
 * Đăng nhập bằng TÀI KHOẢN + phân quyền theo vai trò (tab PHAN_QUYEN trong Sheet).
 * - Mật khẩu lưu dạng băm SHA-256 kèm salt, không lưu mật khẩu gốc.
 * - Phiên đăng nhập = token lưu trong CacheService (6 giờ).
 * - Quyền được chặn PHÍA SERVER: Sales chỉ nhận dữ liệu giá bán + tồn khả dụng, mọi thao tác ghi đều bị từ chối;
 *   duyệt giá + quản lý người dùng chỉ dành cho Admin.
 * - Mọi thao tác ghi ghi vào tab LICH_SU với tên tài khoản đăng nhập (server xác thực, không phải tên tự gõ).
 * Deploy: Deploy > Manage deployments > Edit > Version: New version (giữ nguyên link /exec).
 */
var SHEET_ID = '1CpytSjhaYifMFo8SSCIEJA9GYKdINLEylXMPGpKSvh4';
var PORTAL_URL = 'https://docs.google.com/spreadsheets/d/1bQ-B2fJxBRRq-vYTMGc6BN8WnfMdcSytBUpyqufPSIM/edit';

var TABS = {
  po:        { sheet: 'PO',            edit: [1, 2, 3, 4, 5, 6, 8, 11, 12, 13, 15, 16, 18, 19, 20, 22], label: 'Đơn hàng & PO' },
  nx:        { sheet: 'NHAP_XUAT_KHO', edit: [0, 1, 2, 3, 4, 6, 7, 9, 10], label: 'Nhập / Xuất kho' },
  giavon:    { sheet: 'GIA_VON',       edit: [0, 1, 2, 3, 4, 11, 15], label: 'Giá vốn' },
  chiphi:    { sheet: 'CHI_PHI_NK',    edit: [0, 1, 2, 3, 4], label: 'Chi phí nhập khẩu' },
  ton:       { sheet: 'TON_KHO',       edit: [2, 6, 8, 11, 12], label: 'Kho & tồn kho' },
  congno:    { sheet: 'CONG_NO',       edit: [0, 1, 2, 3, 4, 5, 6, 13], label: 'Công nợ NCC' },
  thanhtoan: { sheet: 'THANH_TOAN',    edit: [0, 1, 2, 3, 4, 5, 6], label: 'Thanh toán' },
  chungtu:   { sheet: 'CHUNG_TU',      edit: [0, 1, 2, 3, 4, 5, 6, 7, 8], label: 'Hồ sơ chứng từ' },
  giaban:    { sheet: 'BANG_GIA_BAN',  edit: [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11], label: 'Bảng giá bán' },
  duyet:     { sheet: 'PHE_DUYET_GIA', edit: [1, 2, 3, 4, 5, 11], label: 'Phê duyệt giá' },
  giahang:   { sheet: 'BANG_GIA_HANG', edit: [0, 1, 2, 3, 5, 6, 8], label: 'Bảng giá hãng' },
  mhang:     { sheet: 'MASTER_HANG',   edit: [0, 1, 2, 3, 4, 5, 6, 7], label: 'Master — Mã hàng' },
  mncc:      { sheet: 'MASTER_NCC',    edit: [0, 1, 2, 3, 4, 5, 6], label: 'Master — Nhà cung cấp' },
  mkhach:    { sheet: 'MASTER_KHACH',  edit: [0, 1, 2, 3], label: 'Master — Khách hàng' }
};

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('TVPE ERP — Quản trị mua hàng & nhập khẩu')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function ss_() { return SpreadsheetApp.openById(SHEET_ID); }

/* ================= ĐĂNG NHẬP & PHÂN QUYỀN ================= */
function hashPw_(salt, pw) {
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + ':' + pw, Utilities.Charset.UTF_8);
  return bytes.map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
}
function permsFrom_(role, g, h, i2, j) {
  // Admin luôn toàn quyền; các vai trò khác đọc 4 cột quyền chi tiết, trống thì dùng mẫu theo vai trò
  if (role === 'Admin') return { cost: true, write: true, del: true, approve: true, kho: false };
  var any = [g, h, i2, j].some(function (x) { return String(x || '') !== ''; });
  if (!any) {
    if (role === 'Mua hàng') return { cost: true, write: true, del: true, approve: false, kho: false };
    if (role === 'Kho') return { cost: false, write: true, del: false, approve: false, kho: true };
    return { cost: false, write: false, del: false, approve: false, kho: false }; // Sales và vai trò khác
  }
  return { cost: g === 'Có', write: h === 'Có', del: i2 === 'Có', approve: j === 'Có', kho: role === 'Kho' };
}
function polines_() {
  // Danh sách mã dòng PO (chỉ mã, KHÔNG giá) — để vai trò Kho gắn phiếu nhập vào đúng dòng PO khi nhận hàng
  var vals = ss_().getSheetByName('PO').getDataRange().getValues();
  var rows = [];
  for (var i = 1; i < vals.length; i++) if (vals[i][0]) rows.push([vals[i][0], vals[i][6]]);
  return { headers: ['Mã dòng PO', 'Product Code'], rows: rows };
}
function tiendo_() {
  // Tiến độ hàng mua đang về — bản rút gọn KHÔNG có giá, dùng chung cho Sales và Kho
  var vals = ss_().getSheetByName('PO').getDataRange().getValues();
  var rows = [];
  for (var i = 1; i < vals.length; i++) {
    var r = vals[i];
    if (!r[6]) continue;
    if (!(Number(r[10]) > 0)) continue; // chỉ dòng còn thiếu hàng = đang về
    rows.push([r[6], r[7], r[2], r[8], r[9], r[10], r[18], fmt_(r[19])]);
  }
  return { headers: ['Mã hàng', 'Tên hàng', 'Khách / Dự án', 'SL đặt', 'SL đã nhận', 'SL còn thiếu', 'Trạng thái', 'ETA dự kiến'], rows: rows };
}
function brands_() {
  // Ánh xạ Mã hàng -> Hãng/NSX (từ MASTER_HANG) — hãng không phải dữ liệu nhạy cảm, gửi cho mọi vai trò để gom tồn kho theo hãng
  var vals = ss_().getSheetByName('MASTER_HANG').getDataRange().getValues();
  var out = {};
  for (var i = 1; i < vals.length; i++) if (vals[i][0]) out[String(vals[i][0])] = String(vals[i][3] || '');
  return out;
}
function users_() {
  var sh = ss_().getSheetByName('PHAN_QUYEN');
  var vals = sh.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < vals.length; i++) {
    if (!vals[i][0]) continue;
    var role = String(vals[i][4]);
    out.push({
      row: i + 1, username: String(vals[i][0]), hash: String(vals[i][1]), salt: String(vals[i][2]),
      name: String(vals[i][3]), role: role, status: String(vals[i][5]),
      perms: permsFrom_(role, vals[i][6], vals[i][7], vals[i][8], vals[i][9])
    });
  }
  return out;
}
function login(username, password) {
  username = String(username || '').trim().toLowerCase();
  var users = users_();
  for (var i = 0; i < users.length; i++) {
    var u = users[i];
    if (u.username.toLowerCase() === username) {
      if (u.status !== 'Hoạt động') throw new Error('Tài khoản đang bị khóa. Liên hệ Admin.');
      if (hashPw_(u.salt, String(password || '')) !== u.hash) throw new Error('Sai mật khẩu. Vui lòng thử lại.');
      var token = Utilities.getUuid();
      CacheService.getScriptCache().put('sess_' + token, JSON.stringify({ u: u.username, name: u.name, role: u.role, perms: u.perms }), 21600);
      log_(u.name + ' (' + u.username + ')', 'hethong', 'Đăng nhập', '', 'Vai trò: ' + u.role);
      return { token: token, username: u.username, name: u.name, role: u.role, perms: u.perms };
    }
  }
  throw new Error('Tài khoản không tồn tại.');
}
function auth_(token) {
  if (!token) throw new Error('AUTH: Chưa đăng nhập');
  var raw = CacheService.getScriptCache().get('sess_' + token);
  if (!raw) throw new Error('AUTH: Phiên đăng nhập đã hết hạn, vui lòng đăng nhập lại');
  return JSON.parse(raw);
}
function logout(token) {
  if (token) CacheService.getScriptCache().remove('sess_' + token);
  return 'Đã đăng xuất';
}
function requireWrite_(sess) {
  if (!sess.perms || !sess.perms.write) throw new Error('Tài khoản của bạn không có quyền Thêm/Sửa dữ liệu (Admin phân quyền ở mục Hệ thống).');
}
function requireDelete_(sess) {
  if (!sess.perms || !sess.perms.del) throw new Error('Tài khoản của bạn không có quyền Xóa dữ liệu (Admin phân quyền ở mục Hệ thống).');
}
function requireApprove_(sess) {
  if (!sess.perms || !sess.perms.approve) throw new Error('Tài khoản của bạn không có quyền Duyệt giá (Admin phân quyền ở mục Hệ thống).');
}
function requireAdmin_(sess) {
  if (sess.role !== 'Admin') throw new Error('Chức năng này chỉ dành cho Admin.');
}
function changePassword(token, oldPw, newPw) {
  var sess = auth_(token);
  if (!newPw || String(newPw).length < 6) throw new Error('Mật khẩu mới phải có ít nhất 6 ký tự.');
  return withLock_(function () {
    var users = users_();
    for (var i = 0; i < users.length; i++) {
      if (users[i].username === sess.u) {
        if (hashPw_(users[i].salt, String(oldPw || '')) !== users[i].hash) throw new Error('Mật khẩu hiện tại không đúng.');
        var salt = Utilities.getUuid().replace(/-/g, '').slice(0, 16);
        var sh = ss_().getSheetByName('PHAN_QUYEN');
        sh.getRange(users[i].row, 2).setValue(hashPw_(salt, String(newPw)));
        sh.getRange(users[i].row, 3).setValue(salt);
        log_(sess.name + ' (' + sess.u + ')', 'hethong', 'Đổi mật khẩu', '', '');
        return 'Đã đổi mật khẩu thành công';
      }
    }
    throw new Error('Không tìm thấy tài khoản.');
  });
}
/* ---- Quản lý người dùng (Admin) ---- */
function listUsers(token) {
  var sess = auth_(token); requireAdmin_(sess);
  return users_().map(function (u) { return { username: u.username, name: u.name, role: u.role, status: u.status, perms: u.perms }; });
}
function flagsArr_(perms) {
  return [perms && perms.cost ? 'Có' : 'Không', perms && perms.write ? 'Có' : 'Không', perms && perms.del ? 'Có' : 'Không', perms && perms.approve ? 'Có' : 'Không'];
}
function addUser(token, username, name, role, pw, perms) {
  var sess = auth_(token); requireAdmin_(sess);
  username = String(username || '').trim().toLowerCase();
  if (!username || !name || !pw) throw new Error('Điền đủ: tên đăng nhập, họ tên, mật khẩu.');
  if (['Admin', 'Mua hàng', 'Sales', 'Kho'].indexOf(role) < 0) throw new Error('Vai trò không hợp lệ.');
  return withLock_(function () {
    var users = users_();
    for (var i = 0; i < users.length; i++) if (users[i].username.toLowerCase() === username) throw new Error('Tên đăng nhập đã tồn tại.');
    var salt = Utilities.getUuid().replace(/-/g, '').slice(0, 16);
    var f = flagsArr_(perms);
    ss_().getSheetByName('PHAN_QUYEN').appendRow([username, hashPw_(salt, String(pw)), salt, name, role, 'Hoạt động', f[0], f[1], f[2], f[3]]);
    log_(sess.name + ' (' + sess.u + ')', 'hethong', 'Tạo tài khoản', '', username + ' — ' + role);
    return 'Đã tạo tài khoản ' + username;
  });
}
function updateUser(token, username, name, role, perms) {
  var sess = auth_(token); requireAdmin_(sess);
  if (['Admin', 'Mua hàng', 'Sales', 'Kho'].indexOf(role) < 0) throw new Error('Vai trò không hợp lệ.');
  return withLock_(function () {
    var users = users_();
    for (var i = 0; i < users.length; i++) {
      if (users[i].username === username) {
        var sh = ss_().getSheetByName('PHAN_QUYEN');
        sh.getRange(users[i].row, 4).setValue(name);
        sh.getRange(users[i].row, 5).setValue(role);
        var f = flagsArr_(perms);
        sh.getRange(users[i].row, 7, 1, 4).setValues([f]);
        log_(sess.name + ' (' + sess.u + ')', 'hethong', 'Sửa & phân quyền TK', '', username + ' — ' + role + ' | Xem giá vốn: ' + f[0] + ', Thêm/Sửa: ' + f[1] + ', Xóa: ' + f[2] + ', Duyệt giá: ' + f[3]);
        return 'Đã cập nhật tài khoản ' + username;
      }
    }
    throw new Error('Không tìm thấy tài khoản.');
  });
}
function setUserStatus(token, username, active) {
  var sess = auth_(token); requireAdmin_(sess);
  if (username === sess.u && !active) throw new Error('Không thể tự khóa tài khoản của chính mình.');
  return withLock_(function () {
    var users = users_();
    for (var i = 0; i < users.length; i++) {
      if (users[i].username === username) {
        ss_().getSheetByName('PHAN_QUYEN').getRange(users[i].row, 6).setValue(active ? 'Hoạt động' : 'Đã khóa');
        log_(sess.name + ' (' + sess.u + ')', 'hethong', active ? 'Mở khóa TK' : 'Khóa TK', '', username);
        return 'Đã ' + (active ? 'mở khóa' : 'khóa') + ' tài khoản ' + username;
      }
    }
    throw new Error('Không tìm thấy tài khoản.');
  });
}
function resetUserPw(token, username, newPw) {
  var sess = auth_(token); requireAdmin_(sess);
  if (!newPw || String(newPw).length < 6) throw new Error('Mật khẩu mới phải có ít nhất 6 ký tự.');
  return withLock_(function () {
    var users = users_();
    for (var i = 0; i < users.length; i++) {
      if (users[i].username === username) {
        var salt = Utilities.getUuid().replace(/-/g, '').slice(0, 16);
        var sh = ss_().getSheetByName('PHAN_QUYEN');
        sh.getRange(users[i].row, 2).setValue(hashPw_(salt, String(newPw)));
        sh.getRange(users[i].row, 3).setValue(salt);
        log_(sess.name + ' (' + sess.u + ')', 'hethong', 'Đặt lại mật khẩu', '', username);
        return 'Đã đặt lại mật khẩu cho ' + username;
      }
    }
    throw new Error('Không tìm thấy tài khoản.');
  });
}

/* ================= DỮ LIỆU ================= */
function fmt_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, 'Asia/Ho_Chi_Minh', 'dd/MM/yyyy');
  return v;
}
function tab_(name) {
  var sh = ss_().getSheetByName(name);
  if (!sh) return { headers: [], rows: [] };
  var vals = sh.getDataRange().getValues();
  var headers = vals.length ? vals[0].map(function (x) { return String(x); }) : [];
  var rows = [];
  for (var i = 1; i < vals.length; i++) {
    var r = vals[i].map(fmt_);
    r.push(i + 1);
    rows.push(r);
  }
  return { headers: headers, rows: rows };
}
function blankCols_(tab, cols) {
  tab.rows.forEach(function (r) { cols.forEach(function (c) { r[c] = ''; }); });
  return tab;
}
function getAllData(token) {
  var sess = auth_(token);
  var out = { portal: PORTAL_URL, me: { username: sess.u, name: sess.name, role: sess.role, perms: sess.perms } };
  out.tiendo = tiendo_();
  out.brands = brands_();
  if (sess.perms && sess.perms.kho) {
    // Vai trò KHO: chỉ kho + nhập/xuất + master (để chọn mã) + đối chiếu MISA + tiến độ — server KHÔNG gửi giá bán, giá vốn, PO, công nợ.
    out.ton = tab_('TON_KHO');
    out.nx = tab_('NHAP_XUAT_KHO');
    out.mhang = tab_('MASTER_HANG');
    out.doichieu = tab_('DOI_CHIEU');
    out.polines = polines_();
    out.kpi = {};
    ['po', 'giavon', 'chiphi', 'congno', 'thanhtoan', 'chungtu', 'duyet', 'giahang', 'giaban', 'mncc', 'mkhach', 'lichsu'].forEach(function (k) { out[k] = { headers: [], rows: [] }; });
    return out;
  }
  if (!sess.perms || !sess.perms.cost) {
    // Tài khoản KHÔNG có quyền xem giá vốn: CHỈ giá bán (đã che giá sàn nội bộ) + tồn khả dụng (chỉ mã, tên, khả dụng) + tiến độ hàng về (không giá) — server không gửi dữ liệu giá vốn/PO/công nợ.
    out.giaban = blankCols_(tab_('BANG_GIA_BAN'), [9]);
    var ton = tab_('TON_KHO');
    ton.rows.forEach(function (r) { for (var c = 0; c < r.length - 1; c++) { if ([0, 1, 7, 15, 16].indexOf(c) < 0) r[c] = ''; } });
    out.ton = ton;
    out.kpi = {};
    ['po', 'nx', 'giavon', 'chiphi', 'congno', 'thanhtoan', 'chungtu', 'duyet', 'giahang', 'mhang', 'mncc', 'mkhach', 'lichsu', 'doichieu'].forEach(function (k) { out[k] = { headers: [], rows: [] }; });
    return out;
  }
  var dash = ss_().getSheetByName('DASHBOARD');
  var kpi = {};
  dash.getRange('A4:B15').getValues().forEach(function (r) { if (r[0]) kpi[String(r[0])] = r[1]; });
  dash.getRange('A28:B37').getValues().forEach(function (r) { if (r[0]) kpi[String(r[0])] = r[1]; });
  out.kpi = kpi;
  out.lichsu = tab_('LICH_SU');
  out.doichieu = tab_('DOI_CHIEU');
  Object.keys(TABS).forEach(function (k) { out[k] = tab_(TABS[k].sheet); });
  return out;
}

/* ================= NHẬT KÝ ================= */
function log_(actor, tabKey, action, row, summary) {
  try {
    var sh = ss_().getSheetByName('LICH_SU');
    if (!sh) return;
    var label = tabKey === 'hethong' ? 'Hệ thống' : (TABS[tabKey] ? TABS[tabKey].label : tabKey);
    sh.appendRow([new Date(), actor || '', label, action, row || '', summary || '']);
  } catch (e) { }
}
function actor_(sess) { return sess.name + ' (' + sess.u + ')'; }

/* ================= GHI DỮ LIỆU ================= */
function withLock_(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}
function guardKho_(sess, tabKey) {
  if (sess.perms && sess.perms.kho && tabKey !== 'nx') throw new Error('Tài khoản Kho chỉ được nhập / sửa dữ liệu Nhập – Xuất kho.');
}
function updateRow(token, tabKey, row, obj) {
  var sess = auth_(token); requireWrite_(sess); guardKho_(sess, tabKey);
  var cfg = TABS[tabKey];
  if (!cfg) throw new Error('Bảng không hợp lệ: ' + tabKey);
  return withLock_(function () {
    var sh = ss_().getSheetByName(cfg.sheet);
    var changed = [];
    Object.keys(obj).forEach(function (k) {
      var col = Number(k);
      if (cfg.edit.indexOf(col) < 0) return;
      var v = obj[k];
      if (typeof v === 'string' && v !== '' && /^-?\d+(\.\d+)?$/.test(v.trim())) v = Number(v);
      sh.getRange(row, col + 1).setValue(v);
      changed.push(sh.getRange(1, col + 1).getValue() + '=' + obj[k]);
    });
    log_(actor_(sess), tabKey, 'Sửa', row, changed.join('; '));
    return 'Đã sửa dòng ' + row + ' (' + cfg.label + ')';
  });
}
function deleteRow(token, tabKey, row) {
  var sess = auth_(token); requireDelete_(sess); guardKho_(sess, tabKey);
  var cfg = TABS[tabKey];
  if (!cfg) throw new Error('Bảng không hợp lệ: ' + tabKey);
  return withLock_(function () {
    var sh = ss_().getSheetByName(cfg.sheet);
    var preview = sh.getRange(row, 1, 1, Math.min(4, sh.getLastColumn())).getValues()[0].join(' | ');
    sh.deleteRow(row);
    log_(actor_(sess), tabKey, 'Xóa', row, preview);
    return 'Đã xóa dòng ' + row + ' (' + cfg.label + ')';
  });
}
function addRow(token, tabKey, values) {
  var sess = auth_(token); requireWrite_(sess); guardKho_(sess, tabKey);
  var cfg = TABS[tabKey];
  if (!cfg) throw new Error('Bảng không hợp lệ: ' + tabKey);
  return withLock_(function () {
    var sh = ss_().getSheetByName(cfg.sheet);
    var n = sh.getLastColumn();
    var row = [];
    for (var c = 0; c < n; c++) {
      var v = values[c] !== undefined ? values[c] : '';
      if (typeof v === 'string' && v !== '' && /^-?\d+(\.\d+)?$/.test(v.trim())) v = Number(v);
      row.push(v);
    }
    sh.appendRow(row);
    var nr = sh.getLastRow();
    if (nr > 2) {
      for (var c2 = 0; c2 < n; c2++) {
        if (cfg.edit.indexOf(c2) >= 0) continue;
        var prev = sh.getRange(nr - 1, c2 + 1);
        if (prev.getFormula()) prev.copyTo(sh.getRange(nr, c2 + 1));
      }
    }
    log_(actor_(sess), tabKey, 'Thêm', nr, row.slice(0, 4).join(' | '));
    return 'Đã thêm dòng mới (' + cfg.label + ')';
  });
}
function addPO(token, o) {
  var sess = auth_(token); requireWrite_(sess); guardKho_(sess, 'po');
  return withLock_(function () {
    var sh = ss_().getSheetByName('PO');
    var r = sh.getLastRow() + 1;
    sh.getRange(r, 1, 1, 7).setValues([[o.po + '/' + o.ma, o.madon || '', o.khach || '', o.po, o.ngay, o.ncc, o.ma]]);
    sh.getRange(r, 8).setFormula('=IF(G' + r + '="";"";IFERROR(INDEX(MASTER_HANG!C:C;MATCH(G' + r + ';MASTER_HANG!A:A;0));""))');
    sh.getRange(r, 9).setValue(Number(o.sl));
    sh.getRange(r, 10).setFormula('=IF(A' + r + '="";"";SUMIFS(NHAP_XUAT_KHO!G:G;NHAP_XUAT_KHO!D:D;A' + r + ';NHAP_XUAT_KHO!B:B;"Nhập"))');
    sh.getRange(r, 11).setFormula('=IF(A' + r + '="";"";MAX(I' + r + '-J' + r + ';0))');
    sh.getRange(r, 12, 1, 3).setValues([[Number(o.list), Number(o.ck1), Number(o.ck2)]]);
    sh.getRange(r, 15).setFormula('=IF(L' + r + '="";"";L' + r + '*(1-M' + r + '/100)*(1-N' + r + '/100))');
    sh.getRange(r, 16).setValue(o.tiente);
    sh.getRange(r, 17).setValue(Number(o.tygia));
    sh.getRange(r, 18).setFormula('=IF(O' + r + '="";"";I' + r + '*O' + r + '*Q' + r + ')');
    sh.getRange(r, 19).setValue(o.trangthai || 'Đã đặt PO');
    sh.getRange(r, 20).setValue(o.eta);
    sh.getRange(r, 21).setValue(o.dieukhoan || '');
    sh.getRange(r, 23).setValue(sess.name);
    log_(actor_(sess), 'po', 'Thêm', r, o.po + '/' + o.ma + ' — SL ' + o.sl);
    return 'Đã thêm dòng PO ' + o.po + '/' + o.ma;
  });
}
function addPhieu(token, o) {
  var sess = auth_(token); requireWrite_(sess);
  return withLock_(function () {
    var sh = ss_().getSheetByName('NHAP_XUAT_KHO');
    var r = sh.getLastRow() + 1;
    sh.getRange(r, 1, 1, 5).setValues([[o.ngay, o.loai, o.maphieu || '', o.madongpo || '', o.ma]]);
    sh.getRange(r, 6).setFormula('=IFERROR(INDEX(MASTER_HANG!C:C;MATCH(E' + r + ';MASTER_HANG!A:A;0));"")');
    sh.getRange(r, 7).setValue(Number(o.sl));
    sh.getRange(r, 8).setValue(o.khach || '');
    sh.getRange(r, 10).setValue(o.note || '');
    sh.getRange(r, 11).setValue(o.kho || '');
    log_(actor_(sess), 'nx', 'Thêm', r, o.loai + ' ' + o.ma + ' — SL ' + o.sl + (o.kho ? ' — Kho ' + o.kho : ''));
    return 'Đã ghi phiếu ' + o.loai + ' ' + o.ma;
  });
}
function addThanhToan(token, o) {
  var sess = auth_(token); requireWrite_(sess); guardKho_(sess, 'thanhtoan');
  return withLock_(function () {
    var sh = ss_().getSheetByName('THANH_TOAN');
    sh.appendRow([o.ngay, o.inv, o.ncc, Number(o.tien), o.ht, o.ck || '', o.note || '']);
    log_(actor_(sess), 'thanhtoan', 'Thêm', sh.getLastRow(), o.inv + ' — ' + Number(o.tien).toLocaleString('vi-VN') + 'đ');
    return 'Đã ghi thanh toán cho ' + o.inv;
  });
}
/* ================= NẠP FILE TỪ MISA ================= */
function importMaster(token, rows) {
  var sess = auth_(token); requireWrite_(sess); guardKho_(sess, 'mhang');
  if (!rows || !rows.length) throw new Error('Không có dòng dữ liệu nào trong file.');
  return withLock_(function () {
    var sh = ss_().getSheetByName('MASTER_HANG');
    var vals = sh.getDataRange().getValues();
    var pos = {};
    for (var i = 1; i < vals.length; i++) if (vals[i][0]) pos[String(vals[i][0]).trim().toUpperCase()] = i + 1;
    var added = 0, updated = 0;
    rows.forEach(function (r) {
      var code = String(r[0] || '').trim();
      if (!code) return;
      var key = code.toUpperCase();
      var line = [code, String(r[1] || code), String(r[2] || ''), String(r[3] || ''), String(r[4] || ''), String(r[5] || ''), String(r[6] || 'Đang bán'), String(r[7] || '')];
      if (pos[key]) {
        var row = pos[key];
        for (var c = 1; c < 8; c++) if (line[c] !== '') sh.getRange(row, c + 1).setValue(line[c]);
        updated++;
      } else {
        sh.appendRow(line);
        pos[key] = sh.getLastRow();
        added++;
      }
    });
    log_(actor_(sess), 'mhang', 'Nạp từ MISA', '', 'Thêm mới ' + added + ' mã, cập nhật ' + updated + ' mã');
    return 'Đã nạp danh sách mã: thêm mới ' + added + ' mã, cập nhật ' + updated + ' mã.';
  });
}
function importDoiChieu(token, rows) {
  var sess = auth_(token); requireWrite_(sess);
  if (!rows || !rows.length) throw new Error('Không có dòng dữ liệu nào trong file.');
  return withLock_(function () {
    var ss = ss_();
    var ton = ss.getSheetByName('TON_KHO').getDataRange().getValues();
    var tonMap = {};
    for (var i = 1; i < ton.length; i++) if (ton[i][0]) tonMap[String(ton[i][0]).trim().toUpperCase()] = { thuc: Number(ton[i][5]) || 0, khadung: Number(ton[i][7]) || 0 };
    var sh = ss.getSheetByName('DOI_CHIEU');
    if (sh.getLastRow() > 1) sh.deleteRows(2, sh.getLastRow() - 1);
    var today = new Date();
    var lech = 0, khongco = 0, n = 0;
    rows.forEach(function (r) {
      var code = String(r[0] || '').trim();
      var qty = Number(r[1]);
      if (!code || isNaN(qty)) return;
      n++;
      var row = sh.getLastRow() + 1;
      var t = tonMap[code.toUpperCase()];
      if (!t) khongco++; else if (t.thuc !== qty) lech++;
      sh.getRange(row, 1).setValue(code);
      sh.getRange(row, 2).setFormula('=IFERROR(INDEX(MASTER_HANG!C:C;MATCH(A' + row + ';MASTER_HANG!A:A;0));"—")');
      sh.getRange(row, 3).setValue(qty);
      sh.getRange(row, 4).setFormula('=IFERROR(INDEX(TON_KHO!F:F;MATCH(A' + row + ';TON_KHO!A:A;0));"")');
      sh.getRange(row, 5).setFormula('=IFERROR(INDEX(TON_KHO!H:H;MATCH(A' + row + ';TON_KHO!A:A;0));"")');
      sh.getRange(row, 6).setFormula('=IF(OR(C' + row + '="";D' + row + '="");"";C' + row + '-D' + row + ')');
      sh.getRange(row, 7).setFormula('=IF(D' + row + '="";"⚠️ Chưa có trong hệ thống";IF(F' + row + '=0;"✔ Khớp";"⚠️ Lệch — cần kiểm tra"))');
      sh.getRange(row, 8).setValue(today);
      sh.getRange(row, 9).setValue(sess.name);
    });
    log_(actor_(sess), 'ton', 'Nạp tồn MISA đối chiếu', '', n + ' mã: khớp ' + (n - lech - khongco) + ', lệch ' + lech + ', chưa có ' + khongco);
    return 'Đã nạp tồn MISA: ' + n + ' mã — khớp ' + (n - lech - khongco) + ', lệch ' + lech + ', chưa có trong hệ thống ' + khongco + '.';
  });
}

/* ================= NẠP PHIẾU NHẬP/XUẤT TỪ MISA + ĐỒNG BỘ DANH MỤC TỪ DRIVE ================= */
function normDateKey_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, 'Asia/Ho_Chi_Minh', 'yyyyMMdd');
  var s = String(v || '').trim();
  var m = s.match(/^(\d{4})-(\d{2})-(\d{2})/); if (m) return m[1] + m[2] + m[3];
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/); if (m) return m[3] + ('0' + m[2]).slice(-2) + ('0' + m[1]).slice(-2);
  return s;
}
function ensureTonRow_(ss, code) {
  var sh = ss.getSheetByName('TON_KHO');
  var vals = sh.getRange(1, 1, Math.max(sh.getLastRow(), 1), 1).getValues();
  for (var i = 1; i < vals.length; i++) if (String(vals[i][0]).trim().toUpperCase() === String(code).trim().toUpperCase()) return false;
  var r = sh.getLastRow() + 1;
  sh.getRange(r, 1).setValue(code);
  sh.getRange(r, 2).setFormula('=IFERROR(INDEX(MASTER_HANG!C:C;MATCH(A' + r + ';MASTER_HANG!A:A;0));"")');
  sh.getRange(r, 3).setValue(0);
  sh.getRange(r, 4).setFormula('=SUMIFS(NHAP_XUAT_KHO!G:G;NHAP_XUAT_KHO!E:E;A' + r + ';NHAP_XUAT_KHO!B:B;"Nhập")');
  sh.getRange(r, 5).setFormula('=SUMIFS(NHAP_XUAT_KHO!G:G;NHAP_XUAT_KHO!E:E;A' + r + ';NHAP_XUAT_KHO!B:B;"Xuất")');
  sh.getRange(r, 6).setFormula('=C' + r + '+D' + r + '-E' + r);
  sh.getRange(r, 7).setValue(0);
  sh.getRange(r, 8).setFormula('=F' + r + '-G' + r);
  sh.getRange(r, 10).setFormula('=IF(F' + r + '<0;"🔴 Tồn âm";IF(H' + r + '<0;"🔴 Giữ vượt tồn";IF(H' + r + '=0;"🟠 Hết khả dụng";"🟢 Còn bán được")))');
  sh.getRange(r, 11).setFormula('=SUMIFS(PO!K:K;PO!G:G;A' + r + ')');
  sh.getRange(r, 12).setValue(0);
  sh.getRange(r, 13).setValue(0);
  sh.getRange(r, 14).setFormula('=L' + r + '+SUMIFS(NHAP_XUAT_KHO!G:G;NHAP_XUAT_KHO!E:E;A' + r + ';NHAP_XUAT_KHO!B:B;"Nhập";NHAP_XUAT_KHO!K:K;"QN")-SUMIFS(NHAP_XUAT_KHO!G:G;NHAP_XUAT_KHO!E:E;A' + r + ';NHAP_XUAT_KHO!B:B;"Xuất";NHAP_XUAT_KHO!K:K;"QN")');
  sh.getRange(r, 15).setFormula('=M' + r + '+SUMIFS(NHAP_XUAT_KHO!G:G;NHAP_XUAT_KHO!E:E;A' + r + ';NHAP_XUAT_KHO!B:B;"Nhập";NHAP_XUAT_KHO!K:K;"HN")-SUMIFS(NHAP_XUAT_KHO!G:G;NHAP_XUAT_KHO!E:E;A' + r + ';NHAP_XUAT_KHO!B:B;"Xuất";NHAP_XUAT_KHO!K:K;"HN")');
  sh.getRange(r, 16).setFormula('=IF(F' + r + '=0;0;ROUND(N' + r + '*H' + r + '/F' + r + ';0))');
  sh.getRange(r, 17).setFormula('=H' + r + '-P' + r);
  return true;
}
function importPhieu(token, rows) {
  var sess = auth_(token); requireWrite_(sess);
  if (!rows || !rows.length) throw new Error('Không có dòng dữ liệu nào trong file.');
  return withLock_(function () {
    var ss = ss_();
    var mh = ss.getSheetByName('MASTER_HANG').getDataRange().getValues();
    var master = {};
    for (var i = 1; i < mh.length; i++) if (mh[i][0]) master[String(mh[i][0]).trim().toUpperCase()] = 1;
    var nxSh = ss.getSheetByName('NHAP_XUAT_KHO');
    var nxVals = nxSh.getDataRange().getValues();
    function keyOf(maphieu, code, loai, sl, ngay) { return String(maphieu || '').trim().toUpperCase() + '|' + String(code || '').trim().toUpperCase() + '|' + loai + '|' + (Number(sl) || 0) + '|' + normDateKey_(ngay); }
    var seen = {};
    for (var j = 1; j < nxVals.length; j++) if (nxVals[j][2]) seen[keyOf(nxVals[j][2], nxVals[j][4], nxVals[j][1], nxVals[j][6], nxVals[j][0])] = 1;
    var poVals = ss.getSheetByName('PO').getDataRange().getValues();
    var openByCode = {};
    for (var p = 1; p < poVals.length; p++) {
      var cd = String(poVals[p][6] || '').trim().toUpperCase();
      var left = Number(poVals[p][10]) || 0;
      if (cd && left > 0) { (openByCode[cd] = openByCode[cd] || []).push({ line: String(poVals[p][0]), left: left }); }
    }
    var added = 0, dup = 0, news = [], ensured = {};
    rows.forEach(function (r) {
      var code = String(r.code || '').trim();
      if (!code) return;
      if (!master[code.toUpperCase()]) { if (news.indexOf(code) < 0) news.push(code); return; }
      var sl = Number(r.sl);
      if (!(sl > 0)) return;
      var loai = r.loai === 'Xuất' ? 'Xuất' : 'Nhập';
      var key = keyOf(r.maphieu, code, loai, sl, r.ngay);
      if (seen[key]) { dup++; return; }
      seen[key] = 1;
      var parts = [{ mdp: String(r.madongpo || ''), sl: sl }];
      if (loai === 'Nhập' && !r.madongpo) {
        parts = [];
        var rest = sl;
        var lines = openByCode[code.toUpperCase()] || [];
        for (var q = 0; q < lines.length && rest > 0; q++) {
          var take = Math.min(rest, lines[q].left);
          if (take > 0) { parts.push({ mdp: lines[q].line, sl: take }); lines[q].left -= take; rest -= take; }
        }
        if (rest > 0) parts.push({ mdp: '', sl: rest });
        if (!parts.length) parts = [{ mdp: '', sl: sl }];
      }
      parts.forEach(function (pt) {
        var row = nxSh.getLastRow() + 1;
        nxSh.getRange(row, 1, 1, 5).setValues([[String(r.ngay || ''), loai, String(r.maphieu || ''), pt.mdp, code]]);
        nxSh.getRange(row, 6).setFormula('=IFERROR(INDEX(MASTER_HANG!C:C;MATCH(E' + row + ';MASTER_HANG!A:A;0));"")');
        nxSh.getRange(row, 7).setValue(pt.sl);
        nxSh.getRange(row, 8).setValue(String(r.khach || ''));
        nxSh.getRange(row, 10).setValue(String(r.ghichu || 'Nạp từ MISA'));
        nxSh.getRange(row, 11).setValue(String(r.kho || ''));
        added++;
      });
      if (!ensured[code.toUpperCase()]) { ensureTonRow_(ss, code); ensured[code.toUpperCase()] = 1; }
    });
    log_(actor_(sess), 'nx', 'Nạp phiếu MISA', '', 'Thêm ' + added + ' dòng phiếu, trùng bỏ qua ' + dup + ', mã mới chưa nạp ' + news.length);
    var msg = 'Đã nạp phiếu: thêm ' + added + ' dòng, bỏ qua ' + dup + ' dòng trùng (đã có trong hệ thống).';
    if (news.length) msg += ' Có ' + news.length + ' mã chưa có trong danh mục nên chưa nạp được: ' + news.slice(0, 10).join(', ') + (news.length > 10 ? '…' : '') + '. Hãy đồng bộ danh mục từ Drive (hoặc nạp danh sách mã) trước, rồi nạp lại file này — các dòng trùng sẽ tự bỏ qua.';
    return msg;
  });
}
function driveId_(url) { var m = String(url || '').match(/[-\w]{25,}/); return m ? m[0] : ''; }
function getDriveCfg(token) {
  auth_(token);
  return { url: PropertiesService.getScriptProperties().getProperty('DRIVE_MASTER_URL') || '' };
}
function setDriveCfg(token, url) {
  var sess = auth_(token); requireWrite_(sess);
  if (!driveId_(url)) throw new Error('Link file Drive chưa đúng — hãy dán link file Google Sheet danh mục.');
  PropertiesService.getScriptProperties().setProperty('DRIVE_MASTER_URL', String(url).trim());
  log_(actor_(sess), 'hethong', 'Cài link Drive danh mục', '', '');
  return 'Đã lưu link file Drive danh mục.';
}
function stripVn_(s) { return String(s || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/đ/g, 'd'); }
function syncDriveMaster(token) {
  var sess = auth_(token); requireWrite_(sess);
  var url = PropertiesService.getScriptProperties().getProperty('DRIVE_MASTER_URL') || '';
  var id = driveId_(url);
  if (!id) throw new Error('Chưa cài link file Drive danh mục. Dán link vào ô phía trên rồi bấm Lưu link trước.');
  var src;
  try { src = SpreadsheetApp.openById(id); } catch (e) { throw new Error('Không mở được file Drive. Kiểm tra lại: file đã được chia sẻ cho tài khoản chủ hệ thống chưa, link có đúng không.'); }
  return withLock_(function () {
    var items = {};
    src.getSheets().forEach(function (sh) {
      var nm = sh.getName();
      var generic = /tong|bao cao|^sheet|demo/i.test(stripVn_(nm));
      var brand = generic ? '' : nm;
      var vals = sh.getDataRange().getValues();
      for (var i = 0; i < vals.length; i++) {
        var r = vals[i];
        var a = String(r[0] === null || r[0] === undefined ? '' : r[0]).trim();
        var b = String(r[1] === null || r[1] === undefined ? '' : r[1]).trim();
        var c = String(r[2] === null || r[2] === undefined ? '' : r[2]).trim();
        if (!b && a && isNaN(Number(a)) && !c) { brand = a; continue; }
        if (!b) continue;
        if (/ma hieu|model/i.test(stripVn_(b)) || /ten hang/i.test(stripVn_(c))) continue;
        if (!c && (r[3] === '' || r[3] === null || r[3] === undefined) && isNaN(Number(b))) { brand = b; continue; }
        items[b.toUpperCase()] = { code: b, name: c, brand: brand, qn: Number(r[3]) || 0, hn: Number(r[4]) || 0 };
      }
    });
    var list = Object.keys(items).map(function (k) { return items[k]; });
    if (!list.length) throw new Error('Không đọc được mã hàng nào trong file Drive. Kiểm tra lại cấu trúc file (cột B = mã hiệu, cột C = tên hàng).');
    var ss = ss_();
    var mhSh = ss.getSheetByName('MASTER_HANG');
    var mh = mhSh.getDataRange().getValues();
    var pos = {};
    for (var i = 1; i < mh.length; i++) if (mh[i][0]) pos[String(mh[i][0]).trim().toUpperCase()] = i + 1;
    var tonSh = ss.getSheetByName('TON_KHO');
    var tonPos = {};
    var tonVals = tonSh.getDataRange().getValues();
    for (var t = 1; t < tonVals.length; t++) if (tonVals[t][0]) tonPos[String(tonVals[t][0]).trim().toUpperCase()] = { row: t + 1, dk: (Number(tonVals[t][11]) || 0) + (Number(tonVals[t][12]) || 0) };
    var added = 0, updated = 0, opened = 0;
    list.forEach(function (it) {
      var key = it.code.toUpperCase();
      if (pos[key]) {
        var row = pos[key];
        var cur = mh[row - 1];
        var ch = false;
        if (it.name && String(cur[2] || '') !== it.name) { mhSh.getRange(row, 3).setValue(it.name); ch = true; }
        if (it.brand && String(cur[3] || '') !== it.brand) { mhSh.getRange(row, 4).setValue(it.brand); ch = true; }
        if (ch) updated++;
      } else {
        mhSh.appendRow([it.code, it.code, it.name || it.code, it.brand || '', '', '', 'Đang bán', 'Đồng bộ từ Drive']);
        pos[key] = mhSh.getLastRow();
        added++;
      }
      if (!tonPos[key]) { ensureTonRow_(ss, it.code); tonPos[key] = { row: tonSh.getLastRow(), dk: 0 }; }
      var tp = tonPos[key];
      if (tp.dk === 0 && (it.qn > 0 || it.hn > 0)) {
        tonSh.getRange(tp.row, 12).setValue(it.qn);
        tonSh.getRange(tp.row, 13).setValue(it.hn);
        tp.dk = it.qn + it.hn;
        opened++;
      }
    });
    log_(actor_(sess), 'mhang', 'Đồng bộ danh mục từ Drive', '', 'Thêm mới ' + added + ' mã, cập nhật ' + updated + ' mã, nạp tồn đầu kỳ ' + opened + ' mã');
    return 'Đồng bộ xong: thêm mới ' + added + ' mã, cập nhật tên/hãng ' + updated + ' mã, nạp tồn đầu kỳ 2 kho cho ' + opened + ' mã (chỉ nạp cho mã chưa có số đầu kỳ).';
  });
}

function setDuyet(token, row, status) {
  var sess = auth_(token); requireApprove_(sess);
  return withLock_(function () {
    var sh = ss_().getSheetByName('PHE_DUYET_GIA');
    sh.getRange(row, 9).setValue(status);
    sh.getRange(row, 10).setValue(sess.name);
    sh.getRange(row, 11).setValue(new Date());
    log_(actor_(sess), 'duyet', status, row, sh.getRange(row, 1).getValue() + ' — ' + sh.getRange(row, 4).getValue());
    return 'Đã cập nhật yêu cầu dòng ' + row + ' → ' + status;
  });
}
