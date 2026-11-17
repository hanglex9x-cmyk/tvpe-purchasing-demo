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
  nx:        { sheet: 'NHAP_XUAT_KHO', edit: [0, 1, 2, 3, 4, 6, 7, 9], label: 'Nhập / Xuất kho' },
  giavon:    { sheet: 'GIA_VON',       edit: [0, 1, 2, 3, 4, 11, 15], label: 'Giá vốn' },
  chiphi:    { sheet: 'CHI_PHI_NK',    edit: [0, 1, 2, 3, 4], label: 'Chi phí nhập khẩu' },
  ton:       { sheet: 'TON_KHO',       edit: [2, 6, 8], label: 'Kho & tồn kho' },
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
  if (role === 'Admin') return { cost: true, write: true, del: true, approve: true };
  var any = [g, h, i2, j].some(function (x) { return String(x || '') !== ''; });
  if (!any) {
    if (role === 'Mua hàng') return { cost: true, write: true, del: true, approve: false };
    return { cost: false, write: false, del: false, approve: false }; // Sales và vai trò khác
  }
  return { cost: g === 'Có', write: h === 'Có', del: i2 === 'Có', approve: j === 'Có' };
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
  if (['Admin', 'Mua hàng', 'Sales'].indexOf(role) < 0) throw new Error('Vai trò không hợp lệ.');
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
  if (['Admin', 'Mua hàng', 'Sales'].indexOf(role) < 0) throw new Error('Vai trò không hợp lệ.');
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
  if (!sess.perms || !sess.perms.cost) {
    // Tài khoản KHÔNG có quyền xem giá vốn: CHỈ giá bán (đã che giá sàn nội bộ) + tồn khả dụng (chỉ mã, tên, khả dụng) — server không gửi dữ liệu giá vốn/PO/công nợ.
    out.giaban = blankCols_(tab_('BANG_GIA_BAN'), [9]);
    var ton = tab_('TON_KHO');
    ton.rows.forEach(function (r) { for (var c = 0; c < r.length - 1; c++) { if ([0, 1, 7].indexOf(c) < 0) r[c] = ''; } });
    out.ton = ton;
    out.kpi = {};
    ['po', 'nx', 'giavon', 'chiphi', 'congno', 'thanhtoan', 'chungtu', 'duyet', 'giahang', 'mhang', 'mncc', 'mkhach', 'lichsu'].forEach(function (k) { out[k] = { headers: [], rows: [] }; });
    return out;
  }
  var dash = ss_().getSheetByName('DASHBOARD');
  var kpi = {};
  dash.getRange('A4:B15').getValues().forEach(function (r) { if (r[0]) kpi[String(r[0])] = r[1]; });
  dash.getRange('A28:B37').getValues().forEach(function (r) { if (r[0]) kpi[String(r[0])] = r[1]; });
  out.kpi = kpi;
  out.lichsu = tab_('LICH_SU');
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
function updateRow(token, tabKey, row, obj) {
  var sess = auth_(token); requireWrite_(sess);
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
  var sess = auth_(token); requireDelete_(sess);
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
  var sess = auth_(token); requireWrite_(sess);
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
  var sess = auth_(token); requireWrite_(sess);
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
    log_(actor_(sess), 'nx', 'Thêm', r, o.loai + ' ' + o.ma + ' — SL ' + o.sl);
    return 'Đã ghi phiếu ' + o.loai + ' ' + o.ma;
  });
}
function addThanhToan(token, o) {
  var sess = auth_(token); requireWrite_(sess);
  return withLock_(function () {
    var sh = ss_().getSheetByName('THANH_TOAN');
    sh.appendRow([o.ngay, o.inv, o.ncc, Number(o.tien), o.ht, o.ck || '', o.note || '']);
    log_(actor_(sess), 'thanhtoan', 'Thêm', sh.getLastRow(), o.inv + ' — ' + Number(o.tien).toLocaleString('vi-VN') + 'đ');
    return 'Đã ghi thanh toán cho ' + o.inv;
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
