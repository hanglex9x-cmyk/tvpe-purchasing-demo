/**
 * TVPE ERP — Web App (Apps Script) v2
 * Giao diện ERP trên link /exec, dữ liệu đọc/ghi trực tiếp vào Google Sheet TVPE.
 * Có đủ Thêm / Sửa / Xóa cho từng module; mọi thao tác ghi đều lưu vào tab LICH_SU kèm tên người thao tác.
 * Deploy: Deploy > Manage deployments > Edit > Version: New version (giữ nguyên link /exec).
 */
var SHEET_ID = '1CpytSjhaYifMFo8SSCIEJA9GYKdINLEylXMPGpKSvh4';
var PORTAL_URL = 'https://docs.google.com/spreadsheets/d/1bQ-B2fJxBRRq-vYTMGc6BN8WnfMdcSytBUpyqufPSIM/edit';

/* Cấu hình từng bảng: tên tab trong Sheet + các cột (0-based) ĐƯỢC PHÉP sửa qua app.
   Các cột công thức không nằm trong danh sách này nên không sửa/xóa nhầm được. */
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
    r.push(i + 1); // _row = số dòng thật trong Sheet
    rows.push(r);
  }
  return { headers: headers, rows: rows };
}

function getAllData() {
  var ss = ss_();
  var dash = ss.getSheetByName('DASHBOARD');
  var kpi = {};
  dash.getRange('A4:B15').getValues().forEach(function (r) { if (r[0]) kpi[String(r[0])] = r[1]; });
  dash.getRange('A28:B37').getValues().forEach(function (r) { if (r[0]) kpi[String(r[0])] = r[1]; });
  var out = { kpi: kpi, portal: PORTAL_URL, lichsu: tab_('LICH_SU') };
  Object.keys(TABS).forEach(function (k) { out[k] = tab_(TABS[k].sheet); });
  return out;
}

/* ---------------- Nhật ký thao tác theo người dùng ---------------- */
function log_(user, tabKey, action, row, summary) {
  try {
    var sh = ss_().getSheetByName('LICH_SU');
    if (!sh) return;
    sh.appendRow([new Date(), user || '(chưa nhập tên)', TABS[tabKey] ? TABS[tabKey].label : tabKey, action, row || '', summary || '']);
  } catch (e) { /* ghi log lỗi không được làm hỏng thao tác chính */ }
}

/* ---------------- Ghi dữ liệu (có khóa chống ghi trùng) ---------------- */
function withLock_(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

/* Sửa 1 dòng: obj = { số cột (0-based) dạng chuỗi: giá trị } — chỉ nhận cột trong danh sách edit */
function updateRow(tabKey, row, obj, user) {
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
    log_(user, tabKey, 'Sửa', row, changed.join('; '));
    return 'Đã sửa dòng ' + row + ' (' + cfg.label + ')';
  });
}

/* Xóa 1 dòng */
function deleteRow(tabKey, row, user) {
  var cfg = TABS[tabKey];
  if (!cfg) throw new Error('Bảng không hợp lệ: ' + tabKey);
  return withLock_(function () {
    var sh = ss_().getSheetByName(cfg.sheet);
    var preview = sh.getRange(row, 1, 1, Math.min(4, sh.getLastColumn())).getValues()[0].join(' | ');
    sh.deleteRow(row);
    log_(user, tabKey, 'Xóa', row, preview);
    return 'Đã xóa dòng ' + row + ' (' + cfg.label + ')';
  });
}

/* Thêm dòng chung cho các bảng nhập tay (master, chứng từ, giá bán, chi phí, giá hãng, duyet...) */
function addRow(tabKey, values, user) {
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
    // Sao chép công thức từ dòng trên cho các cột không nhập tay (giữ nguyên liên kết tự động)
    if (nr > 2) {
      for (var c2 = 0; c2 < n; c2++) {
        if (cfg.edit.indexOf(c2) >= 0) continue;
        var prev = sh.getRange(nr - 1, c2 + 1);
        if (prev.getFormula()) prev.copyTo(sh.getRange(nr, c2 + 1));
      }
    }
    log_(user, tabKey, 'Thêm', nr, row.slice(0, 4).join(' | '));
    return 'Đã thêm dòng mới (' + cfg.label + ')';
  });
}

function addPO(o, user) {
  return withLock_(function () {
    var sh = ss_().getSheetByName('PO');
    var r = sh.getLastRow() + 1;
    // Layout PO: A Mã dòng PO | B Mã ĐH nội bộ | C Khách/Dự án | D Số PO | E Ngày PO | F Mã NCC |
    // G Product Code | H Mô tả (công thức) | I SL đặt | J Đã nhận (công thức) | K Còn thiếu (công thức) |
    // L Giá list | M CK chuẩn % | N CK đặc biệt % | O Giá mua NT (công thức) | P Tiền tệ | Q Tỷ giá |
    // R Thành tiền VND (công thức) | S Trạng thái | T ETA | U Điều khoản TT | V Link hồ sơ | W Người phụ trách
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
    sh.getRange(r, 23).setValue(o.phutrach || user || '');
    log_(user, 'po', 'Thêm', r, o.po + '/' + o.ma + ' — SL ' + o.sl);
    return 'Đã thêm dòng PO ' + o.po + '/' + o.ma;
  });
}

function addPhieu(o, user) {
  return withLock_(function () {
    var sh = ss_().getSheetByName('NHAP_XUAT_KHO');
    var r = sh.getLastRow() + 1;
    // Layout NHAP_XUAT_KHO: A Ngày | B Loại | C Mã phiếu | D Mã dòng PO | E Product Code | F Tên hàng (công thức) | G SL | H Khách/Dự án | I Link chứng từ | J Ghi chú
    sh.getRange(r, 1, 1, 5).setValues([[o.ngay, o.loai, o.maphieu || '', o.madongpo || '', o.ma]]);
    sh.getRange(r, 6).setFormula('=IFERROR(INDEX(MASTER_HANG!C:C;MATCH(E' + r + ';MASTER_HANG!A:A;0));"")');
    sh.getRange(r, 7).setValue(Number(o.sl));
    sh.getRange(r, 8).setValue(o.khach || '');
    sh.getRange(r, 10).setValue(o.note || '');
    log_(user, 'nx', 'Thêm', r, o.loai + ' ' + o.ma + ' — SL ' + o.sl);
    return 'Đã ghi phiếu ' + o.loai + ' ' + o.ma;
  });
}

function addThanhToan(o, user) {
  return withLock_(function () {
    var sh = ss_().getSheetByName('THANH_TOAN');
    sh.appendRow([o.ngay, o.inv, o.ncc, Number(o.tien), o.ht, o.ck || '', o.note || '']);
    log_(user, 'thanhtoan', 'Thêm', sh.getLastRow(), o.inv + ' — ' + Number(o.tien).toLocaleString('vi-VN') + 'đ');
    return 'Đã ghi thanh toán cho ' + o.inv;
  });
}

function setDuyet(row, status, user) {
  return withLock_(function () {
    var sh = ss_().getSheetByName('PHE_DUYET_GIA');
    sh.getRange(row, 9).setValue(status);
    sh.getRange(row, 10).setValue(user || 'Quản lý mua hàng');
    sh.getRange(row, 11).setValue(new Date());
    log_(user, 'duyet', status, row, sh.getRange(row, 1).getValue() + ' — ' + sh.getRange(row, 4).getValue());
    return 'Đã cập nhật yêu cầu dòng ' + row + ' → ' + status;
  });
}
