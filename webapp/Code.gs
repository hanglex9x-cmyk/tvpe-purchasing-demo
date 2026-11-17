/**
 * TVPE ERP — Web App (Apps Script)
 * Giao diện ERP chạy trên link /exec, dữ liệu đọc/ghi trực tiếp vào Google Sheet TVPE.
 * Deploy: Deploy > New deployment > Web app > Execute as: Me > Who has access: Anyone with the link.
 */
var SHEET_ID = '1CpytSjhaYifMFo8SSCIEJA9GYKdINLEylXMPGpKSvh4';
var PORTAL_URL = 'https://docs.google.com/spreadsheets/d/1bQ-B2fJxBRRq-vYTMGc6BN8WnfMdcSytBUpyqufPSIM/edit';

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
  return {
    kpi: kpi,
    portal: PORTAL_URL,
    po: tab_('PO'),
    nx: tab_('NHAP_XUAT_KHO'),
    giaVon: tab_('GIA_VON'),
    ton: tab_('TON_KHO'),
    congNo: tab_('CONG_NO'),
    chungTu: tab_('CHUNG_TU'),
    giaBan: tab_('BANG_GIA_BAN'),
    duyet: tab_('PHE_DUYET_GIA'),
    chiPhi: tab_('CHI_PHI_NK'),
    giaHang: tab_('BANG_GIA_HANG'),
    mHang: tab_('MASTER_HANG'),
    mNCC: tab_('MASTER_NCC'),
    mKhach: tab_('MASTER_KHACH'),
    thanhToan: tab_('THANH_TOAN')
  };
}

/* ---------------- Ghi dữ liệu (có khóa chống ghi trùng) ---------------- */
function withLock_(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function addPO(o) {
  return withLock_(function () {
    var sh = ss_().getSheetByName('PO');
    var r = sh.getLastRow() + 1;
    // Layout PO hiện tại: A Mã dòng PO | B Mã ĐH nội bộ | C Khách/Dự án | D Số PO | E Ngày PO | F Mã NCC |
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
    sh.getRange(r, 23).setValue(o.phutrach || '');
    return 'Đã thêm dòng PO ' + o.po + '/' + o.ma;
  });
}

function addPhieu(o) {
  return withLock_(function () {
    var sh = ss_().getSheetByName('NHAP_XUAT_KHO');
    var r = sh.getLastRow() + 1;
    // Layout NHAP_XUAT_KHO: A Ngày | B Loại | C Mã phiếu | D Mã dòng PO | E Product Code | F Tên hàng (công thức) | G SL | H Khách/Dự án | I Link chứng từ | J Ghi chú
    sh.getRange(r, 1, 1, 5).setValues([[o.ngay, o.loai, o.maphieu || '', o.madongpo || '', o.ma]]);
    sh.getRange(r, 6).setFormula('=IFERROR(INDEX(MASTER_HANG!C:C;MATCH(E' + r + ';MASTER_HANG!A:A;0));"")');
    sh.getRange(r, 7).setValue(Number(o.sl));
    sh.getRange(r, 8).setValue(o.khach || '');
    sh.getRange(r, 10).setValue(o.note || '');
    return 'Đã ghi phiếu ' + o.loai + ' ' + o.ma;
  });
}

function addThanhToan(o) {
  return withLock_(function () {
    var sh = ss_().getSheetByName('THANH_TOAN');
    sh.appendRow([o.ngay, o.inv, o.ncc, Number(o.tien), o.ht, o.ck || '', o.note || '']);
    return 'Đã ghi thanh toán cho ' + o.inv;
  });
}

function setDuyet(row, status, nguoi) {
  return withLock_(function () {
    var sh = ss_().getSheetByName('PHE_DUYET_GIA');
    sh.getRange(row, 9).setValue(status);
    sh.getRange(row, 10).setValue(nguoi || 'Quản lý mua hàng');
    sh.getRange(row, 11).setValue(new Date());
    return 'Đã cập nhật yêu cầu dòng ' + row + ' → ' + status;
  });
}
