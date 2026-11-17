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
  var kpiVals = dash.getRange('A4:B15').getValues();
  var kpi = {};
  kpiVals.forEach(function (r) { if (r[0]) kpi[String(r[0])] = r[1]; });
  return {
    kpi: kpi,
    portal: PORTAL_URL,
    po: tab_('PO'),
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
    sh.getRange(r, 1, 1, 9).setValues([[o.po, o.ngay, o.ncc, o.ma, Number(o.sl), Number(o.list), o.tiente, Number(o.ck1) / 100, Number(o.ck2) / 100]]);
    sh.getRange(r, 10).setFormula('=IF(OR(F' + r + '="",H' + r + '="");"";F' + r + '*(1-H' + r + ')*(1-I' + r + '))');
    sh.getRange(r, 11).setFormula('=IF(OR(E' + r + '="",J' + r + '="");"";E' + r + '*J' + r + ')');
    sh.getRange(r, 12).setValue(Number(o.tygia));
    sh.getRange(r, 13).setFormula('=IF(OR(K' + r + '="",L' + r + '="");"";K' + r + '*L' + r + ')');
    sh.getRange(r, 16).setFormula('=SUMIFS(NHAP_XUAT_KHO!F:F;NHAP_XUAT_KHO!C:C;D' + r + ';NHAP_XUAT_KHO!B:B;"Nhập";NHAP_XUAT_KHO!D:D;A' + r + ')');
    sh.getRange(r, 17).setFormula('=IF(E' + r + '="";"";E' + r + '-P' + r + ')');
    sh.getRange(r, 18).setFormula('=IF(E' + r + '="";"";IF(P' + r + '>=E' + r + ';"Đã về đủ";IF(P' + r + '>0;"Về một phần";"Chưa về")))');
    sh.getRange(r, 19).setValue(o.eta);
    sh.getRange(r, 20).setFormula('=IF(OR(Q' + r + '<=0;S' + r + '="");"";IF(S' + r + '<TODAY();"🔴 Chậm ETA";IF(S' + r + '-TODAY()<=3;"🟠 Sắp đến ETA";"Đúng tiến độ")))');
    sh.getRange(r, 21).setValue(o.kho || 'Kho chính');
    sh.getRange(r, 22).setValue(o.khach || '');
    return 'Đã thêm PO ' + o.po;
  });
}

function addPhieu(o) {
  return withLock_(function () {
    var sh = ss_().getSheetByName('NHAP_XUAT_KHO');
    sh.appendRow([o.ngay, o.loai, o.ma, o.po || '', o.lo || '', Number(o.sl), o.khach || '', o.phieu || '']);
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
