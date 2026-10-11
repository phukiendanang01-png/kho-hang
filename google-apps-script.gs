/**
 * ============================================================
 *  Hậu Store - Quản lý kho hàng
 *  Google Apps Script (Web App) - đồng bộ với Google Sheet
 *
 *  Cách cài: xem file HUONG-DAN.txt
 *  - Mở Google Sheet -> Extensions -> Apps Script
 *  - Dán toàn bộ file này -> Lưu
 *  - Deploy -> New deployment -> Web app
 *      Execute as: Me
 *      Who has access: Anyone
 *  - Copy URL .../exec dán vào app (tab Cài đặt)
 * ============================================================
 */

var SHEET_ID = ''; // để trống nếu script mở từ Extensions > Apps Script của chính Sheet đó
var TAB_IN = 'Nhap';   // sheet dữ liệu nhập kho
var TAB_OUT = 'Xuat';  // sheet dữ liệu xuất kho
var TZ = 'Asia/Ho_Chi_Minh';

var IN_COLS = ['ID', 'ts', 'Thời gian nhập', 'Tên sản phẩm', 'Nhà cung cấp',
  'IMEI', 'Giá nhập', 'Ghi chú', 'Ngày tạo'];
var OUT_COLS = ['ID', 'ts', 'Thời gian xuất', 'Tên sản phẩm', 'Tên khách hàng',
  'Số điện thoại', 'IMEI', 'Bảo hành (tháng)', 'Hạn bảo hành', 'Giá xuất',
  'Ghi chú', 'linkId', 'Ngày tạo'];

/* ---------- tiện ích ---------- */
function fmtDT_(ms) {
  if (!ms) return '';
  return Utilities.formatDate(new Date(Number(ms)), TZ, 'dd/MM/yyyy HH:mm');
}

function warrantyDisplay_(ts, months) {
  if (!ts || !months) return 'Không BH';
  var d = new Date(Number(ts));
  d.setMonth(d.getMonth() + Number(months));
  return Utilities.formatDate(d, TZ, 'dd/MM/yyyy');
}

function getSS_() {
  if (SHEET_ID) return SpreadsheetApp.openById(SHEET_ID);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Không tìm thấy Google Sheet (hãy dùng Extensions > Apps Script, hoặc điền SHEET_ID)');
  return ss;
}

function getTab_(name, cols) {
  var ss = getSS_();
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold');
    try { sh.setFrozenRows(1); } catch (e) {}
  }
  return sh;
}

/* ---------- chuyển đổi record <-> dòng sheet ---------- */
function recToRow_(r) {
  if (r.kind === 'out') {
    return [String(r.id), Number(r.ts) || 0, fmtDT_(r.ts),
      String(r.name || ''), String(r.customer || ''), String(r.phone || ''),
      String(r.imei || ''), Number(r.warrantyMonths) || 0,
      warrantyDisplay_(r.ts, r.warrantyMonths),
      Number(r.price) || 0, String(r.note || ''), String(r.linkId || ''),
      Number(r.created) || 0];
  }
  return [String(r.id), Number(r.ts) || 0, fmtDT_(r.ts),
    String(r.name || ''), String(r.supplier || ''), String(r.imei || ''),
    Number(r.price) || 0, String(r.note || ''), Number(r.created) || 0];
}

function rowToRec_(kind, row) {
  if (kind === 'out') {
    return {
      id: String(row[0]), kind: 'out', ts: Number(row[1]) || 0,
      name: String(row[3] || ''), customer: String(row[4] || ''),
      phone: String(row[5] || ''), imei: String(row[6] || ''),
      warrantyMonths: Number(row[7]) || 0, price: Number(row[9]) || 0,
      note: String(row[10] || ''), linkId: String(row[11] || ''),
      created: Number(row[12]) || 0
    };
  }
  return {
    id: String(row[0]), kind: 'in', ts: Number(row[1]) || 0,
    name: String(row[3] || ''), supplier: String(row[4] || ''),
    imei: String(row[5] || ''), price: Number(row[6]) || 0,
    note: String(row[7] || ''), created: Number(row[8]) || 0
  };
}

/* ---------- CRUD (ghi raw bằng setValues để giữ số điện thoại/IMEI là chuỗi) ---------- */
function findRow_(sh, cols, id) {
  var last = sh.getLastRow();
  if (last < 2) return -1;
  var ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) return 2 + i;
  }
  return -1;
}

function upsert_(kind, r) {
  var cols = kind === 'out' ? OUT_COLS : IN_COLS;
  var sh = getTab_(kind === 'out' ? TAB_OUT : TAB_IN, cols);
  var target = findRow_(sh, cols, r.id);
  if (target < 0) target = sh.getLastRow() + 1;
  sh.getRange(target, 1, 1, cols.length).setValues([recToRow_(r)]);
}

function deleteById_(id) {
  var pairs = [[TAB_IN, IN_COLS], [TAB_OUT, OUT_COLS]];
  for (var i = 0; i < pairs.length; i++) {
    var sh = getTab_(pairs[i][0], pairs[i][1]);
    var row = findRow_(sh, pairs[i][1], id);
    if (row > 0) { sh.deleteRow(row); return true; }
  }
  return false;
}

function readTab_(kind) {
  var cols = kind === 'out' ? OUT_COLS : IN_COLS;
  var sh = getTab_(kind === 'out' ? TAB_OUT : TAB_IN, cols);
  var out = [];
  var last = sh.getLastRow();
  if (last < 2) return out;
  var rows = sh.getRange(2, 1, last - 1, cols.length).getValues();
  for (var i = 0; i < rows.length; i++) {
    if (!rows[i][0]) continue;
    out.push(rowToRec_(kind, rows[i]));
  }
  return out;
}

/* ---------- entry point ---------- */
function doGet(e) {
  var p = (e && e.parameter) || {};
  var result;
  try {
    var op = p.op || 'ping';
    if (op === 'ping') {
      var shIn = getTab_(TAB_IN, IN_COLS);
      var shOut = getTab_(TAB_OUT, OUT_COLS);
      result = {
        ok: 1, ping: 1,
        countIn: Math.max(0, shIn.getLastRow() - 1),
        countOut: Math.max(0, shOut.getLastRow() - 1),
        name: getSS_().getName ? getSS_().getName() : ''
      };
    } else if (op === 'read') {
      result = { ok: 1, 'in': readTab_('in'), out: readTab_('out') };
    } else if (op === 'push') {
      var payload = JSON.parse(p.data || '{}');
      var lock = LockService.getScriptLock();
      lock.waitLock(15000);
      try {
        var deleted = 0, pushed = 0;
        var dels = payload.d || [];
        for (var i = 0; i < dels.length; i++) {
          if (dels[i] && deleteById_(dels[i])) deleted++;
        }
        var recs = payload.r || [];
        for (var j = 0; j < recs.length; j++) {
          var r = recs[j];
          if (!r || !r.id) continue;
          upsert_(r.kind === 'out' ? 'out' : 'in', r);
          pushed++;
        }
        result = { ok: 1, pushed: pushed, deleted: deleted };
      } finally {
        lock.releaseLock();
      }
    } else {
      result = { ok: 0, error: 'op không hợp lệ: ' + op };
    }
  } catch (err) {
    result = { ok: 0, error: String((err && err.message) || err) };
  }

  // JSONP: chỉ chấp nhận callback là tên hàm hợp lệ (chống chèn mã) — check CHUỖI GỐC
  var cb = String(p.cb || '');
  if (/^[A-Za-z_$][A-Za-z0-9_$.]{0,127}$/.test(cb)) {
    return ContentService.createTextOutput(cb + '(' + JSON.stringify(result) + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}
