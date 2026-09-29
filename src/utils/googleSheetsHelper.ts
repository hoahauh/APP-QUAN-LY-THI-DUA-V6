import { WeeklyClassScore, ViolationRecord, ClassInfo, BonusRecord, ViolationRule, BonusRule, GoogleSheetsConfig, SheetsSyncData } from '../types';

export const GOOGLE_APPS_SCRIPT_TEMPLATE = `/**
 * Google Apps Script - Webhook & Database API
 * Quản Lý Thi Đua Liên Đội Trường TH&THCS Phước Hiệp
 * 
 * Hướng dẫn cài đặt 2 chiều (Đẩy dữ liệu lên & Nhận dữ liệu về app):
 * 1. Mở file Google Sheet mới trên Google Drive của trường
 * 2. Vào menu "Tiện ích mở rộng" (Extensions) -> "Apps Script"
 * 3. Xóa code cũ và dán toàn bộ đoạn code này vào file Code.gs
 * 4. Bấm "Triển khai" (Deploy) -> "Tùy chọn triển khai mới" (New deployment)
 * 5. Chọn loại: "Ứng dụng web" (Web app)
 * 6. Cấu hình:
 *    - Thực thi dưới dạng (Execute as): Tôi (Me)
 *    - Ai có quyền truy cập (Who has access): Bất kỳ ai (Anyone)
 * 7. Bấm "Triển khai" (Deploy), cấp quyền và copy Web App URL dán vào ứng dụng.
 */

function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) ? e.parameter.action : 'ping';
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  if (action === 'ping') {
    return ContentService.createTextOutput(JSON.stringify({
      status: 'success',
      message: 'Kết nối Google Sheets thành công! Hệ thống Quản lý thi đua LĐ trường TH&THCS Phước Hiệp sẵn sàng.',
      timestamp: new Date().toISOString()
    })).setMimeType(ContentService.MimeType.JSON);
  }

  // Lấy toàn bộ database để ứng dụng nhận về
  if (action === 'getAllData') {
    var result = {
      status: 'success',
      timestamp: new Date().toISOString(),
      data: {
        violations: readSheetData(ss, 'Bang_Vi_Pham'),
        bonuses: readSheetData(ss, 'Bang_Diem_Thuong'),
        classes: readSheetData(ss, 'Danh_Sach_Lop'),
        violationRules: readSheetData(ss, 'Quy_Dinh_Vi_Pham'),
        bonusRules: readSheetData(ss, 'Quy_Dinh_Diem_Thuong')
      }
    };
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  if (action === 'getViolations') {
    return ContentService.createTextOutput(JSON.stringify({
      status: 'success',
      data: readSheetData(ss, 'Bang_Vi_Pham')
    })).setMimeType(ContentService.MimeType.JSON);
  }

  if (action === 'getBonuses') {
    return ContentService.createTextOutput(JSON.stringify({
      status: 'success',
      data: readSheetData(ss, 'Bang_Diem_Thuong')
    })).setMimeType(ContentService.MimeType.JSON);
  }

  if (action === 'getClasses') {
    return ContentService.createTextOutput(JSON.stringify({
      status: 'success',
      data: readSheetData(ss, 'Danh_Sach_Lop')
    })).setMimeType(ContentService.MimeType.JSON);
  }

  if (action === 'getRules') {
    return ContentService.createTextOutput(JSON.stringify({
      status: 'success',
      data: {
        violations: readSheetData(ss, 'Quy_Dinh_Vi_Pham'),
        bonuses: readSheetData(ss, 'Quy_Dinh_Diem_Thuong')
      }
    })).setMimeType(ContentService.MimeType.JSON);
  }

  return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: 'Action not supported: ' + action }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);
    var action = payload.action;
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Ghi nhận bản ghi vi phạm mới
    if (action === 'recordViolation') {
      var sheet = getOrCreateSheet(ss, 'Bang_Vi_Pham');
      var item = payload.data;
      sheet.appendRow([
        item.id,
        item.weekNumber,
        item.classId,
        item.date,
        item.dayOfWeek,
        item.session,
        item.ruleCode,
        item.ruleName,
        item.category,
        item.pointsDeducted,
        item.studentName || '',
        item.note || '',
        item.reporterName,
        new Date().toISOString(),
        item.status || 'APPROVED',
        item.disputeReason || '',
        item.disputeReply || ''
      ]);
      return ContentService.createTextOutput(JSON.stringify({ status: 'success', id: item.id }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // Ghi nhận điểm thưởng mới
    if (action === 'recordBonus') {
      var sheetBonus = getOrCreateSheet(ss, 'Bang_Diem_Thuong');
      var b = payload.data;
      sheetBonus.appendRow([
        b.id,
        b.weekNumber,
        b.classId,
        b.date,
        b.ruleCode || '',
        b.title,
        b.pointsAdded,
        b.reason || '',
        b.approvedBy,
        new Date().toISOString()
      ]);
      return ContentService.createTextOutput(JSON.stringify({ status: 'success', id: b.id }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    // Đồng bộ bảng tổng hợp xếp hạng tuần
    if (action === 'syncSummary') {
      var sheetSummary = getOrCreateSheet(ss, 'Bang_Tong_Hop_Tuan');
      var summaryList = payload.data; // Mảng điểm 31 lớp
      summaryList.forEach(function(s) {
        sheetSummary.appendRow([
          s.weekNumber,
          s.classId,
          s.className,
          s.gradeLevel,
          s.teacherName,
          s.scoreNeNep,
          s.scoreHocTap,
          s.scorePhongTrao,
          s.totalBonus,
          s.totalScore,
          s.rankInGradeLevel,
          s.rankOverall,
          s.awardTier,
          s.hasRotatingFlag ? 'CỜ LUÂN LƯU' : '',
          new Date().toISOString()
        ]);
      });
      return ContentService.createTextOutput(JSON.stringify({ status: 'success', count: summaryList.length }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // Đồng bộ danh sách lớp Master Data
    if (action === 'syncClasses') {
      var sheetClasses = getOrCreateSheet(ss, 'Danh_Sach_Lop');
      var classList = payload.data;
      sheetClasses.clearContents();
      sheetClasses.appendRow(['id', 'name', 'gradeLevel', 'gradeNumber', 'teacherName', 'teacherEmail', 'studentCount', 'redStarStudent', 'roomNumber']);
      classList.forEach(function(c) {
        sheetClasses.appendRow([c.id, c.name, c.gradeLevel, c.gradeNumber, c.teacherName, c.teacherEmail, c.studentCount, c.redStarStudent, c.roomNumber]);
      });
      return ContentService.createTextOutput(JSON.stringify({ status: 'success', count: classList.length }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: 'Action not supported' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Đọc dữ liệu từ sheet trả về mảng object JSON
function readSheetData(ss, sheetName) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  var headers = data[0];
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    var rowObj = {};
    var hasData = false;
    for (var j = 0; j < headers.length; j++) {
      var key = headers[j].toString().trim();
      if (key) {
        rowObj[key] = data[i][j];
        if (data[i][j] !== '' && data[i][j] !== null) hasData = true;
      }
    }
    if (hasData) rows.push(rowObj);
  }
  return rows;
}

// Khởi tạo các Sheet chuẩn tự động
function getOrCreateSheet(ss, sheetName) {
  var sheet = ss.getSheetByName(sheetName);
  if (sheet) return sheet;
  
  if (sheetName === 'Bang_Vi_Pham') {
    sheet = ss.insertSheet('Bang_Vi_Pham');
    sheet.appendRow([
      'id', 'weekNumber', 'classId', 'date', 'dayOfWeek', 'session', 
      'ruleCode', 'ruleName', 'category', 'pointsDeducted', 'studentName', 
      'note', 'reporterName', 'createdAt', 'status', 'disputeReason', 'disputeReply'
    ]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, 17).setBackground('#1E3A8A').setFontColor('#FFFFFF').setFontWeight('bold');
  } else if (sheetName === 'Bang_Diem_Thuong') {
    sheet = ss.insertSheet('Bang_Diem_Thuong');
    sheet.appendRow(['id', 'weekNumber', 'classId', 'date', 'ruleCode', 'title', 'pointsAdded', 'reason', 'approvedBy', 'createdAt']);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, 10).setBackground('#065F46').setFontColor('#FFFFFF').setFontWeight('bold');
  } else if (sheetName === 'Bang_Tong_Hop_Tuan') {
    sheet = ss.insertSheet('Bang_Tong_Hop_Tuan');
    sheet.appendRow([
      'weekNumber', 'classId', 'className', 'gradeLevel', 'teacherName', 
      'scoreNeNep', 'scoreHocTap', 'scorePhongTrao', 'totalBonus', 
      'totalScore', 'rankInGradeLevel', 'rankOverall', 'awardTier', 'rotatingFlag', 'updatedAt'
    ]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, 15).setBackground('#047857').setFontColor('#FFFFFF').setFontWeight('bold');
  } else if (sheetName === 'Danh_Sach_Lop') {
    sheet = ss.insertSheet('Danh_Sach_Lop');
    sheet.appendRow(['id', 'name', 'gradeLevel', 'gradeNumber', 'teacherName', 'teacherEmail', 'studentCount', 'redStarStudent', 'roomNumber']);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, 9).setBackground('#4338CA').setFontColor('#FFFFFF').setFontWeight('bold');
  } else if (sheetName === 'Quy_Dinh_Vi_Pham') {
    sheet = ss.insertSheet('Quy_Dinh_Vi_Pham');
    sheet.appendRow(['id', 'code', 'name', 'category', 'defaultPoints', 'applicableLevel', 'description']);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, 7).setBackground('#B91C1C').setFontColor('#FFFFFF').setFontWeight('bold');
  } else if (sheetName === 'Quy_Dinh_Diem_Thuong') {
    sheet = ss.insertSheet('Quy_Dinh_Diem_Thuong');
    sheet.appendRow(['id', 'code', 'name', 'defaultPoints', 'description']);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, 5).setBackground('#D97706').setFontColor('#FFFFFF').setFontWeight('bold');
  } else {
    sheet = ss.insertSheet(sheetName);
  }
  return sheet;
}
`;

/**
 * Trích xuất Sheet ID từ URL Google Sheets bất kỳ hoặc ID thuần túy
 */
export function extractSheetIdFromUrl(urlOrId: string): string {
  if (!urlOrId) return '';
  const trimmed = urlOrId.trim();
  // Khớp định dạng: /spreadsheets/d/([a-zA-Z0-9-_]+)
  const match = trimmed.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  if (match && match[1]) {
    return match[1];
  }
  // Nếu là ID thô không chứa URL
  if (/^[a-zA-Z0-9-_]{20,}$/.test(trimmed)) {
    return trimmed;
  }
  return trimmed;
}

/**
 * Phân tích dữ liệu trả về từ Google Visualization API (gviz/tq)
 */
export function parseGvizResponse(rawText: string): any[] {
  try {
    const match = rawText.match(/google\.visualization\.Query\.setResponse\(([\s\S]*)\);?/);
    if (!match || !match[1]) {
      // Trường hợp rawText có thể đã là JSON
      if (rawText.trim().startsWith('{')) {
        const parsed = JSON.parse(rawText);
        if (parsed.data) return Array.isArray(parsed.data) ? parsed.data : [];
        if (parsed.table) return parseGvizTable(parsed.table);
      }
      return [];
    }

    const payload = JSON.parse(match[1]);
    if (!payload.table) return [];
    return parseGvizTable(payload.table);
  } catch (err) {
    console.warn('Lỗi phân tích gviz response:', err);
    return [];
  }
}

function parseGvizTable(table: any): any[] {
  let colNames = (table.cols || []).map((c: any, idx: number) => {
    return (c && c.label && typeof c.label === 'string' && c.label.trim()) ? c.label.trim() : '';
  });

  let startRowIdx = 0;
  // Nếu cols không có nhãn hoặc chỉ là chữ cái A, B, C... thì kiểm tra dòng 0 có phải là tiêu đề cột không
  const hasValidLabels = colNames.some(n => n.length > 0 && !/^[A-Z]{1,3}$/.test(n));
  if (!hasValidLabels && table.rows && table.rows.length > 0) {
    const firstRow = table.rows[0];
    if (firstRow && firstRow.c) {
      const candidateHeaders = firstRow.c.map((cell: any, idx: number) => {
        const v = cell ? (cell.v !== undefined ? cell.v : cell.f) : '';
        return typeof v === 'string' && v.trim().length > 0 ? v.trim() : `col_${idx}`;
      });
      // Nếu có ít nhất 1 header có ý nghĩa
      if (candidateHeaders.some((h: string) => h && !h.startsWith('col_'))) {
        colNames = candidateHeaders;
        startRowIdx = 1;
      }
    }
  }

  // Dự phòng cho cột rỗng
  colNames = colNames.map((name, idx) => name || (table.cols && table.cols[idx] && table.cols[idx].id) || `col_${idx}`);

  const rows: any[] = [];
  for (let i = startRowIdx; i < (table.rows || []).length; i++) {
    const r = table.rows[i];
    if (!r || !r.c) continue;
    const rowObj: Record<string, any> = {};
    let hasVal = false;
    r.c.forEach((cell: any, idx: number) => {
      const colName = colNames[idx] || `col_${idx}`;
      const val = cell ? (cell.v !== undefined ? cell.v : cell.f) : '';
      rowObj[colName] = val;
      rowObj[`col_${idx}`] = val;
      if (val !== '' && val !== null && val !== undefined) hasVal = true;
    });
    if (hasVal) rows.push(rowObj);
  }
  return rows;
}

/**
 * Lấy giá trị trường dữ liệu không phân biệt hoa thường, dấu tiếng Việt hoặc khoảng trắng
 */
function getField(obj: any, ...aliases: string[]): any {
  if (!obj || typeof obj !== 'object') return undefined;
  
  // 1. Khớp chính xác
  for (const alias of aliases) {
    if (obj[alias] !== undefined && obj[alias] !== null && obj[alias] !== '') {
      return obj[alias];
    }
  }

  // 2. Khớp linh hoạt (bỏ dấu tiếng Việt, chữ thường, bỏ ký tự đặc biệt)
  const cleanStr = (s: string) =>
    s
      .toLowerCase()
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .replace(/[^a-z0-9]/g, '');

  const normAliases = aliases.map(cleanStr);
  const keys = Object.keys(obj);
  for (const k of keys) {
    const normK = cleanStr(k);
    if (normAliases.includes(normK)) {
      const val = obj[k];
      if (val !== undefined && val !== null && val !== '') {
        return val;
      }
    }
  }
  return undefined;
}

/**
 * Fetch bảng tính từ Google Sheet công khai (không cần Apps Script Web App)
 */
export async function fetchSheetViaGviz(sheetId: string, sheetName?: string): Promise<any[]> {
  const cleanId = extractSheetIdFromUrl(sheetId);
  if (!cleanId) throw new Error('Mã Google Sheet ID không hợp lệ');

  let url = `https://docs.google.com/spreadsheets/d/${cleanId}/gviz/tq?tqx=out:json`;
  if (sheetName) {
    url += `&sheet=${encodeURIComponent(sheetName)}`;
  }
  let response = await fetch(url);
  
  // Fallback: nếu lỗi với tên sheet cụ thể, thử fetch trang tính đầu tiên
  if (!response.ok && sheetName) {
    const fallbackUrl = `https://docs.google.com/spreadsheets/d/${cleanId}/gviz/tq?tqx=out:json`;
    const fallbackRes = await fetch(fallbackUrl);
    if (fallbackRes.ok) {
      const text = await fallbackRes.text();
      return parseGvizResponse(text);
    }
  }

  if (!response.ok) {
    throw new Error(`Không thể kết nối bảng tính (${sheetName || 'Mặc định'}). Hãy đảm bảo Google Sheet đã mở quyền "Bất kỳ ai có liên kết đều xem được" (HTTP ${response.status})`);
  }
  const text = await response.text();
  return parseGvizResponse(text);
}

/**
 * Quét kiểm tra trạng thái kết nối Google Sheet và thông báo số lượng bản ghi
 */
export async function scanGoogleSheetConnection(config: GoogleSheetsConfig): Promise<{
  connected: boolean;
  message: string;
  source: 'APPS_SCRIPT' | 'GVIZ' | 'NONE';
  counts: { violations: number; bonuses: number; classes: number; rules: number };
}> {
  const counts = { violations: 0, bonuses: 0, classes: 0, rules: 0 };

  // 1. Kiểm tra qua Apps Script Web App nếu có
  if (config.webAppUrl && config.webAppUrl.trim().startsWith('http')) {
    try {
      const url = `${config.webAppUrl.trim()}${config.webAppUrl.includes('?') ? '&' : '?'}action=getAllData&_t=${Date.now()}`;
      const res = await fetch(url, { method: 'GET' });
      if (res.ok) {
        const json = await res.json();
        if (json.status === 'success' && json.data) {
          counts.violations = Array.isArray(json.data.violations) ? json.data.violations.length : 0;
          counts.bonuses = Array.isArray(json.data.bonuses) ? json.data.bonuses.length : 0;
          counts.classes = Array.isArray(json.data.classes) ? json.data.classes.length : 0;
          counts.rules = Array.isArray(json.data.violationRules) ? json.data.violationRules.length : 0;

          return {
            connected: true,
            source: 'APPS_SCRIPT',
            message: `✓ Kết nối Webhook thành công! Tìm thấy: ${counts.violations} vi phạm, ${counts.bonuses} điểm thưởng, ${counts.classes} lớp học.`,
            counts
          };
        }
      }
    } catch (e: any) {
      console.warn('Apps script ping failed, falling back to gviz:', e);
    }
  }

  // 2. Kiểm tra qua Google Sheet ID (Gviz)
  const sheetId = extractSheetIdFromUrl(config.sheetId);
  if (!sheetId) {
    return {
      connected: false,
      source: 'NONE',
      message: 'Vui lòng dán đường link Google Sheet hoặc mã Sheet ID!',
      counts
    };
  }

  try {
    const sheetViPham = config.sheetNameViolations || 'Bang_Vi_Pham';
    const sheetLop = config.sheetNameClasses || 'Danh_Sach_Lop';
    const sheetBonus = config.sheetNameBonuses || 'Bang_Diem_Thuong';

    const [vRes, cRes, bRes] = await Promise.allSettled([
      fetchSheetViaGviz(sheetId, sheetViPham),
      fetchSheetViaGviz(sheetId, sheetLop),
      fetchSheetViaGviz(sheetId, sheetBonus),
    ]);

    let anySuccess = false;
    if (vRes.status === 'fulfilled' && vRes.value.length > 0) {
      counts.violations = mapRawToViolations(vRes.value).length;
      anySuccess = true;
    }
    if (cRes.status === 'fulfilled' && cRes.value.length > 0) {
      counts.classes = mapRawToClasses(cRes.value).length;
      anySuccess = true;
    }
    if (bRes.status === 'fulfilled' && bRes.value.length > 0) {
      counts.bonuses = mapRawToBonuses(bRes.value).length;
      anySuccess = true;
    }

    // Nếu không tìm thấy tên sheet chuẩn, thử quét sheet mặc định (trang tính đầu tiên)
    if (!anySuccess) {
      const defaultRows = await fetchSheetViaGviz(sheetId);
      if (defaultRows.length > 0) {
        // Tự động phân loại
        const mappedViolations = mapRawToViolations(defaultRows);
        const mappedClasses = mapRawToClasses(defaultRows);
        if (mappedViolations.length > 0) {
          counts.violations = mappedViolations.length;
        } else if (mappedClasses.length > 0) {
          counts.classes = mappedClasses.length;
        }
        return {
          connected: true,
          source: 'GVIZ',
          message: `✓ Kết nối Google Sheet thành công! Đã đọc ${defaultRows.length} dòng dữ liệu từ trang tính đầu tiên.`,
          counts
        };
      }
    }

    return {
      connected: true,
      source: 'GVIZ',
      message: `✓ Kết nối Google Sheet thành công! Tìm thấy: ${counts.violations} vi phạm, ${counts.bonuses} điểm thưởng, ${counts.classes} lớp học.`,
      counts
    };
  } catch (err: any) {
    return {
      connected: false,
      source: 'GVIZ',
      message: `Không thể kết nối: ${err?.message || 'Vui lòng kiểm tra quyền chia sẻ bảng tính: "Bất kỳ ai có đường liên kết đều có thể xem"'}`,
      counts
    };
  }
}

/**
 * Nhận toàn bộ database online từ Google Sheets
 * Tự động chọn:
 * 1. Gọi Web App URL (Apps Script API) nếu có
 * 2. Fallback sang Google Sheet ID công khai (gviz)
 */
export async function fetchFullOnlineDatabase(
  config: GoogleSheetsConfig
): Promise<{ success: boolean; data?: SheetsSyncData; source: 'APPS_SCRIPT' | 'GVIZ'; message: string }> {
  // Thử qua Google Apps Script Web App trước nếu có cấu hình
  if (config.webAppUrl && config.webAppUrl.trim().startsWith('http')) {
    try {
      const url = `${config.webAppUrl.trim()}${config.webAppUrl.includes('?') ? '&' : '?'}action=getAllData&_t=${Date.now()}`;
      const res = await fetch(url, { method: 'GET' });
      if (res.ok) {
        const json = await res.json();
        if (json.status === 'success' && json.data) {
          const formatted = normalizeDatabaseRecords(json.data);
          return {
            success: true,
            source: 'APPS_SCRIPT',
            data: formatted,
            message: 'Đã nhận toàn bộ cơ sở dữ liệu qua Google Apps Script Web App thành công!'
          };
        }
      }
    } catch (e: any) {
      console.warn('Lỗi khi fetch qua Apps Script, chuyển sang fetch qua Sheet ID:', e);
    }
  }

  // Fallback: Fetch trực tiếp qua Google Sheet ID (Public Sheet via Gviz)
  const sheetId = extractSheetIdFromUrl(config.sheetId);
  if (!sheetId) {
    return {
      success: false,
      source: 'GVIZ',
      message: 'Vui lòng cung cấp link Google Sheet hoặc Sheet ID để nhận dữ liệu online!'
    };
  }

  try {
    const sheetViPham = config.sheetNameViolations || 'Bang_Vi_Pham';
    const sheetLop = config.sheetNameClasses || 'Danh_Sach_Lop';
    const sheetBonus = config.sheetNameBonuses || 'Bang_Diem_Thuong';
    const sheetRules = config.sheetNameRules || 'Quy_Dinh_Vi_Pham';

    // Tải song song các sheet
    const [rawViolations, rawClasses, rawBonuses, rawRules] = await Promise.allSettled([
      fetchSheetViaGviz(sheetId, sheetViPham),
      fetchSheetViaGviz(sheetId, sheetLop),
      fetchSheetViaGviz(sheetId, sheetBonus),
      fetchSheetViaGviz(sheetId, sheetRules),
    ]);

    const resultData: SheetsSyncData = {};

    if (rawViolations.status === 'fulfilled' && rawViolations.value.length > 0) {
      resultData.violations = mapRawToViolations(rawViolations.value);
    }
    if (rawClasses.status === 'fulfilled' && rawClasses.value.length > 0) {
      resultData.classes = mapRawToClasses(rawClasses.value);
    }
    if (rawBonuses.status === 'fulfilled' && rawBonuses.value.length > 0) {
      resultData.bonuses = mapRawToBonuses(rawBonuses.value);
    }
    if (rawRules.status === 'fulfilled' && rawRules.value.length > 0) {
      resultData.violationRules = mapRawToRules(rawRules.value);
    }

    let totalRecords = 
      (resultData.violations?.length || 0) + 
      (resultData.classes?.length || 0) + 
      (resultData.bonuses?.length || 0) + 
      (resultData.violationRules?.length || 0);

    // Nếu không tìm thấy theo tên sheet chuẩn, thử tải trang tính mặc định đầu tiên
    if (totalRecords === 0) {
      try {
        const defaultRows = await fetchSheetViaGviz(sheetId);
        if (defaultRows.length > 0) {
          const autoV = mapRawToViolations(defaultRows);
          const autoC = mapRawToClasses(defaultRows);
          const autoB = mapRawToBonuses(defaultRows);
          if (autoV.length > 0) {
            resultData.violations = autoV;
          } else if (autoC.length > 0) {
            resultData.classes = autoC;
          } else if (autoB.length > 0) {
            resultData.bonuses = autoB;
          }
          totalRecords = (resultData.violations?.length || 0) + (resultData.classes?.length || 0) + (resultData.bonuses?.length || 0);
        }
      } catch {
        // bỏ qua nếu lỗi
      }
    }

    if (totalRecords === 0) {
      return {
        success: false,
        source: 'GVIZ',
        message: 'Đã kết nối Google Sheet nhưng không tìm thấy dữ liệu phù hợp. Vui lòng kiểm tra lại tên sheet hoặc định dạng cột.'
      };
    }

    return {
      success: true,
      source: 'GVIZ',
      data: resultData,
      message: `Đã nhận thành công dữ liệu từ Google Sheet (${resultData.violations?.length || 0} vi phạm, ${resultData.classes?.length || 0} lớp, ${resultData.bonuses?.length || 0} điểm thưởng)!`
    };
  } catch (err: any) {
    return {
      success: false,
      source: 'GVIZ',
      message: `Lỗi nhận dữ liệu từ Google Sheets: ${err?.message || 'Không thể truy cập Google Sheet. Vui lòng kiểm tra quyền chia sẻ công khai.'}`
    };
  }
}

/**
 * Chuẩn hóa các record từ Apps Script trả về
 */
function normalizeDatabaseRecords(raw: any): SheetsSyncData {
  const res: SheetsSyncData = {};
  if (Array.isArray(raw.violations)) {
    res.violations = mapRawToViolations(raw.violations);
  }
  if (Array.isArray(raw.classes)) {
    res.classes = mapRawToClasses(raw.classes);
  }
  if (Array.isArray(raw.bonuses)) {
    res.bonuses = mapRawToBonuses(raw.bonuses);
  }
  if (Array.isArray(raw.violationRules)) {
    res.violationRules = mapRawToRules(raw.violationRules);
  }
  return res;
}

export function mapRawToViolations(rows: any[]): ViolationRecord[] {
  return rows.map((r, i) => {
    const id = String(getField(r, 'id', 'Mã Bản Ghi', 'Mã Ghi Nhận', 'Mã Vi Phạm', 'Mã') || `v-sheet-${i + 1}`);
    const weekNumber = Number(getField(r, 'weekNumber', 'Tuần', 'Tuần Học', 'Số Tuần', 'Week') || 1);
    const classId = String(getField(r, 'classId', 'Lớp', 'Mã Lớp', 'Tên Lớp', 'Class') || '').trim().toUpperCase();
    const date = String(getField(r, 'date', 'Ngày Vi Phạm', 'Ngày', 'Date') || new Date().toISOString().split('T')[0]);
    const dayOfWeek = String(getField(r, 'dayOfWeek', 'Thứ', 'Ngày Trong Tuần') || 'Thứ Hai');
    const rawSession = String(getField(r, 'session', 'Buổi', 'Ca Trực', 'Session') || 'SANG').toUpperCase();
    const session = rawSession.includes('CHIEU') ? 'CHIEU' : 'SANG';
    const ruleCode = String(getField(r, 'ruleCode', 'Mã Lỗi', 'Mã Quy Định', 'Mã Quy Chế', 'Code') || 'NN01');
    const ruleName = String(getField(r, 'ruleName', 'Tên Lỗi', 'Tên Vi Phạm', 'Nội Dung Vi Phạm', 'Nội Dung', 'Lỗi') || 'Vi phạm nề nếp');
    const rawCategory = String(getField(r, 'category', 'Danh Mục', 'Loại Vi Phạm', 'Nhóm') || 'NE_NEP').toUpperCase();
    const category = rawCategory.includes('HOC') ? 'HOC_TAP' : rawCategory.includes('PHONG') ? 'PHONG_TRAO' : 'NE_NEP';
    const pointsDeducted = Math.abs(Number(getField(r, 'pointsDeducted', 'Điểm Trừ', 'Số Điểm Trừ', 'Điểm') || 1));
    const studentName = String(getField(r, 'studentName', 'Học Sinh Vi Phạm', 'Học Sinh', 'Tên Học Sinh', 'HS') || '');
    const note = String(getField(r, 'note', 'Ghi Chú', 'Chi Tiết', 'Lý Do') || '');
    const reporterName = String(getField(r, 'reporterName', 'Người Ghi Nhận', 'Người Chấm', 'Cờ Đỏ', 'Cờ Đỏ Trực') || 'Cờ Đỏ');
    const createdAt = String(getField(r, 'createdAt', 'Thời Gian Ghi', 'Ngày Tạo') || new Date().toISOString());
    const rawStatus = String(getField(r, 'status', 'Trạng Thái') || 'APPROVED').toUpperCase();
    const status = ['APPROVED', 'DISPUTED', 'RESOLVED', 'CANCELLED'].includes(rawStatus) ? rawStatus as any : 'APPROVED';
    const disputeReason = String(getField(r, 'disputeReason', 'Lý Do Khiếu Nại', 'Khiếu Nại') || '');
    const disputeReply = String(getField(r, 'disputeReply', 'Phản Hồi Khiếu Nại', 'Giải Quyết') || '');

    return {
      id,
      weekNumber: isNaN(weekNumber) ? 1 : weekNumber,
      classId,
      date,
      dayOfWeek,
      session,
      ruleId: ruleCode,
      ruleCode,
      ruleName,
      category,
      pointsDeducted: isNaN(pointsDeducted) ? 1 : pointsDeducted,
      studentName,
      note,
      reporterRole: 'CO_DO',
      reporterName,
      createdAt,
      status,
      disputeReason,
      disputeReply,
    };
  }).filter(v => v.classId && v.classId.length > 0);
}

export function mapRawToClasses(rows: any[]): ClassInfo[] {
  return rows.map((r) => {
    const id = String(getField(r, 'id', 'Mã Lớp', 'Lớp', 'Class') || '').trim().toUpperCase();
    const name = String(getField(r, 'name', 'Tên Lớp', 'Lớp') || `Lớp ${id}`);
    const rawLevel = String(getField(r, 'gradeLevel', 'Cấp Học', 'Khối Cấp', 'Cấp') || '').toUpperCase();
    const gradeLevel = rawLevel === 'THCS' || (id.startsWith('6') || id.startsWith('7') || id.startsWith('8') || id.startsWith('9')) ? 'THCS' : 'TIEU_HOC';
    const gradeNumber = Number(getField(r, 'gradeNumber', 'Khối', 'Khối Lớp') || parseInt(id[0] || '1', 10) || 1);
    const teacherName = String(getField(r, 'teacherName', 'GVCN', 'Giáo Viên Chủ Nhiệm', 'Giáo Viên') || 'Giáo viên chủ nhiệm');
    const teacherEmail = String(getField(r, 'teacherEmail', 'Email', 'Hòm Thư') || '');
    const studentCount = Number(getField(r, 'studentCount', 'Sĩ Số', 'Số Học Sinh') || 30);
    const redStarStudent = String(getField(r, 'redStarStudent', 'Học Sinh Cờ Đỏ', 'Cờ Đỏ') || '');
    const roomNumber = String(getField(r, 'roomNumber', 'Phòng Học', 'Phòng') || '');

    return {
      id,
      name,
      gradeLevel,
      gradeNumber: isNaN(gradeNumber) ? 1 : gradeNumber,
      teacherName,
      teacherEmail,
      studentCount: isNaN(studentCount) ? 30 : studentCount,
      redStarStudent,
      roomNumber
    };
  }).filter(c => c.id && c.id.length > 0);
}

export function mapRawToBonuses(rows: any[]): BonusRecord[] {
  return rows.map((r, i) => {
    const id = String(getField(r, 'id', 'Mã Bản Ghi', 'Mã Thưởng', 'Mã') || `b-sheet-${i + 1}`);
    const weekNumber = Number(getField(r, 'weekNumber', 'Tuần', 'Tuần Học') || 1);
    const classId = String(getField(r, 'classId', 'Lớp', 'Mã Lớp', 'Class') || '').trim().toUpperCase();
    const date = String(getField(r, 'date', 'Ngày', 'Ngày Thưởng') || new Date().toISOString().split('T')[0]);
    const ruleCode = String(getField(r, 'ruleCode', 'Mã', 'Mã Thưởng') || '');
    const title = String(getField(r, 'title', 'Nội Dung', 'Tên Khen Thưởng', 'Tiêu Đề Thưởng', 'Khen Thưởng') || 'Điểm thưởng thi đua');
    const pointsAdded = Math.abs(Number(getField(r, 'pointsAdded', 'Điểm Thưởng', 'Điểm Cộng', 'Điểm') || 1));
    const reason = String(getField(r, 'reason', 'Lý Do', 'Căn Cứ') || '');
    const approvedBy = String(getField(r, 'approvedBy', 'Người Duyệt', 'Người Khen Thưởng') || 'Tổng Phụ Trách');
    const createdAt = String(getField(r, 'createdAt', 'Thời Gian Ghi', 'Ngày Tạo') || new Date().toISOString());

    return {
      id,
      weekNumber: isNaN(weekNumber) ? 1 : weekNumber,
      classId,
      date,
      ruleCode,
      title,
      pointsAdded: isNaN(pointsAdded) ? 1 : pointsAdded,
      reason,
      approvedBy,
      createdAt
    };
  }).filter(b => b.classId && b.classId.length > 0);
}

export function mapRawToRules(rows: any[]): ViolationRule[] {
  return rows.map((r) => {
    const id = String(getField(r, 'id', 'code', 'Mã Lỗi', 'Mã Quy Định', 'Mã') || 'NN01');
    const code = String(getField(r, 'code', 'Mã', 'id') || id);
    const name = String(getField(r, 'name', 'Tên Lỗi', 'Tên Vi Phạm', 'Quy Định') || 'Lỗi vi phạm');
    const rawCat = String(getField(r, 'category', 'Danh Mục', 'Loại') || 'NE_NEP').toUpperCase();
    const category = rawCat.includes('HOC') ? 'HOC_TAP' : rawCat.includes('PHONG') ? 'PHONG_TRAO' : 'NE_NEP';
    const defaultPoints = Math.abs(Number(getField(r, 'defaultPoints', 'Điểm Trừ', 'Điểm Mặc Định') || 1));
    const rawLevel = String(getField(r, 'applicableLevel', 'Áp Dụng Cho', 'Cấp Học') || 'ALL').toUpperCase();
    const applicableLevel = rawLevel.includes('THCS') ? 'THCS' : rawLevel.includes('TIEU') ? 'TIEU_HOC' : 'ALL';
    const description = String(getField(r, 'description', 'Mô Tả', 'Ghi Chú') || '');

    return {
      id,
      code,
      name,
      category,
      defaultPoints: isNaN(defaultPoints) ? 1 : defaultPoints,
      applicableLevel,
      description
    };
  });
}

/**
 * Phân tích tệp CSV thuần túy khi người dùng tải tệp từ Google Sheets về máy
 */
export function parseCSVText(csvText: string): any[] {
  if (!csvText || !csvText.trim()) return [];
  const lines = csvText.trim().split(/\r?\n/).filter(line => line.trim().length > 0 && !line.startsWith('#'));
  if (lines.length < 2) return [];

  // Parse header
  const headers = parseCSVLine(lines[0]);
  const rows: any[] = [];

  for (let i = 1; i < lines.length; i++) {
    const values = parseCSVLine(lines[i]);
    const rowObj: Record<string, any> = {};
    let hasData = false;
    headers.forEach((h, idx) => {
      const val = values[idx] !== undefined ? values[idx].trim() : '';
      rowObj[h] = val;
      if (val) hasData = true;
    });
    if (hasData) rows.push(rowObj);
  }
  return rows;
}

function parseCSVLine(line: string): string[] {
  const result: string[] = [];
  let current = '';
  let inQuotes = false;
  for (let i = 0; i < line.length; i++) {
    const char = line[i];
    if (char === '"') {
      if (inQuotes && line[i + 1] === '"') {
        current += '"';
        i++;
      } else {
        inQuotes = !inQuotes;
      }
    } else if (char === ',' && !inQuotes) {
      result.push(current);
      current = '';
    } else {
      current += char;
    }
  }
  result.push(current);
  return result;
}

export function exportSummaryToCSV(scores: WeeklyClassScore[], weekNumber: number): string {
  const headers = [
    'Hạng Khối',
    'Hạng Trường',
    'Lớp',
    'Khối Cấp',
    'GVCN',
    'Điểm Nề Nếp',
    'Điểm Học Tập',
    'Điểm Phong Trào',
    'Điểm Thưởng (+)',
    'Tổng Điểm Tuần',
    'Xếp Loại',
    'Cờ Luân Lưu',
  ];

  const rows = scores.map((s) => [
    s.rankInGradeLevel,
    s.rankOverall,
    s.className,
    s.gradeLevel === 'TIEU_HOC' ? 'Tiểu học' : 'THCS',
    `"${s.teacherName}"`,
    `${s.scoreNeNep} (/${s.baseNeNep})`,
    `${s.scoreHocTap} (/${s.baseHocTap})`,
    `${s.scorePhongTrao} (/${s.basePhongTrao})`,
    `+${s.totalBonus}`,
    s.totalScore,
    s.awardTier === 'XUAT_SAC' ? 'Xuất sắc' : s.awardTier === 'TOT' ? 'Tốt' : s.awardTier === 'KHA' ? 'Khá' : 'Trung bình',
    s.hasRotatingFlag ? 'CỜ LUÂN LƯU' : '',
  ]);

  return [
    `# BẢNG TỔNG HỢP XẾP HẠNG THI ĐUA LIÊN ĐỘI TRƯỜNG TH&THCS PHƯỚC HIỆP - TUẦN ${weekNumber}`,
    headers.join(','),
    ...rows.map((r) => r.join(',')),
  ].join('\n');
}

export function exportViolationsToCSV(violations: ViolationRecord[], weekNumber: number): string {
  const headers = [
    'Mã Ghi Nhận',
    'Tuần',
    'Lớp',
    'Ngày',
    'Thứ',
    'Buổi',
    'Mã Quy Định',
    'Tên Vi Phạm',
    'Danh Mục',
    'Điểm Trừ',
    'Học Sinh',
    'Ghi Chú',
    'Người Ghi Nhận',
    'Trạng Thái',
  ];

  const rows = violations
    .filter((v) => v.weekNumber === weekNumber)
    .map((v) => [
      v.id,
      v.weekNumber,
      v.classId,
      v.date,
      v.dayOfWeek,
      v.session === 'SANG' ? 'Sáng' : 'Chiều',
      v.ruleCode,
      `"${v.ruleName}"`,
      v.category === 'NE_NEP' ? 'Nề nếp' : v.category === 'HOC_TAP' ? 'Học tập' : 'Phong trào',
      `-${v.pointsDeducted}`,
      `"${v.studentName || ''}"`,
      `"${v.note || ''}"`,
      `"${v.reporterName}"`,
      v.status === 'APPROVED' ? 'Đã duyệt' : v.status === 'DISPUTED' ? 'Khiếu nại' : v.status === 'RESOLVED' ? 'Đã giải quyết' : 'Đã hủy',
    ]);

  return [
    `# SỔ THEO DÕI VI PHẠM THI ĐUA NỀ NẾP - TUẦN ${weekNumber}`,
    headers.join(','),
    ...rows.map((r) => r.join(',')),
  ].join('\n');
}

export function downloadFile(content: string, fileName: string, mimeType = 'text/csv;charset=utf-8;') {
  const blob = new Blob(['\uFEFF' + content], { type: mimeType });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.setAttribute('href', url);
  link.setAttribute('download', fileName);
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}
