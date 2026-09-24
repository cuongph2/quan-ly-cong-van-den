const SPREADSHEET_ID = '1AYQo4Gr6xvz-9V2n2O4Xz1TqzVEUbgaweXPkBRBZs1w';
const SHEET_NAME = 'CongVanNhapMoi';
const ARCHIVE_SHEET_NAME = 'CongVanDen';
const DRIVE_FOLDER_ID = '1X4H2exMveCn9Of_s1B3TsyIi3jlEtJSyfsDo8bwCjqmq0sFIzdRcgsSlkeKayJfgh1TsPsta';
const OCR_TEMP_FOLDER_NAME = 'TAMMI_OCR_TAM';

const HEADERS = [
  'ID',
  'MaCongVan',
  'SoDen',
  'NgayDen',
  'DiaBan',
  'SoKyHieu',
  'NgayVanBan',
  'CoQuanBanHanh',
  'TrichYeu',
  'LoaiVanBan',
  'DoKhan',
  'DoMat',
  'HanXuLy',
  'DonViChuTri',
  'NguoiXuLy',
  'TrangThai',
  'GhiChu',
  'DriveFileIds',
  'DriveFileUrls',
  'CreatedAt',
  'UpdatedAt',
  'CreatedBy',
  'FileScan',
  'OCRText',
  'SoKyHieuAI',
  'NgayVanBanAI',
  'CoQuanBanHanhAI',
  'TrichYeuAI',
  'AIConfidence',
  'AIStatus',
  'AILastRun',
  'AIError'
];

function doGet() {
  return HtmlService
    .createTemplateFromFile('index')
    .evaluate()
    .setTitle('Nhập công văn đến')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function submitCongVanDen(payload) {
  validatePayload_(payload);

  const sheet = getSheet_();
  ensureHeaders_(sheet);

  const now = new Date();
  const id = Utilities.getUuid();
  const maCongVan = buildMaCongVan_(payload.soDen, payload.ngayDen);
  const uploadedFiles = saveFiles_(payload.files || [], maCongVan);
  const primaryFile = uploadedFiles[0] || null;
  const extraction = primaryFile ? extractCongVanFromFile_(primaryFile.id) : emptyExtraction_('NO_FILE');
  const userEmail = Session.getActiveUser().getEmail() || '';

  const row = [
    id,
    maCongVan,
    payload.soDen || '',
    toDateOrBlank_(payload.ngayDen),
    payload.diaBan || '',
    payload.soKyHieu || extraction.soKyHieu || '',
    toDateOrBlank_(payload.ngayVanBan || extraction.ngayVanBan),
    payload.coQuanBanHanh || extraction.coQuanBanHanh || '',
    payload.trichYeu || extraction.trichYeu || '',
    payload.loaiVanBan || '',
    payload.doKhan || '',
    payload.doMat || '',
    toDateOrBlank_(payload.hanXuLy),
    payload.donViChuTri || '',
    payload.nguoiXuLy || '',
    payload.trangThai || 'Mới tiếp nhận',
    payload.ghiChu || '',
    uploadedFiles.map(file => file.id).join(', '),
    uploadedFiles.map(file => file.url).join(', '),
    now,
    now,
    userEmail,
    primaryFile ? primaryFile.url : '',
    extraction.ocrText || '',
    extraction.soKyHieu || '',
    toDateOrBlank_(extraction.ngayVanBan),
    extraction.coQuanBanHanh || '',
    extraction.trichYeu || '',
    extraction.confidence || '',
    extraction.status || '',
    primaryFile ? now : '',
    extraction.error || ''
  ];

  sheet.appendRow(row);

  return {
    ok: true,
    id,
    maCongVan,
    fileCount: uploadedFiles.length,
    aiStatus: extraction.status || ''
  };
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('TAMMI AI')
    .addItem('Quét AI các dòng chưa xử lý', 'processPendingScanFiles')
    .addItem('Cài trigger quét AI mỗi 5 phút', 'installAiTrigger')
    .addToUi();
}

function installAiTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(trigger => trigger.getHandlerFunction() === 'processPendingScanFiles')
    .forEach(trigger => ScriptApp.deleteTrigger(trigger));

  ScriptApp.newTrigger('processPendingScanFiles')
    .timeBased()
    .everyMinutes(5)
    .create();
}

function processPendingScanFiles() {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  [SHEET_NAME, ARCHIVE_SHEET_NAME].forEach(sheetName => {
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) return;
    ensureHeaders_(sheet);
    processPendingScanFilesInSheet_(sheet);
  });
}

function getSheet_() {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = spreadsheet.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(SHEET_NAME);
  }

  return sheet;
}

function ensureHeaders_(sheet) {
  const currentHeaders = sheet
    .getRange(1, 1, 1, HEADERS.length)
    .getValues()[0];

  const hasHeaders = currentHeaders.some(value => String(value).trim() !== '');

  if (!hasHeaders) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.setFrozenRows(1);
    return;
  }

  const missingHeaders = HEADERS.filter(header => !currentHeaders.includes(header));
  if (missingHeaders.length > 0) {
    const startColumn = sheet.getLastColumn() + 1;
    sheet.getRange(1, startColumn, 1, missingHeaders.length).setValues([missingHeaders]);
  }
}

function validatePayload_(payload) {
  if (!payload) {
    throw new Error('Chưa có dữ liệu gửi lên.');
  }

  const requiredFields = [
    ['soDen', 'Số đến'],
    ['ngayDen', 'Ngày đến'],
    ['diaBan', 'Địa bàn'],
    ['soKyHieu', 'Số ký hiệu'],
    ['coQuanBanHanh', 'Cơ quan ban hành'],
    ['trichYeu', 'Trích yếu']
  ];

  const missing = requiredFields
    .filter(([key]) => !String(payload[key] || '').trim())
    .map(([, label]) => label);

  if (missing.length > 0) {
    throw new Error('Vui lòng nhập: ' + missing.join(', '));
  }
}

function buildMaCongVan_(soDen, ngayDen) {
  const date = ngayDen ? new Date(ngayDen) : new Date();
  const year = date.getFullYear();
  const paddedSoDen = String(soDen || '').padStart(4, '0');

  return `CVD-${year}-${paddedSoDen}`;
}

function saveFiles_(files, maCongVan) {
  if (!files.length) {
    return [];
  }

  const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);

  return files.map((file, index) => {
    const bytes = Utilities.base64Decode(file.base64);
    const safeName = sanitizeFileName_(file.name || `tep-${index + 1}`);
    const blob = Utilities.newBlob(bytes, file.mimeType || MimeType.PDF, `${maCongVan}-${safeName}`);
    const driveFile = folder.createFile(blob);

    return {
      id: driveFile.getId(),
      url: driveFile.getUrl()
    };
  });
}

function processPendingScanFilesInSheet_(sheet) {
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return;

  const headerMap = buildHeaderMap_(values[0]);
  const fileScanCol = headerMap.FileScan;
  const statusCol = headerMap.AIStatus;
  const errorCol = headerMap.AIError;

  if (!fileScanCol || !statusCol) return;

  values.slice(1).forEach((row, index) => {
    const rowNumber = index + 2;
    const fileRef = row[fileScanCol - 1];
    const status = String(row[statusCol - 1] || '').trim();

    if (!fileRef || ['DONE', 'REVIEW', 'ERROR'].includes(status)) {
      return;
    }

    try {
      sheet.getRange(rowNumber, statusCol).setValue('PROCESSING');
      const file = findDriveFileFromRef_(fileRef);

      if (!file) {
        throw new Error('Không tìm thấy file scan trên Drive: ' + fileRef);
      }

      moveFileToArchiveFolder_(file);
      const extraction = extractCongVanFromFile_(file.getId());
      updateRowWithExtraction_(sheet, rowNumber, headerMap, extraction, file);
    } catch (error) {
      sheet.getRange(rowNumber, statusCol).setValue('ERROR');
      if (errorCol) sheet.getRange(rowNumber, errorCol).setValue(error.message || String(error));
    }
  });
}

function updateRowWithExtraction_(sheet, rowNumber, headerMap, extraction, file) {
  const now = new Date();
  const updates = {
    DriveFileIds: file.getId(),
    DriveFileUrls: file.getUrl(),
    OCRText: extraction.ocrText || '',
    SoKyHieuAI: extraction.soKyHieu || '',
    NgayVanBanAI: toDateOrBlank_(extraction.ngayVanBan),
    CoQuanBanHanhAI: extraction.coQuanBanHanh || '',
    TrichYeuAI: extraction.trichYeu || '',
    AIConfidence: extraction.confidence || '',
    AIStatus: extraction.status || '',
    AILastRun: now,
    AIError: extraction.error || '',
    UpdatedAt: now
  };

  Object.keys(updates).forEach(header => {
    if (headerMap[header]) sheet.getRange(rowNumber, headerMap[header]).setValue(updates[header]);
  });

  [
    ['SoKyHieu', extraction.soKyHieu],
    ['NgayVanBan', toDateOrBlank_(extraction.ngayVanBan)],
    ['CoQuanBanHanh', extraction.coQuanBanHanh],
    ['TrichYeu', extraction.trichYeu]
  ].forEach(([header, value]) => {
    if (!headerMap[header] || !value) return;
    const cell = sheet.getRange(rowNumber, headerMap[header]);
    if (!String(cell.getValue() || '').trim()) cell.setValue(value);
  });
}

function extractCongVanFromFile_(fileId) {
  try {
    const ocrText = runDriveOcr_(fileId);
    const parsed = parseCongVanText_(ocrText);
    const confidence = calculateConfidence_(parsed);

    return {
      ocrText,
      soKyHieu: parsed.soKyHieu,
      ngayVanBan: parsed.ngayVanBan,
      coQuanBanHanh: parsed.coQuanBanHanh,
      trichYeu: parsed.trichYeu,
      confidence,
      status: confidence >= 0.75 ? 'DONE' : 'REVIEW',
      error: ''
    };
  } catch (error) {
    return emptyExtraction_('ERROR', error.message || String(error));
  }
}

function runDriveOcr_(fileId) {
  const sourceFile = DriveApp.getFileById(fileId);
  const resource = {
    title: 'OCR_' + sourceFile.getName(),
    mimeType: MimeType.GOOGLE_DOCS,
    parents: [{ id: getOrCreateOcrTempFolder_().getId() }]
  };

  const ocrDoc = Drive.Files.copy(resource, fileId, {
    ocr: true,
    ocrLanguage: 'vi'
  });

  try {
    const text = DocumentApp.openById(ocrDoc.id).getBody().getText();
    return normalizeText_(text);
  } finally {
    DriveApp.getFileById(ocrDoc.id).setTrashed(true);
  }
}

function parseCongVanText_(text) {
  const cleaned = normalizeText_(text);
  const lines = cleaned.split('\n').map(line => line.trim()).filter(Boolean);
  const firstLines = lines.slice(0, 20);

  return {
    soKyHieu: extractSoKyHieu_(cleaned),
    ngayVanBan: extractNgayVanBan_(cleaned),
    coQuanBanHanh: extractCoQuanBanHanh_(firstLines),
    trichYeu: extractTrichYeu_(cleaned, lines)
  };
}

function extractSoKyHieu_(text) {
  const patterns = [
    /(?:Số|So)\s*[:：]\s*([0-9]{1,6}\s*\/\s*[A-Z0-9ĐƠƯÂÊÔĂƠƯ\-_.]+)/i,
    /(?:Số|So)\s*[:：]\s*([A-Z0-9ĐƠƯÂÊÔĂƠƯ\-_.]+\/[A-Z0-9ĐƠƯÂÊÔĂƠƯ\-_.]+)/i,
    /\b([0-9]{1,6}\s*\/\s*[A-Z0-9ĐƠƯÂÊÔĂƠƯ\-_.]{2,})\b/i
  ];

  for (const pattern of patterns) {
    const match = text.match(pattern);
    if (match && match[1]) return cleanField_(match[1]).replace(/\s*\/\s*/g, '/');
  }

  return '';
}

function extractNgayVanBan_(text) {
  const patterns = [
    /ngày\s+(\d{1,2})\s+tháng\s+(\d{1,2})\s+năm\s+(\d{4})/i,
    /ngay\s+(\d{1,2})\s+thang\s+(\d{1,2})\s+nam\s+(\d{4})/i,
    /\b(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})\b/
  ];

  for (const pattern of patterns) {
    const match = text.match(pattern);
    if (match) return `${match[3]}-${String(match[2]).padStart(2, '0')}-${String(match[1]).padStart(2, '0')}`;
  }

  return '';
}

function extractCoQuanBanHanh_(lines) {
  const ignored = /(cộng hòa|độc lập|tự do|hạnh phúc|số\s*:|ngày\s+\d|kính gửi)/i;
  const candidates = lines
    .slice(0, 8)
    .filter(line => line.length >= 6 && line.length <= 120)
    .filter(line => !ignored.test(line));

  return cleanField_(candidates[0] || '');
}

function extractTrichYeu_(text, lines) {
  const vvMatch = text.match(/(?:V\/v|Về việc|Ve viec)\s*[:：-]?\s*([^\n\r]{8,300})/i);
  if (vvMatch && vvMatch[1]) return cleanField_(vvMatch[1]);

  const trichMatch = text.match(/(?:Trích yếu|Trich yeu)\s*[:：-]\s*([^\n\r]{8,300})/i);
  if (trichMatch && trichMatch[1]) return cleanField_(trichMatch[1]);

  const keyLine = lines.find(line => /(di dời|phối hợp|triển khai|đề nghị|cung cấp|báo cáo|thông báo|xử lý|hỗ trợ)/i.test(line));
  return cleanField_(keyLine || '');
}

function calculateConfidence_(parsed) {
  let score = 0;
  if (parsed.soKyHieu) score += 0.35;
  if (parsed.ngayVanBan) score += 0.2;
  if (parsed.coQuanBanHanh) score += 0.2;
  if (parsed.trichYeu) score += 0.25;
  return Math.round(score * 100) / 100;
}

function emptyExtraction_(status, error) {
  return {
    ocrText: '',
    soKyHieu: '',
    ngayVanBan: '',
    coQuanBanHanh: '',
    trichYeu: '',
    confidence: 0,
    status,
    error: error || ''
  };
}

function findDriveFileFromRef_(fileRef) {
  const ref = String(fileRef || '').trim();
  const idMatch = ref.match(/[-\w]{25,}/);
  if (idMatch) {
    try {
      return DriveApp.getFileById(idMatch[0]);
    } catch (error) {
      // Continue with filename search below.
    }
  }

  const fileName = ref.split('/').pop();
  if (!fileName) return null;

  const escapedName = fileName.replace(/'/g, "\\'");
  const files = DriveApp.searchFiles(`title = '${escapedName}' and trashed = false`);
  return files.hasNext() ? files.next() : null;
}

function moveFileToArchiveFolder_(file) {
  const targetFolder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  targetFolder.addFile(file);

  const parents = file.getParents();
  while (parents.hasNext()) {
    const parent = parents.next();
    if (parent.getId() !== targetFolder.getId()) parent.removeFile(file);
  }
}

function getOrCreateOcrTempFolder_() {
  const parent = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  const existing = parent.getFoldersByName(OCR_TEMP_FOLDER_NAME);
  return existing.hasNext() ? existing.next() : parent.createFolder(OCR_TEMP_FOLDER_NAME);
}

function buildHeaderMap_(headers) {
  return headers.reduce((map, header, index) => {
    const key = String(header || '').trim();
    if (key) map[key] = index + 1;
    return map;
  }, {});
}

function normalizeText_(value) {
  return String(value || '')
    .replace(/\r/g, '\n')
    .replace(/[ \t]+/g, ' ')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

function cleanField_(value) {
  return String(value || '')
    .replace(/[“”"]/g, '')
    .replace(/\s+/g, ' ')
    .replace(/^[\s:：,\-.]+|[\s,;]+$/g, '')
    .trim();
}

function sanitizeFileName_(name) {
  return String(name)
    .replace(/[\\/:*?"<>|]/g, '-')
    .replace(/\s+/g, ' ')
    .trim();
}

function toDateOrBlank_(value) {
  return value ? new Date(value) : '';
}
