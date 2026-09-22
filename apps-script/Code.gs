const SPREADSHEET_ID = 'PASTE_GOOGLE_SHEET_ID_HERE';
const SHEET_NAME = 'CongVanDen';
const DRIVE_FOLDER_ID = 'PASTE_DRIVE_FOLDER_ID_HERE';

const HEADERS = [
  'ID',
  'MaCongVan',
  'SoDen',
  'NgayDen',
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
  'CreatedBy'
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
  const userEmail = Session.getActiveUser().getEmail() || '';

  const row = [
    id,
    maCongVan,
    payload.soDen || '',
    toDateOrBlank_(payload.ngayDen),
    payload.soKyHieu || '',
    toDateOrBlank_(payload.ngayVanBan),
    payload.coQuanBanHanh || '',
    payload.trichYeu || '',
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
    userEmail
  ];

  sheet.appendRow(row);

  return {
    ok: true,
    id,
    maCongVan,
    fileCount: uploadedFiles.length
  };
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
  }
}

function validatePayload_(payload) {
  if (!payload) {
    throw new Error('Chưa có dữ liệu gửi lên.');
  }

  const requiredFields = [
    ['soDen', 'Số đến'],
    ['ngayDen', 'Ngày đến'],
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

function sanitizeFileName_(name) {
  return String(name)
    .replace(/[\\/:*?"<>|]/g, '-')
    .replace(/\s+/g, ' ')
    .trim();
}

function toDateOrBlank_(value) {
  return value ? new Date(value) : '';
}
