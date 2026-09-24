# Quản lý Công văn đến - V1 Form nhập liệu + AI OCR

Bộ source này dùng để tạo form nhập công văn đến trên Google Apps Script.

## Cấu trúc

- `apps-script/Code.gs`: backend ghi dữ liệu vào Google Sheets, lưu tệp scan vào Drive, OCR file và đề xuất thông tin công văn.
- `apps-script/index.html`: giao diện form nhập công văn đến.
- `apps-script/appsscript.json`: cấu hình Apps Script.
- `schema/google-sheets-schema.csv`: danh sách cột cần tạo trong Google Sheets.

## Cách triển khai nhanh

1. Tạo một Google Sheet mới.
2. Tạo sheet/tab tên `CongVanDen`.
3. Copy dòng header trong `schema/google-sheets-schema.csv` vào dòng 1 của sheet `CongVanDen`.
4. Tạo một folder Google Drive để lưu PDF/scan/phụ lục.
5. Mở Apps Script gắn với Google Sheet.
6. Copy nội dung:
   - `apps-script/Code.gs` vào file `Code.gs`
   - `apps-script/index.html` vào file HTML tên `index`
   - `apps-script/appsscript.json` vào Project Settings / manifest nếu cần
7. Trong `Code.gs`, thay:
   - `SPREADSHEET_ID`
   - `DRIVE_FOLDER_ID`
8. Trong Apps Script, bật **Services → Drive API** để dùng OCR qua Google Drive.
9. Deploy thành Web app.
10. Mở Google Sheet, menu **TAMMI AI → Cài trigger quét AI mỗi 5 phút** để tự xử lý file scan mới.

## Dữ liệu đầu vào

Form đã có sẵn các trường chính:

- Số đến
- Ngày đến
- Số ký hiệu
- Ngày văn bản
- Cơ quan ban hành
- Trích yếu
- Loại văn bản
- Độ khẩn
- Độ mật
- Hạn xử lý
- Đơn vị chủ trì
- Người xử lý
- Trạng thái
- Ghi chú
- Tệp PDF/scan/phụ lục

## Luồng AI OCR

Khi có file scan/PDF/ảnh:

1. File được lưu vào folder Drive cấu hình trong `DRIVE_FOLDER_ID`.
2. Apps Script dùng Google Drive OCR để đọc text tiếng Việt.
3. Script bóc các trường:
   - `SoKyHieuAI`
   - `NgayVanBanAI`
   - `CoQuanBanHanhAI`
   - `TrichYeuAI`
4. Nếu cột chính đang trống, script tự điền sang:
   - `SoKyHieu`
   - `NgayVanBan`
   - `CoQuanBanHanh`
   - `TrichYeu`
5. `AIStatus` sẽ là:
   - `DONE`: đủ tự tin
   - `REVIEW`: cần người dùng kiểm tra
   - `ERROR`: lỗi đọc file hoặc không tìm thấy file
