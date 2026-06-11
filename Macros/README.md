# AUTO DU LIEU COREL V9 - PFI Search Macro

Giao diện VBA cho CorelDRAW để tìm kiếm dữ liệu PFI và hiển thị thông tin sản phẩm.

## Cài đặt

1. Mở CorelDRAW
2. Tools → Macros → Edit Macros
3. File → Open → Chọn file `frmPFISearch.frm`
4. Ctrl + S
5. Đóng VBA Editor

## Sử dụng

1. Tools → Macros → Run Macro
2. Chọn: frmPFISearch
3. Nhập PFI và click "TIM KIEM"
4. Chọn định dạng ngày
5. Copy dữ liệu hoặc insert vào CorelDRAW

## Tính năng

- ✅ Tìm kiếm PFI từ API
- ✅ Hiển thị đầy đủ dữ liệu sản phẩm
- ✅ Linh hoạt chọn định dạng ngày
- ✅ Copy từng ô dữ liệu
- ✅ Insert dữ liệu vào CorelDRAW

## API Endpoints

API chạy tại: `http://localhost:5000`

- GET `/api/data/search/{pfi}` - Tìm PFI
- GET `/api/data/all` - Lấy tất cả dữ liệu
- POST `/api/data/format-date` - Format ngày
- POST `/api/settings/credentials` - Lưu credentials
- GET `/api/settings/status` - Kiểm tra status
