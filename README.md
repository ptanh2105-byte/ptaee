# Tax Decision Extractor - AutoHotkey Tool

Công cụ AutoHotkey để trích xuất dữ liệu từ các file Word chứa Quyết định xử phạt vi phạm hành chính của Cục Thuế Thái Nguyên.

## Tính năng

✅ Trích xuất dữ liệu từ multiple Word files cùng lúc
✅ Tự động phát hiện loại (Cá nhân / Tổ chức)
✅ Hỗ trợ các loại phạt: Phạt tiền, Phạt cảnh cáo
✅ Hiển thị kết quả trên GUI dễ dàng copy
✅ Export kết quả ra Excel hoặc CSV
✅ Xử lý đa tệp (Batch processing)

## Các trường được trích xuất

### Thông tin chung
- Số biên bản
- Ngày lập biên bản
- Hành vi vi phạm
- Quy định tại
- Tình tiết tăng nặng
- Tình tiết giảm nhẹ
- Hình thức xử phạt
- Mức phạt / Loại phạt

### Thông tin cá nhân (nếu là cá nhân)
- Tên cá nhân
- Giới tính
- Ngày, tháng, năm sinh
- Quốc tịch
- Nơi làm việc
- Nơi ở hiện tại
- Số CCCD
- Mã số thuế

### Thông tin tổ chức (nếu là tổ chức)
- Tên tổ chức
- Địa chỉ trụ sở chính
- Mã số thuế
- Số GCN đăng ký kinh doanh
- Người đại diện theo pháp luật
- Chức danh

### Thông tin thanh toán
- Tài khoản thu NSNN
- Đơn vị thu tiền phạt
- Phòng thực hiện
- Người ký quyết định

## Yêu cầu hệ thống

- Windows 7 trở lên
- AutoHotkey v1.1+ (https://www.autohotkey.com/)
- Microsoft Word hoặc LibreOffice (để mở file .docx)

## Cách sử dụng

### 1. Cài đặt AutoHotkey
- Tải từ: https://www.autohotkey.com/
- Cài đặt bình thường

### 2. Chạy script
- Tải file `TaxDecisionExtractor.ahk`
- Double-click để chạy
- Hoặc nhấp phải > Run with AutoHotkey

### 3. Sử dụng giao diện
1. Nhấn nút "Chọn Folder" để chọn thư mục chứa file Word
2. Chọn định dạng output (Excel hoặc CSV)
3. Nhấn "Bắt Đầu Trích Xuất"
4. Xem kết quả trong bảng
5. Copy hoặc Export kết quả

## Cấu trúc File

```
ptaee/
├── TaxDecisionExtractor.ahk       # Script chính
├── README.md                       # Hướng dẫn này
├── CHANGELOG.md                    # Lịch sử thay đổi
└── templates/                      # Thư mục chứa file mẫu
    ├── mau_ca_nhan.txt            # Mô tả mẫu quyết định cá nhân
    ├── mau_to_chuc.txt            # Mô tả mẫu quyết định tổ chức
    └── mau_phat_canh_cao.txt      # Mô tả mẫu quyết định phạt cảnh cáo
```

## Hướng dẫn chi tiết

### Bước 1: Chuẩn bị
1. Cài đặt AutoHotkey từ https://www.autohotkey.com/
2. Tải file `TaxDecisionExtractor.ahk` về máy tính
3. Chuẩn bị thư mục chứa các file Word cần trích xuất

### Bước 2: Chạy tool
1. Chạy file `TaxDecisionExtractor.ahk`
2. Giao diện sẽ hiển thị

### Bước 3: Chọn thư mục
1. Nhấn nút "Chọn Folder"
2. Chọn thư mục chứa file Word
3. Nhấn OK

### Bước 4: Bắt đầu trích xuất
1. Nhấn "Bắt Đầu Trích Xuất"
2. Tool sẽ tự động xử lý từng file
3. Kết quả sẽ hiển thị trong bảng

### Bước 5: Sử dụng kết quả
- **Copy All**: Copy tất cả dữ liệu vào clipboard
- **Export Excel**: Xuất ra file Excel (.xlsx)
- **Export CSV**: Xuất ra file CSV (.csv)
- **Clear**: Xóa tất cả kết quả

## Các tính năng chi tiết

### Tự động phát hiện loại tài liệu
Tool tự động phát hiện:
- Quyết định đối với **cá nhân** - trích xuất: Tên, Ngày sinh, CCCD, ...
- Quyết định đối với **tổ chức** - trích xuất: Tên công ty, Địa chỉ, MST, ...

### Xử lý nhiều file cùng lúc
- Chọn thư mục → Tool sẽ xử lý tất cả file .docx trong thư mục
- Thanh progress hiển thị tiến độ

### Export linh hoạt
- **Excel**: Format đẹp, dễ sắp xếp và lọc
- **CSV**: Dễ import vào các hệ thống khác
- **Clipboard**: Copy trực tiếp để dán vào Word, Excel

## Lưu ý quan trọng

⚠️ **File Word phải có định dạng chuẩn** theo mẫu Quyết định của Cục Thuế
⚠️ **Dữ liệu được trích xuất dựa trên vị trí** cụ thể trong tài liệu
⚠️ **Nếu file có định dạng khác**, kết quả trích xuất có thể không chính xác
⚠️ **Đóng file Word** trước khi chạy tool (tránh lỗi khóa file)
⚠️ **Cần cài Microsoft Word** hoặc LibreOffice để tool hoạt động

## Xử lý sự cố

### Tool không chạy được
- Kiểm tra AutoHotkey đã cài đặt chưa
- Thử chạy với quyền Administrator
- Kiểm tra Windows Defender có chặn không

### Kết quả trích xuất sai
- Kiểm tra file Word có định dạng đúng không
- Thử với file mẫu trước
- Kiểm tra các label ("Tên cá nhân:", "Tên tổ chức:", ...) có giống không

### Không thể export Excel
- Kiểm tra Excel đã cài đặt chưa
- Đóng file Excel đang mở
- Thử export ra CSV thay thế

## Hỗ trợ

Nếu gặp vấn đề:
1. Kiểm tra file Word có định dạng đúng không
2. Đảm bảo AutoHotkey được cài đặt đúng
3. Thử với file mẫu trước
4. Xem lại các lưu ý quan trọng ở trên

## License

Miễn phí sử dụng

---

**Phiên bản:** 1.0
**Cập nhật lần cuối:** 2026-07-10
**Tác giả:** ptanh2105-byte