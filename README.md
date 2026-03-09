# DocTien Excel Add-in

Chuyển số thành chữ tiếng Việt trong Microsoft Excel bằng hàm `DocTien()`.

## Tính năng

- Dùng trực tiếp trong Excel với công thức `=DocTien(A1)`
- Hỗ trợ số lớn và số thập phân
- Có bộ cài `.exe` cho người dùng không muốn thao tác VBA thủ công
- Có file `DocTien.bas` để import thủ công khi cần

## Tải nhanh

File cài mới nhất:

- [`DocTien_AutoInstaller_v2.exe`](dist/DocTien_AutoInstaller_v2.exe)

## Cài đặt

### Cách 1: Dùng file `.exe`

1. Chạy `DocTien_AutoInstaller_v2.exe`
2. Nhấn `Cài đặt tự động`
3. Nếu Excel cảnh báo bảo mật, chọn `Enable Content`

### Cách 2: Import thủ công

1. Mở Microsoft Excel
2. Nhấn `Alt + F11` để mở VBA Editor
3. Chọn `File -> Import File`
4. Chọn file `DocTien.bas`
5. Đóng VBA Editor

## Sử dụng

```excel
=DocTien(A1)
=DocTien(12345)
=DocTien(15000000000)
```

## Ghi chú

- Nếu cài tự động lỗi do quyền VBA/COM của Excel, hãy import `DocTien.bas` thủ công.
- Nếu Excel chưa nhận hàm ngay, hãy đóng và mở lại Excel.

## Cấu trúc chính

- `DocTien.bas`: module VBA đọc số thành chữ
- `installer_auto.py`: bộ cài tự động
- `installer.py`: bộ cài/fallback thủ công
- `dist/DocTien_AutoInstaller_v2.exe`: file cài mới nhất

## Tác giả

- Nguyễn Lê Trường
- Email: `niitbeo28@gmail.com`

## License

MIT
