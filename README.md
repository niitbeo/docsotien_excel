# 📊 DocTien - Excel Add-in Converter

Chuyển đổi số thành chữ tiếng Việt trong Microsoft Excel.

## 🎯 Tính năng

- ✅ Hàm **DocTien()** có thể sử dụng như công thức Excel bình thường
- ✅ Hỗ trợ số từ 0 đến 999,999,999,999,999 (999 nghìn tỷ)
- ✅ Hỗ trợ số thập phân (2 chữ số)
- ✅ Tự động viết hoa chữ cái đầu
- ✅ Cài đặt dễ dàng bằng file .exe

## 📦 Cài đặt

### Cách 1: Sử dụng file .exe (Khuyến nghị)

1. Chạy file **DocTien_Installer.exe**
2. Nhấn nút "Cài đặt"
3. Làm theo hướng dẫn trên màn hình

### Cách 2: Cài đặt thủ công

1. Mở Microsoft Excel
2. Nhấn **Alt + F11** để mở VBA Editor
3. Chọn **File** → **Import File**
4. Chọn file **DocTien.bas**
5. Đóng VBA Editor
6. Hoàn tất!

## 🚀 Sử dụng

Sau khi cài đặt, bạn có thể sử dụng hàm `DocTien()` trong Excel:

```excel
=DocTien(A1)
=DocTien(12345)
=DocTien(B2)
```

### Ví dụ:

| Số | Công thức | Kết quả |
|---|---|---|
| 12345 | =DocTien(A1) | Mười hai nghìn ba trăm bốn mươi lăm đồng |
| 1000000 | =DocTien(A2) | Một triệu đồng |
| 2500.50 | =DocTien(A3) | Hai nghìn năm trăm đồng phẩy năm mươi đồng |
| 0 | =DocTien(A4) | Không đồng |

## 🛠️ Build file .exe (Dành cho Developer)

Nếu bạn muốn tự build file .exe từ source code:

```bash
# Cài đặt PyInstaller
pip install pyinstaller

# Build file .exe
pyinstaller installer.spec
```

File .exe được tạo trong thư mục `dist/`

## 📝 Cấu trúc Project

```
doctien/
├── DocTien.bas          # VBA module chứa hàm DocTien()
├── installer.py         # GUI installer
├── installer.spec       # PyInstaller config
├── README.md            # File này
└── dist/
    └── DocTien_Installer.exe  # File cài đặt
```

## ❓ Hỗ trợ

Nếu gặp vấn đề:

1. Đảm bảo bạn đã cài đặt Microsoft Excel
2. Kiểm tra macro đã được enable trong Excel (File → Options → Trust Center → Trust Center Settings → Macro Settings)
3. Thử import file DocTien.bas thủ công (xem Cách 2)

## 📄 License

MIT License - Sử dụng tự do cho mọi mục đích.

---

Made with ❤️ for Vietnamese Excel users
