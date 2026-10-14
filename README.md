# Hậu Store - Công cụ quản lý & bán hàng

**Kho hàng:** https://phukiendanang01-png.github.io/kho-hang/
**Tạo hoá đơn (offline):** https://phukiendanang01-png.github.io/kho-hang/hoa-don.html

Cả hai app chạy trên Android / iOS / macOS, dùng được offline sau lần mở đầu.

## App kho hàng (`index.html`)
Nhập/xuất hàng (tự trừ kho), tồn kho + tìm kiếm IMEI,
báo cáo lợi nhuận theo kỳ (7 ngày / tháng này / tháng trước),
xuất & đọc file Excel, đồng bộ Google Sheet.

## App tạo hoá đơn (`hoa-don.html`)
Nhập tên khách, SĐT, mẫu máy, IMEI, bảo hành, giá, ghi chú ->
hoá đơn HẬU STORES tự điền, in / lưu PDF (Ctrl+P) hoặc xuất file Word (.doc),
lưu lịch sử hoá đơn trên máy.

- `google-apps-script.gs` — dán vào Google Sheet (Extensions → Apps Script) để bật đồng bộ online cho app kho
- `HUONG-DAN.txt` — hướng dẫn cài đặt & dùng app kho
