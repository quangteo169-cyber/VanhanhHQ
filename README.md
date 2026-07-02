# Trung tâm Báo cáo PVH — HQ Group

Dashboard vận hành nội bộ (PVH1–PVH12 + Tổng quan), lọc theo Ngày / Tháng / Năm.
Trang tĩnh (chỉ `index.html`), deploy thẳng lên Vercel.

## Cấu trúc
```
index.html     # toàn bộ dashboard (HTML + CSS + JS, dùng Chart.js qua CDN)
vercel.json    # cấu hình static + tắt cache để cập nhật nhanh
.gitignore
```

## Chạy thử ở máy
Mở trực tiếp `index.html` bằng trình duyệt là chạy (bản demo dùng số nhúng).

## Deploy lên Vercel qua GitHub (khuyên dùng — tự động deploy khi push)
1. Tạo repo mới trên GitHub, ví dụ `pvh-dashboard`.
2. Ở thư mục này chạy:
   ```bash
   git init
   git add .
   git commit -m "PVH dashboard - giao dien v1"
   git branch -M main
   git remote add origin https://github.com/<tài-khoản>/pvh-dashboard.git
   git push -u origin main
   ```
3. Vào https://vercel.com → **Add New… ▸ Project** → **Import** repo `pvh-dashboard`.
4. Framework Preset để **Other** (static), Root Directory `./`, bấm **Deploy**.
5. Xong — mỗi lần `git push` Vercel tự deploy lại.

## Deploy nhanh bằng Vercel CLI (không cần GitHub)
```bash
npm i -g vercel
vercel        # deploy preview
vercel --prod # deploy production
```

## Bước sau — nối dữ liệu thật realtime (L2_TONGHOP)
Bản hiện tại là **giao diện + số demo** (tháng thật · ngày mẫu). Để chạy realtime:
1. Dựng sheet **L2_TONGHOP** trong Google Sheet (cấp Ngày, long-format:
   `Ngày · Năm · Tháng · Tuần_ISO · SPDV · Nghiệp_vụ · Số_lượng · Giá_trị`).
2. Publish sheet đó ra CSV.
3. Dán link vào biến `CSV_L2` (khối cấu hình đầu thẻ `<script>` trong `index.html`) —
   khi có link, web tự đọc live thay cho số nhúng.

> Bảo mật: chỉ publish sheet L2 tổng hợp (không chứa tài khoản/mã code). Không publish các tab raw.
