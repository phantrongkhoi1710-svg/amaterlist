# ARMATURE IMPORT TOOL - Web Edition

Công cụ hỗ trợ nhập, gom nhóm, ánh xạ bí danh cột và hợp nhất các file Excel Armature vào bảng Master (Total) với xuất file chuẩn định dạng Excel (.xlsx).

---

## 🚀 Hướng Dẫn Deploy Lên GitHub Pages

Dự án đã được cấu hình sẵn để triển khai lên **GitHub Pages** hoàn toàn miễn phí.

### Cách 1: Triển khai Tự động bằng GitHub Actions (Khuyên dùng)

1. **Đưa mã nguồn lên GitHub Repo của bạn:**
   ```bash
   git init
   git add .
   git commit -m "Initial commit"
   git branch -M main
   git remote add origin https://github.com/<tai-khoan-cua-ban>/<ten-repo>.git
   git push -u origin main
   ```

2. **Bật GitHub Pages trong Repo:**
   - Vào repository của bạn trên GitHub.
   - Nhấn tab **Settings** &rarr; Chọn mục **Pages** ở cột bên trái.
   - Tại mục **Build and deployment** &rarr; **Source**: Chọn **`GitHub Actions`**.
   - Ngay sau đó, workflow `.github/workflows/deploy.yml` sẽ tự động chạy và xuất bản website của bạn tại:
     `https://<tai-khoan-cua-ban>.github.io/<ten-repo>/`

---

### Cách 2: Triển khai nhanh bằng lệnh `npm run deploy` (gh-pages)

Nếu bạn muốn deploy trực tiếp từ máy cá nhân lên branch `gh-pages`:

1. Đảm bảo repo đã kết nối với GitHub.
2. Chạy lệnh:
   ```bash
   npm run deploy
   ```
3. Vào **Settings** &rarr; **Pages** &rarr; **Source**: Chọn **Deploy from a branch** &rarr; Branch: `gh-pages` &rarr; Nhấn **Save**.

---

## 🛠️ Chạy ứng dụng ở môi trường Local

```bash
# Cài đặt dependencies
npm install

# Khởi động server phát triển
npm run dev

# Kiểm tra build sản phẩm
npm run build
```
