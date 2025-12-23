# DOCX Form Web (GitHub Pages) — Xuất PDF

Web tĩnh: nhập dữ liệu → thay placeholder `(1)...(17)` trong **template.docx** → **xuất PDF**.

## Chạy thử local
- Mở VS Code → cài extension **Live Server**
- Right click `index.html` → **Open with Live Server**
- Nhập form → bấm **Tạo file PDF**

## Up lên GitHub và bật GitHub Pages
1. Tạo repo mới (public)
2. Upload tất cả file:
   - `index.html`
   - `style.css`
   - `app.js`
   - `template.docx`
3. Vào **Settings → Pages**
   - Source: `Deploy from a branch`
   - Branch: `main` / folder: `/ (root)`

## Cơ chế xuất PDF (quan trọng)
GitHub Pages là web tĩnh (không có server để chạy LibreOffice).
Vì vậy PDF được tạo theo pipeline:
1) Tạo DOCX blob bằng cách replace placeholder trong XML (PizZip)
2) Render DOCX → HTML trong trình duyệt (mammoth)
3) HTML → PDF (html2pdf.js)

Độ giống Word có thể lệch nhẹ tuỳ font/trình duyệt.

## Android APK (dễ nhất)
Nếu bạn muốn APK: bọc web này vào **WebView**.
- Copy toàn bộ web vào `app/src/main/assets/`
- WebView load: `file:///android_asset/index.html`
- Bật `setJavaScriptEnabled(true)` và `setAllowFileAccess(true)`
