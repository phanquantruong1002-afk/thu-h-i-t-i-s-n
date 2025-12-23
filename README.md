# DOCX Form Web (GitHub Pages)

Web tĩnh: nhập dữ liệu -> xuất ra `.docx` mới bằng cách thay placeholder `(1)...(17)` trong **template.docx**.

## 1) Chạy thử local
- Mở VS Code -> cài extension **Live Server**
- Right click `index.html` -> **Open with Live Server**
- Nhập form -> bấm **Tạo file DOCX**

## 2) Up lên GitHub và bật GitHub Pages
1. Tạo repo mới (public)
2. Upload tất cả file trong thư mục này:
   - `index.html`
   - `style.css`
   - `app.js`
   - `template.docx`
3. Vào **Settings → Pages**
   - Source: `Deploy from a branch`
   - Branch: `main` / folder: `/ (root)`
4. Mở link Pages và dùng.

## 3) Android APK (cách dễ nhất)
Nếu bạn vẫn muốn APK: bọc web này vào **WebView** (đỡ phải viết xử lý DOCX native).

Ý tưởng:
- Copy toàn bộ web vào `app/src/main/assets/`
- Tạo `WebViewActivity` load: `file:///android_asset/index.html`
- Bật `setAllowFileAccess(true)` và `setJavaScriptEnabled(true)`

(Lưu ý: tải file DOCX trên Android qua WebView có thể cần thêm xử lý DownloadListener.)

## Placeholder mapping theo mẫu
- Date: `Hà Nội, ngày (1) tháng (2) năm (3)`  fileciteturn1file0L5-L16
- Người nhận + hợp đồng: `(4)(5)(6)(7)`       fileciteturn1file0L39-L51
- Thông tin xe: `(9)-(14)`                    fileciteturn1file1L1-L4
- Liên hệ: `(15)-(17)`                        fileciteturn1file2L11-L18
