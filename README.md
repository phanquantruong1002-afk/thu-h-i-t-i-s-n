# Điền khung → Xuất PDF (không hiện Xem trước)

## Đúng yêu cầu của bạn
- Chỉ hiện form (khung điền)
- Dùng 1 file Word mẫu: `template.docx` (chính file bạn gửi)
- Bấm Tạo -> tải PDF

## Vì sao vẫn phải render HTML?
GitHub Pages là web tĩnh, không có server để convert DOCX->PDF chuẩn Word.
Bản này vẫn dùng pipeline DOCX -> HTML -> PDF, nhưng **không hiển thị** phần HTML ra màn hình.
Đã thêm các bước chống PDF trắng:
- render vào vùng ẩn nhưng nằm trong viewport (opacity:0)
- đợi 2 frame + chờ ảnh load xong
- backgroundColor trắng + ép chữ đen

## Up GitHub Pages
Upload vào root:
- index.html
- style.css
- app.js
- template.docx

---
ver 0.9: Ẩn trang trắng/khối xám khi tạo PDF (overlay đen 100%, exportHost visibility:hidden và chỉ hiện khi in).

---
ver 0.9: Ẩn trang trắng/khối xám khi tạo PDF (overlay đen 100%, exportHost visibility:hidden và chỉ hiện khi in).
