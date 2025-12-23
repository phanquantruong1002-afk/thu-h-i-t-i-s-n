// app.js
// Mục tiêu: bấm "Tạo file PDF" -> xuất PDF.
// Cách làm (chạy tĩnh trên GitHub Pages):
// 1) Dùng template.docx, thay placeholder (1)...(17) => tạo DOCX blob (PizZip)
// 2) Render DOCX blob ra HTML bằng docx-preview
// 3) Dùng html2pdf.js để xuất PDF từ HTML vừa render
//
// Lưu ý quan trọng:
// - DOCX placeholder phải là chuỗi LIỀN (vd "(12)") trong file XML.
//   Nếu Word tách ra thành nhiều đoạn (run) thì replace theo chuỗi sẽ không ăn.

let templateArrayBuffer = null; // cache template hiện tại
const $ = (id) => document.getElementById(id);
const msg = (s) => { $("msg").textContent = s || ""; };

function xmlEscape(str) {
  // Tránh làm hỏng XML khi thay thế text
  return String(str)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;");
}

async function loadDefaultTemplate() {
  const res = await fetch("./template.docx");
  if (!res.ok) throw new Error("Không tải được template.docx (check GitHub Pages / path)");
  templateArrayBuffer = await res.arrayBuffer();
  $("templateStatus").textContent = "Đang dùng: template.docx";
}

function loadTemplateFromFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = () => reject(new Error("Đọc file template thất bại"));
    reader.readAsArrayBuffer(file);
  });
}

function getFormData() {
  const fd = new FormData($("form"));

  const data = {
    "(1)": fd.get("d1"),
    "(2)": fd.get("d2"),
    "(3)": fd.get("d3"),
    "(4)": fd.get("p4"),
    "(5)": fd.get("p5"),
    "(6)": fd.get("p6"),
    "(7)": fd.get("p7"),
    "(9)": fd.get("p9"),
    "(10)": fd.get("p10"),
    "(11)": fd.get("p11"),
    "(12)": fd.get("p12"),
    "(13)": fd.get("p13"),
    "(14)": fd.get("p14"),
    "(15)": fd.get("p15"),
    "(16)": fd.get("p16"),
    "(17)": fd.get("p17"),
  };

  let outName = (fd.get("outName") || "Thong-bao-thu-hoi-tai-san.pdf").trim();
  if (!outName.toLowerCase().endsWith(".pdf")) outName += ".pdf";
  return { data, outName };
}

function replaceInZip(zip, replacements) {
  // Thay trong các XML thuộc "word/" (document, header, footer, ...)
  const targets = Object.keys(zip.files).filter((name) =>
    name.startsWith("word/") &&
    name.endsWith(".xml") &&
    !name.startsWith("word/theme/")
  );

  for (const name of targets) {
    const file = zip.file(name);
    if (!file) continue;

    let xml;
    try {
      xml = file.asText();
    } catch (_) {
      continue;
    }

    let changed = false;
    for (const [ph, value] of Object.entries(replacements)) {
      const safe = xmlEscape(value ?? "");
      if (xml.includes(ph)) {
        xml = xml.split(ph).join(safe);
        changed = true;
      }
    }
    if (changed) zip.file(name, xml);
  }
}

function generateDocxBlob(arrayBuffer, replacements) {
  const zip = new PizZip(arrayBuffer);
  replaceInZip(zip, replacements);

  return zip.generate({
    type: "blob",
    mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  });
}

async function renderDocxToHtml(docxBlob, container) {
  container.innerHTML = ""; // clear
  const ab = await docxBlob.arrayBuffer();

  // Mammoth API: mammoth.convertToHtml({arrayBuffer})
  if (!window.mammoth || !window.mammoth.convertToHtml) {
    throw new Error("Không load được thư viện mammoth (DOCX->HTML)");
  }

  const result = await window.mammoth.convertToHtml({ arrayBuffer: ab }, {
    // styleMap có thể tuỳ biến thêm nếu muốn giống Word hơn
  });

  // Mammoth trả về HTML tương đối "sạch"
  container.innerHTML = result.value || "";
}

async function exportHtmlToPdf(container, filename) {
  // html2pdf options
  const opt = {
    margin:       [10, 10, 10, 10], // mm
    filename:     filename,
    image:        { type: "jpeg", quality: 0.98 },
    html2canvas:  { scale: 2, useCORS: true },
    jsPDF:        { unit: "mm", format: "a4", orientation: "portrait" }
  };

  // html2pdf() trả Promise
  await html2pdf().set(opt).from(container).save();
}

function setBusy(isBusy) {
  $("btnGenerate").disabled = isBusy;
  $("btnFillDemo").disabled = isBusy;
  $("btnUseDefault").disabled = isBusy;
  $("templateFile").disabled = isBusy;
}

(async function init() {
  try {
    await loadDefaultTemplate();
  } catch (e) {
    $("templateStatus").textContent = "Chưa tải được template.docx (bạn vẫn có thể chọn file mẫu ở trên)";
  }

  $("btnUseDefault").addEventListener("click", async () => {
    try {
      await loadDefaultTemplate();
      msg("OK: dùng template.docx");
    } catch (e) {
      msg(e.message);
    }
  });

  $("templateFile").addEventListener("change", async (ev) => {
    const f = ev.target.files?.[0];
    if (!f) return;
    try {
      templateArrayBuffer = await loadTemplateFromFile(f);
      $("templateStatus").textContent = `Đang dùng: ${f.name}`;
      msg("OK: đã nạp template từ máy bạn");
    } catch (e) {
      msg(e.message);
    }
  });

  $("btnFillDemo").addEventListener("click", () => {
    const form = $("form");
    const set = (name, v) => (form.querySelector(`[name="${name}"]`).value = v);

    // Demo (bạn có thể sửa cho đúng dữ liệu thật của bạn)
    set("d1", "05"); set("d2", "11"); set("d3", "2025");
    set("p4", "Nguyễn Văn Thanh");
    set("p5", "Kv Lân Thạnh 1, Phường Trung Kiên, Quận Thốt Nốt, Thành Phố Cần Thơ");
    set("p6", "20285167926");
    set("p7", "24/02/2022");
    set("p9", "HONDA VariO - Click (Nhập khẩu) 150cc Fi 2020");
    set("p10", "4125LK052645");
    set("p11", "KF41E2056854");
    set("p12", "65F1-645.46");
    set("p13", "058656");
    set("p14", "04/12/2020");
    set("p15", "Lê Phước Hữu");
    set("p16", "Trưởng nhóm Khai thác tài sản");
    set("p17", "0913788134");
    set("outName", "Thong-bao-thu-hoi-tai-san.pdf");

    msg("Đã điền demo");
  });

  $("form").addEventListener("submit", async (ev) => {
    ev.preventDefault();
    msg("");

    if (!templateArrayBuffer) {
      msg("Chưa có template. Hãy chọn file template .docx hoặc dùng template.docx");
      return;
    }

    const { data, outName } = getFormData();
    const renderArea = $("renderArea");

    try {
      setBusy(true);
      msg("Đang tạo DOCX…");
      const docxBlob = generateDocxBlob(templateArrayBuffer, data);

      msg("Đang render DOCX…");
      await renderDocxToHtml(docxBlob, renderArea);

      msg("Đang xuất PDF…");
      await exportHtmlToPdf(renderArea, outName);

      msg("Xong: đã tải PDF");
    } catch (e) {
      console.error(e);
      msg("Lỗi xuất PDF: " + (e?.message || e));
    } finally {
      setBusy(false);
      // tránh giữ DOM nặng
      renderArea.innerHTML = "";
    }
  });
})();
