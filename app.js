// app.js
const APP_VERSION = "0.6";

// Không hiện "Xem trước" trên UI.
// Khi bấm Tạo PDF:
// 1) Tải template.docx (file Word bạn gửi)
// 2) Thay placeholder (1)…(17)
// 3) DOCX -> HTML (mammoth) và render vào exportHost (ẩn)
// 4) html2pdf chụp exportHost -> PDF (không trắng)

let templateArrayBuffer = null;

const $ = (id) => document.getElementById(id);
const msg = (s) => ($("msg").textContent = s || "");


function setOverlayText(t){
  const el = document.getElementById("overlayText");
  if (el) el.textContent = t;
}

let _cancelled = false;
function markCancelled(){ _cancelled = true; }

function withTimeout(promise, ms, label){
  return Promise.race([
    promise,
    new Promise((_, reject)=> setTimeout(()=> reject(new Error(label + " timeout (" + Math.round(ms/1000) + "s)")), ms))
  ]);
}

function showOverlay(show) {
  const ov = document.getElementById("loadingOverlay");
  if (!ov) return;
  if (show) ov.hidden = false;
  else ov.hidden = true;
}

function setBusy(b) {
  $("btnGenerate").disabled = b;
  $("btnFillDemo").disabled = b;
  $("btnUseDefault").disabled = b;
  $("templateFile").disabled = b;
}

function xmlEscape(str) {
  return String(str)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;");
}

async function loadDefaultTemplate() {
  const res = await fetch("./template.docx", { cache: "no-store" });
  if (!res.ok) throw new Error("Không tải được template.docx. Kiểm tra file nằm ở root repo.");
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

function getFormMapAndName() {
  const fd = new FormData($("form"));
  const map = {
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
  let outName = (fd.get("outName") || "output.pdf").trim();
  if (!outName.toLowerCase().endsWith(".pdf")) outName += ".pdf";
  return { map, outName };
}

function replaceInZip(zip, replacements) {
  const targets = Object.keys(zip.files).filter((name) =>
    name.startsWith("word/") &&
    name.endsWith(".xml") &&
    !name.startsWith("word/theme/")
  );

  for (const name of targets) {
    const file = zip.file(name);
    if (!file) continue;

    let xml;
    try { xml = file.asText(); } catch { continue; }

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

async function buildFilledDocxArrayBuffer(templateAb, replacements) {
  const zip = new PizZip(templateAb);
  replaceInZip(zip, replacements);
  const blob = zip.generate({
    type: "blob",
    mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  });
  return await blob.arrayBuffer();
}

async function docxToHtml(arrayBuffer) {
  setOverlayText("Đang chuyển DOCX → HTML…");

  if (!window.mammoth?.convertToHtml) throw new Error("Không load được mammoth (DOCX->HTML).");
  const result = await window.mammoth.convertToHtml({ arrayBuffer }, {});
  return result.value || "<p>(Trống)</p>";
}

async function renderHidden(html) {
  setOverlayText("Đang render nội dung…");

  const host = $("exportHost");
  host.innerHTML = html;

  // đợi render xong (2 frame) để tránh html2canvas chụp trắng
  await new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r)));

  // Nếu có ảnh, chờ ảnh load xong
  const imgs = Array.from(host.querySelectorAll("img"));
  await Promise.all(imgs.map(img => {
    if (img.complete) return Promise.resolve();
    return new Promise(res => { img.onload = res; img.onerror = res; });
  }));
}

async function exportPdf(filename) {
  const host = $("exportHost");

  if (!window.html2canvas) throw new Error("Không load được html2canvas");
  const jsPDF = window.jspdf?.jsPDF;
  if (!jsPDF) throw new Error("Không load được jsPDF");

  if (_cancelled) throw new Error("Đã hủy");

  const task = (async () => {
    setOverlayText("Đang chụp trang…");
    const canvas = await window.html2canvas(host, {
      backgroundColor: "#ffffff",
      scale: 0.8,          // giảm RAM để tránh treo
      useCORS: true,
      logging: false,
      scrollX: 0,
      scrollY: 0,
      windowWidth: 794,
    });

    if (_cancelled) throw new Error("Đã hủy");

    setOverlayText("Đang ghép PDF…");
    const imgData = canvas.toDataURL("image/jpeg", 0.9);

    const pdf = new jsPDF({ unit: "mm", format: "a4", orientation: "portrait" });
    const pageW = pdf.internal.pageSize.getWidth();
    const pageH = pdf.internal.pageSize.getHeight();

    const pxPerMm = canvas.width / pageW;
    const pageHPx = Math.floor(pageH * pxPerMm);

    let y = 0;
    let page = 0;

    while (y < canvas.height) {
      const sliceH = Math.min(pageHPx, canvas.height - y);

      const slice = document.createElement("canvas");
      slice.width = canvas.width;
      slice.height = sliceH;
      const ctx = slice.getContext("2d");
      ctx.drawImage(canvas, 0, y, canvas.width, sliceH, 0, 0, canvas.width, sliceH);

      const sliceImg = slice.toDataURL("image/jpeg", 0.9);
      const sliceHMm = (sliceH * pageW) / canvas.width;

      if (page > 0) pdf.addPage();
      pdf.addImage(sliceImg, "JPEG", 0, 0, pageW, sliceHMm);

      y += sliceH;
      page += 1;

      if (_cancelled) throw new Error("Đã hủy");
    }

    setOverlayText("Đang tải file…");
    pdf.save(filename);
  })();

  await withTimeout(task, 45000, "Export PDF");
}


function printToPdf(filename){
  // Fallback chắc chắn nhất trên mọi trình duyệt: mở tab mới và gọi window.print()
  const html = document.getElementById("exportHost").innerHTML || "<p>(Trống)</p>";
  const w = window.open("", "_blank");
  if (!w) {
    throw new Error("Trình duyệt chặn popup. Hãy cho phép popup rồi bấm lại.");
  }
  w.document.open();
  w.document.write(`<!doctype html>
<html>
<head>
<meta charset="utf-8"/>
<title>${filename}</title>
<style>
  body{ margin:0; background:#fff; color:#000; font-family: Arial, sans-serif; }
  .page{ width:794px; min-height:1123px; padding:28px 32px; box-sizing:border-box; margin:0 auto; }
  p{ margin:0 0 8px; line-height:1.45; }
  table{ width:100%; border-collapse:collapse; }
  td,th{ border:1px solid #ddd; padding:6px; }
  @media print{
    body{ -webkit-print-color-adjust: exact; print-color-adjust: exact; }
  }
</style>
</head>
<body>
  <div class="page">${html}</div>
<script>
  window.onload = () => { window.print(); };
<\/script>
</body>
</html>`);
  w.document.close();
}

// ===== UI wiring =====
(async function init() {
  const cancelBtn = document.getElementById('btnCancel');
  if (cancelBtn) cancelBtn.addEventListener('click', ()=>{ markCancelled(); showOverlay(false); msg('Đã hủy.'); });

  try {
    await loadDefaultTemplate();
    msg("Sẵn sàng");
  } catch (e) {
    msg("Chưa tải được template.docx. Bạn có thể chọn template bằng file.");
  }

  $("btnUseDefault").addEventListener("click", async () => {
    try {
      setBusy(true);
      await loadDefaultTemplate();
      msg("OK: dùng template.docx");
    } catch (e) {
      msg("Lỗi: " + (e?.message || e));
    } finally {
      showOverlay(false);
      setBusy(false);
    }
  });

  $("templateFile").addEventListener("change", async (ev) => {
    const f = ev.target.files?.[0];
    if (!f) return;
    try {
      setBusy(true);
      templateArrayBuffer = await loadTemplateFromFile(f);
      $("templateStatus").textContent = "Đang dùng: " + f.name;
      msg("OK: đã nạp template");
    } catch (e) {
      msg("Lỗi: " + (e?.message || e));
    } finally {
      showOverlay(false);
      setBusy(false);
    }
  });

  $("btnFillDemo").addEventListener("click", () => {
    const form = $("form");
    const set = (name, v) => (form.querySelector(`[name="${name}"]`).value = v);
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
    msg("Đã điền demo");
  });

  $("btnPrintPdf").addEventListener("click", async ()=>{
    try{
      _cancelled=false;
      showOverlay(true);
      setOverlayText("Đang chuẩn bị…");
      if (!templateArrayBuffer) throw new Error("Chưa có template.");
      const { map, outName } = getFormMapAndName();
      setOverlayText("Đang áp placeholder…");
      const filledAb = await buildFilledDocxArrayBuffer(templateArrayBuffer, map);
      const html = await withTimeout(docxToHtml(filledAb), 20000, "DOCX->HTML");
      await withTimeout(renderHidden(html), 8000, "Render");
      setOverlayText("Đang mở hộp thoại In…");
      showOverlay(false);
      printToPdf(outName);
    } catch(e){
      console.error(e);
      msg("Lỗi: " + (e?.message || e));
      showOverlay(false);
    }
  });

  
  $("btnTryAuto").addEventListener("click", async ()=>{
    try{
      _cancelled=false;
      showOverlay(true);
      setOverlayText("Đang chuẩn bị…");
      if (!templateArrayBuffer) throw new Error("Chưa có template.");
      const { map, outName } = getFormMapAndName();
      setOverlayText("Đang áp placeholder…");
      const filledAb = await buildFilledDocxArrayBuffer(templateArrayBuffer, map);
      const html = await withTimeout(docxToHtml(filledAb), 20000, "DOCX->HTML");
      await withTimeout(renderHidden(html), 8000, "Render");
      setOverlayText("Đang mở hộp thoại In…");
      await exportPdf(outName);
      showOverlay(false);
      msg("Đã mở hộp thoại In. Chọn Save as PDF để tải.");
    } catch(e){
      console.error(e);
      msg("Lỗi tải PDF tự động: " + (e?.message || e));
      showOverlay(false);
    }
  });

$("form").addEventListener("submit", async (ev) => {
    ev.preventDefault();
    try {
      setBusy(true);
      msg("Đang tạo PDF…");
      _cancelled = false;
      showOverlay(true);
      setOverlayText("Đang chuẩn bị…");

      if (!templateArrayBuffer) throw new Error("Chưa có template. Hãy bấm 'Dùng template.docx' hoặc chọn file template.");

      const { map, outName } = getFormMapAndName();
      const filledAb = await buildFilledDocxArrayBuffer(templateArrayBuffer, map);

      const html = await withTimeout(docxToHtml(filledAb), 20000, "DOCX->HTML");
      await withTimeout(renderHidden(html), 8000, "Render");
      setOverlayText("Đang mở hộp thoại In…");

      showOverlay(false);
      printToPdf(outName);
      msg("Đã mở hộp thoại In. Chọn Save as PDF để tải.");
    } catch (e) {
      console.error(e);
      msg("Lỗi: " + (e?.message || e));
    } finally {
      showOverlay(false);
      setBusy(false);
    }
  });
})();
