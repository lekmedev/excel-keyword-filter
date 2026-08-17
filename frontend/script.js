// Excel Keyword Filter — frontend logic
// Flow: kéo thả/chọn file -> validate .xlsx + <=50MB -> POST /process với FormData
//       -> show kết quả hoặc lỗi -> nút Download file TRUE_Result.xlsx

(function () {
  "use strict";

  const dropZone = document.getElementById("dropZone");
  const fileInput = document.getElementById("fileInput");
  const fileInfo = document.getElementById("fileInfo");
  const fileNameEl = document.getElementById("fileName");
  const clearFileBtn = document.getElementById("clearFileBtn");
  const keywordsEl = document.getElementById("keywords");
  const processBtn = document.getElementById("processBtn");
  const btnSpinner = document.getElementById("btnSpinner");
  const btnText = document.getElementById("btnText");
  const resultBox = document.getElementById("resultBox");

  const MAX_SIZE = 50 * 1024 * 1024; // 50 MB — khớp với giới hạn backend
  const ALLOWED = [".csv"];

  let selectedFile = null;

  // ---- Helpers hiển thị ---------------------------------------------------

  function showResult(html, variant) {
    resultBox.className = "mt-4 d-block alert alert-" + variant;
    resultBox.innerHTML = html;
    resultBox.scrollIntoView({ behavior: "smooth", block: "nearest" });
  }

  function hideResult() {
    resultBox.className = "mt-4 d-none";
    resultBox.innerHTML = "";
  }

  function clearSelection() {
    selectedFile = null;
    fileInput.value = "";
    fileInfo.classList.add("d-none");
    fileInfo.style.display = "none";
    updateButtonState();
    hideResult();
  }

  function updateButtonState() {
    const hasFile = !!selectedFile;
    const hasKeywords = keywordsEl.value.trim().length > 0;
    processBtn.disabled = !(hasFile && hasKeywords);
  }

  // ---- Validate file ------------------------------------------------------

  function pickFile(file) {
    if (!file) return;

    const name = file.name.toLowerCase();
    const ok = ALLOWED.some((ext) => name.endsWith(ext));
    if (!ok) {
      showResult('<i class="bi bi-x-circle me-1"></i>Chỉ chấp nhận file <strong>.csv</strong>.', "danger");
      return;
    }
    if (file.size > MAX_SIZE) {
      showResult('<i class="bi bi-x-circle me-1"></i>File vượt quá giới hạn <strong>50 MB</strong>.', "danger");
      return;
    }

    selectedFile = file;
    fileNameEl.textContent = file.name + " (" + formatBytes(file.size) + ")";
    fileInfo.classList.remove("d-none");
    fileInfo.style.display = "";
    hideResult();
    updateButtonState();
  }

  function formatBytes(bytes) {
    if (bytes < 1024) return bytes + " B";
    if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + " KB";
    return (bytes / (1024 * 1024)).toFixed(1) + " MB";
  }

  // ---- Sự kiện kéo thả & chọn file ----------------------------------------

  dropZone.addEventListener("click", () => fileInput.click());
  dropZone.addEventListener("keydown", (e) => {
    if (e.key === "Enter" || e.key === " ") {
      e.preventDefault();
      fileInput.click();
    }
  });

  // Chống trình duyệt mở file khi thả nhầm ra ngoài vùng upload
  ["dragover", "drop"].forEach((evt) =>
    document.addEventListener(evt, (e) => e.preventDefault())
  );

  dropZone.addEventListener("dragover", (e) => {
    e.preventDefault();
    dropZone.classList.add("dragging");
  });
  dropZone.addEventListener("dragleave", () => dropZone.classList.remove("dragging"));
  dropZone.addEventListener("drop", (e) => {
    e.preventDefault();
    dropZone.classList.remove("dragging");
    pickFile(e.dataTransfer.files[0]);
  });

  fileInput.addEventListener("change", () => pickFile(fileInput.files[0]));
  clearFileBtn.addEventListener("click", clearSelection);
  keywordsEl.addEventListener("input", updateButtonState);

  // ---- Process Excel ------------------------------------------------------

  processBtn.addEventListener("click", async () => {
    if (!selectedFile) return;
    hideResult();

    // Bật trạng thái "đang xử lý"
    processBtn.disabled = true;
    btnSpinner.classList.remove("d-none");
    btnText.textContent = "Đang xử lý...";

    const form = new FormData();
    form.append("file", selectedFile);
    form.append("keywords", keywordsEl.value);

    try {
      const res = await fetch("/process", { method: "POST", body: form });
      const data = await res.json().catch(() => ({}));

      if (!res.ok) {
        showResult(
          '<i class="bi bi-x-circle me-1"></i>' + escapeHtml(data.detail || "Đã có lỗi xảy ra, vui lòng thử lại."),
          "danger"
        );
        return;
      }

      if (data.status === "no_match") {
        showResult(
          '<i class="bi bi-info-circle me-1"></i><strong>No matching data found</strong> — không có dòng nào chứa các từ khóa đã nhập.',
          "warning"
        );
        return;
      }

      // Thành công: hiện nút Download
      showResult(
        '<i class="bi bi-check-circle me-1"></i><strong>Xử lý hoàn tất!</strong> Tìm thấy ' +
          data.match_count +
          " dòng khớp.<br>" +
          '<a class="btn btn-success mt-2" href="' +
          data.download_url +
          '" id="downloadBtn">⬇ Download file kết quả</a>',
        "success"
      );
      const downloadBtn = document.getElementById("downloadBtn");
      downloadBtn.addEventListener("click", () => {
        // Sau khi bấm download, backend tự xóa file tạm — reset form cho lần dùng sau.
        setTimeout(clearSelection, 2000);
      });
    } catch (err) {
      showResult(
        '<i class="bi bi-x-circle me-1"></i>Không kết nối được server. Vui lòng thử lại.',
        "danger"
      );
    } finally {
      processBtn.disabled = false;
      btnSpinner.classList.add("d-none");
      btnText.textContent = "Process Excel";
    }
  });

  // ---- Tiện ích -----------------------------------------------------------

  function escapeHtml(str) {
    const div = document.createElement("div");
    div.textContent = String(str);
    return div.innerHTML;
  }
})();