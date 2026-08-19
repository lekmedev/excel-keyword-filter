"""
FastAPI backend cho app lọc CSV theo từ khóa.

Endpoints:
  POST /process          - upload .csv + keywords -> JSON {status, download_url, match_count}
  GET  /download/{job}   - tải file TRUE_Result.xlsx (file tạm bị xóa sau khi tải xong)
  GET  /health           - liveness probe

Frontend tĩnh được mount tại "/".

Nguyên tắc dữ liệu:
  - File tạm nằm trong /tmp/excel_keyword_filter/<job_id>/, mỗi job 1 thư mục.
  - Xóa file ngay sau khi download, và dọn job cũ hơn 1 giờ -> KHÔNG lưu dữ liệu lâu dài.
"""

import csv
import io
import json
import os
import re
import shutil
import tempfile
import uuid
from datetime import datetime, timedelta
from pathlib import Path

from fastapi import FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
from starlette.exceptions import HTTPException as StarletteHTTPException
from starlette.background import BackgroundTask

from .excel_processor import ExcelProcessingError, NoMatchError, convert_csv_to_xlsx, process_excel
from .stats import Stats

app = FastAPI(title="Excel Keyword Filter", version="1.1.0")

MAX_UPLOAD_BYTES = 50 * 1024 * 1024  # 50 MB
ALLOWED_SUFFIXES = (".csv",)
RESULT_FILENAME = "TRUE_Result.xlsx"
JOB_TTL = timedelta(hours=1)  # file tạm không bao giờ sống lâu hơn 1 giờ
JOB_ID_PATTERN = re.compile(r"^[0-9a-f]{32}$")

TEMP_ROOT = Path(tempfile.gettempdir()) / "excel_keyword_filter"
STATS_DIR = Path(os.environ.get("STATS_DIR", tempfile.gettempdir())) / "excel_stats"
STATS_RETENTION_DAYS = 90
stats = Stats(STATS_DIR, retention_days=STATS_RETENTION_DAYS)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _sweep_jobs() -> None:
    """Dọn các job quá cũ (chạy mỗi lần có request mới)."""
    now = datetime.now()
    if not TEMP_ROOT.exists():
        return
    for job_dir in TEMP_ROOT.iterdir():
        if not job_dir.is_dir():
            continue
        try:
            mtime = datetime.fromtimestamp(job_dir.stat().st_mtime)
        except OSError:
            continue
        if now - mtime > JOB_TTL:
            shutil.rmtree(job_dir, ignore_errors=True)


# ---------------------------------------------------------------------------
# Routes
# ---------------------------------------------------------------------------

@app.get("/health")
async def health():
    return {"status": "ok"}


@app.get("/stats")
async def stats_page():
    """Trang thống kê sử dụng người dùng (static HTML, fetch /stats/data)."""
    stats_file = FRONTEND_DIR / "stats.html"
    if stats_file.is_file():
        return FileResponse(stats_file)
    return HTMLResponse("<h1>stats.html missing</h1>", status_code=404)


@app.get("/stats/data")
async def stats_data():
    """JSON dữ liệu thống kê cho /stats."""
    return stats.summarize(days=STATS_RETENTION_DAYS)


@app.post("/process")
async def process(request: Request, file: UploadFile = File(...), keywords: str = Form(...)):
    """Nhận file Excel + danh sách từ khóa, trả về link tải file kết quả."""
    # Validate định dạng file.
    filename = file.filename or ""
    if not filename.lower().endswith(ALLOWED_SUFFIXES):
        raise HTTPException(status_code=400, detail="Chỉ chấp nhận file .csv")

    # Validate từ khóa.
    if not keywords.strip():
        raise HTTPException(status_code=400, detail="Vui lòng nhập ít nhất một từ khóa")

    _sweep_jobs()

    # Tạo job dir + lưu file với giới hạn 50 MB (stream, không nạp hết vào RAM).
    job_id = uuid.uuid4().hex
    job_dir = TEMP_ROOT / job_id
    job_dir.mkdir(parents=True, exist_ok=True)
    ext = Path(filename).suffix.lower()
    input_path = job_dir / f"input_{job_id[:8]}{ext}"

    try:
        size = 0
        with input_path.open("wb") as out:
            while chunk := await file.read(1024 * 1024):
                size += len(chunk)
                if size > MAX_UPLOAD_BYTES:
                    raise HTTPException(status_code=413, detail="File vượt quá giới hạn 50 MB")
                out.write(chunk)
    except HTTPException:
        shutil.rmtree(job_dir, ignore_errors=True)
        raise
    finally:
        await file.close()

    # CSV được chuyển sang .xlsx trước, rồi xử lý chung bằng openpyxl.
    if ext == ".csv":
        converted = job_dir / f"input_{job_id[:8]}.xlsx"
        try:
            convert_csv_to_xlsx(input_path, converted)
        except ExcelProcessingError as exc:
            shutil.rmtree(job_dir, ignore_errors=True)
            raise HTTPException(status_code=400, detail=str(exc))
        input_path = converted

    # Xử lý Excel.
    try:
        output_path, match_count = process_excel(input_path, keywords)
    except NoMatchError as exc:
        stats.record(
            ip=_client_ip(request),
            method="POST",
            path="/process",
            status=200,
            ua="",
            extra={"keywords": keywords},
        )
        shutil.rmtree(job_dir, ignore_errors=True)
        return {"status": "no_match", "message": str(exc), "download_url": None}
    except ExcelProcessingError as exc:
        stats.record(
            ip=_client_ip(request),
            method="POST",
            path="/process",
            status=400,
            ua="",
            extra={"keywords": keywords},
        )
        shutil.rmtree(job_dir, ignore_errors=True)
        raise HTTPException(status_code=400, detail=str(exc))

    # Log upload thành công kèm keywords + match_count.
    stats.record(
        ip=_client_ip(request),
        method="POST",
        path="/process",
        status=200,
        ua=request.headers.get("user-agent", ""),
        extra={"keywords": keywords, "match_count": match_count},
    )

    return {
        "status": "ok",
        "message": "Xử lý hoàn tất",
        "match_count": match_count,
        "download_url": f"/download/{job_id}",
    }


@app.get("/download/{job_id}")
async def download(job_id: str):
    """Tải file kết quả; sau khi gửi xong sẽ xóa luôn job dir (input + output)."""
    if not JOB_ID_PATTERN.fullmatch(job_id):
        raise HTTPException(status_code=404, detail="Không tìm thấy file kết quả")

    job_dir = TEMP_ROOT / job_id
    output_path = job_dir / RESULT_FILENAME
    if not output_path.is_file():
        raise HTTPException(status_code=404, detail="File đã hết hạn hoặc không tồn tại")

    def cleanup():
        shutil.rmtree(job_dir, ignore_errors=True)

    return FileResponse(
        output_path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=RESULT_FILENAME,
        background=BackgroundTask(cleanup),
    )


# Mount frontend tĩnh — phải khai báo SAU các route API để không nuốt mất chúng.
FRONTEND_DIR = Path(__file__).resolve().parent.parent / "frontend"
app.mount("/", StaticFiles(directory=FRONTEND_DIR, html=True), name="frontend")


def _client_ip(request: Request) -> str:
    """IP thật: Cloudflare set CF-Connecting-IP; fallback X-Forwarded-For; rồi client host."""
    return (
        request.headers.get("cf-connecting-ip")
        or (request.headers.get("x-forwarded-for") or "").split(",")[0].strip()
        or (request.client.host if request.client else "")
    )


def _render_stats_html(data: dict) -> str:
    """(Không còn dùng — /stats trả static stats.html; giữ để tương thích test cũ.)"""
    top_rows = "".join(
        f"<tr><td>{i+1}</td><td>{ip}</td><td class='text-end'>{c}</td></tr>"
        for i, (ip, c) in enumerate(
            [(t["ip"], t["count"]) for t in data.get("top_ips", [])]
        )
    ) or "<tr><td colspan='3' class='text-center'>Chưa có dữ liệu</td></tr>"

    # Biểu đồ theo giờ (đơn giản: bar bằng div)
    hours = data.get("by_hour", {})
    max_h = max(hours.values()) if hours else 1
    bars = ""
    for h in range(24):
        cnt = hours.get(str(h), 0)
        pct = int(cnt / max_h * 100) if max_h else 0
        bars += (
            f"<div class='d-flex align-items-center mb-1'>"
            f"<span class='me-2' style='width:28px'>{h:02d}h</span>"
            f"<div class='flex-grow-1'><div class='bg-primary' style='height:14px;width:{pct}%'></div></div>"
            f"<span class='ms-2'>{cnt}</span></div>"
        )

    return f"""<!doctype html>
<html lang="vi">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Thống kê sử dụng</title>
<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet">
</head>
<body class="bg-light">
<div class="container py-4">
  <h1 class="mb-4">📊 Thống kê sử dụng</h1>
  <div class="row g-3 mb-4">
    <div class="col-md-4"><div class="card"><div class="card-body text-center">
      <div class="fs-3 fw-bold">{data.get('total', 0)}</div><div class="text-secondary">Tổng requests</div>
    </div></div></div>
    <div class="col-md-4"><div class="card"><div class="card-body text-center">
      <div class="fs-3 fw-bold">{data.get('today', 0)}</div><div class="text-secondary">Hôm nay</div>
    </div></div></div>
    <div class="col-md-4"><div class="card"><div class="card-body text-center">
      <div class="fs-3 fw-bold">{data.get('uploads', 0)}</div><div class="text-secondary">Số lần upload</div>
    </div></div></div>
  </div>
  <div class="row g-3">
    <div class="col-lg-6">
      <div class="card"><div class="card-header">Top IP</div>
      <div class="card-body p-0"><table class="table table-sm mb-0">
        <thead><tr><th>#</th><th>IP</th><th class="text-end">Lượt</th></tr></thead>
        <tbody>{top_rows}</tbody></table></div></div>
    </div>
    <div class="col-lg-6">
      <div class="card"><div class="card-header">Theo giờ trong ngày</div>
      <div class="card-body">{bars}</div></div>
    </div>
  </div>
  <p class="text-secondary small mt-4">Giữ log {data.get('retention_days', 90)} ngày, tự xóa. IP lấy từ CF-Connecting-IP (sau Cloudflare).</p>
</div>
</body>
</html>"""


@app.middleware("http")
async def log_access(request: Request, call_next):
    """Ghi log chỉ cho hành động SỬ DỤNG thật (download result).
    Upload (/process) tự log trong endpoint. Chỉ mở trang chủ KHÔNG được đếm."""
    path = request.url.path
    response = await call_next(request)
    # Giữ cache-busting: HTML/CSS/JS không bao giờ cache lâu (CF + browser).
    if path in ("/", "/index.html", "/style.css", "/script.js") or path.endswith((".css", ".js")):
        response.headers["Cache-Control"] = "no-cache, max-age=0"
    # Chỉ log download (hành động dùng thật, xảy ra sau upload).
    if path.startswith("/download/"):
        stats.record(
            ip=_client_ip(request),
            method=request.method,
            path=path,
            status=response.status_code,
            ua=request.headers.get("user-agent", ""),
        )
    return response