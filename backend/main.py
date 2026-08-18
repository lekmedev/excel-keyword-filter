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
import re
import shutil
import tempfile
import uuid
from datetime import datetime, timedelta
from pathlib import Path

from fastapi import FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from starlette.exceptions import HTTPException as StarletteHTTPException
from starlette.background import BackgroundTask

from .excel_processor import ExcelProcessingError, NoMatchError, convert_csv_to_xlsx, process_excel

app = FastAPI(title="Excel Keyword Filter", version="1.1.0")

MAX_UPLOAD_BYTES = 50 * 1024 * 1024  # 50 MB
ALLOWED_SUFFIXES = (".csv",)
RESULT_FILENAME = "TRUE_Result.xlsx"
JOB_TTL = timedelta(hours=1)  # file tạm không bao giờ sống lâu hơn 1 giờ
JOB_ID_PATTERN = re.compile(r"^[0-9a-f]{32}$")

TEMP_ROOT = Path(tempfile.gettempdir()) / "excel_keyword_filter"


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


@app.post("/process")
async def process(file: UploadFile = File(...), keywords: str = Form(...)):
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
        shutil.rmtree(job_dir, ignore_errors=True)
        return {"status": "no_match", "message": str(exc), "download_url": None}
    except ExcelProcessingError as exc:
        shutil.rmtree(job_dir, ignore_errors=True)
        raise HTTPException(status_code=400, detail=str(exc))

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


@app.middleware("http")
async def no_cache_frontend(request: Request, call_next):
    """Chặn cache (browser + Cloudflare) cho HTML/CSS/JS để bản mới hiện ngay."""
    response = await call_next(request)
    path = request.url.path
    if path in ("/", "/index.html", "/style.css", "/script.js") or path.endswith((".css", ".js")):
        response.headers["Cache-Control"] = "no-cache, max-age=0"
    return response