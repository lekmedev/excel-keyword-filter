# Excel Keyword Filter

Webapp lọc file CSV theo từ khóa: upload file `.csv`, nhập danh sách từ khóa,
hệ thống lọc các dòng có cột **N** (cột 14) chứa bất kỳ từ khóa nào,
rồi trả về file `TRUE_Result.xlsx` đã làm sạch theo logic macro VBA gốc.

**Source:** https://github.com/lekmedev/excel-keyword-filter  
**Docker Hub:** https://hub.docker.com/r/thn05/excel-keyword-filter

## Share cho người khác — chạy 1 lệnh Docker

Image đã public trên **Docker Hub**: `thn05/excel-keyword-filter:latest`

Người nhận chỉ cần cài Docker rồi chạy:

```bash
docker run -d -p 8000:8000 --name excel-keyword-filter --restart unless-stopped thn05/excel-keyword-filter:latest
```

Mở trình duyệt: **http://localhost:8000**

Các tag: `latest`, `v1.1.0`
Trang Docker Hub: https://hub.docker.com/r/thn05/excel-keyword-filter

## Tính năng

- Upload bằng **kéo thả** hoặc chọn file (`<input type="file">`)
- Hỗ trợ **`.csv`** (auto-detect encoding UTF-8/cp1252 + delimiter `,` `;` tab)
- Validate: chỉ nhận `.csv`, giới hạn **50 MB** (frontend + backend)
- Xử lý toàn bộ dữ liệu (coi cột 14 = cột N là cột keyword); mỗi keyword 1 dòng, khớp không phân biệt hoa thường
- Kết quả:
  - Sheet mới **TRUE_Result**, copy header + các dòng khớp (giá trị + format cơ bản)
  - **Xóa các cột A, B, G, H, I, K, L, M**
  - Chèn cột **STT** đầu tiên, đánh số `1, 2, 3...`
  - Header **Quantity Replaced** tại cột **H**
  - Auto width + border toàn bộ bảng
- Không tìm thấy dữ liệu → thông báo **"No matching data found"**
- **Không lưu dữ liệu lâu dài**: file tạm nằm trong `/tmp`, bị xóa **ngay sau khi user download**, job cũ hơn 1 giờ bị dọn tự động
- Xử lý lỗi: file Excel hỏng/không mở được, file sai định dạng, quá size, thiếu keyword

## Cấu trúc project

```
excel-keyword-filter/
├── backend/
│   ├── main.py               # FastAPI app: POST /process, GET /download, GET /health
│   ├── excel_processor.py    # Toàn bộ logic Excel (openpyxl)
│   └── requirements.txt
├── frontend/
│   ├── index.html            # Bootstrap 5 + kéo thả upload
│   ├── script.js             # Validate client, fetch /process, hiển thị kết quả
│   └── style.css
├── tests/
│   └── test_app.py           # Unit tests + API tests (pytest)
├── Dockerfile                # Multi-stage build (python:3.12-slim)
├── docker-compose.yml
└── README.md
```

## Cách chạy

```bash
docker compose up -d --build
```

Truy cập:
- http://localhost:8000 (trên máy creek)
- http://100.95.76.53:8000 (qua Tailscale — dùng được từ mọi thiết bị trong tailnet)

Xem log:

```bash
docker compose logs -f
```

Dừng:

```bash
docker compose down
```

## Test

### 1. Test bằng file Excel mẫu

Tạo file mẫu nhanh bằng script Python (chạy trên máy có Python + openpyxl):

```bash
python3 - <<'PY'
from openpyxl import Workbook
wb = Workbook()
ws = wb.active
ws.title = "DataSheet"
headers = ["A","B","C","D","E","F","G","H","I","J","K","L","M","N","O"]
for i, h in enumerate(headers, start=1):
    ws.cell(row=1, column=i, value=f"Col{h}")
rows = [
    ["r2","b","c","d","e","f","g","h","i","j","k","l","m","Logitech keyboard","o2"],
    ["r3","b","c","d","e","f","g","h","i","j","k","l","m","Wireless mouse","o3"],
    ["r4","b","c","d","e","f","g","h","i","j","k","l","m","Bàn phím cơ","o4"],
    ["r5","b","c","d","e","f","g","h","i","j","k","l","m","Monitor 24 inch","o5"],
]
for r, row in enumerate(rows, start=2):
    for c, val in enumerate(row, start=1):
        ws.cell(row=r, column=c, value=val)
wb.save("sample_keyword.xlsx")
print("Đã tạo sample_keyword.xlsx")
PY
```

Sau đó:

1. Mở http://localhost:8000
2. Kéo thả `sample_keyword.xlsx`
3. Nhập keywords:

```
keyboard
bàn phím
```

4. Bấm **Process Excel** → kết quả: `TRUE_Result.xlsx` chứa 2 dòng khớp (Logitech keyboard, Bàn phím cơ),
   có cột STT, cột O thành cột H kèm header "Quantity Replaced" → bấm **Download**.

Thử trường hợp không khớp: keyword `xyzabc` → thông báo **"No matching data found"**.

### 2. Chạy bộ test tự động

```bash
cd ~/apps/excel-keyword-filter
python3 -m venv .venv
. .venv/bin/activate
pip install -r backend/requirements.txt pytest
pytest -q
```

Các test kiểm tra: khớp keyword (case-insensitive), xóa đúng cột + chèn STT,
header Quantity Replaced, border + auto width, không khớp → NoMatchError,
file gốc không bị biến đổi, API reject file không phải .xlsx, reject 50MB+,
download xong thì file tạm bị xóa (404 khi tải lại), file Excel hỏng → 400.

## API

| Method | Path          | Mô tả |
|--------|---------------|-------|
| POST   | `/process`    | `multipart/form-data`: `file` (CSV `.csv`) + `keywords` (text, mỗi keyword 1 dòng) → JSON `{status, message, match_count, download_url}` |
| GET    | `/download/{job_id}` | Tải `TRUE_Result.xlsx`; **file tạm bị xóa ngay sau khi tải** |
| GET    | `/health`     | Liveness check |

JSON trả về khi không khớp: `{"status": "no_match", "message": "No matching data found", "download_url": null}`

## Ghi chú bảo mật / dữ liệu

- Không lưu file người dùng: toàn bộ nằm trong `/tmp/excel_keyword_filter/<job_id>/`, TTL 1 giờ + xóa ngay sau download
- Container chạy bằng user `appuser` (non-root)
- Không mở port ra ngoài nếu không cần: chỉ publish 8000 khi muốn truy cập từ mạng khác