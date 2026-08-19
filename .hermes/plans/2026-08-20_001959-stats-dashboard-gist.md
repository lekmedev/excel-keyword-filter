# Excel Keyword Filter — Plan thay /stats bằng Analytics Dashboard mẫu (gist)

> **Trạng thái:** Chỉ lên plan, chưa code. Marshal đưa link gist `Frontend.html` (mẫu dashboard Tailwind + Chart.js) — muốn update giao diện & thống kê theo đúng code đó.

---

## Goal

Thay toàn bộ trang `/stats` hiện tại bằng **Analytics Dashboard** theo mẫu gist: layout sidebar + KPI cards + line chart + horizontal bar chart keywords + activity log table có search — nhưng **dữ liệu thật từ `/stats/data`** (không phải số demo trong gist).

## Current Context / Assumptions

- File mẫu: `/tmp/ref_frontend.html` (257 dòng, dùng **Tailwind CDN** + **Chart.js CDN**).
- Cấu trúc mẫu:
  - Sidebar trái (tối) — logo "X", menu Tổng quan / Từ khóa & Tệp tin / Bảo mật & Giám sát IP; footer Server Status.
  - Header: tiêu đề + dropdown chọn khoảng thời gian (Hôm nay / 7 ngày / 30 ngày) + nút Xuất báo cáo.
  - **4 KPI cards**: Tổng lượt lọc file, Unique IPs, Tỷ lệ thành công, Thời gian xử lý TB.
  - **2 charts**: line chart "Lượt xử lý theo ngày" (7 ngày) + horizontal bar "Top từ khóa".
  - **Activity log table**: Thời gian / IP / Từ khóa / Tên file & dung lượng / Dòng khớp / Thời gian xử lý / Trạng thái + ô tìm kiếm.
- Hiện tại `/stats` là dark theme Bootstrap tự viết — sẽ **thay hẳn bằng mẫu này** (light theme Tailwind).
- Backend đã có: `GET /stats` (trả stats.html), `GET /stats/data` (JSON: total, today, uploads, top_ips, top_keywords, keyword_count, by_hour, by_day, recent, retention_days).
- Giữ nguyên hành vi log: chỉ đếm upload + download (không đếm lần mở trang) — **QUAN TRỌNG giữ**.

## Dữ liệu thật ánh xạ từ /stats/data

| Ô UI mẫu | Dữ liệu thật |
|---|---|
| Tổng lượt lọc file | `uploads` |
| Unique IPs | `len(top_ips)` hoặc đếm IP unique từ recent — cần thêm `unique_ips` vào summarize |
| Tỷ lệ thành công | từ recent: status 200 / tổng (gần đúng); hoặc giữ "99%" |
| Thời gian xử lý TB | chưa có — cần log `duration_ms` (thêm vào endpoint /process) |
| Line chart theo ngày | `by_day` (label = ngày, value = count) — fallback 7 ngày gần |
| Bar chart top keywords | `top_keywords` (label = keyword, value = count) — **indexAxis 'y' ngang** |
| Activity log table | `recent` (ts, ip, keywords, path…); cần thêm `filename`, `duration_ms`, `match_count` |
| Dropdown khoảng thời gian | lọc client-side theo ts? Hiện summarize cố định `days=90` — có thể thêm param `?days=` hoặc client tự lọc `recent` |

## Backend cần bổ sung (nhỏ)

1. `Stats.summarize()` thêm `unique_ips: int` (đếm IP distinct từ records trong window) + `error_count` (status != 200) + `avg_duration_ms` (nếu có duration).
2. `POST /process` log thêm:
   - `filename`: tên file gốc (đã có `file.filename`)
   - `duration_ms`: đo thời gian xử lý (time.monotonic() quanh process_excel)
   - `total_rows` (tùy chọn — từ excel_processor nếu dễ)
3. `recent` đã có keywords — đảm bảo thêm filename, duration_ms, match_count (match_count đã log).
4. (Tùy chọn) `GET /stats/data?days=7|30|90` — nếu anh muốn dropdown hoạt động thật. Mặc định làm.

## Files Likely to Change

- `frontend/stats.html` — **REWRITE** theo mẫu gist, dùng Tailwind, nối dữ liệu thật từ `/stats/data`, giữ mobile responsive, search client-side trên recent.
- `backend/stats.py` — summarize thêm unique_ips, error_count, avg_duration_ms; hỗ trợ `days` param.
- `backend/main.py` — /process log filename + duration_ms; `/stats/data` nhận `?days=`.
- `tests/test_app.py` — test mới: unique_ips, error_count, duration_ms, filename trong log, `/stats/data?days=`.

## Step-by-Step Plan (TDD)

### Task 1: Backend — log thêm filename + duration_ms
- Test: POST /process → log record có `filename`, `duration_ms` (số dương), keywords, match_count.
- Implement: trong endpoint đo `t0=time.monotonic()`, `duration_ms=round((time.monotonic()-t0)*1000)`; thêm `extra={"keywords", "match_count", "filename", "duration_ms"}`.

### Task 2: Backend — summarize thêm unique_ips, error_count, avg_duration_ms, hỗ trợ days param
- Test: ghi 3 records (2 IP, 1 status 400, có duration 500ms) → summarize(days=…) trả `unique_ips=2`, `error_count=1`, `avg_duration_ms≈500`.
- Implement: đếm set(ip), đếm status != 200, trung bình duration_ms từ records có duration.
- `GET /stats/data?days=30` → `stats.summarize(days=days)`.

### Task 3: Frontend — rewrite stats.html theo gist
- Copy layout gist (Tailwind + Chart.js), thay dữ liệu demo bằng fetch('/stats/data') render dynamic:
  - KPI cards: uploads, unique_ips, success_rate, avg_duration (làm tròn)
  - Line chart: by_day (fallback: 7 ngày gần — lấy last 7 keys)
  - Bar chart ngang: top_keywords (indexAxis 'y')
  - Activity log: recent → bảng (ts format, ip, keywords chips, filename nếu có, match_count, status badge)
  - Search input: lọc client-side theo IP/keyword
- Giữ Sidebar (chỉ để trang trí, không cần điều hướng thật; hoặc đơn giản ẩn trên mobile — mẫu đã có `hidden md:flex`).

### Task 4: Giữ nguyên đếm log
- **Không đổi middleware** — vẫn chỉ log /process + /download. Verify test cũ pass.

### Task 5: Build + deploy + verify E2E
- `pytest -q` → ~33+ pass
- `docker compose build && up -d`
- E2E: POST /process vài lần (IP khác + keywords khác) → /stats (HTML mới có Tailwind + huỷ demo) + /stats/data có unique_ips, error_count, duration_ms
- Push Docker Hub `v1.6.0` + GitHub tag

## Tests / Validation

- `cd ~/apps/excel-keyword-filter && . .venv/bin/activate && python -m pytest -q`
- `curl -s http://100.95.76.53:8000/stats` — HTML chứa `tailwindcss` + `trafficChart` + `keywordsChart`
- `curl -s http://100.95.76.53:8000/stats/data` — có `unique_ips`, `error_count`, `avg_duration_ms`
- Upload thử → bảng recent hiện filename, duration, keywords

## Risks / Tradeoffs / Open Questions

1. **Tailwind CDN** (`cdn.tailwindcss.com`) — dùng production? Nó là dev build chậm hơn một chút, nhưng đơn giản cho app nhỏ. Anh OK?
2. **Dữ liệu demo trong gist** (14,280 lượt, 99.4%…) — sẽ **thay hẳn bằng số thật**, không giữ số ảo. Đúng ý anh chứ?
3. **Phần "Tên file & Dung lượng"** — cần log filename (sẽ thêm) + file size (dễ đo trong endpoint). Muốn hiện cả 2?
4. **"Dòng khớp / Tổng"** — match_count đã có; "Tổng" (tổng dòng file) cần thêm từ excel_processor — thêm hay chỉ hiện match_count?
5. **Dropdown khoảng thời gian** — làm thật (7/30/90 ngày, gọi lại /stats/data?days=) hay chỉ để trưng? Đề xuất làm thật.
6. **Xuất báo cáo** — nút này trong mẫu; muốn hoạt động (tải CSV của stats) hay bỏ?

## Next Steps

- Anh trả lời 3-6 câu hỏi trên (hoặc OK hết) → implement Task 1-5 (~40 phút).
- Nếu anh muốn giữ nguyên toàn bộ mẫu (kể cả sidebar/export), em làm y như mẫu.