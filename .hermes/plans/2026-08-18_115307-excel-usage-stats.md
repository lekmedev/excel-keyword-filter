# Excel Keyword Filter — Plan thống kê người dùng (usage analytics)

> **Trạng thái:** Chỉ lên plan, chưa code. Marshal muốn biết: có ai đang dùng webapp không, dùng lúc nào, từ IP nào.

---

## Goal

Thêm khả năng **theo dõi + thống kê lượt dùng** webapp excel-keyword-filter: số người truy cập, thời điểm truy cập, IP nguồn — nhưng **không lưu dữ liệu nhạy cảm lâu dài**, không phức tạp, và không làm chậm app.

## Current Context / Assumptions

- App hiện tại: FastAPI + uvicorn, Docker compose, **KHÔNG có volume**, dữ liệu người dùng chỉ nằm trong `/tmp` container rồi bị xóa.
- `excel.8biz.net` trỏ về Cloudflare (proxy) → request đến app qua CF, IP thật nằm trong header `CF-Connecting-IP` (hoặc `X-Forwarded-For`), không phải IP docker bridge.
- Traefik đang chạy cho nhiều service nhưng **chưa bật access log** cho app này; docker logs hiện chỉ có lỗi ACME.
- Marshal không cần dashboard phức tạp — muốn biết: **ai truy cập, khi nào, IP nào**.
- Ràng buộc từ trước: app phải **0.0.0.0**, có thể truy cập qua tailnet; không bắt buộc thêm vào dashboard nếu đây là tính năng của app hiện có (không phải service mới).
- `requirements.txt` hiện: fastapi, uvicorn, openpyxl... (nhẹ, không có database ngoài).

## Proposed Approach (2 lựa chọn)

### Option A — Nhẹ nhất: ghi access log vào file JSON/CSV trong container + endpoint đọc thống kê (RECOMMENDED)

- Backend thêm middleware log: mỗi request `GET /`, `POST /process`, `GET /download/*`, `GET /health` (tùy chọn bỏ health) → ghi 1 dòng JSON vào `/app/logs/access.jsonl` trong container.
- Nội dung mỗi dòng: `{ts, ip, method, path, status, match_count?, user_agent, country?}`.
- IP lấy từ: `CF-Connecting-IP` → `X-Forwarded-For` → `request.client.host` (fallback).
- Tự **xóa log > 30 ngày** (giữ ngắn hạn, không phải dữ liệu vĩnh viễn).
- Endpoint mới `GET /stats` (hoặc `/admin/stats`):
  - Tổng số requests hôm nay/tuần/này
  - Top 10 IP
  - Số lượt upload (`POST /process`) + số match/no-match
  - Biểu đồ đơn giản (giờ trong ngày / ngày trong tuần) — trả JSON thuần
- Frontend: thêm 1 trang nhỏ `/stats.html` (Bootstrap, fetch `/stats`, hiển thị bảng + vài số).
- **Không cần database** — file JSONL là đủ cho quy mô này (vài trăm request/ngày). Dễ dọn, dễ backup.
- **Bảo mật**: để `GET /stats` public hay chặn bằng token đơn giản? Đề xuất: token GET param/header `?key=...` (Env var `STATS_KEY`), mặc định off (public) hoặc on khi set.

### Option B — Nặng hơn: gắn với infrastructure sẵn có (Traefik access log / Cloudflare analytics)

- Bật **Traefik access log** (file acccess.log trên host) → parse bằng cron → thống kê.
- Hoặc dùng **Cloudflare Analytics** API (zone 8biz.net) → đã có sẵn số liệu requests/IP, Marshal xem trên CF dashboard luôn.
- Ưu: không thêm code vào app; nhược: khó lọc riêng app này, không có match_count, phụ thuộc CF token (token hiện chỉ tunnel-scoped, không đọc analytics được — sẽ cần token mới).

**Chốt:** Implement **Option A** (tự thân trong app, đơn giản, chủ động), còn **thông tin tham khảo từ Cloudflare/Traefik** có thể bổ sung sau nếu anh muốn.

---

## Files Likely to Change

- `backend/main.py` — thêm middleware log + endpoint `/stats` + config STATS_KEY
- `backend/stats.py` (mới) — module quản lý access log (ghi, đọc, dọn)
- `frontend/stats.html` (mới) — trang thống kê
- `frontend/script.js` — (nếu cần) không đổi; trang stats có script riêng
- `backend/requirements.txt` — không đổi hoặc thêm gì nhẹ (std-only)
- `Dockerfile` — thêm `RUN mkdir /app/logs` (hoặc tạo runtime)
- `docker-compose.yml` — (tùy chọn) mount volume nhỏ nếu muốn log tồn tại qua restart; không bắt buộc
- `tests/test_app.py` — thêm test middleware/stats

---

## Step-by-Step Plan (TDD)

### Task 1: Module stats (`backend/stats.py`)
- Test: `test_stats_record_and_read()` — ghi 1 record, đọc ra đúng
- Implement: `Stats` class: `record(...)`, `read(...)`, `cleanup(older_than_days=30)`, dùng thư mục `STATS_DIR` (default `/tmp/excel_keyword_filter_stats` hoặc `/app/logs` qua env)
- File format: JSONL, mỗi dòng 1 dict `{ts, ip, method, path, status, match_count, ua}`

### Task 2: Middleware log trong main.py
- Test: gọi `/` và `/health`, check file log có dòng tương ứng, ip fallback đúng
- Implement: middleware (`@app.middleware("http")`) ghi log sau khi response; lấy IP từ headers
- Bỏ qua `/stats` và static assets để tránh nhiễu (chỉ log GET /, POST /process, GET /download/*)
- `match_count` lấy từ response body khi path == /process (đọc body tạm, parse JSON) — hoặc đơn giản hóa: không cần match_count, chỉ cần status. **Quyết định:** lấy match_count nếu dễ (đọc JSON body), không thì bỏ qua — chỉ log status.

### Task 3: Endpoint `GET /stats` (+ auth đơn giản)
- Test: `test_stats_endpoint()` — response 200 JSON có `{total, today, top_ips, uploads}`
- Test: nếu STATS_KEY set thì thiếu key → 403
- Implement: đọc `stats.read()` → tổng hợp: total, today, top 10 IP, số POST dấu `uploads`
- Chặn bằng `STATS_KEY` env (nếu set, yêu cầu `?key=`)

### Task 4: Trang `/stats.html`
- Test: không cần unit test (static) — verify curl 200 + JSON fetch
- Implement: HTML + fetch(`/stats`) hiển thị: tổng requests, hôm nay, top IP bảng, uploads
- Style Bootstrap, tối giản, không cần đẹp — đủ đọc

### Task 5: Dọn dữ liệu
- Test: `test_stats_cleanup()` — tạo record cũ (timestamp 31 ngày trước) → cleanup → hết
- Implement: gọi `Stats.cleanup()` sau mỗi `record()` (hoặc 1 watchdog mỗi giờ)
- Không lưu dữ liệu vĩnh viễn — cam kết tối đa 30 ngày

### Task 6: Rebuild + deploy + verify E2E
- `pytest -q` → ~30 tests pass
- `docker compose build && up -d`
- Test: curl `/`, `/process` (dummy), `/stats` → có dữ liệu
- `chmod 644`, push Docker Hub tag `v1.4.0`, commit + push GitHub
- Không thêm dashboard card (app đã có)

---

## Tests / Validation

- `cd ~/apps/excel-keyword-filter && . .venv/bin/activate && python -m pytest -q` → pass
- E2E qua HTTP: gọi vài request từ IP khác nhau (dùng curl `-H "CF-Connecting-IP: 1.2.3.4"` để giả lập Cloudflare) → `/stats` hiện IP đó
- Verify file log sinh ra trong container: `docker exec excel-keyword-filter cat /app/logs/access.jsonl`

## Risks / Tradeoffs / Open Questions

1. **IP thật khi qua Cloudflare**: dùng `CF-Connecting-IP` — tin header này (CF luôn set). Nếu user truy cập trực tiếp tailnet (100.95.x.x) thì fallback client.host → IP tailnet thật. OK.
2. **Không có country geo**: muốn biết quốc gia/khu vực cần thêm GeoIP (maxmind), nặng hơn. Đề xuất bỏ qua v1 (chỉ IP + thời gian), thêm sau nếu cần.
3. **Stats public**: mặc định `GET /stats` public có thể lộ IP người dùng. Đề xuất **set STATS_KEY mặc định ON** (VD `excel_stats_2026`) trong compose env để chặn truy cập trái phép; anh có thể tắt nếu muốn.
4. **Match_count**: hiện tại log không bắt được số dòng match (body trả về lúc response). Có thể parse nhưng thêm phức tạp. Hỏi anh: **cần biết "mỗi lần upload khớp được bao nhiêu dòng" không?** Nếu cần thì OK, thêm đọc body.
5. **Retention**: 30 ngày — anh muốn 7 ngày, 30 ngày, hay 90 ngày?
6. **Stats theo URL trực tiếp hay qua dashboard?** Trang `/stats.html` trong app, or thêm card trên dashboard.8biz.net trỏ tới? (Anh quyết — mặc định làm trang trong app, không thêm dashboard card vì đây là tính năng của app cũ.)

---

## Next Steps (sau khi anh duyệt)

1. Xác nhận các lựa chọn: STATS_KEY (bật/tắt), có match_count không, retention bao lâu.
2. Implement theo Task 1-6.
3. Deploy & verify.

## References

- `backend/main.py`, `backend/excel_processor.py`, `tests/test_app.py` (hiện tại)
- `frontend/index.html`, `frontend/script.js`, `frontend/style.css`