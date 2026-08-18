# Excel Keyword Filter — Plan cải thiện format & print layout

> **Trạng thái:** Chỉ lên plan, chưa code. Marshal yêu cầu dừng code để lên plan trước, cải thiện.
> **Sản phẩm hiện tại:** `TRUE_Result.xlsx` với STT `=ROW()-1`, dòng Total (merge A→G, italic+bold), border, auto width.

---

## Goal

Cải thiện định dạng + print settings cho file `TRUE_Result.xlsx` theo yêu cầu mới của Marshal:

1. Page layout: **Landscape**
2. Print titles: **repeat header row 1** ở mỗi trang in
3. **Wrap text** toàn sheet (giữ nội dung không bị cắt)
4. **Header (row 1)**: tô nền **Đen**, chữ **trắng**, **viết hoa**, **middle align + center** (cả 8 cột)
5. **Cột widths** cố định: A=3, B=11, C=20, D=15, E=35, F=15, G=45, H=8
6. **Cột B**: bỏ phần thời gian — `05-Thg 01-26 08:43 SA` → `05-Thg 01-26` (giữ date, bỏ HH:mm AM/PM "SA/CH")
7. **Toàn sheet (trừ dòng Total)**: middle align + center ngang
8. **Dòng Total**: giữ nguyên định dạng hiện tại (italic + bold, align right, merge A→G, H=SUM)
9. **Row height**: tự vừa chữ, không bị cắt — wrap text + auto height khi in

---

## Current Context / Assumptions

- App: `~/apps/excel-keyword-filter/`
- Backend core: `backend/excel_processor.py` (`process_excel()` + helpers)
- Tests: `tests/test_app.py` (20 tests, hiện pass)
- Kiến trúc hiện tại:
  - `_copy_row()` / `_copy_cell()` copy giá trị + style gốc
  - `_auto_width(ws, last_row)` tự tính width
  - `_apply_borders(ws, last_row)` border thin khắp
  - Header row 1 hiện chỉ có giá trị; Total row được style riêng
- Deploy: Docker compose, image `thn05/excel-keyword-filter`, repo GitHub `lekmedev/excel-keyword-filter`
- **Quy tắc deploy**: bind 0.0.0.0, thêm vào dashboard nếu là service mới (không đổi — app này đã có card)

---

## Proposed Approach

**1. Header row 1 dùng style mới (đè lên style copy từ file gốc):**
- `Font(bold=True, color="FFFFFF")` — chữ trắng
- `PatternFill(fill_type="solid", fgColor="000000")` — nền đen
- `Alignment(horizontal="center", vertical="center", wrap_text=True)` — middle+center
- Giá trị header viết hoa (`upper()`)
- Bỏ qua việc merge cũ ở header nếu có (chỉ áp dụng cho 8 cột kết quả)

**2. Widths cột cố định** — thay `_auto_width()` bằng bảng hằng số (cho 8 cột kết quả):
```
COL_WIDTHS = {1: 3, 2: 11, 3: 20, 4: 15, 5: 35, 6: 15, 7: 45, 8: 8}
```
không dựa vào nội dung; vẫn set min ≥ 1 để không lỗi.

**3. Cột B (out_col=2)** — xử lý date: giá trị dạng `"05-Thg 01-26 08:43 SA"` → `"05-Thg 01-26"`.
- Regex: `^(.*?)\s+\d{1,2}:\d{2}(\s*(SA|CH))?$` → giữ phần date
- Chỉ áp dụng khi giá trị là chuỗi khớp pattern; giá trị khác giữ nguyên
- Nếu là cell datetime thật (date+time) → format `DD-MM-YY`? (open source CSV nên hầu như là text; check thêm)
- Khi copy: áp dụng **trước** khi `_copy_cell`

**4. Alignment toàn sheet (trừ Total row):**
- Mọi cell data (`row ≥ 2`, `row < total_row`): `Alignment(horizontal="center", vertical="middle", wrap_text=True)`
- Total row: giữ nguyên (align right + italic + bold)
- Border giữ nguyên thin khắp

**5. Wrap text toàn sheet:** set `wrap_text=True` cho mọi cell (kể cả Total). Kết hợp `_apply_wrap_text(ws, total_row)`.
- Row heights: openpyxl không tự tính — để Excel tự fit khi mở. Nhưng để in không cắt chữ, set `row[i].height = None`? Thực tế cần đặt height đủ hoặc dùng `ws.row_dimensions[r].customHeight` — sẽ thử nghiệm; với wrap text Excel tự fit khi in nếu không set height cố định. **Kế hoạch:** không set height cứng, chỉ để wrap → Excel tự fit khi in. Kiểm chứng thêm option `ws.sheet_properties.pageSetUpPr.fitToPage` cho in vừa.

**6. Print settings:**
```python
ws.page_setup.orientation = "landscape"
ws.page_setup.paperSize = ws.PAPERSIZE_A4  # (tùy chọn, hợp lý)
ws.print_title_rows = "1:1"               # repeat header row 1
ws.sheet_properties.pageSetUpPr.fitToPage = True
ws.page_setup.fitToWidth = 1
ws.page_setup.fitToHeight = 0             # chiều rộng vừa 1 trang, chiều cao linh hoạt
```

**7. Giữ nguyên:** STT formula, Total row layout, border, xóa cột, Quantity Replaced.

---

## Files Likely to Change

- `backend/excel_processor.py` — nhiều thay đổi nhất (style, widths, date strip, print setup, alignment)
- `tests/test_app.py` — thêm/sửa tests cho từng hành vi mới
- Không đổi: frontend, Dockerfile, compose (pure backend logic — image rebuild lại là đủ)

---

## Step-by-Step Plan (TDD)

### Task 1: Header style mới (đen/trắng/hoa/giữa)
- Test: `test_header_style_black_white_upper()` — kiểm tra row1: cell fill đen, font trắng, align center/middle, giá trị uppercase
- Implement: sau copy row1, ghi đè style từng cell 8 cột bằng hàm `_apply_header_style(ws)`
- Verify: pytest pass

### Task 2: Cột widths cố định
- Test: `test_fixed_column_widths()` — `ws.column_dimensions['A'].width == 3`, ... H == 8
- Implement: `_apply_fixed_widths(ws)` dùng `COL_WIDTHS`
- Thay `_auto_width()` không dùng nữa (hoặc giữ nhưng không gọi)
- Verify: pytest pass

### Task 3: Strip thời gian cột B
- Test: `test_column_b_time_stripped()` — data row2 col2 = "05-Thg 01-26" (bỏ "08:43 SA")
- Implement: hàm `_strip_time_from_date(value)` + áp dụng trong `_copy_row` khi out_col == 2
- Verify: pytest pass

### Task 4: Alignment toàn sheet (trừ Total) + wrap text
- Test: `test_data_cells_centered_wrapped()` — cell data row2..last: align center/middle, wrap True; Total row giữ align right
- Implement: `_apply_alignment(ws, total_row)` + `_apply_wrap_text(ws)`
- Verify: pytest pass

### Task 5: Print settings (landscape, print titles, fit)
- Test: `test_print_settings()` — orientation == "landscape", print_title_rows == "1:1", fitToWidth==1, fitToHeight==0
- Implement: `_apply_print_settings(ws)`
- Verify: pytest pass

### Task 6: Row height — wrap text tự fit
- Test: kiểm tra wrap_text flag đã bật; không set height cứng (hoặc height=auto tùy thử nghiệm)
- Implement: (nếu cần) `ws.row_dimensions[r].height = None`
- Verify: pytest pass

### Task 7: Rebuild + deploy + verify end-to-end
- `pytest -q` → ~26 tests pass
- `chmod 644`, `docker compose build`, `docker compose up -d`
- Test qua HTTP với sample CSV, tải file, mở verify (openpyxl) từng thuộc tính
- Push Docker Hub (tag v1.3.0), commit + push GitHub + tag
- NOT update dashboard (app đã có card; không cần đổi)

---

## Tests / Validation

- `cd ~/apps/excel-keyword-filter && . .venv/bin/activate && python -m pytest -q` → tất cả pass (20 cũ + ~6 mới)
- E2E: `curl -F file=@/tmp/sample.csv -F keywords=keyboard http://100.95.76.53:8000/process` → tải file → mở bằng openpyxl vérify:
  - orientation == "landscape", print_title_rows "1:1"
  - header fill đen font trắng uppercase
  - widths theo bảng
  - col B không còn giờ
  - data align center/middle, wrap True; Total align right
- Mở thủ công bằng LibreOffice/Excel (nếu có) để xem in preview

---

## Risks / Tradeoffs / Open Questions

- **Columns E/G widths lớn + wrap**: các cell 45-char với wrap sẽ cao; nếu không set height Excel tự fit khi in — OK.
- **Pattern date cột B**: regex chỉ bắt format `05-Thg 01-26 08:43 SA`; nếu có định dạng khác (VD `2026-01-05 14:30` hay English `05-Jan-26 08:43 AM`) pattern sẽ không khớp → giữ nguyên. Cần xác nhận Marshal dùng đúng format này (sample file thật).
- **Open question**: Header "Quantity Replaced" (H) — có viết hoa không? (Em sẽ `upper()` như yêu cầu "viết hoa toàn header" → `QUANTITY REPLACED`). Nếu anh muốn giữ "Quantity Replaced" đúng tên thì báo.
- **Open question**: total row wrap text? — giữ align right nhưng wrap vẫn bật cho nhất quán.
- **In vừa khổ giấy**: fitToWidth=1 sẽ co nhỏ để vừa 1 trang ngang; nếu nhiều dòng, height fitToHeight=0 cho phép nhiều trang dọc. Anh thấy OK chứ?