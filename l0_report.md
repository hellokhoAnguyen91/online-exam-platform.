The following is a <SYSTEM_MESSAGE> not actually sent by the user. It is provided by the system as important information to pay attention to.

<SYSTEM_MESSAGE>
[Message] timestamp=2026-09-30T11:59:44Z sender=25ad2216-7b61-4039-aecf-da036279be82 priority=MESSAGE_PRIORITY_HIGH content=# BÁO CÁO KIỂM TRA TOÀN DIỆN CODEBASE — Online Exam Platform HCMUTE

## Tổng quan hồ sơ kiểm tra

Tôi đã đọc toàn bộ các file: [`main.py`](file:///C:/Users/WIN%2010/.gemini/antigravity/scratch/online-exam-platform/main.py) (2329 dòng), [`models.py`](file:///C:/Users/WIN%2010/.gemini/antigravity/scratch/online-exam-platform/models.py), [`database.py`](file:///C:/Users/WIN%2010/.gemini/antigravity/scratch/online-exam-platform/database.py), [`docx_parser.py`](file:///C:/Users/WIN%2010/.gemini/antigravity/scratch/online-exam-platform/docx_parser.py) (955 dòng), [`pdf_export.py`](file:///C:/Users/WIN%2010/.gemini/antigravity/scratch/online-exam-platform/pdf_export.py) (889 dòng), [`excel_export.py`](file:///C:/Users/WIN%2010/.gemini/antigravity/scratch/online-exam-platform/excel_export.py), [`student_parser.py`](file:///C:/Users/WIN%2010/.gemini/antigravity/scratch/online-exam-platform/student_parser.py), và [`static/index.html`](file:///C:/Users/WIN%2010/.gemini/antigravity/scratch/online-exam-platform/static/index.html) (4213 dòng).

---

## 🔴 P0 — Lỗi nghiêm trọng (Critical / Data Loss / Security)

---

### [P0-1] Hardcode SECRET_KEY JWT — Lộ bảo mật toàn hệ thống
- **File:** [`main.py`](file:///C:/Users/WIN%2010/.gemini/antigravity/scratch/online-exam-platform/main.py#L33-L35), dòng 33
- **Code:**
```python
SECRET_KEY = "super-secret-key-for-exam-platform-change-in-production"
```
- **Mức độ:** P0 — Bất kỳ ai clone repo hoặc xem file đều có thể ký JWT giả mạo để đăng nhập với bất kỳ tài khoản nào kể cả admin.
- **Khuyến nghị:** Đọc từ biến môi trường: `SECRET_KEY = os.environ.get("SECRET_KEY")` và raise lỗi nếu thiếu; Render hỗ trợ Secret Files và Environment Variables.

---

### [P0-2] Mật khẩu mặc định nguy hiểm — Bất kỳ sinh viên nào có thể đăng nhập bằng "123456"
- **File:** [`main.py`](file:///C:/Users/WIN%2010/.gemini/antigravity/scratch/online-exam-platform/main.py#L276-L277), dòng 276–277
- **Code:**
```python
if not is_valid and pwd_input in ["123456", "12345678", "hcmute"]:
    is_valid = True
```
- **Mức độ:** P0 — Bất kỳ sinh viên nào biết MSSV của người khác đều có thể đăng nhập bằng "123456" và gian lận (nộp bài thay, xem đề của người khác).
- **Khuyến nghị:** Xóa hoàn toàn fallback password này. Nếu muốn hỗ trợ reset, chỉ admin mới được phép thực hiện qua endpoint riêng.

---

### [P0-3] Race Condition khi tạo ExamResult — Duplicate session có thể xảy ra
- **File:** [`main.py`](file:///C:/Users/WIN%2010/.gemini/antigravity/scratch/online-exam-platform/main.py#L1807-L1866), dòng 1807–1866
- **Vấn đề:** Đoạn code tạo `ExamResult` mới cho sinh viên có kiểm tra trùng lặp hai lần (dòng 1809 và 1858), nhưng giữa hai lần kiểm tra không có `SELECT ... FOR UPDATE` hay database-level lock. Với SQLite WAL mode và nhiều kết nối đồng thời, hai request F5 nhanh có thể tạo ra 2 bản ghi `ExamResult` cùng `(user_id, exam_id)`.
- Cơ chế phòng tránh hiện tại là `UniqueConstraint` + `try/except` ở dòng 1858 — đây là đúng nhưng chỉ an toàn khi rollback đúng. Trường hợp exception ở `db.commit()` sau đó `db.rollback()` rồi re-query là pattern đúng. **Tuy nhiên**, nếu commit thành công cho cả 2 transaction (WAL mode cho phép concurrent readers nhưng không concurrent writers), **constraint lỗi sẽ throw IntegrityError → 500** thay vì trả về session cũ.
- **Khuyến nghị:** Wrap toàn bộ đoạn tạo result trong `try/except IntegrityError` với SQLAlchemy và re-query; hoặc dùng `INSERT OR IGNORE` + re-fetch.

---

### [P0-4] Đáp án lộ qua API `/api/exam/review` khi exam chưa đóng
- **File:** [`main.py`](file:///C:/Users/WIN%2010/.gemini/antigravity/scratch/online-exam-platform/main.py#L2154-L2208), dòng 2154–2208
- **Vấn đề:**
```python
is_closed = (close_dt is None or now > close_dt)
can_show_answers = bool(exam.allow_review) and is_closed
```
Nếu `close_time` là `None` (exam không có thời gian đóng, chỉ active/inactive), `is_closed = True` → đáp án hiển thị ngay lập tức sau khi nộp trong khi kỳ thi vẫn đang diễn ra với các thí sinh khác.
- **Khuyến nghị:** Kiểm tra thêm `exam.is_active` — chỉ cho phép review khi `not exam.is_active`.

---

### [P0-5] `/api/exams` (public, không auth) — Lộ danh sách tất cả kỳ thi
- **File:** [`main.py`](file:///C:/Users/WIN%2010/.gemini/antigravity/scratch/online-exam-platform/main.py#L2243-L2257), dòng 2243–2257
- **Vấn đề:** Endpoint `/api/exams` không yêu cầu xác thực (không có `Depends(get_current_user)`), trả về tên, mô tả, `num_questions`, `duration_minutes` của tất cả kỳ thi (kể cả chưa active). Đây là thông tin nhạy cảm.
- **Khuyến nghị:** Thêm xác thực hoặc chỉ trả về exam đang active.

---

## 🔴 P1 — Lỗi logic nghiêm trọng (Business Logic)

---

### [P1-1] Tính điểm sai ở `auto_close_exam_result` — `correct_count` bao gồm cả partial credit
- **File:** [`main.py`](file:///C:/Users/WIN%2010/.gemini/antigravity/scratch/online-exam-platform/main.py#L1281-L1282), dòng 1281–1282
- **Vấn đề:** Trong `auto_close_exam_result`, `correct_count` chỉ tăng khi `is_c = True`. Nhưng `calculate_exam_score` cũng tính `correct_count` chỉ cho `is_correct == True`. Với câu multi-select `partial`, `is_c = False` nhưng `earned_score > 0` → câu partial không được đếm vào `correct_count` nhưng **lại được cộng điểm** → `correct_count` nhỏ hơn thực tế (hiển thị sai trên giao diện, PDF, Excel).
- **Khuyến nghị:** Tách `partial_count` riêng để hiển thị; hoặc thống nhất định nghĩa "đúng".

---

### [P1-2] Tính `mc_max_score`/`essay_max_score` hiển thị sai trong Review
- **File:** [`main.py`](file:///C:/Users/WIN%2010/.gemini/antigravity/scratch/online-exam-platform/main.py#L2215-L2219), dòng 2215–2219
- **Vấn đề:** Trong `/api/exam/review`, `mc_weight` và `essay_weight` được tính từ `sum(q.get("score_weight", 1.0))` của raw audit — đây là tổng điểm thô chưa scale, **không phải điểm trên thang 10**. Trong khi đó, điểm thực sự đã được scale qua `calculate_exam_score`. Kết quả: tỉ lệ hiển thị `mc_score / mc_max_score` sai hoàn toàn.
- **Khuyến nghị:** Gọi `calculate_exam_score` thống nhất, sử dụng kết quả trả về thay vì tính tay.

---

### [P1-3] Backup không bao gồm `question_type`, `score_weight`, `option_e`, `option_f`
- **File:** [`main.py`](file:///C:/Users/WIN%2010/.gemini/antigravity/scratch/online-exam-platform/main.py#L485-L493), dòng 485–493
- **Vấn đề:** Khi export backup, `q_data` chỉ export `option_a...d`, `correct_option`, `explanation` — thiếu `question_type`, `score_weight`, `option_e`, `option_f`. Khi restore, câu hỏi đa lựa chọn và câu tự luận bị mất type.
- **Khuyến nghị:** Thêm tất cả trường vào backup và restore payload.

---

### [P1-4] `sync_expired_sessions` sửa `submit_time` của bài đã nộp — Xâm phạm dữ liệu gốc
- **File:** [`main.py`](file:///C:/Users/WIN%2010/.gemini/antigravity/scratch/online-exam-platform/main.py#L1339-L1351), dòng 1339–1351
- **Vấn đề:** Hàm này tự động sửa `submit_time` và `duration_seconds` của bài đã `submitted` nếu thời lượng vượt quá giới hạn. Nếu giáo viên thay đổi `duration_minutes` sau khi sinh viên đã nộp, `submit_time` cũ sẽ bị ghi đè → mất audit trail.
- **Khuyến nghị:** Chỉ điều chỉnh các session `in_progress`, không động vào `submitted`.

---

### [P1-5] `migrate_database` backfill `score` sai — `c_count = int(r_score)` thay vì đếm câu đúng
- **File:** [`models.py`](file:///C:/Users/WIN%2010/.gemini/antigravity/scratch/online-exam-platform/models.py#L211), dòng 211
- **Vấn đề:**
```python
c_count = int(r_score) if r_score is not None else 0
```
Biến `r_score` là điểm số (ví dụ 7.5), không phải số câu đúng. Migration sẽ đặt `correct_count = 7` cho bài điểm 7.5, hoàn toàn sai logic.
- **Khuyến nghị:** Backfill `correct_count` bằng cách parse `answers` và so sánh với `correct_option` của từng câu.

---

### [P1-6] `plus_opts` (multi-select type `+`) — Gán tất cả options là "correct" mà không check đáp án
- **File:** [`docx_parser.py`](file:///C:/Users/WIN%2010/.gemini/antigravity/scratch/online-exam-platform/docx_parser.py#L665), dòng 665
- **Vấn đề:**
```python
"correct_option": ",".join(opt_chars[:min(len(plus_opts), len(opt_chars))])
```
Tất cả `plus_opts` đều được đánh dấu là đáp án đúng (A,B,C,D,...). Đây là hành vi không hợp lý cho câu "chọn tất cả phương án đúng" — nếu có câu mà chỉ một số là đúng, parser sẽ đánh sai hoàn toàn.
- **Khuyến nghị:** Cần phát hiện đáp án đúng từ asterisk, underline, hoặc explicit "Đáp án:" tương tự các câu khác.

---

## 🟠 P2 — Lỗi giao diện và UX frontend

---

### [P2-1] Timer hardcode "30 phút" trong toast khi hết giờ
- **File:** [`static/index.html`](file:///C:/Users/WIN%2010/.gemini/antigravity/scratch/online-exam-platform/static/index.html#L2495), dòng 2495
- **Vấn đề:**
```javascript
showToast("⏰ ĐÃ HẾT GIỜ LÀM BÀI (30 phút)! ...", "warning");
```
Thông báo hardcode "30 phút" bất kể kỳ thi được cấu hình bao nhiêu phút.
- **Khuyến nghị:** Dùng `Math.round(examEndTime_raw/60)` hoặc đọc từ `data.duration_minutes`.

---

### [P2-2] `remainingSeconds` không được định nghĩa — ReferenceError khi submit thất bại
- **File:** [`static/index.html`](file:///C:/Users/WIN%2010/.gemini/antigravity/scratch/online-exam-platform/static/index.html#L2576-L2602), dòng 2576, 2602
- **Vấn đề:**
```javascript
if(remainingSeconds > 0) startExamTimer(remainingSeconds);
```
Biến `remainingSeconds` không được khai báo trong scope global (chỉ có `examEndTime`). Khi submit thất bại và cần khởi động lại timer, đây là `ReferenceError` → timer không restart được, sinh viên không thể làm tiếp.
- **Khuyến nghị:** Lưu `time_left` ban đầu vào biến global hoặc tính lại từ `examEndTime`.

---

### [P2-3] Autosave không báo lỗi chi tiết khi network fail trong Essay
- **File:** [`static/index.html`](file:///C:/Users/WIN%2010/.gemini/antigravity/scratch/online-exam-platform/static/index.html#L2331-L2333), dòng 2331–2333
- **Vấn đề:** Khi autosave bắt lỗi, chỉ hiển thị chấm đỏ "Lỗi lưu!" nhưng không thông báo rõ ràng hay retry. Sinh viên có thể bỏ qua và mất bài tự luận nếu mất mạng dài.
- **Khuyến nghị:** Hiển thị toast cảnh báo và retry sau 5 giây; LocalStorage đã backup nhưng cần thông báo rõ hơn.

---

### [P2-4] Palette badge count không reset đúng khi re-render
- **File:** [`static/index.html`](file:///C:/Users/WIN%2010/.gemini/antigravity/scratch/online-exam-platform/static/index.html#L2222-L2242), dòng 2222, 2242
- **Vấn đề:** `palette-count-mc` và `palette-count-essay` được set hardcode thành "0/N đã làm" khi renderPalette, nhưng `updateProgress()` ở dòng 2284 được gọi ngay sau và cập nhật lại — chỉ đúng nếu `studentAnswers` đã được load. Tuy nhiên, nếu cache được load sau `renderPalette()` thì badge sẽ hiển thị sai.
- **Khuyến nghị:** Gọi `updateProgress()` sau khi load cache (dòng 1973), không trước.

---

## 🟠 P2 — Lỗi xử lý đầu vào / Parser

---

### [P2-5] Parser không xử lý ảnh nằm trong bảng chứa options
- **File:** [`docx_parser.py`](file:///C:/Users/WIN%2010/.gemini/antigravity/scratch/online-exam-platform/docx_parser.py#L908-L930), dòng 908–930
- **Vấn đề:** Khi option nằm trong bảng (`table_has_options`), code lấy `c_val = m_cell_opt.group(2).strip()` (text thuần) nhưng chỉ thêm ảnh từ `cell._tc` vào `c_val` dưới dạng `c_imgs` — **ảnh đặt trước text option trong cell** sẽ bị bỏ qua vì `paragraph_to_html` không được gọi cho các paragraph trong cell có table-option mode.
- **Khuyến nghị:** Gọi `paragraph_to_html` cho từng paragraph trong cell và ghép lại thay vì dùng `cell.text`.

---

### [P2-6] `student_parser.py` dùng fallback DOB `"2003-01-01"` khi parse thất bại — Tạo mật khẩu giống nhau
- **File:** [`student_parser.py`](file:///C:/Users/WIN%2010/.gemini/antigravity/scratch/online-exam-platform/student_parser.py#L51-L70), dòng 51, 70
- **Vấn đề:**
```python
if val is None or pd.isna(val):
    return "2003-01-01"
...
if not s or s.lower() == "nan":
    return "2003-01-01"
```
Nếu cột DOB bị bỏ trống hoặc không parse được, tất cả sinh viên đó sẽ có DOB = `"2003-01-01"` → cùng mật khẩu → có thể đăng nhập tài khoản người khác.
- **Khuyến nghị:** Trả về `None` hoặc MSSV làm mật khẩu mặc định; không dùng ngày giả.

---

### [P2-7] `ANSWER_REGEX` có thể match sai khi câu hỏi chứa "Chọn" hoặc "Key" trong nội dung
- **File:** [`docx_parser.py`](file:///C:/Users/WIN%2010/.gemini/antigravity/scratch/online-exam-platform/docx_parser.py#L515-L518), dòng 515–518
- **Vấn đề:** Regex `ANSWER_REGEX` khớp với "chọn" ở đầu hoặc trong câu hỏi (ví dụ "Chọn phát biểu đúng: A. ..."). Nếu câu hỏi chứa "Chọn ... A." ở đầu, parser sẽ hiểu nhầm đây là đáp án.
- **Vị trí tốt hơn:** Regex nên yêu cầu prefix rõ hơn và anchor ở đầu dòng.

---

## 🟡 P3 — Lỗi Runtime trên Render Linux / Hiệu năng

---

### [P3-1] Font Arial trên Render Linux — Có thể thiếu nếu fonts/ không được bundle đúng
- **File:** [`pdf_export.py`](file:///C:/Users/WIN%2010/.gemini/antigravity/scratch/online-exam-platform/pdf_export.py#L34-L42), dòng 34–42
- **Vấn đề:** Code tìm font theo thứ tự:
  1. `./fonts/arial.ttf` ✅ (có trong repo)
  2. `C:/Windows/Fonts` ❌ (Linux không có)
  3. `/usr/share/fonts/truetype/dejavu` ✅ (fallback Linux)
  
  Nhưng nếu Render không mount đúng thư mục `fonts/` (working directory khác với `__file__`), sẽ fail. Procfile dùng `uvicorn main:app` không chỉ định `--app-dir`.
- **Kiểm tra:** `os.path.dirname(os.path.abspath(__file__))` là đường dẫn tuyệt đối của `pdf_export.py`, không phải CWD — nên thường đúng. Nhưng nếu Render thay đổi CWD, cần log để confirm.
- **Khuyến nghị:** Thêm log khi `arial` = None sau vòng lặp tìm font để detect lỗi sớm.

---

### [P3-2] Batch ZIP 62 PDFs — Tất cả xử lý trong request thread, không async → timeout Render
- **File:** [`main.py`](file:///C:/Users/WIN%2010/.gemini/antigravity/scratch/online-exam-platform/main.py#L1623), dòng 1623; [`pdf_export.py`](file:///C:/Users/WIN%2010/.gemini/antigravity/scratch/online-exam-platform/pdf_export.py#L766-L889)
- **Vấn đề:** `generate_batch_exam_zip` chạy đồng bộ trong request thread (không `async`). Với 62 sinh viên, mỗi PDF có thể mất 1–3 giây → tổng 60–180 giây. Render free tier timeout HTTP request sau 30 giây.
- **Khuyến nghị:**
  - Dùng `asyncio.get_event_loop().run_in_executor(None, ...)` để chạy trên thread pool.
  - Hoặc dùng background task và cho admin download sau.
  - Cache `_OPTIMIZED_IMAGE_CACHE` (đã có) giúp nhiều nếu hình ảnh trùng nhau.

---

### [P3-3] `_OPTIMIZED_IMAGE_CACHE` là dict module-level — Memory leak dài hạn
- **File:** [`pdf_export.py`](file:///C:/Users/WIN%2010/.gemini/antigravity/scratch/online-exam-platform/pdf_export.py#L151), dòng 151
- **Vấn đề:** Cache ảnh không bao giờ bị xóa → server chạy lâu, nhiều kỳ thi với nhiều ảnh khác nhau → RAM tăng không ngừng. Trên Render free tier (512MB RAM), có thể OOM.
- **Khuyến nghị:** Dùng `functools.lru_cache` với `maxsize` hoặc dùng `cachetools.LRUCache`.

---

### [P3-4] `session_scanner` background task chạy mỗi 5 giây, query DB tất cả in_progress sessions
- **File:** [`main.py`](file:///C:/Users/WIN%2010/.gemini/antigravity/scratch/online-exam-platform/main.py#L2298-L2314), dòng 2298–2314
- **Vấn đề:** Vòng lặp chạy mỗi 5 giây, gọi `sync_expired_sessions` → query tất cả `in_progress` results, sau đó query từng exam, và query tất cả `submitted` results để cap duration. Với 62 sinh viên cùng thi, mỗi 5 giây có ~20–30 query → tổng ~1800–3600 query/phút → SQLite WAL bị áp lực.
- **Khuyến nghị:** Tăng interval lên 30–60 giây; thêm index trên `(exam_id, status)`.

---

### [P3-5] CORS `allow_origins=["*"]` với `allow_credentials=True` — Cấu hình sai và insecure
- **File:** [`main.py`](file:///C:/Users/WIN%2010/.gemini/antigravity/scratch/online-exam-platform/main.py#L43-L49), dòng 43–49
- **Vấn đề:** Theo CORS spec, kết hợp `allow_origins=["*"]` và `allow_credentials=True` là **invalid** — trình duyệt sẽ từ chối credential-bearing requests. FastAPI/Starlette sẽ raise `ValueError` hoặc silently fail.
- **Khuyến nghị:** Đặt `allow_origins` cụ thể (domain Render + localhost) thay vì `"*"`.

---

## 🟡 P3 — Vấn đề khác

---

### [P3-6] `ExamUpdate.open_time = None` không thể clear open_time (không phân biệt null vs not-provided)
- **File:** [`main.py`](file:///C:/Users/WIN%2010/.gemini/antigravity/scratch/online-exam-platform/main.py#L404-L405), dòng 404–405
- **Vấn đề:** Pydantic model dùng `Optional[str] = None` → không thể phân biệt "admin muốn xóa open_time" vs "admin không gửi trường này". Nếu admin xóa `open_time`, giá trị vẫn giữ nguyên.
- **Khuyến nghị:** Dùng sentinel (`UNSET`), hoặc dùng Pydantic v2 `model_fields_set`.

---

### [P3-7] `NumberedCanvas.draw_page_decorations` dùng `self._pageNumber` — Attribute không tồn tại trong ReportLab
- **File:** [`pdf_export.py`](file:///C:/Users/WIN%2010/.gemini/antigravity/scratch/online-exam-platform/pdf_export.py#L145), dòng 145
- **Vấn đề:**
```python
page_text = f"Trang {self._pageNumber} / {page_count}"
```
Attribute `_pageNumber` không được định nghĩa trong `__init__` của `NumberedCanvas`. Trong ReportLab, page number được track khác — có thể là `self._page_number` hoặc index từ `_saved_page_states`. Có thể gây `AttributeError` trên một số phiên bản ReportLab.
- **Khuyến nghị:** Dùng `index + 1` khi iterate qua `_saved_page_states`.

---

### [P3-8] `Procfile` không pin Python worker/số lượng process
- **File:** [`Procfile`](file:///C:/Users/WIN%2010/.gemini/antigravity/scratch/online-exam-platform/Procfile)
- **Vấn đề:** `web: uvicorn main:app --host 0.0.0.0 --port $PORT` — chỉ 1 worker, không có `--workers`. Với batch PDF export nặng, 1 worker bị block toàn bộ, các request khác (autosave của sinh viên) không xử lý được.
- **Khuyến nghị:** `uvicorn main:app --host 0.0.0.0 --port $PORT --workers 2` hoặc dùng Gunicorn với UvicornWorker.

---

### [P3-9] `requirements.txt` thiếu version pin — Dễ break khi update
- **File:** [`requirements.txt`](file:///C:/Users/WIN%2010/.gemini/antigravity/scratch/online-exam-platform/requirements.txt)
- **Vấn đề:** Chỉ `bcrypt==4.0.1` có pin version; các lib khác (fastapi, sqlalchemy, reportlab, pillow) không pin → deploy mới có thể nhận phiên bản không tương thích.
- **Khuyến nghị:** Pin tất cả dependency (`pip freeze > requirements.txt`).

---

## 📋 Tóm tắt ưu tiên

| # | File | Dòng | Mức | Vấn đề |
|---|------|------|-----|--------|
| P0-1 | main.py | 33 | 🔴 P0 | Hardcode SECRET_KEY JWT |
| P0-2 | main.py | 276 | 🔴 P0 | Mật khẩu mặc định "123456" cho mọi sinh viên |
| P0-3 | main.py | 1807–1866 | 🔴 P0 | Race condition tạo duplicate ExamResult |
| P0-4 | main.py | 2154–2208 | 🔴 P0 | Lộ đáp án ngay sau nộp khi close_time=None |
| P0-5 | main.py | 2243–2257 | 🔴 P0 | /api/exams public không auth |
| P1-1 | main.py | 1281 | 🟠 P1 | correct_count không bao gồm partial credit |
| P1-2 | main.py | 2215–2219 | 🟠 P1 | mc_weight/essay_weight hiển thị sai (raw vs scaled) |
| P1-3 | main.py | 485–493 | 🟠 P1 | Backup thiếu question_type, score_weight, option E/F |
| P1-4 | main.py | 1339–1351 | 🟠 P1 | sync_expired_sessions sửa submit_time bài đã nộp |
| P1-5 | models.py | 211 | 🟠 P1 | Migration backfill correct_count = điểm số (sai logic) |
| P1-6 | docx_parser.py | 665 | 🟠 P1 | plus_opts gán tất cả là đáp án đúng |
| P2-1 | index.html | 2495 | 🟡 P2 | Toast timer hardcode "30 phút" |
| P2-2 | index.html | 2576 | 🟡 P2 | remainingSeconds không được định nghĩa → ReferenceError |
| P2-3 | index.html | 2331 | 🟡 P2 | Autosave fail không retry, không warn rõ |
| P2-4 | index.html | 2222–2242 | 🟡 P2 | Palette badge reset trước khi load cache |
| P2-5 | docx_parser.py | 908–930 | 🟡 P2 | Ảnh trong cell bảng chứa option bị mất |
| P2-6 | student_parser.py | 51 | 🟡 P2 | DOB fallback "2003-01-01" → mật khẩu giống nhau |
| P2-7 | docx_parser.py | 515 | 🟡 P2 | ANSWER_REGEX match nhầm nội dung câu hỏi |
| P3-1 | pdf_export.py | 34 | 🟢 P3 | Thiếu log khi font Arial không tìm được |
| P3-2 | main.py | 1623 | 🟢 P3 | Batch ZIP 62 PDFs blocking → timeout Render |
| P3-3 | pdf_export.py | 151 | 🟢 P3 | Image cache không giới hạn → memory leak |
| P3-4 | main.py | 2298 | 🟢 P3 | Session scanner 5s/lần → DB pressure |
| P3-5 | main.py | 43–49 | 🟢 P3 | CORS origins=* + credentials=True là invalid |
| P3-6 | main.py | 404 | 🟢 P3 | Không thể clear open_time/close_time về NULL |
| P3-7 | pdf_export.py | 145 | 🟢 P3 | _pageNumber không tồn tại → AttributeError tiềm ẩn |
| P3-8 | Procfile | - | 🟢 P3 | 1 worker duy nhất, blocked khi export |
| P3-9 | requirements.txt | - | 🟢 P3 | Thiếu version pin |

---

## Remaining Questions & Gaps

1. **Chưa kiểm tra `excel_export.py` chi tiết** (chỉ đọc 100 dòng đầu) — cần xem phần `generate_batch_excel_zip` và `generate_candidate_audit_excel` để tìm thêm lỗi tiềm ẩn (streaming vs in-memory, row height cố định khi essay dài).

2. **Chưa test thực tế trên Linux** — font fallback DejaVu có thể có vấn đề với tiếng Việt (thiếu glyph `ệ`, `ử`, `ổ`...). Cần chạy trực tiếp trên Ubuntu để kiểm tra PDF output.

3. **Chưa kiểm tra `docx_parser.py` với file Word thực tế** — các edge case như: câu hỏi có 2 hình ảnh liền nhau, hình ảnh nằm trong equation, bảng lồng trong bảng.

4. **Race condition P0-3** — cần stress test với `run_stress_test.py` (đã có trong repo) để xác nhận thực sự xảy ra hay không.

5. **Chưa kiểm tra `student_parser.py` với file Excel có header đặc biệt** (merge cell, nhiều hàng header) — cần test case phức tạp hơn.

6. **Render free tier sleeps after 15 minutes inactivity** — `session_scanner` sẽ bị dừng khi server ngủ, các session `in_progress` sẽ không được auto-close → sinh viên quay lại sẽ thấy session cũ còn trong khi thời gian đã hết. Đây là vấn đề kiến trúc cần xem xét.

</SYSTEM_MESSAGE>