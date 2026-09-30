# KẾ HOẠCH KIỂM THỬ TOÀN DIỆN (TEST PLAN)
**Hệ thống Thi Trực tuyến HCMUTE**

Tài liệu này xác định mục tiêu kiểm thử, phạm vi, danh mục các bộ kiểm thử tự động hiện có, các kịch bản kiểm thử trọng yếu (Test Scenarios), và quy trình kiểm thử trước khi đưa hệ thống vào vận hành thực tế.

---

## 1. Mục tiêu & Phạm vi Kiểm thử (Objectives & Scope)

### 1.1. Mục tiêu
- Đảm bảo tính toàn vẹn dữ liệu bài làm của thí sinh khi có sự cố mạng hoặc tải lại trang.
- Đảm bảo tính chính xác tuyệt đối trong việc chấm điểm trắc nghiệm đơn, trắc nghiệm nhiều đáp án (Partial Credit) và tự luận.
- Đảm bảo an toàn bảo mật, chống mạo danh thí sinh, chống lộ đề thi và đáp án trước khi kết thúc ca thi.
- Đảm bảo độ tin cậy của các bộ bóc tách file Word đề thi và danh sách sinh viên Excel với mọi biến thể định dạng.
- Đảm bảo tính khả dụng và hiệu năng ổn định trên máy chủ đám mây (Render Linux Container).

### 1.2. Phạm vi kiểm thử
- **Đã bao gồm**:
  * Kiểm thử đơn vị (Unit Tests): Bóc tách Word `.docx`, bóc tách Excel `.xlsx`/`.csv`.
  * Kiểm thử tích hợp (Integration Tests): Đăng nhập, phân phối đề, lưu tự động, nộp bài, chấm điểm tự luận, đối chiếu, xuất file.
  * Kiểm thử bảo mật & Chống gian lận (Security & Adversarial Verification).
  * Kiểm thử tương tranh (Concurrency & Stress Testing).
  * Kiểm thử định dạng xuất file (PDF ReportLab, Excel OpenPyXL, đóng gói ZIP).
- **Không bao gồm**:
  * Chỉnh sửa cấu trúc phần cứng máy chủ.
  * Tự ý can thiệp làm sai lệch dữ liệu sản xuất đang chạy.

---

## 2. Danh mục 18 Bộ Kiểm thử Hiện có (Test Inventory)

Trong thư mục dự án đã có sẵn 18 bộ test tự động được xây dựng bằng Python unittest và pytest/fastapi-testclient:

| STT | File Kiểm thử | Loại kiểm thử | Mục đích kiểm thử chính | Tác động DB |
| :-: | :--- | :--- | :--- | :---: |
| 1 | `test_student_parser.py` | Unit Test | Bóc tách Excel/CSV danh sách SV: cột họ/tên lót, ngày sinh nhiều định dạng, số 0 đầu MSSV | Không |
| 2 | `test_docx_parser.py` | Unit Test | Bóc tách câu hỏi Word: ảnh inline/floating, công thức toán OMML, bảng, gạch chân/in đậm/dấu sao | Không |
| 3 | `test_section_headers.py`| Unit Test | Bóc tách nhận diện tiêu đề Phần I (Trắc nghiệm) và Phần II (Tự luận) trong file Word | Không |
| 4 | `test_parser_3cases.py` | Unit Test | Kiểm tra 3 ca đề thi thực tế của giảng viên: Multi-select dấu cộng `+`, Tự luận `(X điểm)`, Trắc nghiệm đơn | Không |
| 5 | `test_parse.py` | Unit Test | Kiểm tra nhanh cấu trúc paragraph và table parser | Không |
| 6 | `test_parse2.py` | Unit Test | Kiểm tra khả năng trích xuất hình ảnh từ khối drawing của OpenXML | Không |
| 7 | `test_edge_cases.py` | Integration | Kiểm tra các trường hợp biên: xóa kỳ thi đang active (bị chặn), đăng nhập bằng ngày sinh, chặn trước giờ mở | Có (tạo/xóa mock) |
| 8 | `test_edge_case_sections.py` | Integration | Kiểm tra bốc đề thi khi ngân hàng câu hỏi chỉ có trắc nghiệm hoặc chỉ có tự luận | Có (tạo/xóa mock) |
| 9 | `test_backend.py` | Smoke Test | Kiểm tra luồng cơ bản: login admin -> config kỳ thi -> xuất báo cáo | Có (tạo/xóa mock) |
| 10 | `test_api_suite.py` | Integration | Kiểm tra 19 API endpoints cốt lõi của hệ thống từ upload đề, sinh viên, thi, audit, backup/restore | Có (tạo/xóa mock) |
| 11 | `test_comprehensive_audit.py` | Full Audit | Bộ kiểm thử toàn diện 21 tiêu chí: P0 bảo mật, chặn mật khẩu yếu, giấu đáp án, backup/restore, PDF/ZIP | Có (dọn dẹp ở finally) |
| 12 | `test_adversarial_verification.py`| Security | Tấn công thử nghiệm: MSSV có tiền tố `sv`, multi-select payload dạng list, xem đáp án lậu khi thi đang mở | Có (tạo/xóa mock) |
| 13 | `test_pending_grading_logic.py` | Integration | Kiểm tra logic chờ chấm điểm: câu tự luận có bài làm, câu trắc nghiệm chưa có đáp án, câu bỏ trắng tự 0đ | Có (tạo/xóa mock) |
| 14 | `test_review_fixes.py` | Integration | Kiểm tra trang xem lại bài thi: giấu đáp án khi ca thi mở, mở đáp án khi ca thi đã đóng | Có (tạo/xóa mock) |
| 15 | `test_mc_essay_separation.py` | Integration | Kiểm tra thuật toán bốc đề: xáo trộn trắc nghiệm nhưng cố định 100% tự luận | Có (tạo/xóa mock) |
| 16 | `test_full_section_flow.py` | End-to-End | Luồng làm bài hoàn chỉnh 2 phần: Trắc nghiệm + Tự luận kèm chấm điểm giáo viên và xuất PDF | Có (tạo/xóa mock) |
| 17 | `test_comprehensive_sections_v2.py` | End-to-End | Luồng kiểm thử chuyên sâu phân tách 2 phần theo chuẩn đào tạo HCMUTE (tỷ lệ 7.0/3.0) | Có (tạo/xóa mock) |
| 18 | `test_new_crud_and_parser.py` | Integration | Kiểm tra CRUD kỳ thi, câu hỏi, sinh viên và tính năng Reset bài thi cho thí sinh gặp sự cố | Có (tạo/xóa mock) |

---

## 3. Ma trận Kịch bản Kiểm thử Trọng yếu (Test Matrix & Scenarios)

### Kịch bản 1: Xác thực & Bảo mật (Authentication & Security)
- **TC-SEC-01**: Đăng nhập với MSSV hợp lệ và mật khẩu là MSSV -> **Kỳ vọng**: 200 OK, trả về Access Token.
- **TC-SEC-02**: Đăng nhập với MSSV hợp lệ và mật khẩu là Ngày sinh chuẩn hóa (`26/05/2005` hoặc `26-05-2005`) -> **Kỳ vọng**: 200 OK.
- **TC-SEC-03**: Đăng nhập thử với mật khẩu yếu mặc định (`123456`, `12345678`, `hcmute`) -> **Kỳ vọng**: 401 Unauthorized (Bị chặn).
- **TC-SEC-04**: Thí sinh gọi API `/api/exam/review` khi ca thi đang mở (`is_active = True`) -> **Kỳ vọng**: `correct = None`, `status = "hidden"`, `explanation = None`.
- **TC-SEC-05**: Thí sinh chưa đăng nhập gọi API `/api/exams` -> **Kỳ vọng**: 401 Unauthorized.
- **TC-SEC-06**: Thí sinh đã đăng nhập gọi API `/api/exams` -> **Kỳ vọng**: 200 OK nhưng chỉ thấy các kỳ thi đang `is_active = True`.

### Kịch bản 2: Bóc tách Đề thi & Danh sách Lớp (Parsers)
- **TC-PAR-01**: Tải lên file Word câu hỏi có đáp án gạch chân, in đậm, asterisk `*` -> **Kỳ vọng**: Nhận diện đúng phương án A/B/C/D.
- **TC-PAR-02**: Tải lên file Word câu hỏi có nhúng hình ảnh trực tiếp và bảng biểu -> **Kỳ vọng**: Hình ảnh được chuyển thành Base64 đầy đủ, bảng biểu giữ nguyên cấu trúc HTML.
- **TC-PAR-03**: Tải lên file Word có câu tự luận ghi rõ `(3 điểm)` -> **Kỳ vọng**: `question_type = "essay"`, `score_weight = 3.0`.
- **TC-PAR-04**: Tải lên file Excel danh sách thí sinh có 2 hàng tiêu đề (Merged Header), cột "Họ và tên lót" và "Tên" riêng biệt -> **Kỳ vọng**: Ghép đúng họ tên, bảo toàn số 0 đầu trong MSSV.

### Kịch bản 3: Phân phối Đề thi & Quy tắc Làm bài (Exam Delivery & Auto-save)
- **TC-DLV-01**: Thí sinh bắt đầu làm bài -> **Kỳ vọng**: Phần I (Trắc nghiệm) được xáo trộn ngẫu nhiên; Phần II (Tự luận) lấy đủ 100% câu và giữ nguyên thứ tự.
- **TC-DLV-02**: Thí sinh F5 trình duyệt -> **Kỳ vọng**: Nạp lại đúng đề cũ, câu trả lời đã chọn được giữ nguyên, đồng hồ đếm ngược tiếp tục tính thời gian thực.
- **TC-DLV-03**: Thí sinh chọn đáp án -> **Kỳ vọng**: Gửi lưu lên server thành công; đồng thời ghi vào `localStorage`.

### Kịch bản 4: Chấm điểm & Quy đổi Thang điểm (Scoring Engine)
- **TC-SCR-01**: Câu hỏi chọn nhiều đáp án có 3 đáp án đúng (A, C, E), thí sinh chọn (A, C, B) -> **Kỳ vọng**: Số đúng = 2, số sai = 1 -> Tỷ lệ = $(2 - 1)/3 = 0.3333$ (Partial Credit).
- **TC-SCR-02**: Thí sinh làm bài tự luận -> **Kỳ vọng**: Trạng thái chuyển sang `pending_grading`, điểm tổng toàn bài là `None` cho tới khi giáo viên chấm bài.
- **TC-SCR-03**: Giáo viên nhập điểm tự luận vượt quá trọng số câu -> **Kỳ vọng**: Hệ thống từ chối với lỗi 400 Bad Request.
- **TC-SCR-04**: Giáo viên chấm xong toàn bộ câu tự luận -> **Kỳ vọng**: Điểm tổng kết tự động tính theo tỷ lệ chuẩn 7.0 Trắc nghiệm / 3.0 Tự luận trên thang 10.

### Kịch bản 5: Quản lý Phiên thi & Hết giờ (Timeouts & Concurrency)
- **TC-TIM-01**: Đồng hồ đếm ngược về 0 -> **Kỳ vọng**: Khóa toàn bộ input, tự động nộp bài và hiển thị kết quả.
- **TC-TIM-02**: Thí sinh bỏ thi giữa chừng -> **Kỳ vọng**: Sau khi quá thời lượng làm bài, tiến trình ngầm `session_scanner` tự động đóng bài thi và tính điểm các câu đã làm.
- **TC-TIM-03**: Bài thi đã nộp thành công -> **Kỳ vọng**: `session_scanner` tuyệt đối không chỉnh sửa `submit_time` hay `duration_seconds` của bài thi đó.

### Kịch bản 6: Xuất bản Hồ sơ (Export & Reporting)
- **TC-EXP-01**: Giáo viên bấm tải ZIP tất cả bài thi PDF -> **Kỳ vọng**: File ZIP tải về chứa 100% file `.pdf` ở thư mục gốc, font tiếng Việt hiển thị sắc nét không lỗi font.
- **TC-EXP-02**: Giáo viên bấm xuất bảng điểm Excel -> **Kỳ vọng**: File ZIP chứa bảng điểm tổng hợp của lớp và file Excel chi tiết của từng thí sinh.

---

## 4. Hướng dẫn Chạy Kiểm thử (Test Execution Guide)

### 4.1. Chạy các bài kiểm thử đơn vị an toàn (Không chạm DB)
```powershell
# Chạy bộ test kiểm tra parser danh sách sinh viên Excel/CSV
.\venv\Scripts\python.exe -m unittest test_student_parser.py

# Chạy bộ test kiểm tra parser đề thi Word (Ảnh, Bảng, OMML)
.\venv\Scripts\python.exe test_docx_parser.py

# Chạy bộ test kiểm tra các trường hợp 3 dạng đề thi thực tế
.\venv\Scripts\python.exe test_parser_3cases.py

# Chạy bộ test kiểm tra phân đoạn Section Headers
.\venv\Scripts\python.exe test_section_headers.py
```

### 4.2. Chạy toàn bộ các bài kiểm thử tích hợp (Integration Tests)
```powershell
# Chạy bộ kiểm thử toàn diện khắt khe nhất
.\venv\Scripts\python.exe test_comprehensive_audit.py

# Chạy kiểm thử bảo mật & chống gian lận
.\venv\Scripts\python.exe test_adversarial_verification.py

# Chạy kiểm thử luồng làm bài 2 phần và chấm điểm tự luận
.\venv\Scripts\python.exe test_full_section_flow.py

# Chạy kiểm thử logic chờ chấm điểm
.\venv\Scripts\python.exe test_pending_grading_logic.py
```

### 4.3. Tiêu chí Đánh giá (Pass/Fail Criteria)
- **ĐẠT (PASS)**: 100% các câu lệnh `assert` thành công, không phát sinh lỗi ngoại lệ `Exception`, `status_code` HTTP trả về đúng chuẩn thiết kế (200, 400, 401, 403, 404).
- **KHÔNG ĐẠT (FAIL)**: Phát sinh lỗi logic, lộ đáp án, sai lệch công thức điểm hoặc vỡ font tài liệu.
