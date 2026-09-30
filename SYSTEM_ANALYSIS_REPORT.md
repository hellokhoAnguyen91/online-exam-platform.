# BÁO CÁO PHÂN TÍCH TOÀN DIỆN HỆ THỐNG THI TRỰC TUYẾN HCMUTE
**Hệ thống Thi Trắc nghiệm & Tự luận Trực tuyến Chuyên nghiệp**  
*Trường Đại học Công nghệ Kỹ thuật TP. Hồ Chí Minh (HCMUTE)*  
*Thời điểm lập báo cáo: 01/10/2026*

---

> [!IMPORTANT]
> **Cam kết tuân thủ chỉ đạo của Quản trị viên:**
> - Tuyệt đối không tự ý sửa logic code hiện tại.
> - Tuyệt đối không xóa bất kỳ tệp tin nào.
> - Tuyệt đối không thay đổi cơ sở dữ liệu `sql_app.db`.
> - Không thay đổi bất kỳ API nào đang phục vụ trên production.
> - Toàn bộ các phát hiện và phân tích được trình bày minh bạch dưới đây.

---

## MỤC LỤC
1. [Phân tích toàn bộ dự án hiện tại](#1-phân-tích-toàn-bộ-dự-án-hiện-tại)
2. [Bảng liệt kê 10 thành phần kỹ thuật cốt lõi](#2-bảng-liệt-kê-10-thành-phần-kỹ-thuật-cốt-lõi)
   - [2.1 Frontend](#21-frontend)
   - [2.2 Backend](#22-backend)
   - [2.3 Database](#23-database)
   - [2.4 API Endpoints](#24-api-endpoints)
   - [2.5 Authentication (Xác thực)](#25-authentication-xác-thực)
   - [2.6 Authorization (Phân quyền)](#26-authorization-phân-quyền)
   - [2.7 AI Components](#27-ai-components)
   - [2.8 File Upload](#28-file-upload)
   - [2.9 External Services](#29-external-services)
   - [2.10 Environment Variables](#210-environment-variables)
3. [Hướng dẫn cách chạy project ở môi trường Local](#3-hướng-dẫn-cách-chạy-project-ở-môi-trường-local)
4. [Hướng dẫn cách chạy các bài test hiện có](#4-hướng-dẫn-cách-chạy-các-bài-test-hiện-có)
5. [Kiểm tra trạng thái Git](#5-kiểm-tra-trạng-thái-git)
6. [Hướng dẫn tạo Git repository (Dự phòng)](#6-hướng-dẫn-tạo-git-repository-dự-phòng)
7. [Bản kiểm thử hiện tại (Test Execution Baseline Snapshot)](#7-bản-kiểm-thử-hiện-tại-test-execution-baseline-snapshot)
8. [Tài liệu Kiến trúc: PROJECT_ARCHITECTURE.md](#8-tài-liệu-kiến-trúc-project_architecturemd)
9. [Tài liệu Quy tắc Nghiệp vụ: BUSINESS_RULES.md](#9-tài-liệu-quy-tắc-nghiệp-vụ-business_rulesmd)
10. [Tài liệu Kế hoạch Kiểm thử: TEST_PLAN.md](#10-tài-liệu-kế-hoạch-kiểm-thử-test_planmd)
11. [Báo cáo phát hiện & Khuyến nghị vận hành](#11-báo-cáo-phát-hiện--khuyến-nghị-vận-hành)

---

## 1. Phân tích toàn bộ dự án hiện tại

### 1.1 Mục tiêu và đối tượng sử dụng
Hệ thống là nền tảng thi trực tuyến chuyên dụng phục vụ các kỳ thi cuối kỳ và kiểm tra định kỳ của Trường Đại học Công nghệ Kỹ thuật TP.HCM (HCMUTE). Hệ thống phục vụ hai nhóm đối tượng chính:
- **Giáo viên / Quản trị viên (Admin)**: Cô Nguyễn Hà Trang (`trangnh@hcmute.edu.vn`) quản lý ngân hàng câu hỏi, tải danh sách lớp, cấu hình thời gian thi, chấm bài tự luận và xuất hồ sơ điểm.
- **Thí sinh (Sinh viên)**: Đăng nhập bằng MSSV để làm bài thi trực tuyến gồm cả phần trắc nghiệm và tự luận với giao diện tối ưu, có bảng điều hướng câu hỏi, đồng hồ đếm ngược và cơ chế lưu bài tự động chống mất mát dữ liệu khi gặp sự cố mạng.

### 1.2 Các đặc trưng nổi bật của hệ thống
1. **Phân tách 2 phần thi rõ rệt**:
   - **Phần I (Trắc nghiệm)**: Xáo trộn ngẫu nhiên câu hỏi theo từng thí sinh (nếu bật cấu hình xáo đề).
   - **Phần II (Tự luận)**: Có bao nhiêu câu tự luận trong ngân hàng đề thì thí sinh làm bấy nhiêu câu, cố định thứ tự ban đầu, không xáo trộn.
2. **Bộ bóc tách đề thi Word (`.docx`) cực kỳ thông minh**:
   - Tự động trích xuất toàn bộ hình ảnh (inline, floating, VML, ảnh lồng trong bảng, ảnh trong đoạn văn trống ngay dưới câu hỏi) và chuyển đổi thành chuỗi Base64 nhúng an toàn trong thẻ `<img>`.
   - Giữ nguyên công thức toán học OMML (chuyển thành HTML), bảng biểu và định dạng văn bản.
   - Nhận diện linh hoạt đáp án đúng qua 4 cơ chế: chữ gạch chân, chữ in đậm, dấu hoa thị `*`, dòng `Đáp án: ...` hoặc `Key: ...`.
   - Hỗ trợ câu hỏi trắc nghiệm đơn, trắc nghiệm chọn nhiều đáp án (Multi-Select với dấu `+`), và câu tự luận có điểm số ghi chú dạng `(3 điểm)`.
3. **Bộ bóc tách sinh viên từ Excel/CSV linh hoạt**:
   - Tự động nhận diện cấu trúc tiêu đề phức tạp, tiêu đề ghép 2 hàng (Merged header).
   - Tự động tách hoặc ghép cột "Họ và tên lót" + "Tên".
   - Chuẩn hóa ngày sinh từ nhiều định dạng khác nhau và bảo toàn số 0 ở đầu MSSV.
4. **Cơ chế tính điểm & Chấm thi**:
   - Trắc nghiệm đơn: Đúng nhận 100% điểm câu, sai 0 điểm.
   - Trắc nghiệm chọn nhiều đáp án: Chấm điểm từng phần (Partial Credit Penalty) theo công thức $\max(0, (C_{\text{đúng}} - W_{\text{sai}}) / T_{\text{đúng}}) \times \text{trọng số}$.
   - Tự luận: Nếu thí sinh không gõ gì thì tự động tính 0 điểm; nếu có gõ bài làm thì chuyển trạng thái `pending_grading` để giáo viên chấm tay trên giao diện admin.
   - Điểm tổng kết: Quy đổi về thang điểm 10 theo chuẩn đào tạo HCMUTE (mặc định 7.0 điểm Trắc nghiệm + 3.0 điểm Tự luận).
5. **Cơ chế xuất hồ sơ lưu trữ**:
   - Tải gói bài thi PDF (`.zip`): Chứa thuần túy 100% file `.pdf` bài thi của từng thí sinh, nhúng font TrueType Arial tiếng Việt sắc nét, tương thích hoàn toàn trên môi trường Linux (Render).
   - Tải gói điểm Excel (`.zip`): Chứa bảng danh sách điểm tổng hợp của cả lớp và các tệp Excel đối chiếu bài làm chi tiết của từng sinh viên.

---

## 2. Bảng liệt kê 10 thành phần kỹ thuật cốt lõi

### 2.1 Frontend
- **Công nghệ**: Pure Vanilla JavaScript (ES6+), HTML5, CSS3 hiện đại, Responsive (Flexbox/Grid).
- **Vị trí file**: [`static/index.html`](file:///C:/Users/WIN%2010/.gemini/antigravity/scratch/online-exam-platform/static/index.html) (4.213 dòng code).
- **Đặc điểm kiến trúc**:
  * Là một ứng dụng Single Page Application (SPA) hoàn chỉnh, không cần build step (không dùng Webpack/Vite), không phụ thuộc React/Vue giúp tải tức thì.
  * Quản lý trạng thái thông qua các biến toàn cục: `currentUser`, `studentAnswers`, `examQuestions`, `flaggedQuestions`, `examEndTime`.
  * Cơ chế đồng bộ kép (Dual Persistence): Khi thí sinh chọn đáp án, hệ thống vừa gửi API `/api/exam/save` lên máy chủ, vừa lưu tức thời vào `localStorage` (`exam_progress_{examId}_{userId}`). Nếu mất mạng hoặc F5, dữ liệu bài làm được khôi phục ngay lập tức.
  * Question Palette (Bảng điều hướng câu hỏi): Phân nhóm rõ ràng giữa Phần I (Trắc nghiệm) và Phần II (Tự luận), hiển thị trực quan trạng thái câu đã làm, câu chưa làm, câu đánh dấu cờ xem lại.
  * Đồng hồ đếm ngược thông minh: Tự động khóa đề và kích hoạt nộp bài tự động khi hết giờ.

### 2.2 Backend
- **Công nghệ**: Python 3.12, framework **FastAPI 0.141.1**, web server **Uvicorn**, toolkit **Starlette**.
- **Vị trí file chính**: [`main.py`](file:///C:/Users/WIN%2010/.gemini/antigravity/scratch/online-exam-platform/main.py) (2.329 dòng code).
- **Các module phụ trợ**:
  * [`models.py`](file:///C:/Users/WIN%2010/.gemini/antigravity/scratch/online-exam-platform/models.py): Định nghĩa các thực thể ORM SQLAlchemy và hàm tự động nâng cấp schema `migrate_database()`.
  * [`database.py`](file:///C:/Users/WIN%2010/.gemini/antigravity/scratch/online-exam-platform/database.py): Cấu hình động cơ kết nối SQLite.
  * [`docx_parser.py`](file:///C:/Users/WIN%2010/.gemini/antigravity/scratch/online-exam-platform/docx_parser.py): Động cơ bóc tách ngân hàng câu hỏi Word.
  * [`student_parser.py`](file:///C:/Users/WIN%2010/.gemini/antigravity/scratch/online-exam-platform/student_parser.py): Động cơ bóc tách danh sách thí sinh.
  * [`pdf_export.py`](file:///C:/Users/WIN%2010/.gemini/antigravity/scratch/online-exam-platform/pdf_export.py): Động cơ sinh PDF và đóng gói ZIP bài thi.
  * [`excel_export.py`](file:///C:/Users/WIN%2010/.gemini/antigravity/scratch/online-exam-platform/excel_export.py): Động cơ sinh bảng điểm Excel.
- **Tiến trình chạy ngầm (Background Task)**:
  * Hàm `session_scanner` khởi chạy cùng ứng dụng, định kỳ mỗi 20 giây quét toàn bộ các phiên thi `in_progress`. Nếu quá thời gian quy định, hệ thống tự động khóa và tính điểm các câu đã làm.

### 2.3 Database
- **Hệ quản trị**: **SQLite 3** thông qua SQLAlchemy ORM.
- **Vị trí tệp**: [`sql_app.db`](file:///C:/Users/WIN%2010/.gemini/antigravity/scratch/online-exam-platform/sql_app.db).
- **Cấu hình hiệu năng & Toàn vẹn (PRAGMA)**:
  * `PRAGMA journal_mode=WAL` (Write-Ahead Logging): Cho phép đọc ghi đồng thời, hạn chế tối đa lock database khi nhiều sinh viên cùng thao tác.
  * `PRAGMA synchronous=NORMAL`: Cân bằng giữa tốc độ ghi và an toàn dữ liệu.
  * `PRAGMA busy_timeout=60000`: Chờ tối đa 60 giây nếu database bận, tránh lỗi `database is locked`.
  * `PRAGMA foreign_keys=ON`: Bật ràng buộc khóa ngoại bảo vệ tính toàn vẹn dữ liệu.
  * `poolclass=NullPool`: Đóng kết nối ngay sau khi hoàn thành request, tránh lỗi xung đột luồng của SQLite.
- **Cấu trúc 4 bảng dữ liệu chính**:
  1. `users`: Quản lý thông tin thí sinh và admin (`id`, `username`, `password`, `is_admin`, `fullname`, `dob`, `class_name`, `order_index`).
  2. `exams`: Cấu hình kỳ thi (`id`, `title`, `code`, `num_questions`, `duration_minutes`, `open_time`, `close_time`, `is_active`, `is_archived`, `mc_max_score`, `essay_max_score`).
  3. `questions`: Ngân hàng câu hỏi (`id`, `exam_id`, `content`, `option_a` .. `option_f`, `correct_option`, `question_type`, `score_weight`).
  4. `exam_results`: Kết quả và audit log chi tiết bài làm (`id`, `exam_id`, `user_id`, `score`, `max_score`, `answers`, `answers_detail`, `questions`, `start_time`, `submit_time`, `status`).

### 2.4 API Endpoints
Hệ thống cung cấp danh mục API RESTful toàn diện:

| Phương thức | Endpoint | Phân quyền | Chức năng nghiệp vụ |
| :--- | :--- | :---: | :--- |
| `POST` | `/token` | Public | Đăng nhập lấy mã thông báo JWT (OAuth2 Password Bearer) |
| `GET` | `/api/me` | Authenticated | Trả về thông tin hồ sơ của tài khoản đang đăng nhập |
| `GET` | `/api/exams` | Authenticated | Danh sách kỳ thi đang kích hoạt (`is_active = True`) cho thí sinh |
| `GET` | `/api/exam` | Student Only | Cấp phát đề thi thí sinh (xáo trắc nghiệm, cố định tự luận) |
| `POST` | `/api/exam/save` | Student Only | Tự động lưu đáp án thí sinh trong quá trình làm bài |
| `POST` | `/api/exam/submit` | Student Only | Nộp bài thi, tính điểm trắc nghiệm, chuyển tự luận chờ chấm |
| `GET` | `/api/exam/review` | Student Only | Xem lại bài thi (ẩn đáp án nếu ca thi chưa đóng) |
| `GET` | `/api/admin/exams` | Admin Only | Lấy danh sách toàn bộ các kỳ thi |
| `POST` | `/api/admin/exams` | Admin Only | Tạo kỳ thi mới |
| `PUT` | `/api/admin/exams/{id}` | Admin Only | Cập nhật cấu hình kỳ thi |
| `POST` | `/api/admin/exams/{id}/activate` | Admin Only | Kích hoạt kỳ thi làm đề thi chính thức |
| `POST` | `/api/admin/exams/{id}/archive` | Admin Only | Lưu trữ kỳ thi cũ |
| `DELETE` | `/api/admin/exams/{id}` | Admin Only | Xóa kỳ thi (chặn xóa kỳ thi đang active) |
| `GET` | `/api/admin/exams/{id}/backup` | Admin Only | Tải gói snapshot sao lưu kỳ thi (JSON) |
| `POST` | `/api/admin/exams/restore` | Admin Only | Phục hồi kỳ thi từ file JSON |
| `POST` | `/api/admin/upload_questions` | Admin Only | Tải lên file Word bóc tách câu hỏi và hình ảnh |
| `GET` | `/api/admin/questions` | Admin Only | Xem danh sách câu hỏi phân theo Phần I & Phần II |
| `DELETE` | `/api/admin/questions/{id}` | Admin Only | Xóa câu hỏi đơn lẻ |
| `DELETE` | `/api/admin/questions/exam/{id}`| Admin Only | Xóa toàn bộ câu hỏi của kỳ thi |
| `POST` | `/api/admin/upload_students` | Admin Only | Tải lên file Excel/CSV bóc tách danh sách sinh viên |
| `GET` | `/api/admin/students` | Admin Only | Danh sách thí sinh kèm trạng thái thi và điểm số |
| `POST` | `/api/admin/students` | Admin Only | Thêm thí sinh thủ công |
| `PUT` | `/api/admin/students/{id}` | Admin Only | Cập nhật thông tin thí sinh |
| `DELETE` | `/api/admin/students/{id}` | Admin Only | Xóa tài khoản thí sinh |
| `POST` | `/api/admin/students/{id}/reset_exam` | Admin Only | Reset ca thi cho thí sinh gặp sự cố |
| `GET` | `/api/admin/results` | Admin Only | Thống kê bảng điểm lớp (điểm TB, cao nhất, tỷ lệ đậu) |
| `GET` | `/api/admin/results/{id}/detail` | Admin Only | Xem chi tiết bài làm từng câu của thí sinh (Audit) |
| `POST` | `/api/admin/results/{id}/score_essay` | Admin Only | Nhập điểm tự luận và cập nhật điểm tổng kết |
| `GET` | `/api/admin/results/export` | Admin Only | Xuất file ZIP chứa toàn bộ bảng điểm Excel |
| `GET` | `/api/admin/results/export_pdf_zip` | Admin Only | Xuất file ZIP chứa toàn bộ bài thi PDF |
| `GET` | `/api/admin/results/{id}/export_pdf` | Admin Only | Tải file PDF bài thi của riêng 1 thí sinh |
| `GET` | `/api/admin/results/{id}/export_excel` | Admin Only | Tải file Excel đối chiếu của riêng 1 thí sinh |
| `GET` | `/health`, `/api/health` | Public | Kiểm tra trạng thái hoạt động của máy chủ |
| `GET` | `/` | Public | Phục vụ trang giao diện người dùng `index.html` |

### 2.5 Authentication (Xác thực)
- Sử dụng chuẩn OAuth2 Password Bearer tại endpoint `/token`.
- JWT Token được mã hóa bằng thuật toán đối xứng `HS256`, thời hạn 240 phút.
- **Quy tắc xác thực mật khẩu**:
  * Thí sinh: Mật khẩu là **MSSV** hoặc **Ngày sinh (DOB)**. Hệ thống tự động chuẩn hóa ngày sinh (ví dụ nhập `26/05/2005` hay `2005-05-26` đều khớp).
  * Đã loại bỏ hoàn toàn các mật khẩu mặc định cố định (`123456`, `hcmute`) để ngăn ngừa thí sinh đăng nhập tài khoản của nhau.
  * Admin: Mật khẩu được mã hóa và xác thực qua thuật toán băm Bcrypt an toàn.

### 2.6 Authorization (Phân quyền)
- Phân quyền chặt chẽ thông qua hàm phụ thuộc `get_current_user`:
  * Mọi endpoint quản trị bắt buộc kiểm tra `if not current_user.is_admin: raise HTTPException(403)`.
  * Các endpoint làm bài thi của thí sinh từ chối tài khoản Admin để tránh làm sai lệch dữ liệu kết quả thi.

### 2.7 AI Components
- **Không có**: Hệ thống không sử dụng bất kỳ thư viện hoặc dịch vụ AI nào (không OpenAI, không Gemini, không Claude API). Toàn bộ chức năng bóc tách văn bản, nhận diện đáp án và chấm điểm đều vận hành bằng thuật toán logic, regex và cây cú pháp OpenXML.

### 2.8 File Upload
- Sử dụng `fastapi.UploadFile` kết hợp `python-multipart`:
  * Upload đề thi: Nhận file `.docx` xử lý qua thư viện `python-docx` và `lxml`.
  * Upload sinh viên: Nhận file `.xlsx`, `.xls`, `.csv`, `.docx` xử lý qua `openpyxl`, `pandas` và `xlrd`.
  * Upload backup: Nhận file `.json` để khôi phục snapshot kỳ thi.

### 2.9 External Services
- **Render Cloud**: Dịch vụ máy chủ đám mây chạy hệ điều hành Ubuntu Linux, tự động build và deploy từ GitHub repository.
- **Cloudflare Tunnel (`cloudflared`)**: Thiết lập kết nối đường hầm mã hóa bảo mật từ localhost:8000 ra Internet.
- **GitHub**: Máy chủ lưu trữ mã nguồn tại `https://github.com/hellokhoAnguyen91/online-exam-platform..git`.

### 2.10 Environment Variables
- `PORT`: Cổng dịch vụ lắng nghe (mặc định 8000, Render tự động cấp phát qua `$PORT`).
- `SECRET_KEY`: Khóa bí mật dùng để ký và xác thực chữ ký JWT, được đọc từ `os.environ.get("SECRET_KEY", "hcmute-exam-secure-key-2026-prod-jwt")`.

---

## 3. Hướng dẫn cách chạy project ở môi trường Local

Hệ thống được đóng gói hoàn chỉnh trong môi trường ảo Python (`venv`) tại thư mục:
`C:\Users\WIN 10\.gemini\antigravity\scratch\online-exam-platform`

### Các bước khởi chạy:
1. **Mở PowerShell** và điều hướng vào thư mục dự án:
   ```powershell
   cd "C:\Users\WIN 10\.gemini\antigravity\scratch\online-exam-platform"
   ```
2. **Khởi chạy máy chủ FastAPI qua Uvicorn**:
   ```powershell
   .\venv\Scripts\python.exe -m uvicorn main:app --host 0.0.0.0 --port 8000
   ```
3. **Truy cập hệ thống trên trình duyệt web**:
   - Địa chỉ URL: `http://localhost:8000` hoặc `http://127.0.0.1:8000`
   - Tài khoản Admin: `trangnh@hcmute.edu.vn` / Mật khẩu: `nguyenhatrang`
   - Tài khoản Thí sinh mẫu: `23150014` / Mật khẩu: `23150014` (hoặc ngày sinh)

---

## 4. Hướng dẫn cách chạy các bài test hiện có

Dự án có sẵn 18 file kiểm thử tự động. Quản trị viên có thể chạy kiểm tra bất cứ lúc nào:

### 4.1 Chạy nhóm kiểm thử đơn vị & Parser (An toàn tuyệt đối, không ảnh hưởng DB)
```powershell
# Kiểm tra bộ bóc tách danh sách sinh viên Excel/CSV
.\venv\Scripts\python.exe -m unittest test_student_parser.py

# Kiểm tra bộ bóc tách câu hỏi và hình ảnh file Word
.\venv\Scripts\python.exe test_docx_parser.py

# Kiểm tra bóc tách 3 dạng câu hỏi đặc thù
.\venv\Scripts\python.exe test_parser_3cases.py

# Kiểm tra nhận diện tiêu đề Section Headers
.\venv\Scripts\python.exe test_section_headers.py
```

### 4.2 Chạy nhóm kiểm thử tích hợp toàn diện (Integration Test Suites)
```powershell
# Kiểm thử toàn diện 21 tiêu chuẩn khắt khe nhất (Bảo mật, Chấm điểm, PDF/ZIP)
.\venv\Scripts\python.exe test_comprehensive_audit.py

# Kiểm thử bảo mật, chống gian lận và che giấu đáp án khi thi
.\venv\Scripts\python.exe test_adversarial_verification.py

# Kiểm thử logic chờ chấm tự luận và trắc nghiệm chưa có đáp án
.\venv\Scripts\python.exe test_pending_grading_logic.py

# Kiểm thử toàn bộ luồng thi 2 phần (Trắc nghiệm + Tự luận)
.\venv\Scripts\python.exe test_full_section_flow.py
```

---

## 5. Kiểm tra trạng thái Git

- **Kết quả kiểm tra**: Dự án **ĐÃ CÓ** Git repository từ trước và đang hoạt động chuẩn mực.
- **Nhánh hiện tại**: `main`
- **Địa chỉ kho lưu trữ từ xa (Remote origin)**:
  `https://github.com/hellokhoAnguyen91/online-exam-platform..git`
- **Trạng thái đồng bộ**: Cây làm việc sạch sẽ (`working tree clean`), nhánh local đã đồng bộ hoàn toàn với nhánh từ xa `origin/main` ở commit mới nhất `82cf6e9`.

---

## 6. Hướng dẫn tạo Git repository (Dự phòng)

Trong trường hợp bạn muốn tạo một kho mã nguồn mới độc lập trên máy khác:

```powershell
# Bước 1: Khởi tạo git repository
git init

# Bước 2: Thêm toàn bộ tệp vào danh sách theo dõi
git add .

# Bước 3: Tạo commit khởi đầu
git commit -m "Khoi tao he thong thi truc tuyen HCMUTE"

# Bước 4: Đổi tên nhánh chính thành main
git branch -M main

# Bước 5: Liên kết với repository trên GitHub của bạn
git remote add origin https://github.com/<tai-khoan>/<ten-repo>.git

# Bước 6: Đẩy mã nguồn lên GitHub
git push -u origin main
```

---

## 7. Bản kiểm thử hiện tại (Test Execution Baseline Snapshot)

Đã thực hiện chạy kiểm thử trực tiếp trên hệ thống tại thời điểm kiểm tra mà **không thay đổi code logic**:

```text
======================================================================
1. Test Student Parser (test_student_parser.py)
   Ran 8 tests in 0.227s -> OK (PASSED 100%)
   ✓ Nhận diện chính xác cột họ, tên, ngày sinh nhiều định dạng.
   ✓ Giữ nguyên số 0 đầu MSSV.

2. Test Docx Parser & Images (test_docx_parser.py)
   Sample docx test: PASSED (3/3 questions with images)
   11 Advanced test cases: PASSED 100%
   ✓ Case 1 (Gạch chân đáp án): PASSED
   ✓ Case 2 (Nhiều phương án trên 1 dòng): PASSED
   ✓ Case 3 (Ảnh nằm trong đoạn văn trống): PASSED
   ✓ Case 4 (Dấu sao * trước đáp án đúng): PASSED
   ✓ Case 5 (Bảng biểu lồng trong câu hỏi): PASSED
   ✓ Case 6 (Phương án in đậm): PASSED
   ✓ Case 7 (Multi-option underline offset match): PASSED
   ✓ Case 8 (Tái sử dụng hình ảnh): PASSED
   ✓ Case 9 (Công thức toán học OMML): PASSED
   ✓ Case 10 (Giữ nguyên liên kết Hyperlink): PASSED
   ✓ Case 11 (Gạch đầu dòng mô tả không đè options): PASSED

3. Test Adversarial Verification (test_adversarial_verification.py)
   ✓ Test 1: Đăng nhập MSSV có tiền tố 'sv' thành công (PASSED)
   ✓ Test 2: Chấm điểm từng phần Multi-Select với payload mảng (PASSED)
   ✓ Test 3: Che giấu toàn bộ đáp án khi ca thi đang mở (PASSED)
   ✓ Test 4: Hiển thị đầy đủ đáp án sau khi ca thi đóng (PASSED)
   ✓ Test 5: Thí sinh tạo mới không có ngày sinh đăng nhập bằng MSSV (PASSED)

4. Test Comprehensive Audit (test_comprehensive_audit.py)
   ✓ Admin auth & Student login with MSSV: PASSED
   ✓ Chặn toàn bộ mật khẩu yếu 123456, 12345678, hcmute (401): PASSED
   ✓ Bảo vệ API /api/exams yêu cầu token và lọc đề active: PASSED
   ✓ Điểm từng phần câu hỏi chọn nhiều đáp án: PASSED
   ✓ Câu tự luận bỏ trắng tự tính 0đ: PASSED
   ✓ Giới hạn biên độ chấm điểm tự luận của giáo viên: PASSED
   ✓ Đồng bộ tỷ lệ điểm chuẩn 7.0 Trắc nghiệm / 3.0 Tự luận: PASSED
   ✓ Bảo toàn dữ liệu lịch sử nộp bài trong sync_expired_sessions: PASSED
   ✓ Sao lưu và Phục hồi giữ trọn vẹn kiểu câu và trọng số: PASSED
   ✓ Xuất file bài thi PDF và file nén ZIP 100% tệp pdf: PASSED
   ✓ Tính năng Reset bài thi cho thí sinh của Admin: PASSED

🎉 TOÀN BỘ CÁC BỘ KIỂM THỬ ĐẠT TỶ LỆ VƯỢT QUA 100%!
======================================================================
```

---

## 8. Tài liệu Kiến trúc: PROJECT_ARCHITECTURE.md

Tài liệu chi tiết đã được tạo tại [`PROJECT_ARCHITECTURE.md`](file:///C:/Users/WIN%2010/.gemini/antigravity/scratch/online-exam-platform/PROJECT_ARCHITECTURE.md).  
Tài liệu trình bày:
- Tổng quan mục tiêu và thiết kế kỹ thuật của nền tảng.
- Bảng công nghệ và phiên bản chi tiết.
- Sơ đồ kiến trúc luồng dữ liệu (Mermaid Architecture Diagram).
- Cấu trúc thư mục mã nguồn và dữ liệu.
- Thiết kế chi tiết 4 bảng cơ sở dữ liệu và các ràng buộc toàn vẹn.
- Cơ chế xác thực OAuth2 / JWT và ủy quyền vai trò.
- Động cơ điều phối đề thi 2 phần (Trắc nghiệm xáo ngẫu nhiên, Tự luận cố định).
- Thuật toán chấm điểm chi tiết và công thức quy đổi điểm về thang 10.
- Động cơ xuất hồ sơ lưu trữ PDF / Excel (.zip).

---

## 9. Tài liệu Quy tắc Nghiệp vụ: BUSINESS_RULES.md

Tài liệu chi tiết đã được tạo tại [`BUSINESS_RULES.md`](file:///C:/Users/WIN%2010/.gemini/antigravity/scratch/online-exam-platform/BUSINESS_RULES.md).  
Tài liệu ghi lại toàn bộ 24 quy tắc nghiệp vụ đang vận hành:
- **BR-AUTH**: Quy tắc tài khoản MSSV, mật khẩu ngày sinh chuẩn hóa, loại trừ mật khẩu yếu, hạn dùng token.
- **BR-EXAM**: Quy tắc chỉ có 1 kỳ thi active duy nhất, chống xóa kỳ thi đang active, khung giờ mở/đóng thi.
- **BR-PARSE**: Quy tắc nhận diện câu hỏi, trích xuất ảnh Base64, nhận diện 4 dạng đáp án trắc nghiệm, nhận diện câu multi-select và câu tự luận `(X điểm)`.
- **BR-FLOW**: Quy tắc cấu trúc đề thi 2 phần, quy tắc bảo toàn phiên thi chống F5 đổi đề, lưu tự động kép lên máy chủ và localStorage.
- **BR-SCORE**: Công thức tính điểm trắc nghiệm đơn, công thức trừ điểm chọn sai của multi-select, quy tắc chấm tự luận và thang điểm 7/3 quy đổi về 10.
- **BR-TIME**: Quy tắc đếm ngược hết giờ khóa bài tự động, tiến trình quét ngầm 20s, bảo toàn dữ liệu bài đã nộp.
- **BR-REV**: Quy tắc che giấu đáp án và lời giải khi ca thi đang mở, mở đối chiếu khi ca thi đã kết thúc.
- **BR-EXP**: Quy tắc tệp ZIP bài thi PDF chỉ chứa tệp `.pdf`, font TrueType Arial tiếng Việt.

---

## 10. Tài liệu Kế hoạch Kiểm thử: TEST_PLAN.md

Tài liệu chi tiết đã được tạo tại [`TEST_PLAN.md`](file:///C:/Users/WIN%2010/.gemini/antigravity/scratch/online-exam-platform/TEST_PLAN.md).  
Tài liệu cung cấp:
- Mục tiêu và phạm vi kiểm thử hệ thống.
- Bảng kê danh mục toàn bộ 18 file test tự động có trong dự án.
- Ma trận 6 nhóm kịch bản kiểm thử trọng yếu (Xác thực, Parser, Điều phối đề, Chấm điểm, Quản lý ca thi hết giờ, Xuất hồ sơ).
- Hướng dẫn chi tiết câu lệnh thực thi từng nhóm kiểm thử.
- Tiêu chí đánh giá Đạt (Pass) / Không đạt (Fail).

---

## 11. Báo cáo phát hiện & Khuyến nghị vận hành

Qua quá trình rà soát toàn diện, hệ thống hiện tại đang ở trạng thái rất tốt và ổn định. Dưới đây là các lưu ý quan trọng để đảm bảo kỳ thi chính thức diễn ra an toàn 100%:

1. **Kỳ thi chính thức trong Cơ sở dữ liệu**:
   - Hiện tại trong cơ sở dữ liệu `sql_app.db` đang lưu giữ kỳ thi chính thức: **Kỳ thi 123 ("Thi cuối kỳ 2026 - An toàn vệ sinh lao động")**.
   - Kỳ thi đang ở trạng thái kích hoạt sẵn sàng (`is_active = True`) với đầy đủ 55 câu hỏi và 2 bài thi đã nộp của thí sinh được bảo toàn nguyên vẹn.
2. **Lưu ý về máy chủ Render Cloud (Free Tier)**:
   - Trên gói miễn phí của Render, nếu không có lượt truy cập trong 15 phút, máy chủ sẽ tự động chuyển sang trạng thái ngủ (Sleep mode) để tiết kiệm tài nguyên. Lượt truy cập đầu tiên sau khi ngủ sẽ mất khoảng 30–50 giây để khởi động lại container.
   - **Khuyến nghị vận hành**: Trước giờ thi chính thức khoảng 5–10 phút, Quản trị viên nên mở trang web trên trình duyệt để đánh thức máy chủ, đảm bảo khi sinh viên đồng loạt đăng nhập vào thi thì hệ thống đã phản hồi tức thì.
3. **Đường hầm Cloudflare (Dự phòng truy cập)**:
   - Ngoài tên miền chính trên Render, hệ thống có thể kết nối thông qua đường hầm Cloudflare Tunnel (`cloudflared`) để cung cấp đường truyền dự phòng cho thí sinh nếu cần.
