# 🎓 Nền tảng Thi Trực tuyến Chuyên nghiệp (HCMUTE Online Exam Platform)

Hệ thống thi trắc nghiệm trực tuyến chuẩn EdTech, được phát triển phục vụ công tác khảo thí và giảng dạy tại Trường ĐH Sư phạm Kỹ thuật TP. Hồ Chí Minh (HCMUTE). Xây dựng trên nền tảng **FastAPI**, **SQLite (SQLAlchemy)** và giao diện hiện đại **Modern Vanilla Web UI**.

---

## 🚀 Các Tính Năng Nâng Cấp Toàn Diện (Phiên Bản 2.0)

### 1. 🧠 Parser Câu Hỏi & Trích Xuất Hình Ảnh Siêu Thông Minh (.docx)
- **Bóc tách 100% hình ảnh trong file Word**: Hỗ trợ mọi cấu trúc ảnh phức tạp của Microsoft Word:
  - Drawing inline (`wp:inline`) và Floating / Anchor (`wp:anchor`)
  - Legacy VML shapes & imagedata (`v:imagedata`, `w:pict`)
  - AlternateContent (`mc:AlternateContent` - ưu tiên Choice, chống trùng lặp Fallback)
  - Bảo toàn 100% hình ảnh khi tái sử dụng cùng 1 ảnh minh họa qua nhiều câu hỏi khác nhau.
  - Ảnh đặt trong bảng (Tables), ảnh nằm trên dòng riêng biệt (empty text paragraph), ảnh trong các phương án A/B/C/D.
  - Chuyển đổi an toàn sang chuẩn Base64 Data URI, tự chứa hoàn chỉnh trong cơ sở dữ liệu.
- **Hỗ trợ công thức toán học Word Equation (OMML / MathML)**:
  - Bóc tách và chuyển đổi các công thức toán `<m:oMath>` (phân số, lũy thừa, chỉ số dưới, căn thức, dấu ngoặc) thành HTML trực quan (`<sup>`, `<sub>`, `&radic;`, phân số), đảm bảo công thức không bao giờ bị biến mất.
- **Bảo toàn siêu liên kết (Hyperlinks)**:
  - Giữ nguyên văn bản và đường link web `<w:hyperlink>` có trong đề thi.
- **Nhận diện linh hoạt mọi định dạng đề thi tiếng Việt & tiếng Anh**:
  - Tiêu đề câu: `Câu 1:`, `Câu 01.`, `Câu hỏi 1:`, `CH 1:`, `1.`, `1/`, `1)`, `Question 1:`, `Q1:`, `Bài 1:`,...
  - Phương án: nằm trên từng dòng riêng (`A. ... \n B. ...`) hoặc gộp trên 1 dòng (`A. ... B. ... C. ... D. ...`) hoặc đặt trong ô bảng (Table cells).
  - Tự động nhận diện đáp án đúng theo:
    - Dòng đáp án: `Đáp án: A`, `Đáp án đúng: B`, `Đ/A: C`, `Key: D`, `Answer: A`, `Hướng dẫn giải: chọn B`.
    - Phương án gạch chân: Giảng viên gạch chân đáp án đúng (`<u>A.</u>` hoặc `<u>Hà Nội</u>`), áp dụng thuật toán so khớp khoảng ký tự chuẩn xác chống nhận nhầm ký tự trong từ.
    - Phương án in đậm: Giảng viên bôi đậm đáp án đúng (tự động nhận diện phương án in đậm trên dòng đơn hoặc dòng gộp).
    - Ký hiệu đánh dấu: `*A. ...` hoặc `[x] A. ...`.

### 2. 🔍 Lưu Vết Bài Làm & Đối Chiếu Phúc Khảo (Audit & Review System)
- **Lưu trữ chi tiết nguyên vẹn từng bài làm**:
  - Ghi nhận chính xác bộ đề câu hỏi đã phân phối cho từng thí sinh tại thời điểm thi.
  - Lưu chi tiết phương án thí sinh chọn, đáp án chính xác, trạng thái đúng/sai từng câu.
  - Cơ chế tự động tái cấu trúc lịch sử phúc khảo dự phòng nếu bản ghi cũ bị thiếu snapshot.
  - Lưu thời gian bắt đầu, thời điểm nộp bài, thời lượng làm bài thực tế (phút:giây), địa chỉ IP và trình duyệt của thí sinh.
- **Giao diện Đối Chiếu Bài Làm (Audit Modal)**:
  - Bấm vào bất kỳ thí sinh nào trong bảng kết quả để xem lại toàn bộ bài làm.
  - Đánh dấu trực quan màu sắc: Màu xanh lá (✓ Đúng - tự tính điểm động theo số câu), Màu đỏ (✗ Sai - 0 đ), Màu vàng (⚪ Chưa làm).
  - Hiển thị rõ ràng phương án thí sinh đã chọn và đáp án đúng.
  - **Phiếu in đối chiếu phúc khảo (Printable Audit Sheet)**: Tích hợp định dạng in `@media print` chuyên nghiệp với quốc hiệu, tiêu ngữ, thông tin thí sinh, bảng câu hỏi và khung ký tên xác nhận của Thí sinh & Giảng viên chấm thi.

### 3. 📦 Quản Lý Nhiều Kỳ Thi & Kho Lưu Trữ Dữ Liệu (Multi-Exam & Archive)
- **Không bao giờ bị ghi đè hay mất dữ liệu**:
  - Quản lý danh sách nhiều kỳ thi độc lập (Kỳ thi giữa kỳ, cuối kỳ, kiểm tra thường xuyên,...).
  - Kích hoạt kỳ thi đang mở thi (Active Exam) hoặc chuyển đổi dễ dàng.
  - **Cơ chế chống Race Condition**: Thí sinh đang làm bài vẫn nộp bài an toàn về đúng kỳ thi của mình ngay cả khi Giảng viên chuyển đổi kích hoạt kỳ thi khác trong phiên làm việc.
  - **Lưu trữ kỳ thi (Archive)**: Đóng băng điểm số và dữ liệu bài làm của kỳ thi đã kết thúc vào kho lưu trữ an toàn.
  - **Xuất gói sao lưu toàn diện (Full Backup JSON)**: Tải file sao lưu chứa toàn bộ cấu hình, ngân hàng câu hỏi kèm ảnh gốc và bài làm của thí sinh.
  - **Khôi phục kỳ thi (Restore)**: Phục hồi nguyên trạng kỳ thi từ file sao lưu bất kỳ lúc nào.

### 4. 🎨 Đại Tu Giao Diện Chuẩn EdTech (UI/UX Hiện Đại & Chuyên Nghiệp)
- **Giao diện Thí sinh (Lấy cảm hứng từ Canvas LMS & Azota)**:
  - Bảng điều hướng câu hỏi bên cạnh (**Question Palette**): Lưới các số 1, 2, 3... Đánh dấu câu đã làm (xanh dương), câu chưa làm (xám), và tính năng **Đánh dấu câu phân vân để xem lại (🚩 Flag)**. Bấm vào số câu để cuộn mượt đến câu đó.
  - **Thanh điều hướng nổi trên thiết bị di động (Mobile Action Bar)**: Tối ưu trên điện thoại, thí sinh có thể xem tiến độ và bấm mở Bảng câu hỏi / Nộp bài bất cứ lúc nào.
  - **Đồng hồ đếm ngược viền card hiện đại**: Tự đổi màu cảnh báo khi còn dưới 5 phút, nhấp nháy đỏ khi còn dưới 1 phút.
  - **Thanh tiến độ làm bài (Progress Bar)**: Hiển thị trực quan tỷ lệ % hoàn thành (`Đã làm 15/20 câu (75%)`).
  - **Tự động lưu tiến độ (Auto-Save)** liên tục sau mỗi thao tác chọn đáp án, báo trạng thái lưu thời gian thực.
  - **Cảnh báo an toàn**: Hộp thoại cảnh báo số câu chưa làm trước khi nộp, cảnh báo chống vô tình tắt/rời tab trình duyệt (`beforeunload`).
  - **Khôi phục phiên thi an toàn**: Xử lý tải lại trang khi kỳ thi chưa mở hoặc khi mất kết nối mà không bị lỗi giao diện.
- **Giao diện Giảng viên (Admin Dashboard)**:
  - Dashboard hiện đại chia 5 tab chuyên biệt: *Kỳ thi & Lưu trữ*, *Ngân hàng đề thi*, *Quản lý thí sinh*, *Cấu hình kỳ thi*, *Bảng điểm & Đối chiếu*.
  - Thống kê phân tích điểm thi: Tổng số bài nộp, Điểm trung bình, Cao nhất/Thấp nhất, Tỷ lệ đạt (>= 5.0).
  - Xuất bảng điểm Excel định dạng chuẩn.

---

## 🔐 Thông Tin Đăng Nhập Mặc Định

| Vai trò | Tên đăng nhập | Mật khẩu mặc định | Ghi chú |
|---------|---------------|-------------------|---------|
| **Giảng viên (Admin)** | `trangnh@hcmute.edu.vn` | `nguyenhatrang` | Tài khoản giảng viên quản trị |
| **Thí sinh (Sinh viên)** | MSSV (vd: `21110001` hoặc `101`) | Ngày sinh DOB (Hỗ trợ linh hoạt `2000-01-01` hoặc `01/01/2000`) | Thí sinh trong danh sách lớp |

---

## 🚀 Hướng Dẫn Cài Đặt & Khởi Động

### Yêu cầu:
- Python 3.8+ (khuyên dùng Python 3.10 - 3.12)

### 1. Cài đặt thư viện:
```bash
pip install -r requirements.txt
```

### 2. Khởi chạy máy chủ:
```bash
python main.py
```
Máy chủ khởi chạy tại: **http://localhost:8000** (hoặc truy cập qua mạng nội bộ: `http://<IP-may-tinh>:8000`).

### 3. Chạy bộ kiểm thử tự động (Unit Tests & Integration Tests):
```bash
python test_docx_parser.py
python test_api_suite.py
python test_edge_cases.py
python test_review_fixes.py
```
Tất cả các bộ test kiểm tra parser, API, xác thực, bóc tách ảnh, công thức toán và đối chiếu bài làm đều đạt 100%.
