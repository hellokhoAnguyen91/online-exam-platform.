# QUY TẮC NGHIỆP VỤ HỆ THỐNG (BUSINESS RULES)
**Hệ thống Thi Trực tuyến HCMUTE**

Tài liệu này ghi lại toàn bộ các quy tắc nghiệp vụ (Business Rules) hiện đang được thực thi trên toàn bộ hệ thống từ Backend, Database đến Frontend.

---

## 1. Quy tắc Tài khoản & Xác thực Thí sinh (Authentication Rules)

- **Quy tắc BR-AUTH-01 (Tài khoản thí sinh)**:
  - Tên đăng nhập (`username`) của thí sinh là Mã số sinh viên (MSSV), ví dụ: `23150014`.
  - Hệ thống tự động xử lý chuẩn hóa nếu thí sinh nhập tiền tố `sv` (ví dụ `sv23150014` tự chuyển thành `23150014`).
- **Quy tắc BR-AUTH-02 (Mật khẩu thí sinh)**:
  - Thí sinh có thể sử dụng chính **Mã số sinh viên (MSSV)** làm mật khẩu mặc định.
  - Thí sinh có thể sử dụng **Ngày tháng năm sinh (DOB)** làm mật khẩu. Hệ thống chấp nhận và tự động chuẩn hóa mọi định dạng: `DD/MM/YYYY`, `DD-MM-YYYY`, `YYYY-MM-DD`, `D/M/YYYY` (ví dụ: `26/05/2005`, `26-5-2005`, `2005-05-26`).
  - Tuyệt đối không cho phép đăng nhập bằng các mật khẩu yếu cố định toàn trường như `123456`, `12345678`, `hcmute` để chống mạo danh.
- **Quy tắc BR-AUTH-03 (Tài khoản Quản trị viên)**:
  - Tài khoản quản trị sử dụng email trường (ví dụ: `trangnh@hcmute.edu.vn`) và mật khẩu mã hóa Bcrypt.
  - Quản trị viên có cờ `is_admin = True`.
- **Quy tắc BR-AUTH-04 (Phiên làm việc JWT)**:
  - Mã thông báo JWT (Access Token) có thời hạn sống tối đa là 240 phút (4 giờ).
  - Quá 4 giờ, token hết hạn và yêu cầu đăng nhập lại.

---

## 2. Quy tắc Quản lý Kỳ thi (Exam Management Rules)

- **Quy tắc BR-EXAM-01 (Kỳ thi kích hoạt duy nhất)**:
  - Tại một thời điểm, chỉ có **1 kỳ thi duy nhất** được ở trạng thái kích hoạt (`is_active = True`).
  - Khi quản trị viên bấm kích hoạt một kỳ thi mới, hệ thống tự động tắt trạng thái kích hoạt của tất cả các kỳ thi còn lại.
- **Quy tắc BR-EXAM-02 (Bảo vệ kỳ thi đang kích hoạt)**:
  - Không được phép xóa một kỳ thi đang ở trạng thái kích hoạt (`is_active = True`).
  - Muốn xóa kỳ thi, quản trị viên bắt buộc phải kích hoạt một kỳ thi khác trước.
- **Quy tắc BR-EXAM-03 (Khung thời gian mở/đóng thi)**:
  - Nếu kỳ thi có cấu hình `open_time`, thí sinh cố gắng vào làm bài trước `open_time` sẽ bị hệ thống chặn lại với thông báo thời gian mở thi cụ thể.
  - Nếu kỳ thi có cấu hình `close_time`, thí sinh cố gắng vào thi sau `close_time` sẽ bị từ chối truy cập.

---

## 3. Quy tắc Bóc tách Ngân hàng Câu hỏi (Question Parsing Rules)

- **Quy tắc BR-PARSE-01 (Nhận diện câu hỏi)**:
  - Hỗ trợ các định dạng mở đầu: `Câu 1:`, `Câu 01.`, `Câu 1 -`, `1.`, `Q1:`, `Question 1:`.
  - Giữ nguyên toàn bộ định dạng văn bản in đậm, in nghiêng, gạch chân, bảng biểu và công thức toán học (OMML).
- **Quy tắc BR-PARSE-02 (Trích xuất hình ảnh)**:
  - Tất cả hình ảnh trong file Word (kể cả inline, floating/anchor, VML, ảnh lồng trong bảng, ảnh nằm trong đoạn văn trống ngay dưới câu hỏi) đều được trích xuất và chuyển đổi thành chuỗi Base64 nhúng trực tiếp vào thẻ `<img>` của câu hỏi.
- **Quy tắc BR-PARSE-03 (Nhận diện đáp án trắc nghiệm đơn)**:
  - Hỗ trợ phương án A, B, C, D (và cả E, F).
  - Nhận diện đáp án đúng qua 4 cơ chế:
    1. Dòng chỉ định tường minh: `Đáp án: A` hoặc `Key: B`.
    2. Phương án được **gạch chân (underline)** trong file Word.
    3. Phương án được **in đậm (bold)** trong file Word.
    4. Phương án có dấu sao `*` đứng trước (ví dụ `*B. Đáp án đúng`).
- **Quy tắc BR-PARSE-04 (Nhận diện câu hỏi chọn nhiều đáp án - Multi-Select)**:
  - Nhận diện khi câu hỏi có danh sách các dòng bắt đầu bằng dấu cộng `+` hoặc đáp án đúng có dạng phân cách dấu phẩy: `MULTI:A,C,E` hoặc `A,C,E`.
- **Quy tắc BR-PARSE-05 (Nhận diện câu hỏi tự luận - Essay)**:
  - Câu hỏi không chứa các phương án A/B/C/D.
  - Tự động bóc tách điểm số từ tiêu đề: `Câu 20 (3 điểm)` hoặc `Câu 20 (3.5 pts)` sẽ gán trọng số điểm `score_weight = 3.0` hoặc `3.5`.
- **Quy tắc BR-PARSE-06 (Phân đoạn đề thi qua tiêu đề Section)**:
  - Nếu file Word có tiêu đề `PHẦN I: TRẮC NGHIỆM` và `PHẦN II: TỰ LUẬN`, hệ thống tự động phân loại chính xác các câu hỏi vào 2 phần tương ứng.

---

## 4. Quy tắc Phân phối Đề thi & Làm bài (Exam Delivery & Taking Rules)

- **Quy tắc BR-FLOW-01 (Cấu trúc đề thi 2 phần)**:
  - **Phần 1 (Trắc nghiệm)**: Lấy ngẫu nhiên $N$ câu (theo `num_questions`), xáo trộn ngẫu nhiên theo từng thí sinh (nếu bật `shuffle_questions`). Thí sinh khác nhau sẽ có thứ tự câu hỏi và mã đề khác nhau.
  - **Phần 2 (Tự luận)**: Giữ nguyên toàn bộ câu tự luận trong ngân hàng đề, giữ nguyên thứ tự cố định, không bỏ sót câu nào.
- **Quy tắc BR-FLOW-02 (Bảo toàn phiên làm bài - Anti-F5 Refresh)**:
  - Khi thí sinh tải lại trang (F5) hoặc vô tình đóng trình duyệt rồi mở lại, hệ thống ưu tiên nạp lại phiên thi đang dang dở (`status = in_progress`).
  - Tuyệt đối không sinh lại đề mới hay xáo trộn lại thứ tự câu hỏi đã phát cho thí sinh.
- **Quy tắc BR-FLOW-03 (Lưu tự động kép - Dual Auto-save)**:
  - Mỗi khi thí sinh chọn đáp án hoặc gõ câu tự luận, hệ thống tự động gửi yêu cầu lưu về server qua `POST /api/exam/save`.
  - Đồng thời sao lưu tức thì vào `localStorage` của trình duyệt. Nếu thí sinh mất mạng internet, bài làm vẫn được giữ an toàn trên trình duyệt và tự động đồng bộ lại khi có mạng.
- **Quy tắc BR-FLOW-04 (Chống lộ đáp án trên giao diện)**:
  - Khi hiển thị các lựa chọn A/B/C/D cho thí sinh làm bài, frontend tự động loại bỏ các thẻ `<strong>`, `<b>` hoặc gạch chân thừa để thí sinh không thể phát hiện đáp án đúng bằng mắt thường.

---

## 5. Quy tắc Tính điểm & Đánh giá (Scoring Rules)

- **Quy tắc BR-SCORE-01 (Trắc nghiệm đơn)**:
  - Chọn đúng đáp án: Nhận đủ $100\%$ trọng số điểm của câu hỏi.
  - Chọn sai hoặc chưa chọn: Nhận $0$ điểm.
- **Quy tắc BR-SCORE-02 (Trắc nghiệm nhiều đáp án - Chấm điểm từng phần)**:
  - Áp dụng công thức tính điểm trừ điểm chọn sai (Partial Credit Penalty):
    $$\text{Tỷ lệ đạt} = \max\left(0.0, \frac{C_{\text{đúng}} - W_{\text{sai}}}{T_{\text{đúng}}}\right)$$
    Trong đó:
    * $C_{\text{đúng}}$: Số đáp án đúng mà thí sinh đã chọn.
    * $W_{\text{sai}}$: Số đáp án sai mà thí sinh chọn nhầm.
    * $T_{\text{đúng}}$: Tổng số đáp án đúng của câu hỏi đó.
  - Điểm thô của câu = $\text{Tỷ lệ đạt} \times \text{score\_weight}$.
- **Quy tắc BR-SCORE-03 (Câu hỏi tự luận)**:
  - Nếu thí sinh không nhập bất kỳ ký tự nào: Tự động chấm $0.0$ điểm (không cần giáo viên chấm).
  - Nếu thí sinh có nhập nội dung: Đánh dấu trạng thái `pending_grading` (Chờ chấm). Điểm tổng kết toàn bài thi tạm thời để `None`.
  - Khi giáo viên chấm điểm qua API `/api/admin/results/{id}/score_essay`:
    * Điểm số nhập vào bắt buộc nằm trong biên độ: $0.0 \le \text{Điểm nhập} \le \text{score\_weight}$.
- **Quy tắc BR-SCORE-04 (Quy chuẩn thang điểm 10 Đào tạo HCMUTE)**:
  - Mặc định cơ cấu điểm: **7.0 điểm Trắc nghiệm** + **3.0 điểm Tự luận**.
  - Điểm Trắc nghiệm quy đổi = $(\text{Tổng điểm thô MC} / \text{Tổng trọng số MC}) \times \text{mc\_max\_score (7.0)}$.
  - Điểm Tự luận quy đổi = $(\text{Tổng điểm thô Essay} / \text{Tổng trọng số Essay}) \times \text{essay\_max\_score (3.0)}$.
  - Điểm toàn bài = $\text{Điểm MC quy đổi} + \text{Điểm Tự luận quy đổi}$ (làm tròn 2 chữ số thập phân).

---

## 6. Quy tắc Thu bài & Hết giờ (Submission & Expiry Rules)

- **Quy tắc BR-TIME-01 (Đếm ngược & Khóa bài tự động)**:
  - Khi đồng hồ đếm ngược trên giao diện về $00:00$, hệ thống tự động khóa mọi ô nhập liệu và gửi yêu cầu nộp bài tự động (`executeSubmitExam(true)`).
- **Quy tắc BR-TIME-02 (Quét dọn ca thi ngầm - Background Scanner)**:
  - Một tiến trình chạy ngầm định kỳ mỗi 20 giây sẽ quét toàn bộ các bài thi `in_progress`.
  - Nếu thời gian trôi qua $(\text{now} - \text{start\_time}) \ge \text{duration\_minutes} \times 60$, hệ thống tự động cưỡng chế kết thúc bài thi, tính điểm các câu đã làm và chuyển trạng thái sang `submitted`.
- **Quy tắc BR-TIME-03 (Bảo toàn dữ liệu đã nộp)**:
  - Tiến trình quét ngầm tuyệt đối không được phép chỉnh sửa hay ghi đè thời gian nộp bài (`submit_time`) và thời lượng làm bài (`duration_seconds`) của những bài thi đã nộp thành công (`submitted`).

---

## 7. Quy tắc Xem lại Bài thi (Review Rules)

- **Quy tắc BR-REV-01 (Bảo mật đề thi khi ca thi chưa kết thúc)**:
  - Khi kỳ thi vẫn đang mở (`is_active = True`) hoặc chưa qua thời điểm `close_time`, thí sinh vào trang Xem lại bài làm (`/api/exam/review`) sẽ **KHÔNG ĐƯỢC XEM**:
    * Đáp án đúng của các câu hỏi (`correct = None`).
    * Trạng thái đúng/sai của từng câu (`status = "hidden"`).
    * Lời giải thích đáp án (`explanation = None`).
    * Điểm đạt được của từng câu (`earned_score = None`).
  - Thí sinh chỉ được xem lại các phương án mình đã chọn và điểm tổng kết toàn bài (nếu đã chấm xong).
- **Quy tắc BR-REV-02 (Mở đáp án khi kỳ thi đã đóng)**:
  - Chỉ khi kỳ thi đã đóng hoàn toàn (`now > close_time` hoặc `is_active = False`) và có bật cờ `allow_review = True`, thí sinh mới được phép xem chi tiết đáp án đúng và lời giải đối chiếu.

---

## 8. Quy tắc Đóng gói Xuất bản Hồ sơ (Export Rules)

- **Quy tắc BR-EXP-01 (Xuất gói bài thi PDF .ZIP)**:
  - Tệp ZIP tải về từ nút "Tải tất cả bài thi PDF (.ZIP)" bắt buộc chỉ chứa thuần túy các tệp tin `.pdf` bài làm của thí sinh ở thư mục gốc của tệp ZIP.
  - Không sinh file `.txt` tổng hợp điểm hay file `.xlsx` trong tệp ZIP này.
- **Quy tắc BR-EXP-02 (Xuất gói điểm Excel .ZIP)**:
  - Tệp ZIP xuất điểm Excel chứa: 1 bảng điểm danh sách tổng hợp của cả lớp + các bảng tính Excel chi tiết bài làm của từng thí sinh.
- **Quy tắc BR-EXP-03 (Font chữ tài liệu PDF)**:
  - Toàn bộ file PDF xuất ra sử dụng bộ font TrueType Arial chuẩn được đóng gói sẵn trong thư mục `./fonts/` của dự án để đảm bảo hiển thị đúng 100% tiếng Việt trên môi trường máy chủ Linux (Render) cũng như Windows.
