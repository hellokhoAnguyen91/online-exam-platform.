import unittest
import io
import docx
import json
import datetime
from fastapi.testclient import TestClient
from main import app, get_password_hash
from database import SessionLocal
from models import User, Exam, Question, ExamResult
import docx_parser
from pdf_export import generate_candidate_pdf

class TestComprehensiveSectionsV2(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.client = TestClient(app)
        cls.db = SessionLocal()
        
        # 1. Admin setup
        admin = cls.db.query(User).filter(User.username == "admin_sec_v2").first()
        if not admin:
            admin = User(
                username="admin_sec_v2",
                password=get_password_hash("admin123"),
                fullname="Admin Section V2",
                is_admin=True
            )
            cls.db.add(admin)
            cls.db.commit()
            cls.db.refresh(admin)
        cls.admin = admin
        
        res = cls.client.post("/token", data={"username": "admin_sec_v2", "password": "admin123"})
        assert res.status_code == 200
        cls.admin_headers = {"Authorization": f"Bearer {res.json()['access_token']}"}

        # 2. Student setup
        student = cls.db.query(User).filter(User.username == "student_sec_v2").first()
        if not student:
            student = User(
                username="student_sec_v2",
                password=get_password_hash("pass123"),
                fullname="Sinh Viên Thử Nghiệm V2",
                dob="2003-05-15",
                is_admin=False
            )
            cls.db.add(student)
            cls.db.commit()
            cls.db.refresh(student)
        cls.student = student

        res = cls.client.post("/token", data={"username": "student_sec_v2", "password": "pass123"})
        assert res.status_code == 200
        cls.student_headers = {"Authorization": f"Bearer {res.json()['access_token']}"}

    @classmethod
    def tearDownClass(cls):
        # Cleanup test data
        cls.db.query(ExamResult).filter(ExamResult.user_id == cls.student.id).delete()
        cls.db.commit()
        cls.db.close()

    def setUp(self):
        # Clean student results before each test
        self.db.query(ExamResult).filter(ExamResult.user_id == self.student.id).delete()
        self.db.commit()

    def test_01_docx_parser_section_headers_and_subitems(self):
        """Test DOCX parser with section headers and essay with sub-parts A, B."""
        doc = docx.Document()
        
        # Section 1 Header
        doc.add_paragraph("MỤC I: CÂU HỎI TRẮC NGHIỆM KHÁCH QUAN")
        doc.add_paragraph("Câu 1. Thủ đô nước CHXHCN Việt Nam là gì?")
        doc.add_paragraph("A. Hà Nội")
        doc.add_paragraph("B. TP Hồ Chí Minh")
        doc.add_paragraph("C. Đà Nẵng")
        doc.add_paragraph("D. Cần Thơ")
        doc.add_paragraph("Đáp án: A")

        # Multi-select under Part I
        doc.add_paragraph("Câu 2: Các giao thức mạng nào thuộc tầng giao vận? (Chọn nhiều đáp án)")
        doc.add_paragraph("A. TCP")
        doc.add_paragraph("B. UDP")
        doc.add_paragraph("C. IP")
        doc.add_paragraph("D. ICMP")
        doc.add_paragraph("Đáp án: A, B")

        # Section 2 Header
        doc.add_paragraph("PHẦN II: BÀI TẬP TỰ LUẬN (5.0 điểm)")
        # Essay question WITH sub-items A. and B. (must NOT be treated as MC!)
        doc.add_paragraph("Câu 3 (3.0 điểm): Cho mạng máy tính doanh nghiệp gồm 3 phân đoạn mạng. Hãy thực hiện:")
        doc.add_paragraph("A. Thiết kế sơ đồ phân chia dải địa chỉ IP cho 3 VLAN.")
        doc.add_paragraph("B. Đề xuất chính sách định tuyến và tường lửa bảo mật liên VLAN.")
        
        # Another essay question
        doc.add_paragraph("Câu 4 (2.0 điểm): Phân tích ưu và nhược điểm của điện toán đám mây so với máy chủ tại chỗ.")

        bio = io.BytesIO()
        doc.save(bio)
        bio.seek(0)

        questions = docx_parser.parse_docx_questions(bio)
        self.assertEqual(len(questions), 4)

        # Q1: Multiple choice
        self.assertEqual(questions[0]["question_type"], "multiple_choice")
        self.assertEqual(questions[0]["correct_option"], "A")
        self.assertIn("Hà Nội", questions[0]["option_a"])

        # Q2: Multi-select
        self.assertEqual(questions[1]["question_type"], "multi_select")
        self.assertEqual(questions[1]["correct_option"], "A,B")

        # Q3: MUST BE ESSAY despite having options A. and B.
        self.assertEqual(questions[2]["question_type"], "essay")
        self.assertEqual(questions[2]["score_weight"], 3.0)
        self.assertEqual(questions[2]["option_a"], "")
        self.assertEqual(questions[2]["correct_option"], "")
        self.assertIn("Thiết kế sơ đồ phân chia", questions[2]["content"])
        self.assertIn("Đề xuất chính sách định tuyến", questions[2]["content"])

        # Q4: Essay
        self.assertEqual(questions[3]["question_type"], "essay")
        self.assertEqual(questions[3]["score_weight"], 2.0)

    def test_02_detect_section_header_variants(self):
        """Test variety of real Vietnamese section header formats."""
        cases = [
            ("PHẦN I: TRẮC NGHIỆM", "multiple_choice"),
            ("PHẦN II: TỰ LUẬN", "essay"),
            ("PHẦN 1. CÂU HỎI TRẮC NGHIỆM (4.0 ĐIỂM)", "multiple_choice"),
            ("PHẦN 2. BÀI TẬP TỰ LUẬN (6.0 ĐIỂM)", "essay"),
            ("I. PHẦN TRẮC NGHIỆM", "multiple_choice"),
            ("II. PHẦN TỰ LUẬN", "essay"),
            ("PHẦN TRẮC NGHIỆM KHÁCH QUAN", "multiple_choice"),
            ("PHẦN TỰ LUẬN", "essay"),
            ("CÂU HỎI TỰ LUẬN", "essay"),
            ("BÀI TẬP TỰ LUẬN", "essay"),
            ("MỤC I: TRẮC NGHIỆM", "multiple_choice"),
            ("MỤC II: TỰ LUẬN", "essay"),
            ("PART I: MULTIPLE CHOICE", "multiple_choice"),
            ("PART II: ESSAY QUESTIONS", "essay"),
            ("A. TRẮC NGHIỆM", "multiple_choice"),
            ("B. TỰ LUẬN", "essay"),
        ]
        for text, expected in cases:
            detected = docx_parser.detect_section_header(text)
            self.assertEqual(detected, expected, f"Failed for header: '{text}' -> expected {expected}, got {detected}")

    def test_03_end_to_end_exam_flow_and_review(self):
        """Test full candidate taking exam, saving progress with multi-select and essay, submitting, and reviewing."""
        # 1. Create Exam
        ex_res = self.client.post("/api/admin/exams", json={
            "title": "Kỳ thi Đánh giá Sections V2",
            "code": "SEC-V2",
            "num_questions": 2,
            "duration_minutes": 25,
            "allow_review": True,
            "shuffle_questions": True
        }, headers=self.admin_headers)
        self.assertEqual(ex_res.status_code, 200)
        exam_id = ex_res.json()["exam_id"]

        # Activate
        self.client.post(f"/api/admin/exams/{exam_id}/activate", headers=self.admin_headers)

        # 2. Add Questions: 3 MC + 1 Multi-Select + 2 Essay
        q1_res = self.client.post("/api/admin/questions", json={
            "exam_id": exam_id,
            "content": "MC Question 1",
            "question_type": "multiple_choice",
            "score_weight": 1.0,
            "option_a": "A1",
            "option_b": "B1",
            "correct_option": "A"
        }, headers=self.admin_headers)
        q1_id = q1_res.json()["question"]["id"]

        q2_res = self.client.post("/api/admin/questions", json={
            "exam_id": exam_id,
            "content": "MC Question 2",
            "question_type": "multiple_choice",
            "score_weight": 1.0,
            "option_a": "A2",
            "option_b": "B2",
            "correct_option": "B"
        }, headers=self.admin_headers)
        q2_id = q2_res.json()["question"]["id"]

        q3_res = self.client.post("/api/admin/questions", json={
            "exam_id": exam_id,
            "content": "Multi Select Question",
            "question_type": "multi_select",
            "score_weight": 2.0,
            "option_a": "Opt A",
            "option_b": "Opt B",
            "option_c": "Opt C",
            "correct_option": "A,C"
        }, headers=self.admin_headers)
        q3_id = q3_res.json()["question"]["id"]

        essay1_res = self.client.post("/api/admin/questions", json={
            "exam_id": exam_id,
            "content": "Essay Question 1",
            "question_type": "essay",
            "score_weight": 3.0
        }, headers=self.admin_headers)
        essay1_id = essay1_res.json()["question"]["id"]

        essay2_res = self.client.post("/api/admin/questions", json={
            "exam_id": exam_id,
            "content": "Essay Question 2",
            "question_type": "essay",
            "score_weight": 3.0
        }, headers=self.admin_headers)
        essay2_id = essay2_res.json()["question"]["id"]

        # 3. Candidate fetches exam
        exam_data_res = self.client.get("/api/exam", headers=self.student_headers)
        self.assertEqual(exam_data_res.status_code, 200)
        exam_data = exam_data_res.json()

        # Capped MC to 2 (from num_questions) + ALL 2 essays retained = 4 questions
        self.assertEqual(exam_data["mc_count"], 2)
        self.assertEqual(exam_data["essay_count"], 2)
        questions = exam_data["questions"]
        self.assertEqual(len(questions), 4)

        # First 2 questions must be MC/Multi-Select
        self.assertNotEqual(questions[0]["question_type"], "essay")
        self.assertNotEqual(questions[1]["question_type"], "essay")
        # Last 2 questions must be Essay in order
        self.assertEqual(questions[2]["question_type"], "essay")
        self.assertEqual(questions[3]["question_type"], "essay")
        self.assertEqual(questions[2]["id"], essay1_id)
        self.assertEqual(questions[3]["id"], essay2_id)

        # 4. Save progress with list (multi-select) and essay string
        mc1_id = questions[0]["id"]
        mc2_id = questions[1]["id"]
        save_res = self.client.post("/api/exam/save_progress", json={
            "answers": {
                str(mc1_id): "A",
                str(mc2_id): ["A", "C"],
                str(essay1_id): "Đây là bài làm tự luận câu 1 chi tiết...",
                str(essay2_id): "Đây là bài làm tự luận câu 2 chi tiết..."
            }
        }, headers=self.student_headers)
        self.assertEqual(save_res.status_code, 200)

        # 5. Submit exam
        sub_res = self.client.post("/api/exam/submit", json={
            "answers": {
                str(mc1_id): "A",
                str(mc2_id): ["A", "C"],
                str(essay1_id): "Đây là bài làm tự luận câu 1 chi tiết...",
                str(essay2_id): "Đây là bài làm tự luận câu 2 chi tiết..."
            }
        }, headers=self.student_headers)
        self.assertEqual(sub_res.status_code, 200)
        sub_data = sub_res.json()
        self.assertTrue(sub_data["needs_grading"])
        self.assertIsNone(sub_data["score"])

        # 6. Student Review endpoint
        rev_res = self.client.get("/api/exam/review", headers=self.student_headers)
        self.assertEqual(rev_res.status_code, 200)
        rev_data = rev_res.json()
        self.assertEqual(rev_data["mc_count"], 2)
        self.assertEqual(rev_data["essay_count"], 2)
        self.assertTrue(rev_data["has_pending_essay"])

        # 7. Admin Result Detail
        results_data = self.client.get(f"/api/admin/results?exam_id={exam_id}", headers=self.admin_headers).json()
        results_list = results_data.get("results", [])
        my_res = next(r for r in results_list if r["user_id"] == self.student.id)
        result_id = my_res["id"]

        detail_res = self.client.get(f"/api/admin/results/{result_id}/detail", headers=self.admin_headers)
        self.assertEqual(detail_res.status_code, 200)
        detail_data = detail_res.json()
        self.assertEqual(detail_data["mc_count"], 2)
        self.assertEqual(detail_data["essay_count"], 2)
        self.assertTrue(detail_data["has_pending_essay"])

        # 8. Grade Essays
        grade_res = self.client.post(f"/api/admin/results/{result_id}/score_essay", json={
            "scores": {
                str(essay1_id): 2.5,
                str(essay2_id): 2.0
            }
        }, headers=self.admin_headers)
        self.assertEqual(grade_res.status_code, 200)

        # 9. Verify updated score
        detail_after = self.client.get(f"/api/admin/results/{result_id}/detail", headers=self.admin_headers).json()
        self.assertFalse(detail_after["has_pending_essay"])
        self.assertIsNotNone(detail_after["score"])
        self.assertEqual(detail_after["essay_score"], 2.25)

        # 10. PDF Export with Part I and Part II
        pdf_bytes = generate_candidate_pdf(
            exam_info={
                "title": "Kỳ thi Sections V2",
                "code": "SEC-V2",
                "duration_minutes": 25
            },
            candidate_info={
                "fullname": self.student.fullname,
                "student_code": self.student.username,
                "dob": self.student.dob,
                "start_time_str": "08:00:00 30/09/2026",
                "submit_time_str": "08:25:00 30/09/2026",
                "duration_str": "25 phút",
                "status_str": "Đã hoàn thành",
                "score": detail_after["score"],
                "max_score": 10.0,
                "correct_cnt": detail_after["correct_count"],
                "total_q": 4
            },
            audit_questions=detail_after["questions"]
        )
        self.assertGreater(len(pdf_bytes), 10000)

        # Also verify the PDF export endpoint directly
        pdf_endpoint_res = self.client.get(f"/api/admin/results/{result_id}/export_pdf", headers=self.admin_headers)
        self.assertEqual(pdf_endpoint_res.status_code, 200)
        self.assertEqual(pdf_endpoint_res.headers["content-type"], "application/pdf")
        self.assertGreater(len(pdf_endpoint_res.content), 10000)

if __name__ == "__main__":
    unittest.main()
