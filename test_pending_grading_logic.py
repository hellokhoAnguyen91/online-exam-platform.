import unittest
import json
import io
import datetime
from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker
from models import Base, User, Exam, Question, ExamResult
from main import auto_close_exam_result, EssayScoreUpdate, score_essay
from pdf_export import generate_candidate_pdf, generate_batch_exam_zip

class TestPendingGrading(unittest.TestCase):
    def setUp(self):
        self.engine = create_engine('sqlite:///:memory:', connect_args={"check_same_thread": False})
        Base.metadata.create_all(bind=self.engine)
        self.Session = sessionmaker(bind=self.engine)
        self.db = self.Session()

        # Create admin and student
        self.admin = User(username="admin", fullname="Teacher", is_admin=True, password="hash")
        self.student = User(username="23150014", fullname="Huynh Vo Phuc An", is_admin=False, password="hash", dob="26/05/2005", class_name="23150C")
        self.db.add_all([self.admin, self.student])
        self.db.commit()

        # Create exam
        self.exam = Exam(
            title="Kỳ thi kiểm tra An Toàn",
            duration_minutes=30,
            num_questions=3,
            is_active=True
        )
        self.db.add(self.exam)
        self.db.commit()

        # Q1: Multiple choice WITH answer
        self.q1 = Question(
            exam_id=self.exam.id,
            content="Q1: Đâu là thiết bị PPE?",
            option_a="Mũ", option_b="Nón lá", option_c="Áo thun", option_d="Quần short",
            correct_option="A",
            question_type="multiple_choice",
            score_weight=1.0
        )
        # Q2: Essay question
        self.q2 = Question(
            exam_id=self.exam.id,
            content="Q2: Trình bày quy trình LOTO?",
            option_a="", option_b="", option_c="", option_d="",
            correct_option="",
            question_type="essay",
            score_weight=2.0
        )
        # Q3: Multiple choice WITHOUT answer key in question bank
        self.q3 = Question(
            exam_id=self.exam.id,
            content="Q3: Biện pháp xử lý sự cố tràn dầu?",
            option_a="Cách ly", option_b="Dùng phao", option_c="Báo cáo", option_d="Tất cả",
            correct_option="",
            question_type="multiple_choice",
            score_weight=1.0
        )
        self.db.add_all([self.q1, self.q2, self.q3])
        self.db.commit()

    def tearDown(self):
        self.db.close()

    def test_auto_close_with_pending_questions_yields_none_score(self):
        # Student answered all questions
        answers = {
            str(self.q1.id): "A",
            str(self.q2.id): "Quy trình LOTO gồm 6 bước...",
            str(self.q3.id): "D"
        }
        res = ExamResult(
            user_id=self.student.id,
            exam_id=self.exam.id,
            start_time=datetime.datetime.now() - datetime.timedelta(minutes=35),
            questions=json.dumps([self.q1.id, self.q2.id, self.q3.id]),
            answers=json.dumps(answers),
            status="in_progress"
        )
        self.db.add(res)
        self.db.commit()

        # Auto-close
        closed = auto_close_exam_result(res, self.db, self.exam)
        self.assertIsNone(closed.score, "Score must be None when essay or unkeyed questions exist")
        self.assertEqual(closed.correct_count, 1)

    def test_score_essay_updates_score_and_clears_pending(self):
        # Create result with pending status
        audit_details = [
            {"id": self.q1.id, "q_idx": 1, "is_correct": True, "status": "correct"},
            {"id": self.q2.id, "q_idx": 2, "is_correct": False, "status": "pending_grading"},
            {"id": self.q3.id, "q_idx": 3, "is_correct": False, "status": "pending_grading"}
        ]
        res = ExamResult(
            user_id=self.student.id,
            exam_id=self.exam.id,
            start_time=datetime.datetime.now() - datetime.timedelta(minutes=20),
            submit_time=datetime.datetime.now(),
            questions=json.dumps([self.q1.id, self.q2.id, self.q3.id]),
            answers=json.dumps({str(self.q1.id): "A", str(self.q2.id): "Text", str(self.q3.id): "D"}),
            answers_detail=json.dumps(audit_details),
            score=None,
            correct_count=1,
            total_questions=3,
            max_score=4.0,
            status="submitted"
        )
        self.db.add(res)
        self.db.commit()

        # Grade Q2 only (Q3 still pending)
        update1 = EssayScoreUpdate(scores={str(self.q2.id): 2.0})
        res1 = score_essay(res.id, update1, self.db, self.admin)
        self.assertTrue(res1["needs_grading"])
        self.assertIsNone(res1["new_score"], "Score should remain None until all pending questions are graded")

        # Grade Q3 (now all pending are graded)
        update2 = EssayScoreUpdate(scores={str(self.q3.id): 1.0})
        res2 = score_essay(res.id, update2, self.db, self.admin)
        self.assertFalse(res2["needs_grading"])
        # Total score: Q1(1.0) + Q2(2.0) + Q3(1.0) = 4.0 / 4.0 * 10 = 10.0
        self.assertEqual(res2["new_score"], 10.0)

        # Check DB
        self.db.refresh(res)
        self.assertEqual(res.score, 10.0)

    def test_pdf_export_with_none_score(self):
        candidate_info = {
            "username": "23150014",
            "fullname": "Huynh Vo Phuc An",
            "dob": "26/05/2005",
            "score": None,
            "max_score": 10.0,
            "correct_count": 0,
            "total_questions": 3,
            "duration_str": "15p 30s",
            "start_time_str": "30/09/2026 08:00:00",
            "submit_time_str": "30/09/2026 08:15:30",
            "status_text": "Đã nộp bài"
        }
        exam_info = {
            "title": "Kiểm tra giữa kỳ",
            "code": "ATLD2026"
        }
        audit_questions = [
            {
                "q_idx": 1,
                "content": "Câu 1: An toàn lao động?",
                "option_a": "A", "option_b": "B", "option_c": "C", "option_d": "D",
                "selected": "A", "correct": "",
                "status": "pending_grading"
            }
        ]
        # Must not raise any exceptions
        pdf_bytes = generate_candidate_pdf(exam_info, candidate_info, audit_questions)
        self.assertGreater(len(pdf_bytes), 1000)

if __name__ == '__main__':
    unittest.main()
