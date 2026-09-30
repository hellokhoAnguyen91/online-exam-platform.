import sys
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
import json
from fastapi.testclient import TestClient
from main import app, SessionLocal, get_password_hash
from models import User, Exam, Question, ExamResult
from pdf_export import generate_candidate_pdf

client = TestClient(app)

def test_mc_essay_separation():
    print("\n==========================================")
    print("VERIFYING MC & ESSAY SEPARATION LOGIC")
    print("==========================================")
    db = SessionLocal()
    prev_active_exam = db.query(Exam).filter(Exam.is_active == True).first()
    prev_active_id = prev_active_exam.id if prev_active_exam else None
    exam_id = None

    try:
        # 1. Setup admin user
        admin = db.query(User).filter(User.username == "sep_admin").first()
        if not admin:
            admin = User(
                username="sep_admin",
                password=get_password_hash("admin123"),
                fullname="Separation Admin",
                is_admin=True
            )
            db.add(admin)
            db.commit()
            db.refresh(admin)

        login_res = client.post("/token", data={"username": "sep_admin", "password": "admin123"})
        assert login_res.status_code == 200
        admin_token = login_res.json()["access_token"]
        admin_headers = {"Authorization": f"Bearer {admin_token}"}
        print("  ✓ Admin authenticated")

        # 2. Setup student user
        student = db.query(User).filter(User.username == "sep_student").first()
        if not student:
            student = User(
                username="sep_student",
                password=get_password_hash("student123"),
                fullname="Separation Student",
                dob="2004-01-01",
                is_admin=False
            )
            db.add(student)
            db.commit()
            db.refresh(student)

        # Clean up any prior test results for clean state
        db.query(ExamResult).filter(ExamResult.user_id == student.id).delete()
        db.commit()

        stu_login = client.post("/token", data={"username": "sep_student", "password": "student123"})
        assert stu_login.status_code == 200
        stu_token = stu_login.json()["access_token"]
        stu_headers = {"Authorization": f"Bearer {stu_token}"}
        print("  ✓ Student authenticated")

        # 3. Create Exam: num_questions=3 (for MC), shuffle_questions=True
        exam_res = client.post("/api/admin/exams", json={
            "title": "Kỳ thi Phân chia TN và TL",
            "code": "SEP-2026",
            "description": "Kiểm tra phân chia trắc nghiệm và tự luận",
            "num_questions": 3,
            "duration_minutes": 30,
            "allow_review": True,
            "shuffle_questions": True
        }, headers=admin_headers)
        assert exam_res.status_code == 200
        exam_id = exam_res.json()["exam_id"]
        print(f"  ✓ Created Exam ID: {exam_id}")

        # Activate Exam
        act_res = client.post(f"/api/admin/exams/{exam_id}/activate", headers=admin_headers)
        assert act_res.status_code == 200
        print("  ✓ Exam activated")

        # 4. Add 5 MC questions and 3 Essay questions
        mc_ids = []
        for i in range(1, 6):
            q_res = client.post("/api/admin/questions", json={
                "exam_id": exam_id,
                "content": f"Câu hỏi trắc nghiệm {i} về kiến thức ABC",
                "question_type": "multiple_choice",
                "score_weight": 1.0,
                "option_a": f"Phương án A{i}",
                "option_b": f"Phương án B{i}",
                "option_c": f"Phương án C{i}",
                "option_d": f"Phương án D{i}",
                "correct_option": "A",
                "explanation": f"Giải thích câu {i}"
            }, headers=admin_headers)
            assert q_res.status_code == 200
            mc_ids.append(q_res.json()["question"]["id"])

        essay_ids = []
        for i in range(1, 4):
            q_res = client.post("/api/admin/questions", json={
                "exam_id": exam_id,
                "content": f"Câu hỏi tự luận {i}: Phân tích vấn đề XYZ",
                "question_type": "essay",
                "score_weight": 2.5,
                "correct_option": f"Hướng dẫn chấm câu tự luận {i}",
                "explanation": f"Gợi ý đáp án tự luận {i}"
            }, headers=admin_headers)
            assert q_res.status_code == 200
            essay_ids.append(q_res.json()["question"]["id"])

        print(f"  ✓ Added 5 MC questions ({mc_ids}) and 3 Essay questions ({essay_ids})")

        # 5. Check Admin Questions API
        admin_q_res = client.get(f"/api/admin/questions?exam_id={exam_id}", headers=admin_headers)
        assert admin_q_res.status_code == 200
        admin_q_data = admin_q_res.json()
        assert admin_q_data["total"] == 8
        assert admin_q_data["mc_count"] == 5
        assert admin_q_data["essay_count"] == 3
        questions_list = admin_q_data["questions"]
        assert len(questions_list) == 8

        # Verify ordering: all MCs first, all Essays second
        for idx in range(5):
            assert questions_list[idx]["question_type"] == "multiple_choice", f"Expected MC at index {idx}"
        for idx in range(5, 8):
            assert questions_list[idx]["question_type"] == "essay", f"Expected Essay at index {idx}"

        # Verify exact essay order matches insertion order
        for idx, expected_id in enumerate(essay_ids):
            assert questions_list[5 + idx]["id"] == expected_id, f"Essay order mismatch at {idx}"

        print("  ✓ Admin Questions API returned correct counts and MC-first/Essay-second ordering")

        # 6. Check Student Exam Start API
        exam_get_res = client.get("/api/exam", headers=stu_headers)
        assert exam_get_res.status_code == 200
        candidate_data = exam_get_res.json()
        cand_questions = candidate_data["questions"]

        # Critical Rules verification:
        # Rule: num_questions=3 only caps MC questions
        # Rule: ALL essay questions are retained (never truncated)
        # Total candidate questions = 3 (MC) + 3 (Essay) = 6
        assert len(cand_questions) == 6, f"Expected 6 candidate questions, got {len(cand_questions)}"
        assert candidate_data["mc_count"] == 3
        assert candidate_data["essay_count"] == 3

        # First 3 MUST be multiple_choice
        for idx in range(3):
            assert cand_questions[idx]["question_type"] == "multiple_choice", f"Cand Q {idx} must be MC"
            assert cand_questions[idx]["id"] in mc_ids, f"cand question id {cand_questions[idx]['id']} not in {mc_ids}"

        # Last 3 MUST be essay
        for idx in range(3, 6):
            assert cand_questions[idx]["question_type"] == "essay", f"Cand Q {idx} must be essay"

        # Essay questions MUST retain exact bank order
        cand_essay_ids = [q["id"] for q in cand_questions[3:]]
        assert cand_essay_ids == essay_ids, f"Candidate essay order {cand_essay_ids} must match bank order {essay_ids}"
        print("  ✓ Candidate exam allocation strictly followed: MC capped to num_questions (3) & all Essay (3) preserved in exact bank order")

        # 7. Student saves progress and submits
        # Answer 2 MC and 2 Essay questions
        save_res = client.post("/api/exam/save_progress", json={
            "answers": {
                str(cand_questions[0]["id"]): "A",
                str(cand_questions[1]["id"]): "B",
                str(cand_questions[3]["id"]): "Bài làm tự luận câu 1 của sinh viên...",
                str(cand_questions[4]["id"]): "Bài làm tự luận câu 2 của sinh viên..."
            }
        }, headers=stu_headers)
        assert save_res.status_code == 200
        assert save_res.json()["status"] == "saved"
        print("  ✓ Candidate saved MC and Essay answers")

        # Submit exam
        submit_res = client.post("/api/exam/submit", json={
            "answers": {
                str(cand_questions[0]["id"]): "A",
                str(cand_questions[1]["id"]): "B",
                str(cand_questions[3]["id"]): "Bài làm tự luận câu 1 của sinh viên...",
                str(cand_questions[4]["id"]): "Bài làm tự luận câu 2 của sinh viên..."
            },
            "tab_switch_count": 0
        }, headers=stu_headers)
        assert submit_res.status_code == 200
        submit_data = submit_res.json()
        assert submit_data["status"] == "submitted"
        assert submit_data["needs_grading"] is True
        assert submit_data["score"] is None

        # Fetch result record from DB
        exam_result = db.query(ExamResult).filter(
            ExamResult.user_id == student.id,
            ExamResult.exam_id == exam_id
        ).order_by(ExamResult.id.desc()).first()
        assert exam_result is not None
        result_id = exam_result.id
        print(f"  ✓ Exam submitted with needs_grading=True, result_id: {result_id}")

        # 8. Check Student Review API (/api/exam/review)
        rev_res = client.get("/api/exam/review", headers=stu_headers)
        assert rev_res.status_code == 200
        rev_data = rev_res.json()
        assert rev_data["has_pending_essay"] is True
        assert rev_data["mc_count"] == 3
        assert rev_data["essay_count"] == 3
        rev_questions = rev_data["questions"]
        assert len(rev_questions) == 6
        # MC first, Essay second
        for idx in range(3):
            assert rev_questions[idx]["question_type"] == "multiple_choice"
        for idx in range(3, 6):
            assert rev_questions[idx]["question_type"] == "essay"
        print("  ✓ Review API maintained MC-first/Essay-second ordering and counts")

        # 9. Check Admin Result Detail API
        detail_res = client.get(f"/api/admin/results/{result_id}/detail", headers=admin_headers)
        assert detail_res.status_code == 200
        detail_data = detail_res.json()
        assert detail_data["mc_count"] == 3
        assert detail_data["essay_count"] == 3
        detail_q = detail_data["questions"]
        assert len(detail_q) == 6
        for idx in range(3):
            assert detail_q[idx]["question_type"] == "multiple_choice"
        for idx in range(3, 6):
            assert detail_q[idx]["question_type"] == "essay"
        print("  ✓ Admin result detail API verified")

        # 10. Admin grades essay question
        grade_res = client.post(f"/api/admin/results/{result_id}/score_essay", json={
            "scores": {
                str(cand_questions[3]["id"]): 2.5
            }
        }, headers=admin_headers)
        assert grade_res.status_code == 200
        grade_data = grade_res.json()
        assert "new_score" in grade_data
        print(f"  ✓ Admin essay grading verified (needs_grading: {grade_data['needs_grading']})")

        # 11. Test PDF Export
        pdf_res = client.get(f"/api/admin/results/{result_id}/export_pdf", headers=admin_headers)
        assert pdf_res.status_code == 200
        assert pdf_res.headers["content-type"] == "application/pdf"
        pdf_bytes = pdf_res.content
        assert len(pdf_bytes) > 1000
        print(f"  ✓ PDF export successfully generated via API ({len(pdf_bytes)} bytes)")

        # 12. Verify Frontend static/index.html UI elements
        with open("static/index.html", "r", encoding="utf-8") as f:
            html_content = f.read()

        assert "id=\"q-mc-badge\"" in html_content
        assert "id=\"q-essay-badge\"" in html_content
        assert "id=\"q-total-badge\"" in html_content
        assert "id=\"qbank-tab-mc\"" in html_content
        assert "id=\"qbank-tab-essay\"" in html_content
        assert "id=\"qbank-tab-all\"" in html_content
        assert "switchQBankTab" in html_content
        assert "palette-btn is-essay" in html_content or "is-essay" in html_content
        assert "Phần I: Trắc nghiệm" in html_content
        assert "Phần II: Tự luận" in html_content
        assert "TL" in html_content
        print("  ✓ static/index.html frontend elements verified")

        print("\n==========================================")
        print("ALL MC & ESSAY SEPARATION TESTS PASSED!")
        print("==========================================\n")

    finally:
        try:
            if exam_id:
                db.query(ExamResult).filter(ExamResult.exam_id == exam_id).delete(synchronize_session=False)
                db.query(Question).filter(Question.exam_id == exam_id).delete(synchronize_session=False)
                db.query(Exam).filter(Exam.id == exam_id).delete(synchronize_session=False)
                db.commit()
            if prev_active_id:
                db.query(Exam).filter(Exam.id != prev_active_id).update({"is_active": False})
                db.query(Exam).filter(Exam.id == prev_active_id).update({"is_active": True})
                db.commit()
        except Exception as e:
            print(f"Cleanup error in test: {e}")
        finally:
            db.close()

if __name__ == "__main__":
    test_mc_essay_separation()
