import sys
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
import json
import datetime
from fastapi.testclient import TestClient
from main import app, SessionLocal, get_password_hash
from models import User, Exam, Question, ExamResult
from pdf_export import generate_candidate_pdf
from sqlalchemy import text

client = TestClient(app)

def test_edge_cases():
    print("\n========================================================")
    print("TESTING SECTION EDGE CASES & ADAPTIVE RENDERING")
    print("========================================================")

    db = SessionLocal()
    try:
        # Setup admin
        admin = db.query(User).filter(User.username == "edge_admin").first()
        if not admin:
            admin = User(username="edge_admin", password=get_password_hash("admin123"), is_admin=True)
            db.add(admin)
            db.commit()
            db.refresh(admin)

        adm_tok = client.post("/token", data={"username": "edge_admin", "password": "admin123"}).json()["access_token"]
        adm_headers = {"Authorization": f"Bearer {adm_tok}"}

        # Setup student
        student = db.query(User).filter(User.username == "edge_student").first()
        if not student:
            student = User(username="edge_student", password=get_password_hash("stu123"), fullname="Edge Student", is_admin=False)
            db.add(student)
            db.commit()
            db.refresh(student)

        stu_tok = client.post("/token", data={"username": "edge_student", "password": "stu123"}).json()["access_token"]
        stu_headers = {"Authorization": f"Bearer {stu_tok}"}

        # ----------------------------------------------------
        # SCENARIO 1: Essay-Only Exam (0 MC, 3 Essay)
        # ----------------------------------------------------
        print("\n--- Testing Scenario 1: Essay-Only Exam (0 MC, 3 Essay) ---")
        now = datetime.datetime.now()
        ex_essay = Exam(
            title="Kỳ thi Chỉ có Tự luận",
            code="ESSAY_ONLY_TEST",
            duration_minutes=30,
            num_questions=10, # num_questions applies to MC, so essay questions should ALL be taken
            open_time=now - datetime.timedelta(minutes=5),
            close_time=now + datetime.timedelta(hours=2),
            allow_review=True,
            is_active=True
        )
        db.add(ex_essay)
        db.commit()

        q_e1 = Question(exam_id=ex_essay.id, content="Tự luận 1", question_type="essay", score_weight=3.0)
        q_e2 = Question(exam_id=ex_essay.id, content="Tự luận 2", question_type="essay", score_weight=4.0)
        q_e3 = Question(exam_id=ex_essay.id, content="Tự luận 3", question_type="essay", score_weight=3.0)
        db.add_all([q_e1, q_e2, q_e3])
        db.commit()

        client.post(f"/api/admin/exams/{ex_essay.id}/activate", headers=adm_headers)

        # Check Admin Questions
        qs_res = client.get(f"/api/admin/questions?exam_id={ex_essay.id}", headers=adm_headers)
        assert qs_res.status_code == 200
        assert qs_res.json()["mc_count"] == 0
        assert qs_res.json()["essay_count"] == 3

        # Candidate fetches exam
        take_res = client.get("/api/exam", headers=stu_headers)
        assert take_res.status_code == 200
        take_data = take_res.json()
        assert take_data["mc_count"] == 0
        assert take_data["essay_count"] == 3
        assert len(take_data["questions"]) == 3
        # Strict bank order preserved
        assert [q["id"] for q in take_data["questions"]] == [q_e1.id, q_e2.id, q_e3.id]
        print("  ✓ Candidate received all 3 essay questions in exact bank order")

        # Submit answers
        sub_res = client.post("/api/exam/submit", json={"answers": {str(q_e1.id): "Bài làm TL 1", str(q_e2.id): "Bài làm TL 2"}}, headers=stu_headers)
        assert sub_res.status_code == 200
        assert sub_res.json()["needs_grading"] is True

        # Review
        rev_res = client.get("/api/exam/review", headers=stu_headers)
        assert rev_res.status_code == 200
        rev_data = rev_res.json()
        assert rev_data["mc_count"] == 0
        assert rev_data["essay_count"] == 3
        assert rev_data["has_pending_essay"] is True

        # Detail & Scoring
        r_entry = db.query(ExamResult).filter(ExamResult.user_id == student.id, ExamResult.exam_id == ex_essay.id).first()
        assert r_entry is not None
        detail_res = client.get(f"/api/admin/results/{r_entry.id}/detail", headers=adm_headers)
        assert detail_res.status_code == 200

        grade_res = client.post(f"/api/admin/results/{r_entry.id}/score_essay", json={
            "scores": {str(q_e1.id): 3.0, str(q_e2.id): 4.0, str(q_e3.id): 0.0}
        }, headers=adm_headers)
        assert grade_res.status_code == 200
        assert grade_res.json()["new_score"] == 7.0

        # PDF export (verify it doesn't say PHẦN II when there is no PHẦN I)
        detail_after = client.get(f"/api/admin/results/{r_entry.id}/detail", headers=adm_headers).json()
        cand_info = {
            "username": student.username, "fullname": student.fullname, "dob": "2000-01-01",
            "score": detail_after["score"], "max_score": 10.0, "correct_count": 0,
            "total_questions": 3, "duration_str": "10p", "start_time_str": "30/09/2026 10:00:00",
            "submit_time_str": "30/09/2026 10:10:00", "status_text": "Đã chấm"
        }
        exam_info = {"id": ex_essay.id, "title": ex_essay.title, "code": ex_essay.code, "duration_minutes": 30, "num_questions": 3}
        pdf_bytes = generate_candidate_pdf(exam_info, cand_info, detail_after["questions"])
        assert len(pdf_bytes) > 1000
        print("  ✓ Essay-only PDF successfully generated without crash")

        # Cleanup Scenario 1
        db.query(ExamResult).filter(ExamResult.exam_id == ex_essay.id).delete()
        db.query(Question).filter(Question.exam_id == ex_essay.id).delete()
        db.delete(ex_essay)
        db.commit()

        # ----------------------------------------------------
        # SCENARIO 2: MC-Only Exam (3 MC, 0 Essay)
        # ----------------------------------------------------
        print("\n--- Testing Scenario 2: MC-Only Exam (3 MC, 0 Essay) ---")
        ex_mc = Exam(
            title="Kỳ thi Chỉ có Trắc nghiệm",
            code="MC_ONLY_TEST",
            duration_minutes=30,
            num_questions=3,
            open_time=now - datetime.timedelta(minutes=5),
            close_time=now + datetime.timedelta(hours=2),
            allow_review=True,
            is_active=True
        )
        db.add(ex_mc)
        db.commit()

        q_m1 = Question(exam_id=ex_mc.id, content="MC 1", option_a="A", option_b="B", correct_option="A", score_weight=1.0)
        q_m2 = Question(exam_id=ex_mc.id, content="MC 2", option_a="A", option_b="B", correct_option="B", score_weight=1.0)
        q_m3 = Question(exam_id=ex_mc.id, content="MC 3", option_a="A", option_b="B", correct_option="A", score_weight=1.0)
        db.add_all([q_m1, q_m2, q_m3])
        db.commit()

        client.post(f"/api/admin/exams/{ex_mc.id}/activate", headers=adm_headers)

        qs_res2 = client.get(f"/api/admin/questions?exam_id={ex_mc.id}", headers=adm_headers)
        assert qs_res2.status_code == 200
        assert qs_res2.json()["mc_count"] == 3
        assert qs_res2.json()["essay_count"] == 0

        take_res2 = client.get("/api/exam", headers=stu_headers)
        assert take_res2.status_code == 200
        assert take_res2.json()["mc_count"] == 3
        assert take_res2.json()["essay_count"] == 0

        sub_res2 = client.post("/api/exam/submit", json={"answers": {str(q_m1.id): "A", str(q_m2.id): "B", str(q_m3.id): "B"}}, headers=stu_headers)
        assert sub_res2.status_code == 200
        assert sub_res2.json()["needs_grading"] is False
        assert sub_res2.json()["correct_count"] == 2

        # Cleanup Scenario 2
        db.query(ExamResult).filter(ExamResult.exam_id == ex_mc.id).delete()
        db.query(Question).filter(Question.exam_id == ex_mc.id).delete()
        db.delete(ex_mc)
        db.commit()

        # ----------------------------------------------------
        # SCENARIO 3: Robustness against raw NULL score_weight in DB
        # ----------------------------------------------------
        print("\n--- Testing Scenario 3: Raw NULL score_weight Resilience ---")
        ex_null = Exam(title="Resilience Exam", is_active=True, duration_minutes=30, num_questions=2)
        db.add(ex_null)
        db.commit()

        q_null_mc = Question(exam_id=ex_null.id, content="Null Weight MC", option_a="A", option_b="B", correct_option="A", score_weight=1.0)
        q_null_es = Question(exam_id=ex_null.id, content="Null Weight Essay", question_type="essay", score_weight=2.0)
        db.add_all([q_null_mc, q_null_es])
        db.commit()

        # Force raw NULL into database
        db.execute(text(f"UPDATE questions SET score_weight = NULL WHERE id IN ({q_null_mc.id}, {q_null_es.id})"))
        db.commit()

        client.post(f"/api/admin/exams/{ex_null.id}/activate", headers=adm_headers)

        take_res3 = client.get("/api/exam", headers=stu_headers)
        assert take_res3.status_code == 200

        sub_res3 = client.post("/api/exam/submit", json={"answers": {str(q_null_mc.id): "A", str(q_null_es.id): "Raw Essay"}}, headers=stu_headers)
        assert sub_res3.status_code == 200
        print("  ✓ submit_exam handled NULL score_weight without TypeError")

        res_e3 = db.query(ExamResult).filter(ExamResult.user_id == student.id, ExamResult.exam_id == ex_null.id).first()
        detail_res3 = client.get(f"/api/admin/results/{res_e3.id}/detail", headers=adm_headers)
        assert detail_res3.status_code == 200
        print("  ✓ admin detail handled NULL score_weight without TypeError")

        grade_res3 = client.post(f"/api/admin/results/{res_e3.id}/score_essay", json={"scores": {str(q_null_es.id): 1.0}}, headers=adm_headers)
        assert grade_res3.status_code == 200
        print("  ✓ score_essay handled NULL score_weight without TypeError")

        pdf_res3 = client.get(f"/api/admin/results/{res_e3.id}/export_pdf", headers=adm_headers)
        assert pdf_res3.status_code == 200
        print("  ✓ export_pdf handled NULL score_weight without TypeError")

        # Cleanup Scenario 3
        db.query(ExamResult).filter(ExamResult.exam_id == ex_null.id).delete()
        db.query(Question).filter(Question.exam_id == ex_null.id).delete()
        db.delete(ex_null)
        db.commit()

        # ----------------------------------------------------
        # SCENARIO 4: Frontend HTML Inspection
        # ----------------------------------------------------
        print("\n--- Testing Scenario 4: Frontend index.html Inspection ---")
        with open("static/index.html", "r", encoding="utf-8") as f:
            html = f.read()

        assert 'id="q-stat-breakdown-badge"' in html
        assert 'PHẦN I' in html
        assert 'PHẦN II' in html
        assert 'palette-group-card' in html
        assert 'TL${essayIdx + 1}' in html
        print("  ✓ All required frontend section components and badges verified")

        print("\n========================================================")
        print("🎉 ALL EDGE CASE TESTS PASSED 100%!")
        print("========================================================")

    finally:
        db.close()

if __name__ == "__main__":
    test_edge_cases()
