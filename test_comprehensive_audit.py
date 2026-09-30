import sys
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
import os
import io
import json
import zipfile
import datetime
from fastapi.testclient import TestClient
from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker

from main import app, get_db, get_password_hash
from database import Base
from models import User, Exam, Question, ExamResult
from pdf_export import generate_candidate_pdf, generate_batch_exam_zip

client = TestClient(app)

def test_audit_suite():
    print("\n==========================================")
    print("RUNNING COMPREHENSIVE AUDIT VERIFICATION")
    print("==========================================")

    # 1. Setup in-memory DB or use test app DB session
    from main import SessionLocal
    db = SessionLocal()

    try:
        # Create test admin
        admin = db.query(User).filter(User.username == "audit_admin").first()
        if not admin:
            admin = User(
                username="audit_admin",
                password=get_password_hash("admin123"),
                fullname="Audit Admin",
                is_admin=True
            )
            db.add(admin)
            db.commit()
            db.refresh(admin)

        admin_login = client.post("/token", data={"username": "audit_admin", "password": "admin123"})
        assert admin_login.status_code == 200, f"Admin login failed: {admin_login.text}"
        admin_token = admin_login.json()["access_token"]
        admin_headers = {"Authorization": f"Bearer {admin_token}"}
        print("  ✓ Admin auth verified")

        # 2. Test Student Login with MSSV as username AND MSSV as password
        mssv_stu = db.query(User).filter(User.username == "21110099").first()
        if not mssv_stu:
            mssv_stu = User(
                username="21110099",
                password=get_password_hash("21110099"),
                fullname="MSSV Test Student",
                dob="2003-05-15",
                is_admin=False
            )
            db.add(mssv_stu)
            db.refresh(mssv_stu)

        # Clear any prior results for test student to ensure clean state
        db.query(ExamResult).filter(ExamResult.user_id == mssv_stu.id).delete()
        db.commit()

        # Login using MSSV as password
        stu_login = client.post("/token", data={"username": "21110099", "password": "21110099"})
        assert stu_login.status_code == 200, f"Student login with MSSV as password failed: {stu_login.text}"
        stu_token = stu_login.json()["access_token"]
        stu_headers = {"Authorization": f"Bearer {stu_token}"}
        print("  ✓ Student login with MSSV as password verified")

        # Test Default passwords (123456, 12345678, hcmute) must be rejected with 401
        for bad_pwd in ["123456", "12345678", "hcmute"]:
            bad_login = client.post("/token", data={"username": "21110099", "password": bad_pwd})
            assert bad_login.status_code == 401, f"Default password '{bad_pwd}' was accepted, should be 401!"
        print("  ✓ Default fallback passwords (123456, 12345678, hcmute) successfully blocked (401)")

        # Test /api/exams requires authentication and filters inactive exams
        no_auth_exams = client.get("/api/exams")
        assert no_auth_exams.status_code == 401, f"Expected 401 for unauthenticated /api/exams, got {no_auth_exams.status_code}"
        stu_exams_res = client.get("/api/exams", headers=stu_headers)
        assert stu_exams_res.status_code == 200
        assert all(e["is_active"] for e in stu_exams_res.json()), "Student should only see active exams"
        print("  ✓ /api/exams security (auth required & active-only for student) verified")

        # 3. Create a dedicated Audit Exam
        now = datetime.datetime.now()
        open_time = (now - datetime.timedelta(minutes=10)).isoformat()
        close_time = (now + datetime.timedelta(hours=2)).isoformat()

        exam_res = client.post("/api/admin/exams", json={
            "title": "Comprehensive Audit Exam",
            "code": "AUDIT101",
            "duration_minutes": 45,
            "num_questions": 4,
            "open_time": open_time,
            "close_time": close_time,
            "allow_review": True,
            "shuffle_questions": False
        }, headers=admin_headers)
        assert exam_res.status_code == 200, f"Create exam failed: {exam_res.text}"
        exam_id = exam_res.json()["exam_id"]

        # Activate exam
        client.post(f"/api/admin/exams/{exam_id}/activate", headers=admin_headers)

        # Add 4 types of questions:
        # Q1: Multi-select with options A, B, C, D, E, correct = A,C,E (weight 3.0)
        q1 = Question(
            exam_id=exam_id,
            content="Q1: Chọn các đáp án đúng (A, C, E)",
            option_a="A đúng",
            option_b="B sai",
            option_c="C đúng",
            option_d="D sai",
            option_e="E đúng",
            option_f="F sai",
            correct_option="A,C,E",
            question_type="multi_select",
            score_weight=3.0
        )
        # Q2: Single choice with option E correct (weight 2.0)
        q2 = Question(
            exam_id=exam_id,
            content="Q2: Chọn đáp án E",
            option_a="A sai",
            option_b="B sai",
            option_c="C sai",
            option_d="D sai",
            option_e="E đúng nhất",
            correct_option="E",
            question_type="multiple_choice",
            score_weight=2.0
        )
        # Q3: Essay question (weight 2.0)
        q3 = Question(
            exam_id=exam_id,
            content="Q3: Trình bày quy trình an toàn lao động",
            question_type="essay",
            correct_option="",
            score_weight=2.0
        )
        # Q4: Essay question left completely blank (weight 3.0)
        q4 = Question(
            exam_id=exam_id,
            content="Q4: Phân tích sự cố hóa chất",
            question_type="essay",
            correct_option="",
            score_weight=3.0
        )
        db.add_all([q1, q2, q3, q4])
        db.commit()
        db.refresh(q1)
        db.refresh(q2)
        db.refresh(q3)
        db.refresh(q4)
        print("  ✓ Questions including options A-F, multi_select, essay created")

        # 4. Student takes exam:
        take_res = client.get("/api/exam", headers=stu_headers)
        assert take_res.status_code == 200, f"Get exam failed: {take_res.text}"
        data = take_res.json()
        assert data["status"] == "ongoing"
        assert len(data["questions"]) == 4

        # Test partial credit formula for Q1:
        # Correct set: {A, C, E} (total_corr = 3)
        # Student chooses: "A,C,D" (num_corr = 2 (A,C), num_wrong = 1 (D))
        # Formula: max(0.0, (2 - 1) / 3) * 3.0 = 1/3 * 3.0 = 1.0 earned pts.
        # Q2: Student chooses "E" -> 2.0 earned pts.
        # Q3: Student answers "Quy trình gồm 4 bước..." -> pending grading
        # Q4: Student leaves blank "" -> should be auto-zeroed to 0.0 pts!
        answers_payload = {
            str(q1.id): "A,C,D",
            str(q2.id): "E",
            str(q3.id): "Quy trình gồm 4 bước...",
            str(q4.id): ""
        }

        submit_res = client.post("/api/exam/submit", json={"answers": answers_payload}, headers=stu_headers)
        assert submit_res.status_code == 200, f"Submit failed: {submit_res.text}"
        sub_data = submit_res.json()
        assert sub_data["status"] == "submitted"
        assert sub_data["needs_grading"] is True
        assert sub_data["score"] is None, "Score should be None because Q3 is pending grading"
        print("  ✓ Exam submitted with pending essay and auto-zeroed blank essay")

        # Inspect ExamResult in DB
        result_entry = db.query(ExamResult).filter(ExamResult.user_id == mssv_stu.id, ExamResult.exam_id == exam_id).first()
        assert result_entry is not None
        details = json.loads(result_entry.answers_detail)

        # Verify Q1 partial credit
        d1 = next(d for d in details if d["id"] == q1.id)
        assert d1["status"] == "partial", f"Q1 status expected partial, got {d1['status']}"
        assert abs(d1["earned_score"] - 1.0) < 1e-4, f"Q1 earned score expected 1.0, got {d1['earned_score']}"
        print(f"  ✓ Multi-select partial credit formula verified: {d1['earned_score']} / {q1.score_weight} pts")

        # Verify Q2 single choice with option E
        d2 = next(d for d in details if d["id"] == q2.id)
        assert d2["status"] == "correct"
        assert abs(d2["earned_score"] - 2.0) < 1e-4
        print(f"  ✓ Option E single-choice verified: {d2['earned_score']} / {q2.score_weight} pts")

        # Verify Q3 pending essay
        d3 = next(d for d in details if d["id"] == q3.id)
        assert d3["status"] == "pending_grading"
        print("  ✓ Filled essay status verified: pending_grading")

        # Verify Q4 blank essay auto-zeroed
        d4 = next(d for d in details if d["id"] == q4.id)
        assert d4["status"] == "unanswered"
        assert d4["earned_score"] == 0.0
        print("  ✓ Blank essay auto-zeroed to 0.0 pts verified")

        # 5. Teacher Essay Scoring & Validation Bounds:
        # Score weight for Q3 is 2.0. Trying to give 2.5 or -0.5 must be rejected!
        bad_score_res1 = client.post(f"/api/admin/results/{result_entry.id}/score_essay", json={
            "scores": {str(q3.id): 2.5}
        }, headers=admin_headers)
        assert bad_score_res1.status_code == 400, "Should reject score > score_weight"

        bad_score_res2 = client.post(f"/api/admin/results/{result_entry.id}/score_essay", json={
            "scores": {str(q3.id): -1.0}
        }, headers=admin_headers)
        assert bad_score_res2.status_code == 400, "Should reject negative score"
        print("  ✓ Manual essay grading bounds (0.0 <= score <= score_weight) verified")

        # Now score Q3 with valid score 1.5
        good_score_res = client.post(f"/api/admin/results/{result_entry.id}/score_essay", json={
            "scores": {str(q3.id): 1.5}
        }, headers=admin_headers)
        assert good_score_res.status_code == 200, f"Valid essay score failed: {good_score_res.text}"
        
        # With 7.0 MC / 3.0 Essay default split:
        # MC: raw earned (1.0 + 2.0) / raw total 5.0 * 7.0 = 4.2
        # Essay: raw earned 1.5 / raw total 5.0 * 3.0 = 0.9
        # Total = 4.2 + 0.9 = 5.1
        db.refresh(result_entry)
        assert abs(result_entry.score - 5.1) < 1e-4, f"Final score expected 5.1, got {result_entry.score}"
        assert result_entry.max_score == 10.0
        print(f"  ✓ Essay scoring complete: final score = {result_entry.score} / {result_entry.max_score}")

        # 6. Test Review Masking when exam close_time is in the future
        # Exam close_time is still in the future! Correct answer must be masked (None)
        review_res1 = client.get("/api/exam/review", headers=stu_headers)
        assert review_res1.status_code == 200, f"Review failed: {review_res1.text}"
        rev_data1 = review_res1.json()
        assert len(rev_data1["questions"]) == 4
        for rq in rev_data1["questions"]:
            assert rq["correct"] is None, f"Correct answer was NOT masked while exam is still open! {rq}"
        print("  ✓ Review masking verified: correct answers hidden while exam is open")

        # Now update exam close_time to past and re-verify review unmasks correct answers
        exam_obj = db.query(Exam).filter(Exam.id == exam_id).first()
        exam_obj.close_time = now - datetime.timedelta(minutes=5)
        db.commit()

        review_res2 = client.get("/api/exam/review", headers=stu_headers)
        assert review_res2.status_code == 200
        rev_data2 = review_res2.json()
        rq1 = next(q for q in rev_data2["questions"] if q["id"] == q1.id)
        assert rq1["correct"] == "A,C,E", f"Expected 'A,C,E', got {rq1['correct']}"
        print("  ✓ Review unmasking verified: correct answers visible after exam close_time")

        # Test Review Masking when close_time is None:
        # If exam is still active, answers MUST be masked
        exam_obj.close_time = None
        exam_obj.is_active = True
        db.commit()
        review_res_none_active = client.get("/api/exam/review", headers=stu_headers)
        assert review_res_none_active.status_code == 200
        for rq in review_res_none_active.json()["questions"]:
            assert rq["correct"] is None, "Answers leaked when close_time is None and exam is active!"
        print("  ✓ Review masking verified: correct answers hidden when close_time is None and exam is active")

        # When deactivated (not active), answers should be unmasked
        exam_obj.is_active = False
        db.commit()
        review_res_none_inactive = client.get("/api/exam/review", headers=stu_headers)
        assert review_res_none_inactive.status_code == 200
        rq1_none = next(q for q in review_res_none_inactive.json()["questions"] if q["id"] == q1.id)
        assert rq1_none["correct"] == "A,C,E"
        print("  ✓ Review unmasking verified: correct answers visible when close_time is None and exam is inactive")

        # Restore exam_obj state for remaining tests
        exam_obj.close_time = now - datetime.timedelta(minutes=5)
        exam_obj.is_active = True
        db.commit()

        # Test sync_expired_sessions does NOT alter submitted session duration or submit_time
        from main import sync_expired_sessions
        res_submitted = db.query(ExamResult).filter(ExamResult.user_id == mssv_stu.id, ExamResult.exam_id == exam_id).first()
        assert res_submitted is not None
        test_duration = 7200
        test_submit = datetime.datetime(2026, 1, 1, 12, 0, 0)
        res_submitted.duration_seconds = test_duration
        res_submitted.submit_time = test_submit
        db.commit()

        sync_expired_sessions(db, exam_id)
        db.refresh(res_submitted)
        assert res_submitted.duration_seconds == test_duration, "sync_expired_sessions illegally capped submitted duration!"
        assert res_submitted.submit_time == test_submit, "sync_expired_sessions illegally modified submitted submit_time!"
        print("  ✓ Data preservation verified: sync_expired_sessions does NOT modify submitted sessions")

        # Test Backup & Restore preserves question_type, score_weight, option_e, option_f
        backup_res = client.get(f"/api/admin/exams/{exam_id}/backup", headers=admin_headers)
        assert backup_res.status_code == 200
        backup_json = backup_res.json()
        assert any(q.get("question_type") == "multi_select" and q.get("option_e") for q in backup_json["questions"])
        assert any(q.get("question_type") == "essay" and q.get("score_weight") == 2.0 for q in backup_json["questions"])

        restore_payload = io.BytesIO(json.dumps(backup_json).encode("utf-8"))
        restore_res = client.post("/api/admin/exams/restore", files={"file": ("backup.json", restore_payload, "application/json")}, headers=admin_headers)
        assert restore_res.status_code == 200
        restored_exam_id = restore_res.json()["exam_id"]
        restored_qs = db.query(Question).filter(Question.exam_id == restored_exam_id).all()
        assert len(restored_qs) == 4
        rq1_restored = next(q for q in restored_qs if q.question_type == "multi_select")
        assert rq1_restored.option_e == "E đúng"
        assert rq1_restored.score_weight == 3.0
        db.query(Question).filter(Question.exam_id == restored_exam_id).delete()
        db.query(Exam).filter(Exam.id == restored_exam_id).delete()
        db.commit()
        print("  ✓ Backup & Restore preserves question_type, score_weight, option_e, option_f verified")

        # 7. Test PDF Generation with Options A-F and Multi-Select & Essay
        cand_info = {
            "username": mssv_stu.username,
            "fullname": mssv_stu.fullname,
            "dob": mssv_stu.dob,
            "score": result_entry.score,
            "max_score": 10.0,
            "correct_count": result_entry.correct_count,
            "total_questions": 4,
            "duration_str": "12p 30s",
            "start_time_str": "30/09/2026 02:00:00",
            "submit_time_str": "30/09/2026 02:12:30",
            "status_text": "Đã nộp bài"
        }
        exam_info = {
            "id": exam_id,
            "title": exam_obj.title,
            "code": exam_obj.code,
            "duration_minutes": exam_obj.duration_minutes,
            "num_questions": 4
        }
        pdf_bytes = generate_candidate_pdf(exam_info, cand_info, details)
        assert len(pdf_bytes) > 1000, "PDF generation produced empty or truncated output"
        assert pdf_bytes.startswith(b"%PDF"), "Output is not a valid PDF document"
        print(f"  ✓ Candidate PDF export generated successfully ({len(pdf_bytes)} bytes)")

        # Test Batch ZIP Generation
        zip_bytes, exported_count = generate_batch_exam_zip(exam_obj, [result_entry], db)
        assert exported_count == 1
        with zipfile.ZipFile(io.BytesIO(zip_bytes)) as zf:
            namelist = zf.namelist()
            assert all(n.endswith(".pdf") for n in namelist), "PDF ZIP must contain only .pdf files!"
            assert any(n.endswith(".pdf") for n in namelist)
        print(f"  ✓ Batch ZIP export generated successfully ({len(zip_bytes)} bytes)")

        # 8. Test Reset Student Exam endpoint
        reset_res = client.post(f"/api/admin/students/{mssv_stu.id}/reset_exam?exam_id={exam_id}", headers=admin_headers)
        assert reset_res.status_code == 200, f"Reset exam failed: {reset_res.text}"
        
        # Verify ExamResult is deleted
        check_res = db.query(ExamResult).filter(ExamResult.user_id == mssv_stu.id, ExamResult.exam_id == exam_id).first()
        assert check_res is None, "ExamResult should be deleted after reset"
        print("  ✓ Admin reset student exam session verified")

        print("\n🎉 ALL AUDIT REQUIREMENTS FULLY PASSED AND VERIFIED 100%!")

    finally:
        # Clean up test exam and reactivate Exam 123
        try:
            db.query(ExamResult).filter(ExamResult.exam_id == exam_id).delete()
            db.query(Question).filter(Question.exam_id == exam_id).delete()
            db.query(Exam).filter(Exam.id == exam_id).delete()
            ex123 = db.query(Exam).filter(Exam.id == 123).first()
            if ex123:
                ex123.is_active = True
            db.commit()
        except Exception:
            pass
        db.close()

if __name__ == "__main__":
    test_audit_suite()
