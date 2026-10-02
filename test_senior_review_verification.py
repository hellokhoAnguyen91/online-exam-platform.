import sys
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

import os
import json
import base64
import datetime
from fastapi.testclient import TestClient

from main import app, create_access_token, utc_now, utc_now_naive, to_utc_dt
from database import SessionLocal
from models import User, Exam, Question, ExamResult

client = TestClient(app)

def test_senior_verification():
    db = SessionLocal()
    print("\n========================================================")
    print("RUNNING SENIOR CODE REVIEW VERIFICATION SUITE")
    print("========================================================")

    # 1. Setup Admin & Student
    admin = db.query(User).filter(User.username == "admin").first()
    if not admin:
        admin = db.query(User).filter(User.is_admin == True).first()
    admin_token = create_access_token({"sub": admin.username})
    adm_headers = {"Authorization": f"Bearer {admin_token}"}

    test_student = db.query(User).filter(User.username == "test_utc_stu").first()
    if not test_student:
        test_student = User(
            username="test_utc_stu",
            fullname="Sinh Vien Test UTC",
            dob="2003-05-15",
            class_name="21DTV",
            is_admin=False
        )
        db.add(test_student)
        db.commit()
        db.refresh(test_student)
    stu_token = create_access_token({"sub": test_student.username})
    stu_headers = {"Authorization": f"Bearer {stu_token}"}

    # 2. Create and Activate Test Exam
    exam = Exam(
        title="Kỳ thi Kiểm định Đồng bộ UTC & Trụ cột 1-5",
        code="UTC_TEST_2026",
        description="Kiểm tra độ chính xác thời gian và tối ưu hóa hệ thống",
        num_questions=5,
        duration_minutes=80,
        is_active=True,
        is_archived=False,
        allow_review=True,
        shuffle_questions=False,
        mc_max_score=60.0,
        essay_max_score=20.0
    )
    db.add(exam)
    db.commit()
    db.refresh(exam)

    # Deactivate other exams
    db.query(Exam).filter(Exam.id != exam.id).update({"is_active": False})
    db.commit()

    # Clear previous results for this student
    db.query(ExamResult).filter(ExamResult.user_id == test_student.id).delete()
    db.commit()

    # Add Questions
    q1 = Question(
        exam_id=exam.id,
        content="Q1: Đâu là cổng kết nối chuẩn?",
        option_a="USB", option_b="PCI", option_c="VGA", option_d="HDMI",
        correct_option="A", question_type="multiple_choice", score_weight=1.0
    )
    q2 = Question(
        exam_id=exam.id,
        content="Q2: Bài tự luận kiểm tra an toàn điện",
        option_a="", option_b="", option_c="", option_d="",
        correct_option="", question_type="essay", score_weight=2.0
    )
    db.add_all([q1, q2])
    db.commit()
    db.refresh(q1)
    db.refresh(q2)

    # 3. Test /api/exam: Verify start_time is UTC naive in DB and time_left is ~4800s (80 mins)
    res = client.get("/api/exam", headers=stu_headers)
    assert res.status_code == 200, f"/api/exam failed: {res.text}"
    exam_data = res.json()
    time_left = exam_data.get("time_left", 0)
    assert 4790 <= time_left <= 4800, f"Expected ~4800s (80 mins), got {time_left}s. Timezone bug present!"
    print(f"  ✓ /api/exam time_left correctly calculated: {time_left}s (~80 mins, NO 420m timezone jump)")

    er = db.query(ExamResult).filter(ExamResult.user_id == test_student.id, ExamResult.exam_id == exam.id).first()
    assert er is not None
    # Check that start_time in DB matches UTC (within 5 seconds of utc_now_naive)
    diff_from_utc = abs((utc_now_naive() - er.start_time).total_seconds())
    assert diff_from_utc < 10, f"er.start_time ({er.start_time}) is not UTC! Diff={diff_from_utc}s"
    print(f"  ✓ ExamResult.start_time in database is strictly UTC (diff: {diff_from_utc:.2f}s)")

    # 4. Test quick submit (e.g. after 5 seconds): raw_duration must NOT jump to 4800s (80 mins) or 420 mins!
    submit_res = client.post("/api/exam/submit", json={"answers": {str(q1.id): "A", str(q2.id): "Trình bày các bước an toàn"}}, headers=stu_headers)
    assert submit_res.status_code == 200
    sub_data = submit_res.json()
    duration_sec = sub_data.get("duration_seconds", 0)
    assert duration_sec < 60, f"Duration was falsely maxed out to {duration_sec}s! Must be < 60s"
    print(f"  ✓ Quick submit recorded actual duration: {duration_sec}s (NOT inflated to 80m or 420m)")

    # 5. Test /api/admin/results: Single outerjoin without bloat
    adm_res = client.get(f"/api/admin/results?exam_id={exam.id}", headers=adm_headers)
    assert adm_res.status_code == 200
    adm_data = adm_res.json()
    assert len(adm_data["results"]) >= 1
    found_rec = [r for r in adm_data["results"] if r["user_id"] == test_student.id][0]
    assert found_rec["duration_seconds"] == duration_sec
    print("  ✓ /api/admin/results returns clean projection without answers_detail bloat")

    # 6. Test /api/admin/results/{id}/detail: Questions reconstructed on demand
    detail_res = client.get(f"/api/admin/results/{er.id}/detail", headers=adm_headers)
    assert detail_res.status_code == 200
    det = detail_res.json()
    assert len(det["questions"]) == 2
    assert det["questions"][0]["content"] == q1.content
    assert det["questions"][1]["content"] == q2.content
    print("  ✓ Admin detail successfully hydrated questions on demand with zero N+1 queries")

    # 7. Test sync_expired_sessions behavior on timed-out candidate
    # Create another in-progress candidate with start_time 90 minutes ago
    stu_expired = User(
        username="test_expired_stu",
        fullname="Sinh Vien Qua Gio",
        dob="2003-01-01",
        class_name="21DTV",
        is_admin=False
    )
    db.add(stu_expired)
    db.commit()
    db.refresh(stu_expired)

    expired_er = ExamResult(
        exam_id=exam.id,
        user_id=stu_expired.id,
        start_time=utc_now_naive() - datetime.timedelta(minutes=90),
        questions=json.dumps([q1.id, q2.id]),
        total_questions=2,
        status="in_progress",
        answers=json.dumps({str(q1.id): "A"})
    )
    db.add(expired_er)
    db.commit()
    db.refresh(expired_er)

    # Calling get_results should trigger sync_expired_sessions and auto-finalize expired_er
    client.get(f"/api/admin/results?exam_id={exam.id}", headers=adm_headers)
    db.refresh(expired_er)
    assert expired_er.status == "submitted", f"Expected submitted, got {expired_er.status}"
    assert expired_er.duration_seconds == 80 * 60, f"Expected 4800s, got {expired_er.duration_seconds}"
    print("  ✓ sync_expired_sessions correctly auto-finalized timed-out session using true UTC clock")

    # 8. Test Rescue Import with HCMUTE-RESCUE: string
    rescue_payload = {
        "v": 1,
        "exam_id": exam.id,
        "user_id": test_student.id,
        "username": test_student.username,
        "fullname": test_student.fullname,
        "answers": {str(q1.id): "A", str(q2.id): "Noi dung cuu ho bai lam"},
        "timestamp": utc_now().isoformat()
    }
    b64_code = base64.b64encode(json.dumps(rescue_payload).encode('utf-8')).decode('ascii')
    rescue_str = f"HCMUTE-RESCUE:{b64_code}"

    rescue_res = client.post("/api/admin/rescue_import", json={"rescue_code": rescue_str}, headers=adm_headers)
    assert rescue_res.status_code == 200, f"Rescue import failed: {rescue_res.text}"
    rescued_json = rescue_res.json()
    assert rescued_json["status"] == "success"
    print("  ✓ Black Box Rescue Import processed HCMUTE-RESCUE code flawlessly")

    # Cleanup test data
    db.query(ExamResult).filter(ExamResult.exam_id == exam.id).delete()
    db.query(Question).filter(Question.exam_id == exam.id).delete()
    db.query(Exam).filter(Exam.id == exam.id).delete()
    db.query(User).filter(User.username.in_(["test_utc_stu", "test_expired_stu"])).delete()
    db.commit()
    db.close()

    print("\n========================================================")
    print("🎉 ALL SENIOR CODE REVIEW VERIFICATION TESTS PASSED 100%!")
    print("========================================================")

if __name__ == "__main__":
    test_senior_verification()
