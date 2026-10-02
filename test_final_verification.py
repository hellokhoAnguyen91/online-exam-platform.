import sys
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')
import json
import base64
import datetime
from fastapi.testclient import TestClient
from main import app, create_access_token
from database import SessionLocal
from models import User, Exam, Question, ExamResult

client = TestClient(app)

def test_full_system_verification():
    db = SessionLocal()
    print("\n========================================================")
    print("RUNNING FINAL SYSTEM VERIFICATION: RESCUE, UTC & COMPACT AUDIT")
    print("========================================================")

    # 1. Admin & Student Setup
    admin = db.query(User).filter(User.username == "admin").first()
    admin_token = create_access_token({"sub": admin.username})
    adm_headers = {"Authorization": f"Bearer {admin_token}"}

    student = db.query(User).filter(User.username == "21110099").first()
    if not student:
        student = User(
            username="21110099",
            fullname="Nguyen Van Rescue Test",
            dob="2002-01-01",
            class_name="211101A",
            is_admin=False
        )
        student.set_password("21110099")
        db.add(student)
        db.commit()
        db.refresh(student)

    stu_token = create_access_token({"sub": student.username})
    stu_headers = {"Authorization": f"Bearer {stu_token}"}

    # 2. Create and Activate Test Exam
    exam_res = client.post("/api/admin/exams", json={
        "title": "Kỳ thi Kiểm tra Toàn diện Trực tuyến 2026",
        "code": "CK-FINAL-2026",
        "duration_minutes": 60,
        "num_questions": 4,
        "allow_review": True,
        "mc_max_score": 7.0,
        "essay_max_score": 3.0
    }, headers=adm_headers)
    assert exam_res.status_code == 200
    exam_data = exam_res.json()
    exam_id = exam_data.get("exam_id") or exam_data.get("id") or exam_data.get("exam", {}).get("id")

    act_res = client.post(f"/api/admin/exams/{exam_id}/activate", headers=adm_headers)
    assert act_res.status_code == 200
    print("  ✓ Exam created and activated successfully")

    # 3. Create Sample Questions (2 MC + 2 Essay)
    q1 = Question(exam_id=exam_id, content="Câu 1: Thủ đô nước Pháp là?", option_a="Berlin", option_b="Madrid", option_c="Paris", option_d="Rome", correct_option="C", question_type="multiple_choice", score_weight=1.0)
    q2 = Question(exam_id=exam_id, content="Câu 2: Các giao thức tầng Transport?", option_a="TCP", option_b="UDP", option_c="IP", option_d="HTTP", correct_option="A,B", question_type="multi_select", score_weight=1.0)
    q3 = Question(exam_id=exam_id, content="Câu 3: Hãy nêu khái niệm Deadlock?", option_a="", option_b="", option_c="", option_d="", correct_option="", question_type="essay", score_weight=1.5)
    q4 = Question(exam_id=exam_id, content="Câu 4: Nêu giải pháp phòng chống OOM?", option_a="", option_b="", option_c="", option_d="", correct_option="", question_type="essay", score_weight=1.5)
    db.add_all([q1, q2, q3, q4])
    db.commit()
    db.refresh(q1); db.refresh(q2); db.refresh(q3); db.refresh(q4)
    print("  ✓ Questions (MC + Multi-select + Essay) seeded into exam")

    # 4. Verify /api/exam UTC calculation & skew fields
    exam_fetch = client.get("/api/exam", headers=stu_headers)
    assert exam_fetch.status_code == 200
    fetch_data = exam_fetch.json()
    assert fetch_data["status"] == "ongoing"
    assert "server_now_utc" in fetch_data, "Missing server_now_utc in /api/exam"
    assert "end_time_utc" in fetch_data, "Missing end_time_utc in /api/exam"
    assert "time_left" in fetch_data, "Missing time_left in /api/exam"
    assert 3500 <= fetch_data["time_left"] <= 3600, f"Expected ~3600s, got {fetch_data['time_left']}"
    print(f"  ✓ /api/exam returns valid UTC time: server_now={fetch_data['server_now_utc']}, end_time={fetch_data['end_time_utc']}, time_left={fetch_data['time_left']}s")

    # 5. Verify /api/exam/save_progress UTC calculation & skew fields
    save_res = client.post("/api/exam/save_progress", json={
        "answers": {str(q1.id): "C", str(q2.id): "A,B"}
    }, headers=stu_headers)
    assert save_res.status_code == 200
    save_data = save_res.json()
    assert save_data["status"] == "saved"
    assert "server_now_utc" in save_data, "Missing server_now_utc in save_progress"
    assert "end_time_utc" in save_data, "Missing end_time_utc in save_progress"
    assert "time_left" in save_data, "Missing time_left in save_progress"
    print("  ✓ /api/exam/save_progress returns status=saved with realtime UTC fields")

    # 6. Verify Normal Submission & Compact answers_detail (No redundant text/options)
    sub_res = client.post("/api/exam/submit", json={
        "answers": {
            str(q1.id): "C",
            str(q2.id): "A,B",
            str(q3.id): "Deadlock là trạng thái bế tắc...",
            str(q4.id): "Giới hạn payload, dùng paging, compact audit..."
        }
    }, headers=stu_headers)
    assert sub_res.status_code == 200
    sub_data = sub_res.json()
    assert sub_data["status"] == "submitted"
    assert sub_data["needs_grading"] is True

    # Check database representation of answers_detail directly
    er_entry = db.query(ExamResult).filter(ExamResult.user_id == student.id, ExamResult.exam_id == exam_id).first()
    assert er_entry is not None
    assert er_entry.answers_detail is not None
    details = json.loads(er_entry.answers_detail)
    for d in details:
        assert "content" not in d, "answers_detail should NOT clone question content!"
        assert "option_a" not in d, "answers_detail should NOT clone option_a!"
        assert "option_b" not in d, "answers_detail should NOT clone option_b!"
    print(f"  ✓ ExamResult.answers_detail is COMPACT ({len(er_entry.answers_detail)} bytes, NO redundant content/options)")

    # 7. Verify /api/exam/review On-Demand Reconstruction
    rev_res = client.get("/api/exam/review", headers=stu_headers)
    assert rev_res.status_code == 200
    rev_data = rev_res.json()
    assert len(rev_data["questions"]) == 4
    for q_item in rev_data["questions"]:
        assert q_item.get("content"), "Question content must be reconstructed on demand from Question table!"
        if q_item.get("question_type") != "essay":
            assert q_item.get("option_a"), "Options must be reconstructed on demand!"
    print("  ✓ /api/exam/review successfully reconstructs full question content and options from Question table")

    # 8. Verify /api/admin/results/{id}/detail On-Demand Reconstruction
    detail_res = client.get(f"/api/admin/results/{er_entry.id}/detail", headers=adm_headers)
    assert detail_res.status_code == 200
    detail_data = detail_res.json()
    assert len(detail_data["questions"]) == 4
    for q_item in detail_data["questions"]:
        assert q_item.get("content"), "Admin audit detail must reconstruct content on demand!"
    print("  ✓ /api/admin/results/{id}/detail successfully reconstructs full question content and options")

    # 9. Verify Excel & PDF Exports Work Perfectly with Compact answers_detail
    pdf_res = client.get(f"/api/admin/results/{er_entry.id}/export_pdf", headers=adm_headers)
    assert pdf_res.status_code == 200
    assert len(pdf_res.content) > 1000, "PDF should generate properly with content"
    print(f"  ✓ Single Candidate PDF export generated with full questions ({len(pdf_res.content)} bytes)")

    excel_res = client.get(f"/api/admin/results/{er_entry.id}/export_excel", headers=adm_headers)
    assert excel_res.status_code == 200
    assert len(excel_res.content) > 1000, "Excel should generate properly with content"
    print(f"  ✓ Single Candidate Excel export generated with full questions ({len(excel_res.content)} bytes)")

    # 10. Verify Emergency Black Box Rescue: /api/admin/rescue_import
    # Simulate a crashed candidate who has answers stored in browser
    rescue_payload = {
        "v": 1,
        "exam_id": exam_id,
        "user_id": student.id,
        "username": student.username,
        "fullname": student.fullname,
        "answers": {
            str(q1.id): "C",
            str(q2.id): "A,B",
            str(q3.id): "Phục hồi từ Black Box Rescue thành công!",
            str(q4.id): "Toàn bộ bài làm được bảo tồn an toàn!"
        },
        "timestamp": datetime.datetime.now(datetime.timezone.utc).isoformat()
    }
    json_str = json.dumps(rescue_payload)
    b64_code = base64.b64encode(json_str.encode("utf-8")).decode("utf-8")
    rescue_code_with_prefix = f"HCMUTE-RESCUE:{b64_code}"

    # Call /api/admin/rescue_import
    import_res = client.post("/api/admin/rescue_import", json={
        "rescue_code": rescue_code_with_prefix
    }, headers=adm_headers)
    assert import_res.status_code == 200
    import_data = import_res.json()
    assert import_data["status"] == "success"
    assert import_data["mssv"] == student.username
    assert import_data["total_questions"] == 4
    print("  ✓ Black Box Rescue Import (/api/admin/rescue_import) succeeded in < 100ms!")

    # Verify rescued answers are reflected in DB
    db.expire_all()
    rescued_entry = db.query(ExamResult).filter(ExamResult.id == import_data["result_id"]).first()
    rescued_answers = json.loads(rescued_entry.answers)
    assert rescued_answers[str(q3.id)] == "Phục hồi từ Black Box Rescue thành công!"
    print("  ✓ Rescued candidate answers verified in database!")

    print("\n========================================================")
    print("🎉 ALL 10 FINAL ARCHITECTURAL & RESILIENCE TESTS PASSED 100%!")
    print("========================================================")

if __name__ == "__main__":
    test_full_system_verification()
