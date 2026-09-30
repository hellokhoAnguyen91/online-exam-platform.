import sys
import os
import io
import json
import datetime
from fastapi.testclient import TestClient

sys.stdout.reconfigure(encoding='utf-8')

from main import app, get_password_hash, normalize_dob
from database import SessionLocal
from models import User, Exam, Question, ExamResult
from docx_parser import parse_docx_questions
import docx
from docx.oxml import parse_xml

client = TestClient(app)

def test_dob_normalization_and_login():
    print("\n--- 1. Testing DOB Normalization & Flexible Login ---")
    db = SessionLocal()
    # Create test student with DOB "2002-08-25"
    stu = db.query(User).filter(User.username == "test_dob_stu").first()
    if not stu:
        stu = User(
            username="test_dob_stu",
            password=get_password_hash("2002-08-25"),
            fullname="Test DOB Student",
            is_admin=False
        )
        db.add(stu)
        db.commit()
    db.close()

    # Login with standard format YYYY-MM-DD
    r1 = client.post("/token", data={"username": "test_dob_stu", "password": "2002-08-25"})
    assert r1.status_code == 200, f"Expected 200, got {r1.status_code}: {r1.text}"
    print("  ✓ Login with 2002-08-25: SUCCESS")

    # Login with Vietnamese format DD/MM/YYYY
    r2 = client.post("/token", data={"username": "test_dob_stu", "password": "25/08/2002"})
    assert r2.status_code == 200, f"Expected 200, got {r2.status_code}: {r2.text}"
    print("  ✓ Login with 25/08/2002 (Vietnamese slash format): SUCCESS")

    # Login with hyphen format DD-MM-YYYY
    r3 = client.post("/token", data={"username": "test_dob_stu", "password": "25-08-2002"})
    assert r3.status_code == 200, f"Expected 200, got {r3.status_code}: {r3.text}"
    print("  ✓ Login with 25-08-2002 (hyphen format): SUCCESS")

def test_submission_race_condition():
    print("\n--- 2. Testing Exam Submission Race Condition ---")
    # Admin login
    admin_login = client.post("/token", data={"username": "trangnh@hcmute.edu.vn", "password": "nguyenhatrang"})
    admin_token = admin_login.json()["access_token"]
    admin_headers = {"Authorization": f"Bearer {admin_token}"}

    # Student login
    stu_login = client.post("/token", data={"username": "test_dob_stu", "password": "2002-08-25"})
    stu_token = stu_login.json()["access_token"]
    stu_headers = {"Authorization": f"Bearer {stu_token}"}

    # Admin creates Exam A & Exam B
    res_a = client.post("/api/admin/exams", json={"title": "Exam A for Race Test", "num_questions": 2, "duration_minutes": 15}, headers=admin_headers)
    exam_a_id = res_a.json()["exam_id"]
    client.post(f"/api/admin/exams/{exam_a_id}/activate", headers=admin_headers)

    # Add questions to Exam A
    db = SessionLocal()
    q1 = Question(exam_id=exam_a_id, content="Race Q1", option_a="1", option_b="2", option_c="3", option_d="4", correct_option="A")
    q2 = Question(exam_id=exam_a_id, content="Race Q2", option_a="1", option_b="2", option_c="3", option_d="4", correct_option="B")
    db.add_all([q1, q2])
    db.commit()
    db.refresh(q1)
    db.refresh(q2)
    db.close()

    # Student starts Exam A
    r_start = client.get("/api/exam", headers=stu_headers)
    assert r_start.status_code == 200
    exam_data = r_start.json()
    assert exam_data["status"] == "ongoing"
    assert exam_data["exam_title"] == "Exam A for Race Test"
    print("  ✓ Student started Exam A successfully")

    # Now Admin creates Exam B and activates Exam B WHILE student is taking Exam A!
    res_b = client.post("/api/admin/exams", json={"title": "Exam B Active", "num_questions": 2, "duration_minutes": 15}, headers=admin_headers)
    exam_b_id = res_b.json()["exam_id"]
    client.post(f"/api/admin/exams/{exam_b_id}/activate", headers=admin_headers)
    print("  ✓ Admin activated Exam B while student is in the middle of Exam A")

    # Student saves progress on Exam A
    r_save = client.post("/api/exam/save_progress", json={"answers": {str(q1.id): "A"}}, headers=stu_headers)
    assert r_save.status_code == 200
    assert r_save.json()["status"] == "saved"
    print("  ✓ Auto-save progress preserved student session")

    # Student submits Exam A!
    r_sub = client.post("/api/exam/submit", json={"answers": {str(q1.id): "A", str(q2.id): "B"}}, headers=stu_headers)
    assert r_sub.status_code == 200, f"Expected 200, got {r_sub.status_code}: {r_sub.text}"
    sub_data = r_sub.json()
    assert sub_data["status"] == "submitted"
    assert sub_data["correct_count"] == 2
    assert sub_data["score"] == 10.0
    print("  ✓ Exam submission SUCCEEDED despite active exam changing! Score: 10.0 (2/2)")

def test_student_review_fallback():
    print("\n--- 3. Testing Student Review Fallback Reconstruction ---")
    stu_login = client.post("/token", data={"username": "test_dob_stu", "password": "2002-08-25"})
    stu_token = stu_login.json()["access_token"]
    stu_headers = {"Authorization": f"Bearer {stu_token}"}

    # Simulate a result with answers_detail = None
    db = SessionLocal()
    stu = db.query(User).filter(User.username == "test_dob_stu").first()
    res = db.query(ExamResult).filter(ExamResult.user_id == stu.id).order_by(ExamResult.submit_time.desc()).first()
    assert res is not None
    res.answers_detail = None # Clear snapshot
    db.commit()
    db.close()

    r_rev = client.get("/api/exam/review", headers=stu_headers)
    assert r_rev.status_code == 200, f"Expected 200, got {r_rev.status_code}: {r_rev.text}"
    rev_data = r_rev.json()
    assert len(rev_data["questions"]) > 0, "Expected fallback questions reconstruction!"
    print(f"  ✓ Fallback reconstruction successfully restored {len(rev_data['questions'])} review questions!")

def test_index_html_structure():
    print("\n--- 4. Testing Frontend index.html Structure ---")
    with open("static/index.html", encoding="utf-8") as f:
        html_content = f.read()

    assert 'id="student-notice-box"' in html_content, "Missing student-notice-box"
    assert 'id="student-exam-body"' in html_content, "Missing student-exam-body"
    assert 'id="mobile-exam-bar"' in html_content, "Missing mobile-exam-bar"
    assert 'data-tab="exams"' in html_content, "Missing data-tab attributes"
    assert 'data-tab="questions"' in html_content, "Missing data-tab questions"
    assert 'onbeforeunload' in html_content, "Missing beforeunload warning safeguard"
    assert '.math-expr' in html_content, "Missing math-expr CSS"
    assert '.math-frac' in html_content, "Missing math-frac CSS"
    print("  ✓ All required frontend HTML containers, mobile bar, and safety guards verified!")

if __name__ == "__main__":
    test_dob_normalization_and_login()
    test_submission_race_condition()
    test_student_review_fallback()
    test_index_html_structure()
    print("\n🎉 ALL REVIEW FIXES AND CRITICAL PATHS VERIFIED 100%!")
