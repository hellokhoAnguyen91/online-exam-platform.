import sys
import os
import json
from fastapi.testclient import TestClient
from main import app

sys.stdout.reconfigure(encoding='utf-8')
client = TestClient(app)

def test_full_system():
    print("=== 1. Testing Admin Authentication ===")
    res = client.post("/token", data={"username": "trangnh@hcmute.edu.vn", "password": "nguyenhatrang"})
    assert res.status_code == 200, f"Admin login failed: {res.text}"
    admin_token = res.json()["access_token"]
    admin_headers = {"Authorization": f"Bearer {admin_token}"}
    print("  ✓ Admin login successful!")

    # /api/me
    res = client.get("/api/me", headers=admin_headers)
    assert res.status_code == 200
    user_info = res.json()
    assert user_info["is_admin"] is True
    assert "Trang" in user_info["fullname"]
    print("  ✓ /api/me verified:", user_info["fullname"])

    print("\n=== 2. Testing Exam Management & Archive ===")
    res = client.get("/api/admin/exams", headers=admin_headers)
    assert res.status_code == 200
    exams = res.json()
    print(f"  Existing exams: {len(exams)}")
    active_exam = next((e for e in exams if e["is_active"]), exams[0] if exams else None)
    assert active_exam is not None, "Expected at least one active exam"
    exam_id = active_exam["id"]
    print(f"  Active exam ID: {exam_id} - '{active_exam['title']}'")

    # Create new exam
    res = client.post("/api/admin/exams", json={
        "title": "Kỳ thi Kiểm tra Đánh giá Định kỳ",
        "code": "KT-2026",
        "description": "Kỳ thi thử nghiệm kiểm thử tự động",
        "num_questions": 5,
        "duration_minutes": 15,
        "allow_review": True,
        "shuffle_questions": True
    }, headers=admin_headers)
    assert res.status_code == 200
    new_exam_id = res.json()["exam_id"]
    print("  ✓ Created new exam ID:", new_exam_id)

    # Activate new exam
    res = client.post(f"/api/admin/exams/{new_exam_id}/activate", headers=admin_headers)
    assert res.status_code == 200
    print("  ✓ Activated new exam ID:", new_exam_id)

    print("\n=== 3. Testing Docx Upload with Images ===")
    # Upload sample_questions.docx to the new exam
    with open("sample_questions.docx", "rb") as f:
        res = client.post(
            f"/api/admin/upload_questions?exam_id={new_exam_id}",
            files={"file": ("sample_questions.docx", f, "application/vnd.openxmlformats-officedocument.wordprocessingml.document")},
            data={"replace_existing": "true"},
            headers=admin_headers
        )
    assert res.status_code == 200, f"Upload questions failed: {res.text}"
    upload_res = res.json()
    print("  ✓ Upload questions result:", upload_res["detail"])
    assert upload_res["count"] >= 3

    # Check questions list
    res = client.get(f"/api/admin/questions?exam_id={new_exam_id}", headers=admin_headers)
    assert res.status_code == 200
    q_data = res.json()
    assert q_data["count"] >= 3
    # Check that image data URI is present in questions
    has_img = any("data:image" in q["content"] for q in q_data["questions"])
    assert has_img, "Expected at least one question to contain base64 image data URI"
    print("  ✓ Verified questions with images successfully retrieved!")

    print("\n=== 4. Testing Student Management & Upload ===")
    # Check student list
    res = client.get(f"/api/admin/students?exam_id={new_exam_id}", headers=admin_headers)
    assert res.status_code == 200
    students = res.json()
    print(f"  Existing students count: {len(students)}")
    assert len(students) >= 1, "Expected at least one student in database"

    test_student = students[0]
    student_mssv = test_student["username"]
    print(f"  Selected student for test: {student_mssv} - {test_student['fullname']}")

    print("\n=== 5. Testing Candidate Exam Taking Flow ===")
    # Get student password (or login with sample student)
    # Student 21110001 DOB is '2003-01-15' or similar, let's login with student
    # If student password isn't known, let's upload sample_students.xlsx to be certain
    if os.path.exists("sample_students.xlsx"):
        with open("sample_students.xlsx", "rb") as f:
            res = client.post(
                "/api/admin/upload_students",
                files={"file": ("sample_students.xlsx", f, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
                headers=admin_headers
            )
            assert res.status_code == 200
            print("  ✓ Re-uploaded sample_students.xlsx:", res.json()["detail"])

    # Login as student 101 / 2000-01-01 or 21110001
    res = client.post("/token", data={"username": "101", "password": "2000-01-01"})
    if res.status_code != 200:
        # Try 21110001 with common DOB
        res = client.post("/token", data={"username": "21110001", "password": "2003-01-15"})
    assert res.status_code == 200, f"Student login failed: {res.text}"
    student_token = res.json()["access_token"]
    student_headers = {"Authorization": f"Bearer {student_token}"}
    print("  ✓ Student login successful!")

    # Student starts exam
    res = client.get("/api/exam", headers=student_headers)
    assert res.status_code == 200, f"Get exam failed: {res.text}"
    exam_start = res.json()
    print("  ✓ Candidate exam started. Status:", exam_start["status"], "| Questions count:", len(exam_start["questions"]))
    student_questions = exam_start["questions"]
    assert len(student_questions) > 0

    # Student saves answers
    first_q_id = student_questions[0]["id"]
    res = client.post("/api/exam/save_progress", json={"answers": {str(first_q_id): "C"}}, headers=student_headers)
    assert res.status_code == 200
    print("  ✓ Auto-save progress successful!")

    # Student submits exam
    all_answers = {}
    for q in student_questions:
        all_answers[str(q["id"])] = "B" # Pick B for test
    res = client.post("/api/exam/submit", json={"answers": all_answers}, headers=student_headers)
    assert res.status_code == 200
    submit_res = res.json()
    assert submit_res["status"] == "submitted"
    print("  ✓ Exam submitted! Score:", submit_res["score"], f"({submit_res['correct_count']}/{submit_res['total_questions']})")

    # Student reviews own submission
    res = client.get("/api/exam/review", headers=student_headers)
    assert res.status_code == 200
    review_data = res.json()
    assert len(review_data["questions"]) > 0
    print("  ✓ Student review endpoint verified! Questions in review:", len(review_data["questions"]))

    print("\n=== 6. Testing Admin Audit Review & Results Export ===")
    res = client.get(f"/api/admin/results?exam_id={new_exam_id}", headers=admin_headers)
    assert res.status_code == 200
    res_list = res.json()
    assert len(res_list["results"]) >= 1
    submitted_result = res_list["results"][0]
    result_id = submitted_result["id"]
    print("  ✓ Results list verified. Stats:", res_list["stats"])

    # Detailed Audit Review for Teacher
    res = client.get(f"/api/admin/results/{result_id}/detail", headers=admin_headers)
    assert res.status_code == 200
    audit = res.json()
    assert audit["candidate"]["mssv"] == "101" or audit["candidate"]["mssv"] == "21110001"
    assert len(audit["questions"]) > 0
    q1_audit = audit["questions"][0]
    assert "selected" in q1_audit
    assert "correct" in q1_audit
    assert "is_correct" in q1_audit
    assert "status" in q1_audit
    print(f"  ✓ Full candidate audit review verified! Q1 status: {q1_audit['status']}, Student chose: {q1_audit['selected']}, Correct: {q1_audit['correct']}")

    # Export Excel
    res = client.get(f"/api/admin/results/export?exam_id={new_exam_id}", headers=admin_headers)
    assert res.status_code == 200
    assert len(res.content) > 100
    print("  ✓ Export Excel spreadsheet verified! Size:", len(res.content), "bytes")

    print("\n=== 7. Testing Archive & Backup / Restore ===")
    # Backup exam package
    res = client.get(f"/api/admin/exams/{new_exam_id}/backup", headers=admin_headers)
    assert res.status_code == 200
    backup_json_bytes = res.content
    backup_data = json.loads(backup_json_bytes.decode('utf-8'))
    assert "exam" in backup_data and "questions" in backup_data and "results" in backup_data
    print("  ✓ Exam backup generated! Questions in backup:", len(backup_data["questions"]), "| Results in backup:", len(backup_data["results"]))

    # Archive exam
    res = client.post(f"/api/admin/exams/{new_exam_id}/archive", headers=admin_headers)
    assert res.status_code == 200
    print("  ✓ Exam archived successfully!")

    # Restore exam from backup
    res = client.post(
        "/api/admin/exams/restore",
        files={"file": ("backup.json", backup_json_bytes, "application/json")},
        headers=admin_headers
    )
    assert res.status_code == 200
    restore_res = res.json()
    print("  ✓ Exam restored successfully:", restore_res["detail"])

    # Switch back to original active exam
    client.post(f"/api/admin/exams/{exam_id}/activate", headers=admin_headers)

    print("\n🎉 ALL 19 API ENDPOINTS AND SYSTEM WORKFLOWS VERIFIED 100%!")

if __name__ == "__main__":
    test_full_system()
