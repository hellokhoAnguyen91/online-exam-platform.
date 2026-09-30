import sys
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
import io
import json
import datetime
import docx
from fastapi.testclient import TestClient
from main import app, get_password_hash, SessionLocal
from models import User, Exam, Question, ExamResult
from pdf_export import generate_candidate_pdf

client = TestClient(app)

def test_full_section_workflow():
    print("\n========================================================")
    print("TESTING FULL SECTION WORKFLOW: PART I (MC) & PART II (ESSAY)")
    print("========================================================")

    db = SessionLocal()
    try:
        # 1. Setup Admin
        admin = db.query(User).filter(User.username == "sec_admin").first()
        if not admin:
            admin = User(username="sec_admin", password=get_password_hash("admin123"), fullname="Section Admin", is_admin=True)
            db.add(admin)
            db.commit()
            db.refresh(admin)

        adm_login = client.post("/token", data={"username": "sec_admin", "password": "admin123"})
        assert adm_login.status_code == 200
        adm_token = adm_login.json()["access_token"]
        adm_headers = {"Authorization": f"Bearer {adm_token}"}
        print("  ✓ Admin authenticated")

        # 2. Setup Student
        student = db.query(User).filter(User.username == "sec_student01").first()
        if not student:
            student = User(username="sec_student01", password=get_password_hash("sec_student01"), fullname="Nguyen Van Section", dob="2003-01-01", is_admin=False)
            db.add(student)
            db.commit()
            db.refresh(student)

        stu_login = client.post("/token", data={"username": "sec_student01", "password": "sec_student01"})
        assert stu_login.status_code == 200
        stu_token = stu_login.json()["access_token"]
        stu_headers = {"Authorization": f"Bearer {stu_token}"}
        print("  ✓ Candidate authenticated")

        # 3. Create Exam
        now = datetime.datetime.now()
        open_time = (now - datetime.timedelta(minutes=5)).isoformat()
        close_time = (now + datetime.timedelta(hours=2)).isoformat()
        exam_res = client.post("/api/admin/exams", json={
            "title": "Kỳ thi Phân đoạn Trắc nghiệm & Tự luận",
            "code": "SEC2026",
            "duration_minutes": 60,
            "num_questions": 10,
            "open_time": open_time,
            "close_time": close_time,
            "allow_review": True,
            "shuffle_questions": True
        }, headers=adm_headers)
        assert exam_res.status_code == 200
        exam_id = exam_res.json()["exam_id"]

        client.post(f"/api/admin/exams/{exam_id}/activate", headers=adm_headers)
        print("  ✓ Exam created and activated")

        # 4. Generate Docx with Section Headers and upload
        doc = docx.Document()
        doc.add_paragraph("PHẦN I: TRẮC NGHIỆM")
        doc.add_paragraph("Câu 1: 1 + 1 bằng mấy?")
        doc.add_paragraph("A. 1")
        doc.add_paragraph("B. 2")
        doc.add_paragraph("C. 3")
        doc.add_paragraph("D. 4")
        doc.add_paragraph("Đáp án: B")

        doc.add_paragraph("Câu 2: Chọn các hệ điều hành mã nguồn mở (multi select):")
        doc.add_paragraph("+ Linux")
        doc.add_paragraph("+ FreeBSD")
        doc.add_paragraph("+ Windows 11")
        doc.add_paragraph("Đáp án: A, B")

        doc.add_paragraph("PHẦN II: TỰ LUẬN")
        doc.add_paragraph("Câu 3: Hãy nêu kiến trúc cơ bản của hệ điều hành Linux (3.0 điểm).")
        doc.add_paragraph("Câu 4: Trình bày quy trình xử lý lỗi bộ nhớ (2.0 điểm).")

        docx_bio = io.BytesIO()
        doc.save(docx_bio)
        docx_bio.seek(0)

        upload_res = client.post(
            f"/api/admin/upload_questions?exam_id={exam_id}&replace_existing=true",
            files={"file": ("sections_test.docx", docx_bio.getvalue(), "application/vnd.openxmlformats-officedocument.wordprocessingml.document")},
            headers=adm_headers
        )
        assert upload_res.status_code == 200
        print("  ✓ Docx with Section Headers parsed and uploaded successfully:", upload_res.json()["detail"])

        # 5. Verify /api/admin/questions returns separated count
        qs_res = client.get(f"/api/admin/questions?exam_id={exam_id}", headers=adm_headers)
        assert qs_res.status_code == 200
        qs_data = qs_res.json()
        assert qs_data["mc_count"] == 2
        assert qs_data["essay_count"] == 2
        assert len(qs_data["questions"]) == 4
        # Verify order: MC questions first, Essay questions second
        assert qs_data["questions"][0]["question_type"] in ["multiple_choice", "multi_select"]
        assert qs_data["questions"][1]["question_type"] in ["multiple_choice", "multi_select"]
        assert qs_data["questions"][2]["question_type"] == "essay"
        assert qs_data["questions"][3]["question_type"] == "essay"
        print("  ✓ Admin questions endpoint verified (order: Part I then Part II)")

        # 6. Candidate fetches exam via /api/exam
        take_res = client.get("/api/exam", headers=stu_headers)
        assert take_res.status_code == 200
        exam_data = take_res.json()
        assert exam_data["mc_count"] == 2
        assert exam_data["essay_count"] == 2
        c_questions = exam_data["questions"]
        assert len(c_questions) == 4
        assert c_questions[0]["question_type"] in ["multiple_choice", "multi_select"]
        assert c_questions[1]["question_type"] in ["multiple_choice", "multi_select"]
        assert c_questions[2]["question_type"] == "essay"
        assert c_questions[3]["question_type"] == "essay"
        print("  ✓ Candidate exam fetch verified (Part I MC first, Part II Essay second, counts correct)")

        # 7. Candidate answers questions and submits
        q1_id = c_questions[0]["id"]
        q2_id = c_questions[1]["id"]
        q3_id = c_questions[2]["id"]
        q4_id = c_questions[3]["id"]

        submit_payload = {
            "answers": {
                str(q1_id): "B",
                str(q2_id): "A,B",
                str(q3_id): "Kiến trúc Linux gồm: Kernel, System libraries, Shell và System utilities...",
                str(q4_id): "Quy trình xử lý lỗi bộ nhớ..."
            }
        }
        sub_res = client.post("/api/exam/submit", json=submit_payload, headers=stu_headers)
        assert sub_res.status_code == 200
        sub_data = sub_res.json()
        assert sub_data["needs_grading"] is True
        print("  ✓ Exam submitted with pending essay grading")

        # 8. Check Student Review endpoint (/api/exam/review)
        rev_res = client.get("/api/exam/review", headers=stu_headers)
        assert rev_res.status_code == 200
        rev_data = rev_res.json()
        assert rev_data["mc_count"] == 2
        assert rev_data["essay_count"] == 2
        assert rev_data["has_pending_essay"] is True
        assert len(rev_data["questions"]) == 4
        print("  ✓ Student review endpoint verified with section breakdown")

        # 9. Check Admin Audit Detail endpoint
        res_entry = db.query(ExamResult).filter(ExamResult.user_id == student.id, ExamResult.exam_id == exam_id).first()
        assert res_entry is not None
        audit_res = client.get(f"/api/admin/results/{res_entry.id}/detail", headers=adm_headers)
        assert audit_res.status_code == 200
        audit_data = audit_res.json()
        assert audit_data["mc_count"] == 2
        assert audit_data["essay_count"] == 2
        assert audit_data["has_pending_essay"] is True
        print("  ✓ Admin audit detail endpoint verified with section breakdown")

        # 10. Grade Essay Questions
        grade_res = client.post(f"/api/admin/results/{res_entry.id}/score_essay", json={
            "scores": {
                str(q3_id): 2.5,
                str(q4_id): 1.5
            }
        }, headers=adm_headers)
        assert grade_res.status_code == 200
        print("  ✓ Teacher essay grading applied successfully")

        # Re-check audit detail after grading
        audit_res2 = client.get(f"/api/admin/results/{res_entry.id}/detail", headers=adm_headers)
        audit_data2 = audit_res2.json()
        assert audit_data2["has_pending_essay"] is False
        assert audit_data2["essay_score"] == 2.4
        print(f"  ✓ Graded score verified: {audit_data2['score']} / {audit_data2['max_score']}")

        # 11. PDF Export Verification
        cand_info = {
            "username": student.username,
            "fullname": student.fullname,
            "dob": student.dob,
            "score": audit_data2["score"],
            "max_score": 10.0,
            "correct_count": audit_data2["correct_count"],
            "total_questions": 4,
            "duration_str": "15p 00s",
            "start_time_str": "30/09/2026 09:00:00",
            "submit_time_str": "30/09/2026 09:15:00",
            "status_text": "Đã chấm xong"
        }
        exam_info = {
            "id": exam_id,
            "title": "Kỳ thi Phân đoạn Trắc nghiệm & Tự luận",
            "code": "SEC2026",
            "duration_minutes": 60,
            "num_questions": 4
        }
        pdf_bytes = generate_candidate_pdf(exam_info, cand_info, audit_data2["questions"])
        assert len(pdf_bytes) > 2000
        assert pdf_bytes.startswith(b"%PDF")
        print(f"  ✓ PDF generated with Part I and Part II sections ({len(pdf_bytes)} bytes)")

        # 12. Frontend index.html validation
        with open("static/index.html", "r", encoding="utf-8") as f:
            html = f.read()

        assert 'id="palette-tabs"' in html, "Missing #palette-tabs in index.html"
        assert 'palette-group-card' in html, "Missing .palette-group-card in index.html"
        assert 'PHẦN I' in html, "Missing PHẦN I header in index.html"
        assert 'PHẦN II' in html, "Missing PHẦN II header in index.html"
        assert 'teacher-evaluation-box' in html, "Missing .teacher-evaluation-box in index.html"
        assert 'Ô CHẤM ĐIỂM & NHẬN XÉT CỦA GIẢNG VIÊN' in html
        print("  ✓ Frontend UI components verified in index.html")

        print("\n========================================================")
        print("🎉 ALL SECTION INTEGRATION TESTS PASSED 100%!")
        print("========================================================")

    finally:
        db.close()

if __name__ == "__main__":
    test_full_section_workflow()
