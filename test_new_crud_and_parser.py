import unittest
import io
import pandas as pd
from fastapi.testclient import TestClient
from main import app
from database import SessionLocal
from models import User, Exam, Question, ExamResult

class TestNewCrudAndParser(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.client = TestClient(app)
        # Login as Admin
        res = cls.client.post("/token", data={"username": "trangnh@hcmute.edu.vn", "password": "nguyenhatrang"})
        assert res.status_code == 200, f"Admin login failed: {res.text}"
        cls.token = res.json()["access_token"]
        cls.headers = {"Authorization": f"Bearer {cls.token}"}

    def setUp(self):
        db = SessionLocal()
        test_usernames = ["21119998", "21118801", "21118802"]
        test_user_ids = [u.id for u in db.query(User).filter(User.username.in_(test_usernames)).all()]
        if test_user_ids:
            db.query(ExamResult).filter(ExamResult.user_id.in_(test_user_ids)).delete(synchronize_session=False)
            db.query(User).filter(User.id.in_(test_user_ids)).delete(synchronize_session=False)
            db.commit()
        db.close()

    def test_01_manual_question_crud(self):
        # 1. Create Question Manually
        payload = {
            "content": "Thủ đô của Việt Nam là thành phố nào?",
            "option_a": "Hồ Chí Minh",
            "option_b": "Hà Nội",
            "option_c": "Đà Nẵng",
            "option_d": "Huế",
            "correct_option": "B",
            "explanation": "Hà Nội là thủ đô của nước CHXHCN Việt Nam."
        }
        res = self.client.post("/api/admin/questions", json=payload, headers=self.headers)
        self.assertEqual(res.status_code, 200)
        data = res.json()
        q_id = data["question"]["id"]
        self.assertEqual(data["question"]["correct_option"], "B")
        self.assertEqual(data["question"]["explanation"], "Hà Nội là thủ đô của nước CHXHCN Việt Nam.")

        # 2. Update Question Manually
        update_payload = {
            "content": "Thủ đô của nước CHXHCN Việt Nam là gì?",
            "option_a": "TP. Hồ Chí Minh",
            "option_b": "Thủ đô Hà Nội",
            "option_c": "Đà Nẵng",
            "option_d": "Cần Thơ",
            "correct_option": "B",
            "explanation": "Hà Nội là thủ đô nghìn năm văn hiến."
        }
        res_up = self.client.put(f"/api/admin/questions/{q_id}", json=update_payload, headers=self.headers)
        self.assertEqual(res_up.status_code, 200)
        up_data = res_up.json()
        self.assertEqual(up_data["question"]["content"], "Thủ đô của nước CHXHCN Việt Nam là gì?")
        self.assertEqual(up_data["question"]["option_b"], "Thủ đô Hà Nội")

        # 3. Verify in List
        res_list = self.client.get("/api/admin/questions", headers=self.headers)
        self.assertEqual(res_list.status_code, 200)
        questions = res_list.json()["questions"]
        matched = [q for q in questions if q["id"] == q_id]
        self.assertEqual(len(matched), 1)

        # 4. Delete Question Manually
        res_del = self.client.delete(f"/api/admin/questions/{q_id}", headers=self.headers)
        self.assertEqual(res_del.status_code, 200)
        
        # Verify deleted
        res_list_after = self.client.get("/api/admin/questions", headers=self.headers)
        matched_after = [q for q in res_list_after.json()["questions"] if q["id"] == q_id]
        self.assertEqual(len(matched_after), 0)

    def test_02_manual_student_crud(self):
        # 1. Create Student Manually
        payload = {
            "username": "21119998",
            "fullname": "Trần Thị Thủ Công",
            "dob": "25/11/2003"
        }
        res = self.client.post("/api/admin/students", json=payload, headers=self.headers)
        self.assertEqual(res.status_code, 200)
        stu_id = res.json()["student"]["id"]

        # Verify Student can log in with DOB
        res_login = self.client.post("/token", data={"username": "21119998", "password": "25/11/2003"})
        self.assertEqual(res_login.status_code, 200)
        # Also check ISO DOB format
        res_login_iso = self.client.post("/token", data={"username": "21119998", "password": "2003-11-25"})
        self.assertEqual(res_login_iso.status_code, 200)

        # 2. Update Student Manually (Change name and new DOB password)
        up_payload = {
            "username": "21119998",
            "fullname": "Trần Thị Thủ Công (Đã Cập Nhật)",
            "dob": "10/10/2003"
        }
        res_up = self.client.put(f"/api/admin/students/{stu_id}", json=up_payload, headers=self.headers)
        self.assertEqual(res_up.status_code, 200)

        # Verify login with NEW password works and OLD password fails
        res_old = self.client.post("/token", data={"username": "21119998", "password": "25/11/2003"})
        self.assertEqual(res_old.status_code, 401)
        res_new = self.client.post("/token", data={"username": "21119998", "password": "10/10/2003"})
        self.assertEqual(res_new.status_code, 200)

        # 3. Delete Student Manually
        res_del = self.client.delete(f"/api/admin/students/{stu_id}", headers=self.headers)
        self.assertEqual(res_del.status_code, 200)

        # Verify Student can no longer log in
        res_gone = self.client.post("/token", data={"username": "21119998", "password": "10/10/2003"})
        self.assertEqual(res_gone.status_code, 401)

    def test_03_smart_student_upload_with_irrelevant_columns(self):
        # Create Excel with title rows, split name columns, and ignored columns
        rows = [
            ["TRƯỜNG ĐH SƯ PHẠM KỸ THUẬT TP.HCM", "", "", "", "", "", ""],
            ["DANH SÁCH SINH VIÊN THI HỌC KỲ", "", "", "", "", "", ""],
            ["STT", "Mã SV", "Họ và chữ lót", "Tên", "Ngày sinh", "Lớp", "Phòng thi"],
            [1, "21118801", "Lê Văn", "Hải", "12/03/2003", "21110A", "P101"],
            [2, "21118802", "Phạm Quỳnh", "Nga", "18/07/2003", "21110A", "P101"]
        ]
        df = pd.DataFrame(rows)
        buf = io.BytesIO()
        df.to_excel(buf, index=False, header=False)

        res = self.client.post(
            "/api/admin/upload_students",
            files={"file": ("danh_sach_thi.xlsx", buf.getvalue(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
            headers=self.headers
        )
        self.assertEqual(res.status_code, 200)
        data = res.json()
        self.assertEqual(data["count"], 2)
        self.assertIn("STT", data["ignored_columns"])
        self.assertEqual(data["preview"][0]["class_name"], "21110A")
        self.assertIn("Phòng thi", data["ignored_columns"])

        # Check student 21118801 can login with their DOB
        res_log = self.client.post("/token", data={"username": "21118801", "password": "12/03/2003"})
        self.assertEqual(res_log.status_code, 200)

        # Check student 21118802 combined name
        res_list = self.client.get("/api/admin/students", headers=self.headers)
        students = res_list.json()
        nga = [s for s in students if s["username"] == "21118802"]
        self.assertEqual(len(nga), 1)
        self.assertEqual(nga[0]["fullname"], "Phạm Quỳnh Nga")

        # Cleanup test students
        db = SessionLocal()
        db.query(User).filter(User.username.in_(["21118801", "21118802"])).delete(synchronize_session=False)
        db.commit()
        db.close()

if __name__ == "__main__":
    unittest.main()
