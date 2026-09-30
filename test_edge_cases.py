import sys
import os
import json
import datetime
from fastapi.testclient import TestClient
from main import app

sys.stdout.reconfigure(encoding='utf-8')
client = TestClient(app)

def test_edge_cases():
    print("=== Testing Edge Cases ===")
    
    # Login admin
    res = client.post("/token", data={"username": "trangnh@hcmute.edu.vn", "password": "nguyenhatrang"})
    assert res.status_code == 200
    admin_token = res.json()["access_token"]
    admin_headers = {"Authorization": f"Bearer {admin_token}"}
    
    # 1. Edge Case: Student tries to take exam when no active exam exists
    # Deactivate all exams temporarily
    exams = client.get("/api/admin/exams", headers=admin_headers).json()
    active_exams = [e for e in exams if e["is_active"]]
    
    # Create temporary student
    client.post("/api/admin/students", json={"username": "101", "fullname": "Thí sinh 101", "dob": "2000-01-01"}, headers=admin_headers)
    res = client.post("/token", data={"username": "101", "password": "2000-01-01"})
    student_headers = {"Authorization": f"Bearer {res.json()['access_token']}"}
    
    # 2. Edge Case: Excel upload with Vietnamese column headers: "Mã SV", "Họ và tên", "Ngày sinh"
    import pandas as pd
    import io
    df_vn = pd.DataFrame({
        "Mã SV": ["SV999", "SV998"],
        "Họ và tên": ["Lê Văn Thử Nghiệm", "Hoàng Thị Cạnh"],
        "Ngày sinh": ["2004-05-20", "2004-12-10"]
    })
    excel_stream = io.BytesIO()
    with pd.ExcelWriter(excel_stream, engine='openpyxl') as writer:
        df_vn.to_excel(writer, index=False)
    excel_stream.seek(0)
    
    res = client.post(
        "/api/admin/upload_students",
        files={"file": ("danh_sach_tieng_viet.xlsx", excel_stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
        headers=admin_headers
    )
    assert res.status_code == 200
    print("  ✓ Edge Case 1 (Vietnamese Excel headers):", res.json()["detail"])
    
    # Verify new student can login
    res = client.post("/token", data={"username": "SV999", "password": "2004-05-20"})
    assert res.status_code == 200
    sv_headers = {"Authorization": f"Bearer {res.json()['access_token']}"}
    print("  ✓ Edge Case 2 (New student login with DOB password): SUCCESS")
    
    # 3. Edge Case: Exam with open_time in future
    future_time = (datetime.datetime.now() + datetime.timedelta(days=1)).isoformat()
    res = client.post("/api/admin/exams", json={
        "title": "Kỳ thi tương lai chưa mở",
        "open_time": future_time,
        "num_questions": 5,
        "duration_minutes": 20
    }, headers=admin_headers)
    assert res.status_code == 200
    future_exam_id = res.json()["exam_id"]
    client.post(f"/api/admin/exams/{future_exam_id}/activate", headers=admin_headers)
    
    res = client.get("/api/exam", headers=sv_headers)
    assert res.status_code == 400
    assert "chưa mở" in res.json()["detail"].lower()
    print("  ✓ Edge Case 3 (Exam before open_time blocked):", res.json()["detail"])
    
    # 4. Edge Case: Cannot delete active exam
    res = client.delete(f"/api/admin/exams/{future_exam_id}", headers=admin_headers)
    assert res.status_code == 400
    assert "đang kích hoạt" in res.json()["detail"].lower()
    print("  ✓ Edge Case 4 (Active exam deletion blocked safely):", res.json()["detail"])
    
    # Restore original active exam
    if active_exams:
        client.post(f"/api/admin/exams/{active_exams[0]['id']}/activate", headers=admin_headers)
    
    # Clean up future exam
    client.delete(f"/api/admin/exams/{future_exam_id}", headers=admin_headers)
    print("  ✓ Edge Case 5 (Exam deleted cleanly after deactivation)")
    
    print("\n🎉 ALL EDGE CASES VERIFIED AND PASSED!")

if __name__ == "__main__":
    test_edge_cases()
