import sys
from fastapi.testclient import TestClient
from main import app, SessionLocal, get_password_hash
from models import User

client = TestClient(app)

def test_flow():
    db = SessionLocal()
    admin = db.query(User).filter(User.username == "admin").first()
    if not admin:
        admin = User(
            username="admin",
            password=get_password_hash("admin123"),
            fullname="Administrator",
            is_admin=True
        )
        db.add(admin)
        db.commit()
    db.close()

    # 1. Login
    res = client.post("/token", data={"username": "admin", "password": "admin123"})
    assert res.status_code == 200
    token = res.json()["access_token"]
    headers = {"Authorization": f"Bearer {token}"}
    
    # 2. Config
    res = client.post("/api/admin/config", data={"num_questions": 2, "duration_minutes": 1}, headers=headers)
    assert res.status_code == 200
    
    # 3. Export
    res = client.get("/api/admin/results/export", headers=headers)
    assert res.status_code == 200
    assert len(res.content) > 0

    print("Test passed!")

if __name__ == "__main__":
    test_flow()
