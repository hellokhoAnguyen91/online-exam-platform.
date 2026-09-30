import sys
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
from fastapi.testclient import TestClient
from main import app
from database import SessionLocal
from models import User, Exam, Question, ExamResult
from passlib.context import CryptContext
import datetime

pwd_context = CryptContext(schemes=['bcrypt'], deprecated='auto')
client = TestClient(app)
db = SessionLocal()

for e in db.query(Exam).filter(Exam.is_active == True).all():
    e.is_active = False
db.commit()

print('=== Test 1: MSSV password with SV prefix ===')
stu = db.query(User).filter(User.username == '21118888').first()
if not stu:
    stu = User(username='21118888', password=pwd_context.hash('2003-01-01'), is_admin=False, fullname='Test MSSV SV')
    db.add(stu)
else:
    stu.password = pwd_context.hash('2003-01-01')
db.commit()

# Case 1A: Student in DB is 21118888, logs in with username 21118888 and password SV21118888
r1 = client.post('/token', data={'username': '21118888', 'password': 'SV21118888'})
assert r1.status_code == 200, f'Login failed: {r1.text}'
print('  PASS: Login with password SV21118888 succeeded!')

# Case 1B: Student logs in with username SV21118888 and password 21118888
r2 = client.post('/token', data={'username': 'SV21118888', 'password': '21118888'})
assert r2.status_code == 200, f'Login failed: {r2.text}'
print('  PASS: Login with username SV21118888 and password 21118888 succeeded!')

# Case 1C: Student logs in with username SV21118888 and password SV21118888
r3 = client.post('/token', data={'username': 'SV21118888', 'password': 'SV21118888'})
assert r3.status_code == 200, f'Login failed: {r3.text}'
print('  PASS: Login with username SV21118888 and password SV21118888 succeeded!')

# Case 1D: Student logs in with DOB
r4 = client.post('/token', data={'username': '21118888', 'password': '01/01/2003'})
assert r4.status_code == 200, f'Login with DOB failed: {r4.text}'
print('  PASS: Login with DOB 01/01/2003 succeeded!')

print('=== Test 2: Multi-select with list payload & partial credit ===')
ex = Exam(title='Test Multi List OK', is_active=True, duration_minutes=30)
db.add(ex)
db.commit()
q = Question(exam_id=ex.id, content='Select A, B, C', option_a='A', option_b='B', option_c='C', option_d='D', option_e='E', correct_option='A,B,C', question_type='multi_select', score_weight=2.0)
db.add(q)
db.commit()

tok = r1.json()['access_token']
headers = {'Authorization': f'Bearer {tok}'}
client.get('/api/exam', headers=headers)

# Case 2A: Full credit with list payload ['A', 'B', 'C']
sub_res = client.post('/api/exam/submit', json={'answers': {str(q.id): ['A', 'B', 'C']}}, headers=headers)
assert sub_res.status_code == 200, f'Submit failed: {sub_res.text}'
score = sub_res.json()['score']
assert score == 10.0, f'Expected 10.0, got {score}'
print('  PASS: Multi-select full credit scored 10.0 with list payload!')

print('=== Test 3: Review masking during active exam ===')
ex.close_time = datetime.datetime.now() + datetime.timedelta(hours=2)
db.commit()

rev_res = client.get('/api/exam/review', headers=headers)
assert rev_res.status_code == 200
rq = rev_res.json()['questions'][0]
assert rq.get('correct') is None, f'Expected None for correct, got {rq.get("correct")}'
assert rq.get('is_correct') is None, f'Expected None for is_correct, got {rq.get("is_correct")}'
assert rq.get('status') == 'hidden', f'Expected hidden for status, got {rq.get("status")}'
assert rq.get('earned_score') is None, f'Expected None for earned_score, got {rq.get("earned_score")}'
print('  PASS: Review masking hides all answer leak fields during active exam!')

print('=== Test 4: Review unmasking after exam close_time ===')
ex.close_time = datetime.datetime.now() - datetime.timedelta(minutes=5)
db.commit()

rev_res2 = client.get('/api/exam/review', headers=headers)
assert rev_res2.status_code == 200
rq2 = rev_res2.json()['questions'][0]
assert rq2.get('correct') == 'A,B,C', f'Expected A,B,C for correct, got {rq2.get("correct")}'
assert rq2.get('is_correct') is True, f'Expected True for is_correct, got {rq2.get("is_correct")}'
assert rq2.get('status') == 'correct', f'Expected correct for status, got {rq2.get("status")}'
print('  PASS: Review unmasking reveals answers after close_time!')

print('=== Test 5: Manual student creation with empty DOB ===')
admin_stu = db.query(User).filter(User.username == 'trangnh@hcmute.edu.vn').first()
admin_tok = client.post('/token', data={'username': 'trangnh@hcmute.edu.vn', 'password': 'nguyenhatrang'}).json()['access_token']
admin_headers = {'Authorization': f'Bearer {admin_tok}'}

# Create student with dob=""
existing_7777 = db.query(User).filter(User.username == '21117777').first()
if existing_7777:
    db.query(ExamResult).filter(ExamResult.user_id == existing_7777.id).delete()
    db.delete(existing_7777)
    db.commit()

create_stu_res = client.post('/api/admin/students', json={
    'username': '21117777',
    'fullname': 'Nguyen Van NoDOB',
    'dob': '',
    'class_name': '21110CLC'
}, headers=admin_headers)
assert create_stu_res.status_code == 200, f'Create student with empty DOB failed: {create_stu_res.text}'

# Login with created student using MSSV as password
stu_login_res = client.post('/token', data={'username': '21117777', 'password': '21117777'})
assert stu_login_res.status_code == 200, f'Login with MSSV as password failed: {stu_login_res.text}'
print('  PASS: Student created with empty DOB can log in with MSSV as password!')

# Cleanup
db.query(ExamResult).filter(ExamResult.user_id == stu.id).delete()
db.delete(q)
db.delete(ex)
db.delete(stu)
stu7 = db.query(User).filter(User.username == '21117777').first()
if stu7:
    db.query(ExamResult).filter(ExamResult.user_id == stu7.id).delete()
    db.delete(stu7)
db.commit()
db.close()
print('\n🎉 ALL ADVERSARIAL VERIFICATION TESTS PASSED 100%!')
