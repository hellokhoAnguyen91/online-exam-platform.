import os
import json
import datetime
import random
import io
import re
import asyncio
from typing import List, Optional, Dict, Any

from fastapi import FastAPI, Depends, HTTPException, UploadFile, File, Form, status, Request, Query
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse, StreamingResponse
from fastapi.security import OAuth2PasswordBearer, OAuth2PasswordRequestForm
from fastapi.middleware.cors import CORSMiddleware
from sqlalchemy.orm import Session
from sqlalchemy import desc, func
from passlib.context import CryptContext
from pydantic import BaseModel
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import uvicorn
from jose import JWTError, jwt

from database import engine, SessionLocal
from models import User, Exam, Question, ExamResult, migrate_database
from docx_parser import parse_docx_questions, strip_question_prefix
from student_parser import parse_student_file
from pdf_export import generate_batch_exam_zip, generate_candidate_pdf, sanitize_filename
from excel_export import generate_candidate_excel, generate_batch_excel_zip, generate_candidate_audit_excel

SECRET_KEY = os.environ.get("SECRET_KEY", "hcmute-exam-secure-key-2026-prod-jwt")
ALGORITHM = "HS256"
ACCESS_TOKEN_EXPIRE_MINUTES = 240

# Run schema migrations on boot
migrate_database()

# --- FastAPI App ---
app = FastAPI(title="Hệ thống Thi Trắc nghiệm Trực tuyến Chuyên nghiệp", version="2.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)

# --- Auth Setup ---
pwd_context = CryptContext(schemes=["bcrypt"], deprecated="auto")
oauth2_scheme = OAuth2PasswordBearer(tokenUrl="token", auto_error=False)

def verify_password(plain_password, hashed_password):
    return pwd_context.verify(plain_password, hashed_password)

def get_password_hash(password):
    return pwd_context.hash(password)

def get_db():
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()

# Create default admin if not exists
def create_default_admin():
    db = SessionLocal()
    admin = db.query(User).filter(User.username == "trangnh@hcmute.edu.vn").first()
    if not admin:
        new_admin = User(
            username="trangnh@hcmute.edu.vn",
            password=get_password_hash("nguyenhatrang"),
            is_admin=True,
            fullname="Cô Nguyễn Hà Trang"
        )
        db.add(new_admin)
        db.commit()
    db.close()

create_default_admin()

# --- Schemas ---
class Token(BaseModel):
    access_token: str
    token_type: str

class ExamCreate(BaseModel):
    title: str
    code: Optional[str] = ""
    description: Optional[str] = ""
    num_questions: int = 10
    duration_minutes: int = 30
    open_time: Optional[str] = None
    close_time: Optional[str] = None
    allow_review: bool = True
    shuffle_questions: bool = True
    mc_max_score: float = 7.0
    essay_max_score: float = 3.0

class ExamUpdate(BaseModel):
    title: Optional[str] = None
    code: Optional[str] = None
    description: Optional[str] = None
    num_questions: Optional[int] = None
    duration_minutes: Optional[int] = None
    open_time: Optional[str] = None
    close_time: Optional[str] = None
    allow_review: Optional[bool] = None
    shuffle_questions: Optional[bool] = None
    mc_max_score: Optional[float] = None
    essay_max_score: Optional[float] = None

class QuestionCreate(BaseModel):
    exam_id: Optional[int] = None
    content: str
    option_a: Optional[str] = ""
    option_b: Optional[str] = ""
    option_c: Optional[str] = ""
    option_d: Optional[str] = ""
    option_e: Optional[str] = ""
    option_f: Optional[str] = ""
    correct_option: Optional[str] = "" # "A", "B", "C", "D" or "A,B,C" or ""
    explanation: Optional[str] = None
    question_type: Optional[str] = "multiple_choice"
    score_weight: Optional[float] = 1.0

class QuestionUpdate(BaseModel):
    content: Optional[str] = None
    option_a: Optional[str] = None
    option_b: Optional[str] = None
    option_c: Optional[str] = None
    option_d: Optional[str] = None
    option_e: Optional[str] = None
    option_f: Optional[str] = None
    correct_option: Optional[str] = None
    explanation: Optional[str] = None
    question_type: Optional[str] = None
    score_weight: Optional[float] = None

class StudentCreate(BaseModel):
    username: str # MSSV
    fullname: str
    dob: Optional[str] = ""      # Ngày sinh / Password (nếu để trống mặc định là MSSV)
    class_name: Optional[str] = ""

class StudentUpdate(BaseModel):
    username: Optional[str] = None
    fullname: Optional[str] = None
    dob: Optional[str] = None
    class_name: Optional[str] = None

class AnswerPayload(BaseModel):
    answers: dict

class EssayScoreUpdate(BaseModel):
    scores: dict # {"q_id": score}

def create_access_token(data: dict, expires_delta: Optional[datetime.timedelta] = None):
    to_encode = data.copy()
    if expires_delta:
        expire = datetime.datetime.now() + expires_delta
    else:
        expire = datetime.datetime.now() + datetime.timedelta(minutes=ACCESS_TOKEN_EXPIRE_MINUTES)
    to_encode.update({"exp": expire})
    return jwt.encode(to_encode, SECRET_KEY, algorithm=ALGORITHM)

def get_current_user(
    request: Request,
    token_header: Optional[str] = Depends(oauth2_scheme),
    token_query: Optional[str] = Query(None, alias="token"),
    db: Session = Depends(get_db)
):
    credentials_exception = HTTPException(
        status_code=status.HTTP_401_UNAUTHORIZED,
        detail="Could not validate credentials",
        headers={"WWW-Authenticate": "Bearer"},
    )
    token = token_header or token_query or request.query_params.get("token")
    if not token:
        auth_header = request.headers.get("Authorization") or request.headers.get("authorization")
        if auth_header and auth_header.startswith("Bearer "):
            token = auth_header[7:].strip()
    if not token:
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Not authenticated",
            headers={"WWW-Authenticate": "Bearer"},
        )
    try:
        payload = jwt.decode(token, SECRET_KEY, algorithms=[ALGORITHM])
        username: str = payload.get("sub")
        if username is None:
            raise credentials_exception
    except JWTError:
        raise credentials_exception
    user = db.query(User).filter(User.username == username).first()
    if user is None:
        raise credentials_exception
    return user

def parse_iso_dt(dt_str: Optional[str]) -> Optional[datetime.datetime]:
    if not dt_str:
        return None
    try:
        dt = datetime.datetime.fromisoformat(dt_str.replace("Z", "+00:00"))
        if dt.tzinfo is not None:
            dt = dt.replace(tzinfo=None)
        return dt
    except Exception:
        return None

def normalize_dob(dob_str: str) -> str:
    if not dob_str:
        return ""
    s = dob_str.strip()
    m1 = re.match(r'^(\d{1,2})[/.-](\d{1,2})[/.-](\d{4})$', s)
    if m1:
        d, m, y = m1.groups()
        return f"{y}-{int(m):02d}-{int(d):02d}"
    m2 = re.match(r'^(\d{4})[/.-](\d{1,2})[/.-](\d{1,2})$', s)
    if m2:
        y, m, d = m2.groups()
        return f"{y}-{int(m):02d}-{int(d):02d}"
    return s

# ==========================================
# AUTH ENDPOINTS
# ==========================================
@app.post("/token", response_model=Token)
def login_for_access_token(form_data: OAuth2PasswordRequestForm = Depends(), db: Session = Depends(get_db)):
    clean_username = form_data.username.strip()
    
    # 1. Flexible Username Resolution
    # Exact match (case-insensitive)
    user = db.query(User).filter(func.lower(User.username) == clean_username.lower()).first()
    
    # Prefix tolerance: with or without 'sv' / 'SV'
    if not user:
        if clean_username.lower().startswith("sv"):
            alt_username = clean_username[2:].strip()
        else:
            alt_username = "sv" + clean_username
        user = db.query(User).filter(func.lower(User.username) == alt_username.lower()).first()
        
    if not user:
        raise HTTPException(status_code=401, detail="Mã số sinh viên/Tên đăng nhập không tồn tại trong hệ thống. Vui lòng kiểm tra lại MSSV!")
        
    pwd_input = form_data.password.strip()
    is_valid = False
    
    # 2. Multi-Tier Password Verification
    # Tier 1: Fast O(1) Student Plaintext Matches (Applies ONLY to Students, never Admin!)
    if not user.is_admin:
        def strip_sv_prefix(val: str) -> str:
            v = val.strip().lower()
            return v[2:] if v.startswith("sv") else v

        u_stripped = strip_sv_prefix(user.username)
        c_stripped = strip_sv_prefix(clean_username)
        p_stripped = strip_sv_prefix(pwd_input)

        # Fallback 1: Student uses their own MSSV as password!
        if p_stripped and (p_stripped == u_stripped or p_stripped == c_stripped):
            is_valid = True
            
        # Match against stored plain DOB if present
        if not is_valid and user.dob:
            clean_stored_dob = user.dob.strip()
            if pwd_input.lower() == clean_stored_dob.lower() or normalize_dob(pwd_input) == normalize_dob(clean_stored_dob):
                is_valid = True

    # Tier 2: Cryptographic Bcrypt Hash Matches (For Admin or hashed accounts)
    if not is_valid and user.password:
        if verify_password(pwd_input, user.password):
            is_valid = True
            
        if not is_valid:
            norm_pwd = normalize_dob(pwd_input)
            if norm_pwd and norm_pwd != pwd_input and verify_password(norm_pwd, user.password):
                is_valid = True
                
        if not is_valid:
            m_8d = re.match(r'^(\d{2})(\d{2})(\d{4})$', pwd_input)
            if m_8d:
                iso_d = f"{m_8d.group(3)}-{m_8d.group(2)}-{m_8d.group(1)}"
                if verify_password(iso_d, user.password):
                    is_valid = True
            
    if not is_valid:
        raise HTTPException(
            status_code=401, 
            detail="Mật khẩu (Ngày sinh) không chính xác! Bạn có thể nhập ngày sinh dạng DD/MM/YYYY (vd: 26/05/2005), hoặc nhập chính Mã số sinh viên của bạn làm mật khẩu."
        )
    
    access_token = create_access_token(data={"sub": user.username})
    return {"access_token": access_token, "token_type": "bearer"}

@app.get("/api/me")
def read_users_me(current_user: User = Depends(get_current_user)):
    return {
        "id": current_user.id,
        "username": current_user.username,
        "fullname": current_user.fullname,
        "is_admin": current_user.is_admin
    }

# ==========================================
# EXAM MANAGEMENT (MULTI-EXAM & ARCHIVE)
# ==========================================
@app.get("/api/admin/exams")
def list_exams(db: Session = Depends(get_db), current_user: User = Depends(get_current_user)):
    if not current_user.is_admin:
        raise HTTPException(status_code=403, detail="Không có quyền truy cập")
    
    exams = db.query(Exam).order_by(desc(Exam.is_active), desc(Exam.created_at)).all()
    out = []
    for ex in exams:
        q_count = db.query(Question).filter(Question.exam_id == ex.id).count()
        results = db.query(ExamResult).filter(ExamResult.exam_id == ex.id, ExamResult.status == "submitted").all()
        candidate_count = len(results)
        scored_results = [r.score for r in results if r.score is not None]
        avg_score = round(sum(scored_results) / len(scored_results), 2) if scored_results else 0.0
        
        now = datetime.datetime.now()
        exam_status = "active" if ex.is_active else ("archived" if ex.is_archived else "inactive")
        open_t = ex.open_time.replace(tzinfo=None) if (ex.open_time and ex.open_time.tzinfo) else ex.open_time
        close_t = ex.close_time.replace(tzinfo=None) if (ex.close_time and ex.close_time.tzinfo) else ex.close_time
        if ex.is_active:
            if open_t and now < open_t:
                exam_status = "upcoming"
            elif close_t and now > close_t:
                exam_status = "closed"

        out.append({
            "id": ex.id,
            "title": ex.title,
            "code": ex.code,
            "description": ex.description,
            "num_questions": ex.num_questions,
            "duration_minutes": ex.duration_minutes,
            "open_time": ex.open_time.isoformat() if ex.open_time else None,
            "close_time": ex.close_time.isoformat() if ex.close_time else None,
            "is_active": ex.is_active,
            "is_archived": ex.is_archived,
            "archived_at": ex.archived_at.isoformat() if ex.archived_at else None,
            "allow_review": ex.allow_review,
            "shuffle_questions": ex.shuffle_questions,
            "mc_max_score": getattr(ex, 'mc_max_score', 7.0) if getattr(ex, 'mc_max_score', None) is not None else 7.0,
            "essay_max_score": getattr(ex, 'essay_max_score', 3.0) if getattr(ex, 'essay_max_score', None) is not None else 3.0,
            "created_at": ex.created_at.strftime("%Y-%m-%d %H:%M:%S") if ex.created_at else None,
            "question_count": q_count,
            "candidate_count": candidate_count,
            "avg_score": avg_score,
            "status": exam_status
        })
    return out

@app.post("/api/admin/exams")
def create_exam(payload: ExamCreate, db: Session = Depends(get_db), current_user: User = Depends(get_current_user)):
    if not current_user.is_admin:
        raise HTTPException(status_code=403, detail="Không có quyền truy cập")
    
    new_exam = Exam(
        title=payload.title.strip(),
        code=payload.code.strip() if payload.code else f"EXAM-{random.randint(100, 999)}",
        description=payload.description,
        num_questions=max(1, payload.num_questions),
        duration_minutes=max(1, payload.duration_minutes),
        open_time=parse_iso_dt(payload.open_time),
        close_time=parse_iso_dt(payload.close_time),
        allow_review=payload.allow_review,
        shuffle_questions=payload.shuffle_questions,
        mc_max_score=payload.mc_max_score if payload.mc_max_score is not None else 7.0,
        essay_max_score=payload.essay_max_score if payload.essay_max_score is not None else 3.0,
        is_active=False,
        is_archived=False,
        created_at=datetime.datetime.now()
    )
    db.add(new_exam)
    db.commit()
    db.refresh(new_exam)
    return {"detail": "Tạo kỳ thi thành công", "exam_id": new_exam.id}

@app.put("/api/admin/exams/{exam_id}")
def update_exam(exam_id: int, payload: ExamUpdate, db: Session = Depends(get_db), current_user: User = Depends(get_current_user)):
    if not current_user.is_admin:
        raise HTTPException(status_code=403, detail="Không có quyền truy cập")
    exam = db.query(Exam).filter(Exam.id == exam_id).first()
    if not exam:
        raise HTTPException(status_code=404, detail="Không tìm thấy kỳ thi")
        
    if payload.title is not None: exam.title = payload.title.strip()
    if payload.code is not None: exam.code = payload.code.strip()
    if payload.description is not None: exam.description = payload.description
    if payload.num_questions is not None: exam.num_questions = max(1, payload.num_questions)
    if payload.duration_minutes is not None: exam.duration_minutes = max(1, payload.duration_minutes)
    if payload.open_time is not None: exam.open_time = parse_iso_dt(payload.open_time)
    if payload.close_time is not None: exam.close_time = parse_iso_dt(payload.close_time)
    if payload.allow_review is not None: exam.allow_review = payload.allow_review
    if payload.shuffle_questions is not None: exam.shuffle_questions = payload.shuffle_questions
    if payload.mc_max_score is not None: exam.mc_max_score = max(0.0, float(payload.mc_max_score))
    if payload.essay_max_score is not None: exam.essay_max_score = max(0.0, float(payload.essay_max_score))
    
    db.commit()
    return {"detail": "Cập nhật kỳ thi thành công"}

@app.post("/api/admin/exams/{exam_id}/activate")
def activate_exam(exam_id: int, db: Session = Depends(get_db), current_user: User = Depends(get_current_user)):
    if not current_user.is_admin:
        raise HTTPException(status_code=403, detail="Không có quyền truy cập")
    exam = db.query(Exam).filter(Exam.id == exam_id).first()
    if not exam:
        raise HTTPException(status_code=404, detail="Không tìm thấy kỳ thi")
        
    # Deactivate all other exams
    db.query(Exam).update({Exam.is_active: False})
    exam.is_active = True
    exam.is_archived = False
    db.commit()
    return {"detail": f"Đã kích hoạt kỳ thi '{exam.title}'"}

@app.post("/api/admin/exams/{exam_id}/archive")
def archive_exam(exam_id: int, db: Session = Depends(get_db), current_user: User = Depends(get_current_user)):
    if not current_user.is_admin:
        raise HTTPException(status_code=403, detail="Không có quyền truy cập")
    exam = db.query(Exam).filter(Exam.id == exam_id).first()
    if not exam:
        raise HTTPException(status_code=404, detail="Không tìm thấy kỳ thi")
        
    exam.is_active = False
    exam.is_archived = True
    exam.archived_at = datetime.datetime.now()
    db.commit()
    return {"detail": f"Đã lưu trữ thành công kỳ thi '{exam.title}'"}

@app.post("/api/admin/exams/{exam_id}/unarchive")
def unarchive_exam(exam_id: int, db: Session = Depends(get_db), current_user: User = Depends(get_current_user)):
    if not current_user.is_admin:
        raise HTTPException(status_code=403, detail="Không có quyền truy cập")
    exam = db.query(Exam).filter(Exam.id == exam_id).first()
    if not exam:
        raise HTTPException(status_code=404, detail="Không tìm thấy kỳ thi")
        
    exam.is_archived = False
    exam.archived_at = None
    db.commit()
    return {"detail": f"Đã mở khóa lưu trữ kỳ thi '{exam.title}'"}

@app.delete("/api/admin/exams/{exam_id}")
def delete_exam(exam_id: int, db: Session = Depends(get_db), current_user: User = Depends(get_current_user)):
    if not current_user.is_admin:
        raise HTTPException(status_code=403, detail="Không có quyền truy cập")
    exam = db.query(Exam).filter(Exam.id == exam_id).first()
    if not exam:
        raise HTTPException(status_code=404, detail="Không tìm thấy kỳ thi")
    if exam.is_active:
        raise HTTPException(status_code=400, detail="Không thể xóa kỳ thi đang kích hoạt. Vui lòng kích hoạt kỳ thi khác trước!")
        
    # Delete questions and results associated
    db.query(Question).filter(Question.exam_id == exam_id).delete()
    db.query(ExamResult).filter(ExamResult.exam_id == exam_id).delete()
    db.delete(exam)
    db.commit()
    return {"detail": "Đã xóa kỳ thi và dữ liệu liên quan thành công"}

@app.get("/api/admin/exams/{exam_id}/backup")
def export_exam_backup(exam_id: int, db: Session = Depends(get_db), current_user: User = Depends(get_current_user)):
    """Export complete snapshot package of the exam (config, questions with images, candidate audit records)."""
    if not current_user.is_admin:
        raise HTTPException(status_code=403, detail="Không có quyền truy cập")
    exam = db.query(Exam).filter(Exam.id == exam_id).first()
    if not exam:
        raise HTTPException(status_code=404, detail="Không tìm thấy kỳ thi")
        
    questions = db.query(Question).filter(Question.exam_id == exam_id).all()
    results = db.query(ExamResult).filter(ExamResult.exam_id == exam_id).all()
    
    q_data = [{
        "content": q.content,
        "option_a": q.option_a,
        "option_b": q.option_b,
        "option_c": q.option_c,
        "option_d": q.option_d,
        "option_e": getattr(q, 'option_e', '') or '',
        "option_f": getattr(q, 'option_f', '') or '',
        "question_type": getattr(q, 'question_type', 'multiple_choice') or 'multiple_choice',
        "score_weight": getattr(q, 'score_weight', 1.0) if getattr(q, 'score_weight', None) is not None else 1.0,
        "correct_option": q.correct_option,
        "explanation": q.explanation
    } for q in questions]
    
    r_data = []
    for r in results:
        u = db.query(User).filter(User.id == r.user_id).first()
        r_data.append({
            "username": u.username if u else "",
            "fullname": u.fullname if u else "",
            "score": r.score,
            "max_score": r.max_score,
            "correct_count": r.correct_count,
            "total_questions": r.total_questions,
            "answers": r.answers,
            "answers_detail": r.answers_detail,
            "start_time": r.start_time.isoformat() if r.start_time else None,
            "submit_time": r.submit_time.isoformat() if r.submit_time else None,
            "duration_seconds": r.duration_seconds,
            "status": r.status
        })
        
    backup_pkg = {
        "version": "2.0",
        "exported_at": datetime.datetime.now().isoformat(),
        "exam": {
            "title": exam.title,
            "code": exam.code,
            "description": exam.description,
            "num_questions": exam.num_questions,
            "duration_minutes": exam.duration_minutes,
            "open_time": exam.open_time.isoformat() if exam.open_time else None,
            "close_time": exam.close_time.isoformat() if exam.close_time else None,
            "allow_review": exam.allow_review,
            "shuffle_questions": exam.shuffle_questions,
            "shuffle_options": getattr(exam, 'shuffle_options', False),
            "mc_max_score": getattr(exam, 'mc_max_score', 7.0),
            "essay_max_score": getattr(exam, 'essay_max_score', 3.0)
        },
        "questions": q_data,
        "results": r_data
    }
    
    json_str = json.dumps(backup_pkg, ensure_ascii=False, indent=2)
    stream = io.BytesIO(json_str.encode('utf-8'))
    filename = f"sao_luu_{re.sub(r'[^a-zA-Z0-9_-]', '_', exam.title)}_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
    return StreamingResponse(
        stream,
        media_type="application/json",
        headers={"Content-Disposition": f"attachment; filename={filename}"}
    )

@app.post("/api/admin/exams/restore")
async def restore_exam_backup(file: UploadFile = File(...), db: Session = Depends(get_db), current_user: User = Depends(get_current_user)):
    if not current_user.is_admin:
        raise HTTPException(status_code=403, detail="Không có quyền truy cập")
    try:
        content = await file.read()
        pkg = json.loads(content.decode('utf-8'))
        exam_info = pkg.get("exam", {})
        
        new_exam = Exam(
            title=exam_info.get("title", "Kỳ thi phục hồi") + " (Đã phục hồi)",
            code=exam_info.get("code", "") + "_restored",
            description=exam_info.get("description", ""),
            num_questions=exam_info.get("num_questions", 10),
            duration_minutes=exam_info.get("duration_minutes", 30),
            open_time=parse_iso_dt(exam_info.get("open_time")),
            close_time=parse_iso_dt(exam_info.get("close_time")),
            allow_review=exam_info.get("allow_review", True),
            shuffle_questions=exam_info.get("shuffle_questions", True),
            shuffle_options=exam_info.get("shuffle_options", False),
            mc_max_score=exam_info.get("mc_max_score", 7.0),
            essay_max_score=exam_info.get("essay_max_score", 3.0),
            is_active=False,
            is_archived=True,
            archived_at=datetime.datetime.now()
        )
        db.add(new_exam)
        db.commit()
        db.refresh(new_exam)
        
        # Restore questions
        for q in pkg.get("questions", []):
            db.add(Question(
                exam_id=new_exam.id,
                content=q.get("content", ""),
                option_a=q.get("option_a", ""),
                option_b=q.get("option_b", ""),
                option_c=q.get("option_c", ""),
                option_d=q.get("option_d", ""),
                option_e=q.get("option_e", ""),
                option_f=q.get("option_f", ""),
                question_type=q.get("question_type", "multiple_choice") or "multiple_choice",
                score_weight=q.get("score_weight", 1.0) if q.get("score_weight") is not None else 1.0,
                correct_option=q.get("correct_option", "A"),
                explanation=q.get("explanation")
            ))
        db.commit()
        
        return {"detail": f"Phục hồi kỳ thi '{new_exam.title}' thành công!", "exam_id": new_exam.id}
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Lỗi khi đọc file sao lưu: {str(e)}")

# ==========================================
# QUESTION BANK MANAGEMENT (.DOCX PARSER)
# ==========================================
@app.post("/api/admin/upload_questions")
async def upload_questions(
    file: UploadFile = File(...),
    exam_id: Optional[int] = Form(None),
    replace_existing: bool = Form(True),
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user)
):
    if not current_user.is_admin:
        raise HTTPException(status_code=403, detail="Không có quyền truy cập")
        
    # Target exam
    target_exam = None
    if exam_id:
        target_exam = db.query(Exam).filter(Exam.id == exam_id).first()
    if not target_exam:
        target_exam = db.query(Exam).filter(Exam.is_active == True).first()
    if not target_exam:
        target_exam = db.query(Exam).first()
    if not target_exam:
        raise HTTPException(status_code=400, detail="Chưa có kỳ thi nào để gán câu hỏi. Vui lòng tạo kỳ thi trước!")
        
    contents = await file.read()
    file_stream = io.BytesIO(contents)
    
    try:
        parsed_questions = parse_docx_questions(file_stream)
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Lỗi khi phân tích file Word: {str(e)}")
        
    if not parsed_questions:
        raise HTTPException(
            status_code=400,
            detail="Không tìm thấy câu hỏi trắc nghiệm hợp lệ trong file. Vui lòng kiểm tra định dạng (Câu 1: ... A. B. C. D. Đáp án: ...)"
        )
        
    if replace_existing:
        db.query(Question).filter(Question.exam_id == target_exam.id).delete()
        
    for q in parsed_questions:
        db.add(Question(
            exam_id=target_exam.id,
            content=strip_question_prefix(q['content']),
            option_a=q.get('option_a', ''),
            option_b=q.get('option_b', ''),
            option_c=q.get('option_c', ''),
            option_d=q.get('option_d', ''),
            option_e=q.get('option_e', ''),
            option_f=q.get('option_f', ''),
            correct_option=q['correct_option'],
            question_type=q.get('question_type', 'multiple_choice'),
            score_weight=q.get('score_weight', 1.0),
            created_at=datetime.datetime.now()
        ))
    db.commit()
    
    return {
        "detail": f"Đã bóc tách thành công {len(parsed_questions)} câu hỏi kèm hình ảnh đầy đủ vào kỳ thi '{target_exam.title}'!",
        "count": len(parsed_questions),
        "exam_id": target_exam.id
    }

@app.get("/api/admin/questions")
def get_questions(exam_id: Optional[int] = None, db: Session = Depends(get_db), current_user: User = Depends(get_current_user)):
    if not current_user.is_admin:
        raise HTTPException(status_code=403, detail="Không có quyền truy cập")
    
    query = db.query(Question)
    if exam_id:
        query = query.filter(Question.exam_id == exam_id)
    else:
        active_exam = db.query(Exam).filter(Exam.is_active == True).first()
        if active_exam:
            query = query.filter(Question.exam_id == active_exam.id)
            
    questions = query.order_by(Question.id.asc()).all()
    mc_qs = [q for q in questions if q.question_type != 'essay']
    essay_qs = [q for q in questions if q.question_type == 'essay']
    sorted_questions = mc_qs + essay_qs

    q_list = [{
        "id": q.id,
        "exam_id": q.exam_id,
        "content": q.content,
        "option_a": q.option_a,
        "option_b": q.option_b,
        "option_c": q.option_c,
        "option_d": q.option_d,
        "option_e": getattr(q, 'option_e', '') or '',
        "option_f": getattr(q, 'option_f', '') or '',
        "correct_option": q.correct_option,
        "explanation": q.explanation,
        "question_type": q.question_type,
        "score_weight": q.score_weight
    } for q in sorted_questions]
    return {"total": len(q_list), "count": len(q_list), "mc_count": len(mc_qs), "essay_count": len(essay_qs), "questions": q_list}

@app.post("/api/admin/questions")
def create_question_manual(payload: QuestionCreate, db: Session = Depends(get_db), current_user: User = Depends(get_current_user)):
    if not current_user.is_admin:
        raise HTTPException(status_code=403, detail="Không có quyền truy cập")
    
    target_exam_id = payload.exam_id
    if not target_exam_id:
        active_exam = db.query(Exam).filter(Exam.is_active == True).first()
        if not active_exam:
            active_exam = db.query(Exam).first()
        if active_exam:
            target_exam_id = active_exam.id

    content = strip_question_prefix(payload.content.strip())
    if not content:
        raise HTTPException(status_code=400, detail="Nội dung câu hỏi không được để trống")
    
    q_type = payload.question_type or "multiple_choice"
    correct_opt = payload.correct_option.strip().upper() if payload.correct_option else ""
    if q_type == "multiple_choice":
        if correct_opt and correct_opt not in ["A", "B", "C", "D", "E", "F"]:
            raise HTTPException(status_code=400, detail="Đáp án đúng phải là một trong các giá trị: A, B, C, D, E, F hoặc để trống")
    elif q_type == "multi_select":
        ans_letters = re.findall(r'[A-Fa-f]', correct_opt)
        correct_opt = ",".join(sorted(list(set(ch.upper() for ch in ans_letters))))
    elif q_type == "essay":
        correct_opt = ""

    q = Question(
        exam_id=target_exam_id,
        content=content,
        option_a=payload.option_a.strip() if payload.option_a else "",
        option_b=payload.option_b.strip() if payload.option_b else "",
        option_c=payload.option_c.strip() if payload.option_c else "",
        option_d=payload.option_d.strip() if payload.option_d else "",
        option_e=payload.option_e.strip() if payload.option_e else "",
        option_f=payload.option_f.strip() if payload.option_f else "",
        correct_option=correct_opt,
        explanation=payload.explanation.strip() if payload.explanation else None,
        question_type=q_type,
        score_weight=payload.score_weight if payload.score_weight is not None else 1.0
    )
    db.add(q)
    db.commit()
    db.refresh(q)
    return {
        "detail": "Đã thêm câu hỏi mới thành công",
        "question": {
            "id": q.id,
            "exam_id": q.exam_id,
            "content": q.content,
            "option_a": q.option_a,
            "option_b": q.option_b,
            "option_c": q.option_c,
            "option_d": q.option_d,
            "option_e": getattr(q, 'option_e', '') or '',
            "option_f": getattr(q, 'option_f', '') or '',
            "correct_option": q.correct_option,
            "explanation": q.explanation,
            "question_type": q.question_type,
            "score_weight": q.score_weight
        }
    }

@app.put("/api/admin/questions/{question_id}")
def update_question_manual(question_id: int, payload: QuestionUpdate, db: Session = Depends(get_db), current_user: User = Depends(get_current_user)):
    if not current_user.is_admin:
        raise HTTPException(status_code=403, detail="Không có quyền truy cập")
    q = db.query(Question).filter(Question.id == question_id).first()
    if not q:
        raise HTTPException(status_code=404, detail="Không tìm thấy câu hỏi")

    if payload.content is not None:
        c = strip_question_prefix(payload.content.strip())
        if not c:
            raise HTTPException(status_code=400, detail="Nội dung câu hỏi không được để trống")
        q.content = c
    if payload.option_a is not None:
        q.option_a = payload.option_a.strip()
    if payload.option_b is not None:
        q.option_b = payload.option_b.strip()
    if payload.option_c is not None:
        q.option_c = payload.option_c.strip()
    if payload.option_d is not None:
        q.option_d = payload.option_d.strip()
    if payload.option_e is not None:
        q.option_e = payload.option_e.strip()
    if payload.option_f is not None:
        q.option_f = payload.option_f.strip()
    if payload.question_type is not None:
        q.question_type = payload.question_type
    if payload.score_weight is not None:
        q.score_weight = payload.score_weight

    if payload.correct_option is not None:
        opt = payload.correct_option.strip().upper()
        if q.question_type == "multi_select":
            ans_letters = re.findall(r'[A-Fa-f]', opt)
            q.correct_option = ",".join(sorted(list(set(ch.upper() for ch in ans_letters))))
        elif q.question_type == "essay":
            q.correct_option = ""
        else:
            q.correct_option = opt
    elif q.question_type == "essay":
        q.correct_option = ""
    if payload.explanation is not None:
        q.explanation = payload.explanation.strip() if payload.explanation else None

    db.commit()
    db.refresh(q)
    return {
        "detail": "Đã cập nhật câu hỏi thành công",
        "question": {
            "id": q.id,
            "exam_id": q.exam_id,
            "content": q.content,
            "option_a": q.option_a,
            "option_b": q.option_b,
            "option_c": q.option_c,
            "option_d": q.option_d,
            "option_e": getattr(q, 'option_e', '') or '',
            "option_f": getattr(q, 'option_f', '') or '',
            "correct_option": q.correct_option,
            "explanation": q.explanation,
            "question_type": q.question_type,
            "score_weight": q.score_weight
        }
    }

@app.delete("/api/admin/questions/{question_id}")
def delete_question(question_id: int, db: Session = Depends(get_db), current_user: User = Depends(get_current_user)):
    if not current_user.is_admin:
        raise HTTPException(status_code=403, detail="Không có quyền truy cập")
    q = db.query(Question).filter(Question.id == question_id).first()
    if not q:
        raise HTTPException(status_code=404, detail="Không tìm thấy câu hỏi")
    db.delete(q)
    db.commit()
    return {"detail": "Đã xóa câu hỏi"}

@app.delete("/api/admin/questions")
def clear_all_questions(exam_id: Optional[int] = None, db: Session = Depends(get_db), current_user: User = Depends(get_current_user)):
    if not current_user.is_admin:
        raise HTTPException(status_code=403, detail="Không có quyền truy cập")
    query = db.query(Question)
    if exam_id:
        query = query.filter(Question.exam_id == exam_id)
    count = query.delete()
    db.commit()
    return {"detail": f"Đã xóa toàn bộ {count} câu hỏi trong ngân hàng đề thi"}

# ==========================================
# STUDENT MANAGEMENT & SMART EXTRACTION
# ==========================================
@app.post("/api/admin/upload_students")
async def upload_students(file: UploadFile = File(...), db: Session = Depends(get_db), current_user: User = Depends(get_current_user)):
    if not current_user.is_admin:
        raise HTTPException(status_code=403, detail="Không có quyền truy cập")
    
    contents = await file.read()
    try:
        parsed = parse_student_file(contents, file.filename)
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Lỗi khi bóc tách danh sách sinh viên: {str(e)}")
        
    students = parsed["students"]
    if not students:
        raise HTTPException(status_code=400, detail="Không tìm thấy bản ghi sinh viên hợp lệ nào trong tệp tải lên.")

    count = 0
    new_count = 0
    updated_count = 0
    
    for idx, s in enumerate(students):
        raw_mssv = s["mssv"]
        raw_name = s["fullname"]
        raw_dob = s.get("dob") or ""
        raw_class = s.get("class_name", "")
        raw_order = s.get("order_index", idx + 1)
        
        pwd_val = raw_dob.strip() if (raw_dob and raw_dob.strip()) else raw_mssv
        user = db.query(User).filter(User.username == raw_mssv).first()
        if not user:
            user = User(
                username=raw_mssv,
                password=get_password_hash(pwd_val),
                is_admin=False,
                fullname=raw_name,
                dob=raw_dob if raw_dob else None,
                class_name=raw_class,
                order_index=raw_order
            )
            db.add(user)
            new_count += 1
        else:
            if not user.is_admin:
                user.password = get_password_hash(pwd_val)
                user.fullname = raw_name
                user.dob = raw_dob if raw_dob else None
                user.class_name = raw_class
                user.order_index = raw_order
                updated_count += 1
        count += 1
        
    db.commit()
    return {
        "detail": f"Đã bóc tách thông minh thành công {count} thí sinh (Thêm mới: {new_count}, Cập nhật: {updated_count})!",
        "count": count,
        "new_count": new_count,
        "updated_count": updated_count,
        "detected_columns": parsed["detected_columns"],
        "ignored_columns": parsed["ignored_columns"],
        "header_row_index": parsed["header_row_index"],
        "preview": students[:5]
    }

@app.post("/api/admin/students")
def create_student_manual(payload: StudentCreate, db: Session = Depends(get_db), current_user: User = Depends(get_current_user)):
    if not current_user.is_admin:
        raise HTTPException(status_code=403, detail="Không có quyền truy cập")
        
    mssv = payload.username.strip()
    if not mssv:
        raise HTTPException(status_code=400, detail="Mã số sinh viên (MSSV) không được để trống")
    fullname = payload.fullname.strip()
    if not fullname:
        raise HTTPException(status_code=400, detail="Họ và tên không được để trống")
    dob_raw = (payload.dob or "").strip()
    if not dob_raw:
        dob_norm = None
        pwd_val = mssv
    else:
        dob_norm = normalize_dob(dob_raw)
        pwd_val = dob_norm

    existing = db.query(User).filter(User.username == mssv).first()
    if existing:
        raise HTTPException(status_code=400, detail=f"Mã số sinh viên '{mssv}' đã tồn tại trong hệ thống")

    new_user = User(
        username=mssv,
        password=get_password_hash(pwd_val),
        is_admin=False,
        fullname=fullname,
        dob=dob_norm,
        class_name=payload.class_name.strip() if payload.class_name else ""
    )
    db.add(new_user)
    db.commit()
    db.refresh(new_user)
    return {
        "detail": f"Đã thêm thí sinh '{fullname}' ({mssv}) thành công",
        "student": {
            "id": new_user.id,
            "username": new_user.username,
            "fullname": new_user.fullname,
            "dob": new_user.dob or "",
            "class_name": new_user.class_name or ""
        }
    }

@app.put("/api/admin/students/{user_id}")
def update_student_manual(user_id: int, payload: StudentUpdate, db: Session = Depends(get_db), current_user: User = Depends(get_current_user)):
    if not current_user.is_admin:
        raise HTTPException(status_code=403, detail="Không có quyền truy cập")
        
    student = db.query(User).filter(User.id == user_id, User.is_admin == False).first()
    if not student:
        raise HTTPException(status_code=404, detail="Không tìm thấy thí sinh")

    if payload.username is not None:
        new_mssv = payload.username.strip()
        if not new_mssv:
            raise HTTPException(status_code=400, detail="MSSV không được để trống")
        if new_mssv != student.username:
            check_dup = db.query(User).filter(User.username == new_mssv).first()
            if check_dup:
                raise HTTPException(status_code=400, detail=f"Mã số sinh viên '{new_mssv}' đã tồn tại")
            student.username = new_mssv

    if payload.fullname is not None:
        new_name = payload.fullname.strip()
        if not new_name:
            raise HTTPException(status_code=400, detail="Họ và tên không được để trống")
        student.fullname = new_name

    if payload.dob is not None and payload.dob.strip():
        norm_dob = normalize_dob(payload.dob.strip())
        student.password = get_password_hash(norm_dob)
        student.dob = norm_dob

    if payload.class_name is not None:
        student.class_name = payload.class_name.strip()

    db.commit()
    db.refresh(student)
    return {
        "detail": f"Đã cập nhật thông tin thí sinh '{student.fullname}' ({student.username}) thành công",
        "student": {
            "id": student.id,
            "username": student.username,
            "fullname": student.fullname,
            "dob": student.dob or "",
            "class_name": student.class_name or ""
        }
    }

@app.delete("/api/admin/students/{user_id}")
def delete_student_manual(user_id: int, db: Session = Depends(get_db), current_user: User = Depends(get_current_user)):
    if not current_user.is_admin:
        raise HTTPException(status_code=403, detail="Không có quyền truy cập")
        
    student = db.query(User).filter(User.id == user_id).first()
    if not student:
        raise HTTPException(status_code=404, detail="Không tìm thấy thí sinh")
    if student.is_admin:
        raise HTTPException(status_code=400, detail="Không thể xóa tài khoản Quản trị viên (Admin)")

    # Delete any exam results associated with student
    db.query(ExamResult).filter(ExamResult.user_id == user_id).delete()
    db.delete(student)
    db.commit()
    return {"detail": f"Đã xóa thí sinh '{student.fullname}' ({student.username}) và bài làm liên quan thành công"}

@app.delete("/api/admin/students")
def clear_all_students(db: Session = Depends(get_db), current_user: User = Depends(get_current_user)):
    if not current_user.is_admin:
        raise HTTPException(status_code=403, detail="Không có quyền truy cập")
    non_admins = db.query(User).filter(User.is_admin == False).all()
    count = len(non_admins)
    non_admin_ids = [u.id for u in non_admins]
    if non_admin_ids:
        db.query(ExamResult).filter(ExamResult.user_id.in_(non_admin_ids)).delete(synchronize_session=False)
        db.query(User).filter(User.is_admin == False).delete(synchronize_session=False)
        db.commit()
    return {"detail": f"Đã xóa toàn bộ {count} thí sinh khỏi hệ thống"}

@app.get("/api/admin/students")
def list_students(exam_id: Optional[int] = None, db: Session = Depends(get_db), current_user: User = Depends(get_current_user)):
    if not current_user.is_admin:
        raise HTTPException(status_code=403, detail="Không có quyền truy cập")
    
    target_exam_id = exam_id
    if not target_exam_id:
        active_exam = db.query(Exam).filter(Exam.is_active == True).first()
        if active_exam:
            target_exam_id = active_exam.id
            
    students = db.query(User).filter(User.is_admin == False).order_by(
        User.order_index.asc(),
        User.id.asc()
    ).all()
    out = []
    for s in students:
        res = None
        if target_exam_id:
            res = db.query(ExamResult).filter(ExamResult.exam_id == target_exam_id, ExamResult.user_id == s.id).first()
        
        exam_status = "not_started"
        score = None
        submit_time = None
        result_id = None
        if res:
            exam_status = res.status
            score = res.score
            submit_time = res.submit_time.strftime("%H:%M:%S %d/%m/%Y") if res.submit_time else None
            result_id = res.id
            
        out.append({
            "id": s.id,
            "username": s.username,
            "fullname": s.fullname,
            "dob": s.dob or "",
            "class_name": s.class_name or "",
            "order_index": s.order_index or 0,
            "exam_status": exam_status,
            "score": score,
            "submit_time": submit_time,
            "result_id": result_id
        })
    return out

@app.post("/api/admin/students/{user_id}/reset_exam")
def reset_student_exam(user_id: int, exam_id: Optional[int] = None, db: Session = Depends(get_db), current_user: User = Depends(get_current_user)):
    """Allows teacher to reset a candidate's exam session in case of power loss or disconnect."""
    if not current_user.is_admin:
        raise HTTPException(status_code=403, detail="Không có quyền truy cập")
        
    if exam_id:
        results = db.query(ExamResult).filter(ExamResult.user_id == user_id, ExamResult.exam_id == exam_id).all()
    else:
        results = db.query(ExamResult).filter(ExamResult.user_id == user_id).all()
    for res in results:
        db.delete(res)
    if results:
        db.commit()
    return {"detail": "Đã đặt lại trạng thái làm bài cho thí sinh thành công"}

def calculate_exam_score(audit_details: list, q_dict: dict, exam: Optional[Exam]) -> dict:
    """
    Computes candidate score respecting the configured score split between 
    Trắc nghiệm (MC) and Tự luận (Essay).
    """
    mc_items = [d for d in audit_details if q_dict.get(d.get("id")) and q_dict[d["id"]].question_type != "essay"]
    essay_items = [d for d in audit_details if q_dict.get(d.get("id")) and q_dict[d["id"]].question_type == "essay"]
    
    # 1. Effective max points
    if not essay_items:
        eff_mc_max = 10.0
        eff_essay_max = 0.0
    elif not mc_items:
        eff_mc_max = 0.0
        eff_essay_max = 10.0
    else:
        cfg_mc = float(getattr(exam, 'mc_max_score', 7.0) if exam and getattr(exam, 'mc_max_score', None) is not None else 7.0)
        cfg_essay = float(getattr(exam, 'essay_max_score', 3.0) if exam and getattr(exam, 'essay_max_score', None) is not None else 3.0)
        tot_cfg = cfg_mc + cfg_essay
        if tot_cfg <= 0:
            cfg_mc = 7.0
            cfg_essay = 3.0
            tot_cfg = 10.0
        eff_mc_max = round((cfg_mc / tot_cfg) * 10.0, 2)
        eff_essay_max = round(10.0 - eff_mc_max, 2)

    # 2. Scale factor for MC
    sum_mc_weights = sum(float(q_dict[d["id"]].score_weight or 1.0) for d in mc_items)
    mc_scale = (eff_mc_max / sum_mc_weights) if sum_mc_weights > 0 else 0.0

    # 3. Scale factor for Essay
    sum_essay_weights = sum(float(q_dict[d["id"]].score_weight or 1.0) for d in essay_items)
    essay_scale = (eff_essay_max / sum_essay_weights) if sum_essay_weights > 0 else 0.0

    has_pending = False
    mc_earned = 0.0
    correct_count = 0
    for d in mc_items:
        q_obj = q_dict.get(d.get("id"))
        weight = float(q_obj.score_weight or 1.0) if q_obj else 1.0
        d["score_weight"] = round(weight * mc_scale, 2)
        
        if d.get("status") == "pending_grading":
            has_pending = True
            d["points"] = 0.0
        elif d.get("status") == "graded" and "essay_score" in d and d["essay_score"] is not None:
            raw_sc = float(d["essay_score"])
            item_pts = round(raw_sc * mc_scale, 2)
            d["earned_score"] = raw_sc
            d["points"] = item_pts
            mc_earned += item_pts
            if raw_sc > 0:
                correct_count += 1
        else:
            raw_earned = float(d.get("earned_score", 0.0) if d.get("earned_score") is not None else (weight if d.get("is_correct") else 0.0))
            item_pts = round(raw_earned * mc_scale, 2)
            d["earned_score"] = raw_earned
            d["points"] = item_pts
            mc_earned += item_pts
            if d.get("is_correct") or d.get("status") == "correct":
                correct_count += 1
    essay_earned = 0.0
    for d in essay_items:
        q_obj = q_dict.get(d["id"])
        raw_weight = float(q_obj.score_weight or 1.0) if q_obj else 1.0
        d["score_weight"] = round(raw_weight * essay_scale, 2)
        
        if "essay_score" in d and d["essay_score"] is not None:
            raw_sc = float(d["essay_score"])
            item_pts = round(raw_sc * essay_scale, 2)
            d["points"] = item_pts
            essay_earned += item_pts
        elif d.get("status") == "unanswered":
            d["essay_score"] = 0.0
            d["points"] = 0.0
        else:
            has_pending = True
            d["points"] = 0.0

    total_score = round(mc_earned + essay_earned, 2) if not has_pending else None

    return {
        "score": total_score,
        "mc_score": round(mc_earned, 2),
        "mc_max_score": eff_mc_max,
        "essay_score": round(essay_earned, 2) if not has_pending else None,
        "essay_max_score": eff_essay_max,
        "has_pending": has_pending,
        "correct_count": correct_count,
        "mc_scale": round(mc_scale, 4),
        "essay_scale": round(essay_scale, 4)
    }

@app.post("/api/admin/results/{result_id}/score_essay")
def score_essay(result_id: int, payload: EssayScoreUpdate, db: Session = Depends(get_db), current_user: User = Depends(get_current_user)):
    if not current_user.is_admin:
        raise HTTPException(status_code=403, detail="Không có quyền truy cập")
        
    res = db.query(ExamResult).filter(ExamResult.id == result_id).first()
    if not res:
        raise HTTPException(status_code=404, detail="Không tìm thấy bài thi")
    exam = db.query(Exam).filter(Exam.id == res.exam_id).first()
        
    audit_details = json.loads(res.answers_detail) if res.answers_detail else []
    q_ids = json.loads(res.questions) if res.questions else []
    questions = db.query(Question).filter(Question.id.in_(q_ids)).all()
    q_dict = {q.id: q for q in questions}

    for detail in audit_details:
        q_id = detail.get("id")
        q = q_dict.get(q_id)
        if not q:
            continue
            
        score_val = payload.scores.get(str(q_id)) if str(q_id) in payload.scores else payload.scores.get(q_id)
        if score_val is not None:
            score = float(score_val)
            max_w = q.score_weight if (q and q.score_weight is not None) else 1.0
            if score < 0.0 or score > (max_w + 1e-5):
                raise HTTPException(status_code=400, detail=f"Điểm số ({score}) không hợp lệ! Điểm phải nằm trong khoảng từ 0.0 đến {max_w}")
            detail["essay_score"] = score
            detail["status"] = "graded"

    scoring = calculate_exam_score(audit_details, q_dict, exam)
    res.score = scoring["score"]
    res.correct_count = scoring["correct_count"]
    res.max_score = 10.0
    res.answers_detail = json.dumps(audit_details, ensure_ascii=False)
    db.commit()
    
    return {"detail": "Đã cập nhật điểm thành công", "new_score": res.score, "needs_grading": scoring["has_pending"]}

# ==========================================
# RESULTS & AUDIT REVIEW SYSTEM
# ==========================================
def auto_close_exam_result(result: ExamResult, db: Session, exam: Optional[Exam] = None) -> ExamResult:
    """Auto-finalize and grade an exam session whose time limit has expired."""
    if not exam:
        exam = db.query(Exam).filter(Exam.id == result.exam_id).first()
        
    q_ids = json.loads(result.questions) if result.questions else []
    questions = db.query(Question).filter(Question.id.in_(q_ids)).all()
    q_dict = {q.id: q for q in questions}
    
    # Ensure MC questions are grouped first, then Essay questions
    mc_ids = [qid for qid in q_ids if q_dict.get(qid) and q_dict[qid].question_type != 'essay']
    essay_ids = [qid for qid in q_ids if q_dict.get(qid) and q_dict[qid].question_type == 'essay']
    ordered_q_ids = mc_ids + essay_ids if (mc_ids or essay_ids) else q_ids
    
    answers_dict = json.loads(result.answers) if result.answers else {}
    correct_count = 0
    audit_details = []
    
    for idx, qid in enumerate(ordered_q_ids):
        q = q_dict.get(qid)
        if not q:
            continue
        ans = answers_dict.get(str(qid)) or answers_dict.get(qid)
        
        if isinstance(ans, (list, tuple, set)):
            ans_str = ",".join(str(x).strip() for x in ans if str(x).strip())
        else:
            ans_str = str(ans).strip() if ans is not None else ""
        is_c = False
        st = "unanswered"
        weight = q.score_weight or 1.0
        earned_score = 0.0

        if q.question_type == "essay":
            if ans_str:
                st = "pending_grading"
            else:
                st = "unanswered"
                earned_score = 0.0
        else:
            correct_val = (q.correct_option or "").strip()
            if not correct_val:
                is_c = False
                st = "unanswered" if not ans_str else "pending_grading"
            elif q.question_type == "multi_select":
                if isinstance(ans, (list, tuple, set)):
                    sel_set = {str(x).strip().upper() for x in ans if str(x).strip()}
                elif isinstance(ans, str):
                    sel_set = set(re.findall(r'[A-Za-z]', ans.upper())) if ans.strip() else set()
                else:
                    sel_set = set()
                corr_set = set(re.findall(r'[A-Za-z]', correct_val.upper())) if correct_val else set()
                if not sel_set:
                    st = "unanswered"
                    is_c = False
                    earned_score = 0.0
                else:
                    num_corr = len(sel_set & corr_set)
                    num_wrong = len(sel_set - corr_set)
                    total_corr = len(corr_set)
                    fraction = max(0.0, (num_corr - num_wrong) / total_corr) if total_corr > 0 else 0.0
                    earned_score = fraction * weight
                    if fraction >= 1.0:
                        is_c = True
                        st = "correct"
                    elif fraction > 0.0:
                        is_c = False
                        st = "partial"
                    else:
                        is_c = False
                        st = "incorrect"
            else:
                if isinstance(ans, (list, tuple, set)):
                    chosen_opt = str(list(ans)[0]).strip().upper() if ans else ""
                else:
                    chosen_opt = str(ans).strip().upper() if ans is not None else ""
                is_c = bool(chosen_opt and chosen_opt == correct_val.upper())
                st = "correct" if is_c else ("incorrect" if chosen_opt else "unanswered")
                earned_score = weight if is_c else 0.0
                
            if is_c:
                correct_count += 1
            
        audit_details.append({
            "q_idx": idx + 1,
            "id": q.id,
            "content": q.content,
            "option_a": q.option_a,
            "option_b": q.option_b,
            "option_c": q.option_c,
            "option_d": q.option_d,
            "option_e": getattr(q, 'option_e', '') or '',
            "option_f": getattr(q, 'option_f', '') or '',
            "question_type": q.question_type or 'multiple_choice',
            "score_weight": weight,
            "selected": ans,
            "correct": q.correct_option,
            "is_correct": is_c,
            "status": st,
            "earned_score": round(earned_score, 4)
        })
        
    dur_limit_sec = (exam.duration_minutes * 60) if exam else 1800
    
    scoring = calculate_exam_score(audit_details, q_dict, exam)
    result.score = scoring["score"]
    result.max_score = 10.0
    result.correct_count = scoring["correct_count"]
    result.total_questions = len(q_ids)
    result.answers_detail = json.dumps(audit_details, ensure_ascii=False)
    result.duration_seconds = dur_limit_sec
    result.submit_time = result.start_time + datetime.timedelta(seconds=dur_limit_sec)
    result.status = "submitted"
    db.commit()
    return result

def sync_expired_sessions(db: Session, exam_id: Optional[int] = None):
    """Scan and automatically finalize any student session that has exceeded exam duration.
    Only processes in-progress sessions; does not modify already-submitted sessions.
    """
    now = datetime.datetime.now()
    # Auto-close in-progress sessions that have timed out
    query = db.query(ExamResult).filter(ExamResult.status == "in_progress")
    if exam_id:
        query = query.filter(ExamResult.exam_id == exam_id)
    in_progress_results = query.all()
    
    for r in in_progress_results:
        exam = db.query(Exam).filter(Exam.id == r.exam_id).first()
        if not exam:
            continue
        duration_sec = exam.duration_minutes * 60
        if r.start_time:
            elapsed = (now - r.start_time).total_seconds()
            if elapsed >= duration_sec:
                auto_close_exam_result(r, db, exam)

@app.get("/api/admin/results")
def get_results(exam_id: Optional[int] = None, db: Session = Depends(get_db), current_user: User = Depends(get_current_user)):
    if not current_user.is_admin:
        raise HTTPException(status_code=403, detail="Không có quyền truy cập")
        
    target_exam_id = exam_id
    if not target_exam_id:
        active_exam = db.query(Exam).filter(Exam.is_active == True).first()
        if active_exam:
            target_exam_id = active_exam.id

    # Sync and auto-close any sessions that have exceeded time limit
    sync_expired_sessions(db, target_exam_id)
            
    query = db.query(ExamResult)
    if target_exam_id:
        query = query.filter(ExamResult.exam_id == target_exam_id)
        
    results = query.order_by(desc(ExamResult.submit_time)).all()
    
    scores = [r.score for r in results if r.score is not None]
    candidate_count = len(scores)
    avg_score = round(sum(scores) / candidate_count, 2) if candidate_count > 0 else 0.0
    highest = max(scores) if candidate_count > 0 else 0.0
    lowest = min(scores) if candidate_count > 0 else 0.0
    pass_count = sum(1 for s in scores if s >= 5.0)
    pass_rate = round(pass_count / candidate_count * 100, 1) if candidate_count > 0 else 0.0
    
    out = []
    for r in results:
        user = db.query(User).filter(User.id == r.user_id).first()
        duration_fmt = ""
        if r.duration_seconds:
            m = r.duration_seconds // 60
            s = r.duration_seconds % 60
            duration_fmt = f"{m}p {s}s"
            
        out.append({
            "id": r.id,
            "exam_id": r.exam_id,
            "user_id": r.user_id,
            "username": user.username if user else "",
            "fullname": user.fullname if user else "",
            "score": r.score,
            "max_score": r.max_score or 10.0,
            "correct_count": r.correct_count,
            "total_questions": r.total_questions,
            "duration_str": duration_fmt,
            "duration_seconds": r.duration_seconds,
            "start_time": r.start_time.strftime("%H:%M:%S %d/%m/%Y") if r.start_time else "",
            "submit_time": r.submit_time.strftime("%H:%M:%S %d/%m/%Y") if r.submit_time else "",
            "status": r.status
        })
        
    return {
        "stats": {
            "total_candidates": candidate_count,
            "avg_score": avg_score,
            "highest_score": highest,
            "lowest_score": lowest,
            "pass_count": pass_count,
            "pass_rate": pass_rate
        },
        "results": out
    }

@app.get("/api/admin/results/{result_id}/detail")
def get_result_detail(result_id: int, db: Session = Depends(get_db), current_user: User = Depends(get_current_user)):
    """Comprehensive candidate submission audit record (Đối chiếu bài làm chi tiết từng câu)."""
    if not current_user.is_admin:
        raise HTTPException(status_code=403, detail="Không có quyền truy cập")
        
    res = db.query(ExamResult).filter(ExamResult.id == result_id).first()
    if not res:
        raise HTTPException(status_code=404, detail="Không tìm thấy bài thi")
        
    student = db.query(User).filter(User.id == res.user_id).first()
    exam = db.query(Exam).filter(Exam.id == res.exam_id).first()
    
    # Audit detail: inspect answers_detail or build from snapshot
    audit_questions = []
    if res.answers_detail:
        try:
            audit_questions = json.loads(res.answers_detail)
        except Exception:
            audit_questions = []
            
    # If answers_detail not present (e.g. legacy result), construct from raw answers & questions
    if not audit_questions and res.questions:
        try:
            q_ids = json.loads(res.questions)
            raw_answers = json.loads(res.answers) if res.answers else {}
            questions = db.query(Question).filter(Question.id.in_(q_ids)).all()
            q_dict = {q.id: q for q in questions}
            
            for idx, qid in enumerate(q_ids):
                q = q_dict.get(qid)
                if not q:
                    continue
                selected = raw_answers.get(str(qid)) or raw_answers.get(qid)
                is_correct = (selected == q.correct_option)
                st = "correct" if is_correct else ("incorrect" if selected else "unanswered")
                audit_questions.append({
                    "q_idx": idx + 1,
                    "id": q.id,
                    "content": q.content,
                    "option_a": q.option_a,
                    "option_b": q.option_b,
                    "option_c": q.option_c,
                    "option_d": q.option_d,
                    "option_e": getattr(q, 'option_e', '') or '',
                    "option_f": getattr(q, 'option_f', '') or '',
                    "question_type": q.question_type or 'multiple_choice',
                    "score_weight": q.score_weight or 1.0,
                    "selected": selected,
                    "correct": q.correct_option,
                    "is_correct": is_correct,
                    "status": st
                })
        except Exception as e:
            print(f"Error rebuilding audit detail: {e}")

    duration_str = ""
    if res.duration_seconds:
        m = res.duration_seconds // 60
        s = res.duration_seconds % 60
        duration_str = f"{m} phút {s} giây"
        
    mc_qs = [q for q in audit_questions if q.get("question_type") != "essay"]
    essay_qs = [q for q in audit_questions if q.get("question_type") == "essay"]
    audit_questions = mc_qs + essay_qs
    for idx, q in enumerate(audit_questions):
        q["q_idx"] = idx + 1
        
    q_dict_audit = {}
    for q in audit_questions:
        q_obj = db.query(Question).filter(Question.id == q.get("id")).first()
        if q_obj:
            q_dict_audit[q["id"]] = q_obj
        else:
            q_dict_audit[q["id"]] = Question(id=q["id"], question_type=q.get("question_type", "multiple_choice"), score_weight=q.get("score_weight", 1.0))
            
    scoring = calculate_exam_score(audit_questions, q_dict_audit, exam)

    return {
        "result_id": res.id,
        "candidate": {
            "mssv": student.username if student else "",
            "fullname": student.fullname if student else ""
        },
        "exam": {
            "title": exam.title if exam else "Kỳ thi trắc nghiệm",
            "duration_minutes": exam.duration_minutes if exam else 30
        },
        "score": scoring["score"] if res.score is not None else None,
        "max_score": res.max_score or 10.0,
        "correct_count": scoring["correct_count"],
        "total_questions": res.total_questions or len(audit_questions),
        "incorrect_count": ((res.total_questions or len(audit_questions)) - scoring["correct_count"]) if scoring["correct_count"] is not None else None,
        "start_time": res.start_time.strftime("%H:%M:%S %d/%m/%Y") if res.start_time else "",
        "submit_time": res.submit_time.strftime("%H:%M:%S %d/%m/%Y") if res.submit_time else "",
        "duration_str": duration_str,
        "client_ip": res.client_ip or "N/A",
        "user_agent": res.user_agent or "N/A",
        "mc_count": len(mc_qs),
        "mc_correct": scoring["correct_count"],
        "mc_score": scoring["mc_score"],
        "mc_max_score": scoring["mc_max_score"],
        "essay_count": len(essay_qs),
        "essay_score": scoring["essay_score"],
        "essay_max_score": scoring["essay_max_score"],
        "has_pending_essay": scoring["has_pending"],
        "questions": audit_questions
    }

@app.get("/api/admin/results/export")
def export_results_excel(exam_id: Optional[str] = None, db: Session = Depends(get_db), current_user: User = Depends(get_current_user)):
    """Export complete Excel package (.ZIP) containing class summary sheet and all candidate audit sheets."""
    if not current_user.is_admin:
        raise HTTPException(status_code=403, detail="Không có quyền truy cập")
        
    target_exam_id = None
    if exam_id is not None and str(exam_id).strip() and str(exam_id).strip().isdigit():
        target_exam_id = int(str(exam_id).strip())
        
    if not target_exam_id:
        active_exam = db.query(Exam).filter(Exam.is_active == True).first()
        if active_exam:
            target_exam_id = active_exam.id

    # Auto-close any expired sessions before exporting Excel
    if target_exam_id:
        sync_expired_sessions(db, target_exam_id)
            
    query = db.query(ExamResult)
    if target_exam_id:
        query = query.filter(ExamResult.exam_id == target_exam_id)
        
    results = query.order_by(desc(ExamResult.score)).all()
    exam = db.query(Exam).filter(Exam.id == target_exam_id).first() if target_exam_id else None
    
    zip_bytes, exported_count = generate_batch_excel_zip(exam, results, db)
    
    exam_title_slug = sanitize_filename(exam.title) if exam else "ky_thi"
    now_str = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"Goi_Bang_Diem_Excel_{exam_title_slug}_{now_str}.zip"
    
    return StreamingResponse(
        io.BytesIO(zip_bytes),
        media_type="application/zip",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'}
    )

@app.get("/api/admin/results/export_pdf_zip")
def export_results_pdf_zip(
    exam_id: Optional[str] = None,
    date: Optional[str] = None,
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user)
):
    """
    Download a ZIP archive containing individual PDF exam papers of all candidates 
    who participated in the specified exam (optionally filtered by test date).
    """
    if not current_user.is_admin:
        raise HTTPException(status_code=403, detail="Không có quyền truy cập")
        
    target_exam = None
    clean_exam_id = None
    if exam_id is not None and str(exam_id).strip() and str(exam_id).strip().isdigit():
        clean_exam_id = int(str(exam_id).strip())
        target_exam = db.query(Exam).filter(Exam.id == clean_exam_id).first()

    if not target_exam:
        target_exam = db.query(Exam).filter(Exam.is_active == True, Exam.is_archived == False).first()
    if not target_exam:
        target_exam = db.query(Exam).filter(Exam.is_active == True).first()
    if not target_exam:
        target_exam = db.query(Exam).order_by(Exam.id.desc()).first()
    if not target_exam:
        raise HTTPException(status_code=404, detail="Không tìm thấy kỳ thi phù hợp để xuất bài thi")

    # Auto-close any expired sessions before generating ZIP of PDFs
    sync_expired_sessions(db, target_exam.id)

    query = db.query(ExamResult).filter(ExamResult.exam_id == target_exam.id)
    if date:
        date_clean = date.strip().lower()
        target_dt = None
        if date_clean in ["today", "hôm nay"]:
            target_dt = datetime.date.today()
        else:
            try:
                target_dt = datetime.date.fromisoformat(date_clean)
            except ValueError:
                try:
                    parts = date_clean.split('/')
                    if len(parts) == 3:
                        target_dt = datetime.date(int(parts[2]), int(parts[1]), int(parts[0]))
                except Exception:
                    pass
        if target_dt:
            s_dt = datetime.datetime.combine(target_dt, datetime.time.min)
            e_dt = datetime.datetime.combine(target_dt, datetime.time.max)
            query = query.filter(ExamResult.start_time >= s_dt, ExamResult.start_time <= e_dt)

    results = query.order_by(ExamResult.id.asc()).all()
    if not results:
        raise HTTPException(status_code=400, detail="Kỳ thi này hiện chưa có bài thi nào được nộp để xuất file PDF.")

    zip_bytes, exported_count = generate_batch_exam_zip(target_exam, results, db)
    if exported_count == 0:
        raise HTTPException(status_code=500, detail="Không thể tạo file PDF bài thi nào cho thí sinh. Vui lòng kiểm tra định dạng dữ liệu.")
    safe_title = sanitize_filename(target_exam.title)
    now_str = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"Bai_thi_PDF_{safe_title}_{now_str}.zip"

    return StreamingResponse(
        io.BytesIO(zip_bytes),
        media_type="application/zip",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'}
    )

@app.get("/api/admin/results/{result_id}/export_pdf")
def export_single_result_pdf(
    result_id: int,
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user)
):
    """Download a single candidate's exam submission as an official PDF."""
    if not current_user.is_admin:
        raise HTTPException(status_code=403, detail="Không có quyền truy cập")
        
    res = db.query(ExamResult).filter(ExamResult.id == result_id).first()
    if not res:
        raise HTTPException(status_code=404, detail="Không tìm thấy bài làm của thí sinh")
        
    exam = db.query(Exam).filter(Exam.id == res.exam_id).first()
    student = db.query(User).filter(User.id == res.user_id).first()
    
    # Audit detail reconstruction
    audit_questions = []
    if res.answers_detail:
        try:
            audit_questions = json.loads(res.answers_detail)
        except Exception:
            audit_questions = []
            
    if not audit_questions and res.questions:
        try:
            q_ids = json.loads(res.questions) if isinstance(res.questions, str) else res.questions
            raw_answers = json.loads(res.answers) if (res.answers and isinstance(res.answers, str)) else (res.answers or {})
            questions = db.query(Question).filter(Question.id.in_(q_ids)).all()
            q_dict = {q.id: q for q in questions}
            for q_pos, qid in enumerate(q_ids):
                q = q_dict.get(qid)
                if not q:
                    continue
                selected = raw_answers.get(str(qid)) or raw_answers.get(qid)
                is_corr = (selected == q.correct_option) if selected else False
                st = "correct" if is_corr else ("incorrect" if selected else "unanswered")
                audit_questions.append({
                    "q_idx": q_pos + 1,
                    "id": q.id,
                    "content": q.content,
                    "option_a": q.option_a,
                    "option_b": q.option_b,
                    "option_c": q.option_c,
                    "option_d": q.option_d,
                    "selected": selected,
                    "correct": q.correct_option,
                    "is_correct": is_corr,
                    "status": st
                })
        except Exception as e:
            print(f"Error rebuilding audit: {e}")

    dur_sec = res.duration_seconds or 0
    dur_str = f"{dur_sec // 60}p {dur_sec % 60:02d}s" if dur_sec > 0 else "-"
    start_str = res.start_time.strftime("%d/%m/%Y %H:%M:%S") if res.start_time else "-"
    submit_str = res.submit_time.strftime("%d/%m/%Y %H:%M:%S") if res.submit_time else "-"
    status_text = "Đã nộp bài" if res.status == "submitted" else "Đang làm dở"
    if res.status == "auto_submitted":
        status_text = "Nộp tự động (Hết giờ)"

    exam_info = {
        "id": exam.id if exam else 1,
        "title": exam.title if exam else "Kỳ thi trắc nghiệm",
        "code": getattr(exam, 'code', '') or '',
        "duration_minutes": getattr(exam, 'duration_minutes', 30),
        "num_questions": getattr(exam, 'num_questions', 10)
    }
    candidate_info = {
        "username": student.username if student else "unknown",
        "fullname": student.fullname if student else "Thí sinh",
        "dob": getattr(student, 'dob', None) or "-",
        "score": res.score,
        "max_score": res.max_score or 10.0,
        "correct_count": res.correct_count or 0,
        "total_questions": res.total_questions or len(audit_questions),
        "duration_str": dur_str,
        "start_time_str": start_str,
        "submit_time_str": submit_str,
        "status_text": status_text
    }

    pdf_bytes = generate_candidate_pdf(exam_info, candidate_info, audit_questions)
    safe_name = sanitize_filename(f"BaiThi_{candidate_info['username']}_{candidate_info['fullname']}")
    filename = f"{safe_name}.pdf"

    return StreamingResponse(
        io.BytesIO(pdf_bytes),
        media_type="application/pdf",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'}
    )

@app.get("/api/admin/results/{result_id}/export_excel")
def export_single_result_excel(
    result_id: int,
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user)
):
    """Download a single candidate's exam submission audit as an official Excel (.xlsx) file."""
    if not current_user.is_admin:
        raise HTTPException(status_code=403, detail="Không có quyền truy cập")
        
    res = db.query(ExamResult).filter(ExamResult.id == result_id).first()
    if not res:
        raise HTTPException(status_code=404, detail="Không tìm thấy bài làm của thí sinh")
        
    exam = db.query(Exam).filter(Exam.id == res.exam_id).first()
    student = db.query(User).filter(User.id == res.user_id).first()
    
    excel_bytes = generate_candidate_audit_excel(res, exam, student, db)
    uname = student.username if student else f"user_{res.user_id}"
    fname = student.fullname if student else "thi_sinh"
    safe_name = sanitize_filename(f"BaiThi_{uname}_{fname}")
    filename = f"{safe_name}.xlsx"

    return StreamingResponse(
        io.BytesIO(excel_bytes),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'}
    )

# ==========================================
# CANDIDATE / STUDENT EXAM FLOW
# ==========================================
@app.get("/api/exam")
def get_exam(request: Request, db: Session = Depends(get_db), current_user: User = Depends(get_current_user)):
    if current_user.is_admin:
        raise HTTPException(status_code=400, detail="Tài khoản quản trị viên không tham gia làm bài thi.")
        
    now = datetime.datetime.now()
    
    # 1. Prioritize any existing in_progress session for this student so F5 does not re-roll questions or create duplicate
    active_exam = db.query(Exam).filter(Exam.is_active == True, Exam.is_archived == False).order_by(desc(Exam.id)).first()
    if not active_exam:
        raise HTTPException(status_code=400, detail="Hiện tại chưa có kỳ thi nào đang mở. Vui lòng liên hệ giáo viên!")
        
    result = db.query(ExamResult).filter(
        ExamResult.user_id == current_user.id,
        ExamResult.exam_id == active_exam.id,
        ExamResult.status == "in_progress"
    ).order_by(desc(ExamResult.start_time)).first()
    
    if not result:
        result = db.query(ExamResult).filter(
            ExamResult.user_id == current_user.id,
            ExamResult.exam_id == active_exam.id
        ).order_by(desc(ExamResult.start_time)).first()

    open_t = active_exam.open_time.replace(tzinfo=None) if (active_exam.open_time and active_exam.open_time.tzinfo) else active_exam.open_time
    close_t = active_exam.close_time.replace(tzinfo=None) if (active_exam.close_time and active_exam.close_time.tzinfo) else active_exam.close_time

    if open_t and now < open_t:
        raise HTTPException(status_code=400, detail=f"Kỳ thi chưa mở. Thời gian mở thi: {open_t.strftime('%H:%M %d/%m/%Y')}")
    if close_t and now > close_t:
        raise HTTPException(status_code=400, detail=f"Kỳ thi đã kết thúc vào lúc {close_t.strftime('%H:%M %d/%m/%Y')}")
        
    # If already submitted
    if result and result.status == "submitted":
        return {
            "status": "submitted",
            "score": result.score,
            "max_score": 10.0,
            "correct_count": result.correct_count,
            "total_questions": result.total_questions,
            "allow_review": active_exam.allow_review,
            "exam_title": active_exam.title,
            "needs_grading": (result.score is None)
        }
        
    # If new exam session for this candidate
    if not result:
        # Check again to guard against duplicate creation
        existing = db.query(ExamResult).filter(
            ExamResult.user_id == current_user.id,
            ExamResult.exam_id == active_exam.id
        ).first()
        if existing:
            result = existing
        else:
            questions = db.query(Question).filter(Question.exam_id == active_exam.id).order_by(Question.id.asc()).all()
            if not questions:
                raise HTTPException(status_code=400, detail="Chưa có câu hỏi nào trong ngân hàng đề thi của kỳ thi này. Vui lòng liên hệ giáo viên!")
                
            mc_questions = [q for q in questions if q.question_type != 'essay']
            essay_questions = [q for q in questions if q.question_type == 'essay']

            # MC logic: Take active_exam.num_questions if set, else all; shuffle ONLY if shuffle_questions is True
            if mc_questions:
                target_mc_count = min(active_exam.num_questions, len(mc_questions)) if (active_exam.num_questions and active_exam.num_questions > 0) else len(mc_questions)
                if active_exam.shuffle_questions:
                    selected_mc = random.sample(mc_questions, target_mc_count)
                else:
                    selected_mc = mc_questions[:target_mc_count]
            else:
                selected_mc = []

            # Essay logic: ALWAYS take ALL essay questions from bank, NEVER truncate, NEVER shuffle (fixed bank order)
            selected_essay = list(essay_questions)

            # Combined exam: Part 1 (MC) followed by Part 2 (Essay)
            selected = selected_mc + selected_essay
            if not selected:
                raise HTTPException(status_code=400, detail="Chưa có câu hỏi nào trong ngân hàng đề thi của kỳ thi này. Vui lòng liên hệ giáo viên!")

            q_ids = [q.id for q in selected]
            num_to_take = len(q_ids)
            
            client_ip = request.client.host if request.client else ""
            user_agent = request.headers.get("user-agent", "")
            
            result = ExamResult(
                exam_id=active_exam.id,
                user_id=current_user.id,
                start_time=now,
                questions=json.dumps(q_ids),
                total_questions=num_to_take,
                client_ip=client_ip,
                user_agent=user_agent,
                status="in_progress"
            )
            db.add(result)
            try:
                db.commit()
                db.refresh(result)
            except Exception:
                db.rollback()
                result = db.query(ExamResult).filter(
                    ExamResult.user_id == current_user.id,
                    ExamResult.exam_id == active_exam.id
                ).first()
        
    # Load candidate questions in their specific order: Part 1 (MC) first, then Part 2 (Essay)
    q_ids = json.loads(result.questions) if result.questions else []
    questions = db.query(Question).filter(Question.id.in_(q_ids)).all()
    q_dict = {q.id: q for q in questions}
    
    # Ensure ordered MC first, then Essay
    mc_ids = [qid for qid in q_ids if q_dict.get(qid) and q_dict[qid].question_type != 'essay']
    essay_ids = [qid for qid in q_ids if q_dict.get(qid) and q_dict[qid].question_type == 'essay']
    ordered_q_ids = mc_ids + essay_ids if (mc_ids or essay_ids) else q_ids
    
    q_list = []
    for qid in ordered_q_ids:
        q = q_dict.get(qid)
        if q:
            q_list.append({
                "id": q.id,
                "content": q.content,
                "option_a": q.option_a,
                "option_b": q.option_b,
                "option_c": q.option_c,
                "option_d": q.option_d,
                "option_e": getattr(q, 'option_e', '') or '',
                "option_f": getattr(q, 'option_f', '') or '',
                "question_type": q.question_type or 'multiple_choice',
                "score_weight": q.score_weight
            })
            
    # Calculate time left
    elapsed = (datetime.datetime.now() - result.start_time).total_seconds()
    time_left = max(0, active_exam.duration_minutes * 60 - elapsed)
    
    # Auto submit if time completely ran out
    if time_left <= 0:
        result = auto_close_exam_result(result, db, active_exam)
        return {
            "status": "submitted",
            "exam_id": active_exam.id,
            "score": result.score,
            "max_score": 10.0,
            "correct_count": result.correct_count,
            "total_questions": result.total_questions,
            "allow_review": active_exam.allow_review,
            "exam_title": active_exam.title,
            "needs_grading": (result.score is None)
        }
        
    return {
        "status": "ongoing",
        "exam_id": active_exam.id,
        "exam_title": active_exam.title,
        "duration_minutes": active_exam.duration_minutes,
        "questions": q_list,
        "mc_count": len(mc_ids),
        "essay_count": len(essay_ids),
        "time_left": int(time_left),
        "answers": json.loads(result.answers) if result.answers else {}
    }

@app.post("/api/exam/save_progress")
@app.post("/api/exam/autosave")
def save_progress(payload: AnswerPayload, db: Session = Depends(get_db), current_user: User = Depends(get_current_user)):
    result = db.query(ExamResult).filter(
        ExamResult.user_id == current_user.id,
        ExamResult.status == "in_progress"
    ).order_by(desc(ExamResult.start_time)).first()
    
    if not result:
        active_exam = db.query(Exam).filter(Exam.is_active == True, Exam.is_archived == False).first()
        if active_exam:
            result = db.query(ExamResult).filter(
                ExamResult.user_id == current_user.id,
                ExamResult.exam_id == active_exam.id
            ).first()
            
    if not result:
        return {"status": "no_active_session"}
        
    exam = db.query(Exam).filter(Exam.id == result.exam_id).first()
    if not exam:
        return {"status": "no_exam"}
        
    if result.status == "submitted":
        return {
            "status": "expired",
            "detail": f"Đã hết thời gian làm bài ({exam.duration_minutes} phút)! Bài thi của bạn đã được hệ thống tự động khóa và nộp điểm."
        }
        
    elapsed = (datetime.datetime.now() - result.start_time).total_seconds()
    duration_limit = exam.duration_minutes * 60
    if elapsed < duration_limit:
        result.answers = json.dumps(payload.answers)
        db.commit()
        return {"status": "saved"}
    else:
        # Time has expired! Auto-finalize and reject any late edits!
        auto_close_exam_result(result, db, exam)
        return {
            "status": "expired",
            "detail": f"Đã hết thời gian làm bài ({exam.duration_minutes} phút)! Bài thi của bạn đã được hệ thống tự động khóa và nộp điểm."
        }

@app.post("/api/exam/submit")
def submit_exam(payload: AnswerPayload, db: Session = Depends(get_db), current_user: User = Depends(get_current_user)):
    if current_user.is_admin:
        raise HTTPException(status_code=400, detail="Admin không tham gia làm bài thi")
        
    # Prioritize user's in-progress exam session to prevent race condition if admin changes active exam
    result = db.query(ExamResult).filter(
        ExamResult.user_id == current_user.id,
        ExamResult.status == "in_progress"
    ).order_by(desc(ExamResult.start_time)).first()
    
    if not result:
        active_exam = db.query(Exam).filter(Exam.is_active == True, Exam.is_archived == False).first()
        if active_exam:
            result = db.query(ExamResult).filter(
                ExamResult.user_id == current_user.id,
                ExamResult.exam_id == active_exam.id
            ).first()
            
    if not result:
        raise HTTPException(status_code=400, detail="Bài thi chưa được bắt đầu hoặc đã kết thúc")
        
    exam = db.query(Exam).filter(Exam.id == result.exam_id).first()
    if not exam:
        raise HTTPException(status_code=400, detail="Không tìm thấy kỳ thi tương ứng")
        
    if result.status == "submitted":
        return {
            "status": "submitted",
            "score": result.score,
            "max_score": result.max_score or 10.0,
            "correct_count": result.correct_count,
            "total_questions": result.total_questions,
            "allow_review": exam.allow_review,
            "needs_grading": (result.score is None)
        }
        
    answers_dict = payload.answers
    q_ids = json.loads(result.questions) if result.questions else []
    questions = db.query(Question).filter(Question.id.in_(q_ids)).all()
    q_dict = {q.id: q for q in questions}
    
    # Ensure ordered MC first, then Essay
    mc_ids = [qid for qid in q_ids if q_dict.get(qid) and q_dict[qid].question_type != 'essay']
    essay_ids = [qid for qid in q_ids if q_dict.get(qid) and q_dict[qid].question_type == 'essay']
    ordered_q_ids = mc_ids + essay_ids if (mc_ids or essay_ids) else q_ids
    
    correct_count = 0
    audit_details = []
    
    for idx, qid in enumerate(ordered_q_ids):
        q = q_dict.get(qid)
        if not q:
            continue
        selected = answers_dict.get(str(qid))
        
        if isinstance(selected, (list, tuple, set)):
            sel_str = ",".join(str(x).strip() for x in selected if str(x).strip())
        else:
            sel_str = str(selected).strip() if selected is not None else ""
        is_c = False
        st = "unanswered"
        weight = q.score_weight or 1.0
        earned_score = 0.0

        if q.question_type == "essay":
            if sel_str:
                st = "pending_grading"
            else:
                st = "unanswered"
                earned_score = 0.0
        else:
            correct_val = (q.correct_option or "").strip()
            if not correct_val:
                is_c = False
                st = "unanswered" if not sel_str else "pending_grading"
            elif q.question_type == "multi_select":
                if isinstance(selected, (list, tuple, set)):
                    sel_set = {str(x).strip().upper() for x in selected if str(x).strip()}
                elif isinstance(selected, str):
                    sel_set = set(re.findall(r'[A-Za-z]', selected.upper())) if selected.strip() else set()
                else:
                    sel_set = set()
                corr_set = set(re.findall(r'[A-Za-z]', correct_val.upper())) if correct_val else set()
                if not sel_set:
                    st = "unanswered"
                    is_c = False
                    earned_score = 0.0
                else:
                    num_corr = len(sel_set & corr_set)
                    num_wrong = len(sel_set - corr_set)
                    total_corr = len(corr_set)
                    fraction = max(0.0, (num_corr - num_wrong) / total_corr) if total_corr > 0 else 0.0
                    earned_score = fraction * weight
                    if fraction >= 1.0:
                        is_c = True
                        st = "correct"
                    elif fraction > 0.0:
                        is_c = False
                        st = "partial"
                    else:
                        is_c = False
                        st = "incorrect"
            else:
                if isinstance(selected, (list, tuple, set)):
                    chosen_opt = str(list(selected)[0]).strip().upper() if selected else ""
                else:
                    chosen_opt = str(selected).strip().upper() if selected is not None else ""
                is_c = bool(chosen_opt and chosen_opt == correct_val.upper())
                st = "correct" if is_c else ("incorrect" if chosen_opt else "unanswered")
                earned_score = weight if is_c else 0.0
                
            if is_c:
                correct_count += 1
            
        audit_details.append({
            "q_idx": idx + 1,
            "id": q.id,
            "content": q.content,
            "option_a": q.option_a,
            "option_b": q.option_b,
            "option_c": q.option_c,
            "option_d": q.option_d,
            "option_e": getattr(q, 'option_e', '') or '',
            "option_f": getattr(q, 'option_f', '') or '',
            "question_type": q.question_type or 'multiple_choice',
            "score_weight": weight,
            "selected": selected,
            "correct": q.correct_option,
            "is_correct": is_c,
            "status": st,
            "earned_score": round(earned_score, 4)
        })
        
    now = datetime.datetime.now()
    raw_duration = int((now - result.start_time).total_seconds())
    max_duration_sec = exam.duration_minutes * 60
    if raw_duration >= max_duration_sec:
        duration_sec = max_duration_sec
        final_submit_time = result.start_time + datetime.timedelta(seconds=max_duration_sec)
    else:
        duration_sec = raw_duration
        final_submit_time = now

    total_q = len(q_ids)
    scoring = calculate_exam_score(audit_details, q_dict, exam)
    scaled_score = scoring["score"]
    result.score = scaled_score
    result.answers = json.dumps(answers_dict)
    result.answers_detail = json.dumps(audit_details, ensure_ascii=False)
    result.max_score = 10.0
    result.correct_count = scoring["correct_count"]
    result.total_questions = total_q
    result.duration_seconds = duration_sec
    result.submit_time = final_submit_time
    result.status = "submitted"
    db.commit()
    
    return {
        "status": "submitted",
        "score": scaled_score,
        "max_score": 10.0,
        "correct_count": scoring["correct_count"],
        "total_questions": total_q,
        "duration_seconds": duration_sec,
        "allow_review": exam.allow_review,
        "needs_grading": scoring["has_pending"]
    }

@app.get("/api/exam/review")
def review_my_exam(db: Session = Depends(get_db), current_user: User = Depends(get_current_user)):
    """Allow candidate to review their own submitted exam if allowed by exam config."""
    # Find user's latest submitted exam result
    res = db.query(ExamResult).filter(
        ExamResult.user_id == current_user.id,
        ExamResult.status == "submitted"
    ).order_by(desc(ExamResult.submit_time)).first()
    
    if not res:
        raise HTTPException(status_code=404, detail="Bạn chưa nộp bài thi")
        
    exam = db.query(Exam).filter(Exam.id == res.exam_id).first()
    if not exam:
        raise HTTPException(status_code=400, detail="Không tìm thấy kỳ thi")
        
    now = datetime.datetime.now()
    close_dt = exam.close_time.replace(tzinfo=None) if (exam.close_time and exam.close_time.tzinfo) else exam.close_time
    if close_dt is not None:
        is_closed = (now > close_dt)
    else:
        is_closed = not bool(exam.is_active)
    can_show_answers = bool(exam.allow_review) and is_closed

    audit_questions = []
    if res.answers_detail:
        try:
            audit_questions = json.loads(res.answers_detail)
        except:
            audit_questions = []
            
    # Fallback reconstruction if answers_detail was missing
    if not audit_questions and res.questions:
        try:
            q_ids = json.loads(res.questions)
            raw_answers = json.loads(res.answers) if res.answers else {}
            questions = db.query(Question).filter(Question.id.in_(q_ids)).all()
            q_dict = {q.id: q for q in questions}
            for idx, qid in enumerate(q_ids):
                q = q_dict.get(qid)
                if not q:
                    continue
                selected = raw_answers.get(str(qid)) or raw_answers.get(qid)
                is_correct = (selected == q.correct_option)
                st = "correct" if is_correct else ("incorrect" if selected else "unanswered")
                audit_questions.append({
                    "q_idx": idx + 1,
                    "id": q.id,
                    "content": q.content,
                    "option_a": q.option_a,
                    "option_b": q.option_b,
                    "option_c": q.option_c,
                    "option_d": q.option_d,
                    "option_e": getattr(q, 'option_e', '') or '',
                    "option_f": getattr(q, 'option_f', '') or '',
                    "question_type": q.question_type or 'multiple_choice',
                    "score_weight": q.score_weight or 1.0,
                    "selected": selected,
                    "correct": q.correct_option,
                    "is_correct": is_correct,
                    "status": st
                })
        except Exception as ex:
            print(f"Error rebuilding student review audit detail: {ex}")
            
    safe_audit_questions = []
    for item in audit_questions:
        item_copy = dict(item)
        if not can_show_answers:
            item_copy["correct"] = None
            item_copy["is_correct"] = None
            item_copy["status"] = "hidden"
            item_copy["earned_score"] = None
            item_copy["explanation"] = None
        safe_audit_questions.append(item_copy)

    mc_qs = [q for q in safe_audit_questions if q.get("question_type") != "essay"]
    essay_qs = [q for q in safe_audit_questions if q.get("question_type") == "essay"]
    safe_audit_questions = mc_qs + essay_qs
    for idx, q in enumerate(safe_audit_questions):
        q["q_idx"] = idx + 1
    mc_correct = sum(1 for q in mc_qs if q.get("is_correct") or q.get("status") == "correct")
    mc_score = sum(q.get("earned_score", 0.0) or 0.0 for q in mc_qs)
    mc_weight = sum(q.get("score_weight", 1.0) or 1.0 for q in mc_qs)
    essay_score = sum(q.get("essay_score", 0.0) or 0.0 for q in essay_qs if "essay_score" in q and q.get("essay_score") is not None)
    essay_weight = sum(q.get("score_weight", 1.0) or 1.0 for q in essay_qs)
    has_pending_essay = any(q.get("status") == "pending_grading" or ("essay_score" not in q and q.get("status") != "unanswered") for q in essay_qs)

    return {
        "exam_title": exam.title,
        "fullname": current_user.fullname,
        "mssv": current_user.username,
        "score": res.score,
        "max_score": 10.0,
        "correct_count": res.correct_count,
        "total_questions": res.total_questions or len(safe_audit_questions),
        "duration_seconds": res.duration_seconds,
        "submit_time": res.submit_time.strftime("%H:%M:%S %d/%m/%Y") if res.submit_time else "",
        "mc_count": len(mc_qs),
        "mc_correct": mc_correct,
        "mc_score": round(mc_score, 2),
        "mc_max_score": round(mc_weight, 2),
        "essay_count": len(essay_qs),
        "essay_score": round(essay_score, 2) if not has_pending_essay else None,
        "essay_max_score": round(essay_weight, 2),
        "has_pending_essay": has_pending_essay,
        "questions": safe_audit_questions
    }

@app.api_route("/api/exams", methods=["GET", "HEAD"])
def get_public_exams(db: Session = Depends(get_db), current_user: User = Depends(get_current_user)):
    query = db.query(Exam).filter(Exam.is_archived == False)
    if not current_user.is_admin:
        query = query.filter(Exam.is_active == True)
    exams = query.all()
    return [{
        "id": ex.id,
        "title": ex.title,
        "code": ex.code,
        "description": ex.description,
        "num_questions": ex.num_questions,
        "duration_minutes": ex.duration_minutes,
        "is_active": ex.is_active,
        "open_time": ex.open_time.isoformat() if ex.open_time else None,
        "close_time": ex.close_time.isoformat() if ex.close_time else None,
        "allow_review": ex.allow_review
    } for ex in exams]

@app.api_route("/health", methods=["GET", "HEAD"])
@app.api_route("/api/health", methods=["GET", "HEAD"])
def health_check():
    return {"status": "ok"}

@app.api_route("/api/admin/config", methods=["GET", "POST"])
def admin_config_endpoint(
    request: Request,
    num_questions: Optional[int] = Form(None),
    duration_minutes: Optional[int] = Form(None),
    open_time: Optional[str] = Form(None),
    close_time: Optional[str] = Form(None),
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user)
):
    if not current_user.is_admin:
        raise HTTPException(status_code=403, detail="Not authorized")
    active_exam = db.query(Exam).filter(Exam.is_active == True).first()
    if not active_exam:
        active_exam = db.query(Exam).first()
    if request.method == "POST":
        if active_exam:
            if num_questions is not None: active_exam.num_questions = num_questions
            if duration_minutes is not None: active_exam.duration_minutes = duration_minutes
            if open_time is not None: active_exam.open_time = parse_iso_dt(open_time)
            if close_time is not None: active_exam.close_time = parse_iso_dt(close_time)
            db.commit()
        return {"detail": "Config updated"}
    else:
        return {
            "num_questions": active_exam.num_questions if active_exam else 10,
            "duration_minutes": active_exam.duration_minutes if active_exam else 30,
            "open_time": active_exam.open_time.isoformat() if active_exam and active_exam.open_time else None,
            "close_time": active_exam.close_time.isoformat() if active_exam and active_exam.close_time else None
        }

# ==========================================
# BACKGROUND WORKER FOR EXAM EXPIRATION
# ==========================================
@app.on_event("startup")
async def start_background_session_scanner():
    async def session_scanner():
        # Delay slightly on startup to let DB initialize
        await asyncio.sleep(2)
        while True:
            try:
                db = SessionLocal()
                try:
                    sync_expired_sessions(db)
                finally:
                    db.close()
            except Exception:
                pass
            await asyncio.sleep(20)
            
    asyncio.create_task(session_scanner())

# ==========================================
# SERVE STATIC ASSETS & ROOT
# ==========================================
app.mount("/static", StaticFiles(directory="static"), name="static")

@app.api_route("/", methods=["GET", "HEAD"])
def read_root():
    return FileResponse("static/index.html")

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8000))
    uvicorn.run("main:app", host="0.0.0.0", port=port, reload=False)

