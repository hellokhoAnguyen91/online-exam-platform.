import datetime
import json
from sqlalchemy import (
    Column, Integer, String, Boolean, DateTime, ForeignKey, Text, Float, text, UniqueConstraint
)
from database import Base, engine

class User(Base):

    __tablename__ = "users"
    id = Column(Integer, primary_key=True, index=True)
    username = Column(String, unique=True, index=True)
    password = Column(String)
    is_admin = Column(Boolean, default=False)
    fullname = Column(String)
    dob = Column(String, nullable=True)
    class_name = Column(String, nullable=True)
    order_index = Column(Integer, default=0, nullable=True)

class Exam(Base):
    __tablename__ = "exams"
    id = Column(Integer, primary_key=True, index=True)
    title = Column(String, default="Kỳ thi trắc nghiệm")
    code = Column(String, default="")
    description = Column(Text, nullable=True)
    num_questions = Column(Integer, default=10)
    duration_minutes = Column(Integer, default=30)
    open_time = Column(DateTime, nullable=True)
    close_time = Column(DateTime, nullable=True)
    is_active = Column(Boolean, default=True)
    is_archived = Column(Boolean, default=False)
    archived_at = Column(DateTime, nullable=True)
    allow_review = Column(Boolean, default=True) # Sinh viên được xem lại đáp án sau khi nộp
    shuffle_questions = Column(Boolean, default=True)
    shuffle_options = Column(Boolean, default=False)
    mc_max_score = Column(Float, default=50.0) # Điểm tối đa phần trắc nghiệm / lý thuyết
    essay_max_score = Column(Float, default=30.0) # Điểm tối đa phần tự luận
    created_at = Column(DateTime, default=datetime.datetime.now)

    @property
    def total_max_score(self):
        mc = self.mc_max_score if self.mc_max_score is not None else 50.0
        essay = self.essay_max_score if self.essay_max_score is not None else 30.0
        return round(mc + essay, 2)

class Question(Base):
    __tablename__ = "questions"
    id = Column(Integer, primary_key=True, index=True)
    exam_id = Column(Integer, ForeignKey("exams.id", ondelete="CASCADE"), nullable=True, index=True)
    content = Column(Text)
    option_a = Column(Text)
    option_b = Column(Text)
    option_c = Column(Text)
    option_d = Column(Text)
    option_e = Column(Text, nullable=True)
    option_f = Column(Text, nullable=True)
    correct_option = Column(String) # A, B, C, D
    explanation = Column(Text, nullable=True)
    question_type = Column(String, default="multiple_choice")
    score_weight = Column(Float, default=1.0)
    created_at = Column(DateTime, default=datetime.datetime.now)

class ExamConfig(Base):
    # Kept for backward compatibility
    __tablename__ = "exam_config"
    id = Column(Integer, primary_key=True, index=True)
    num_questions = Column(Integer, default=10)
    duration_minutes = Column(Integer, default=30)
    open_time = Column(DateTime, nullable=True)
    close_time = Column(DateTime, nullable=True)

class ExamResult(Base):
    __tablename__ = "exam_results"
    __table_args__ = (
        UniqueConstraint('user_id', 'exam_id', name='uq_exam_results_user_exam'),
    )
    id = Column(Integer, primary_key=True, index=True)
    exam_id = Column(Integer, ForeignKey("exams.id", ondelete="CASCADE"), nullable=True, index=True)
    user_id = Column(Integer, ForeignKey("users.id", ondelete="CASCADE"), index=True)
    score = Column(Float, default=0.0)
    max_score = Column(Float, default=80.0)
    correct_count = Column(Integer, default=0)
    total_questions = Column(Integer, default=0)
    answers = Column(Text, nullable=True) # JSON: {"<q_id>": "A"}
    answers_detail = Column(Text, nullable=True) # Full audit trail snapshot JSON
    questions = Column(Text, nullable=True) # Shuffled list of q_ids for candidate
    start_time = Column(DateTime, default=datetime.datetime.now)
    submit_time = Column(DateTime, nullable=True)
    duration_seconds = Column(Integer, default=0)
    client_ip = Column(String, nullable=True)
    user_agent = Column(String, nullable=True)
    status = Column(String, default="in_progress") # "in_progress", "submitted", "timed_out"


def migrate_database():
    """Ensure all tables and columns exist safely without data loss."""
    # 1. Create any missing tables
    Base.metadata.create_all(bind=engine)
    
    with engine.connect() as conn:
        # Check questions columns
        res = conn.execute(text("PRAGMA table_info(questions);")).fetchall()
        q_cols = [r[1] for r in res]
        if "exam_id" not in q_cols:
            conn.execute(text("ALTER TABLE questions ADD COLUMN exam_id INTEGER;"))
        if "explanation" not in q_cols:
            conn.execute(text("ALTER TABLE questions ADD COLUMN explanation TEXT;"))
        if "question_type" not in q_cols:
            conn.execute(text("ALTER TABLE questions ADD COLUMN question_type TEXT DEFAULT 'multiple_choice';"))
        if "score_weight" not in q_cols:
            conn.execute(text("ALTER TABLE questions ADD COLUMN score_weight REAL DEFAULT 1.0;"))
        if "created_at" not in q_cols:
            conn.execute(text("ALTER TABLE questions ADD COLUMN created_at DATETIME;"))
            
        # Check exam_results columns
        res = conn.execute(text("PRAGMA table_info(exam_results);")).fetchall()
        er_cols = [r[1] for r in res]
        if "exam_id" not in er_cols:
            conn.execute(text("ALTER TABLE exam_results ADD COLUMN exam_id INTEGER;"))
        if "max_score" not in er_cols:
            conn.execute(text("ALTER TABLE exam_results ADD COLUMN max_score REAL DEFAULT 10.0;"))
        if "correct_count" not in er_cols:
            conn.execute(text("ALTER TABLE exam_results ADD COLUMN correct_count INTEGER DEFAULT 0;"))
        if "total_questions" not in er_cols:
            conn.execute(text("ALTER TABLE exam_results ADD COLUMN total_questions INTEGER DEFAULT 0;"))
        if "answers_detail" not in er_cols:
            conn.execute(text("ALTER TABLE exam_results ADD COLUMN answers_detail TEXT;"))
        if "duration_seconds" not in er_cols:
            conn.execute(text("ALTER TABLE exam_results ADD COLUMN duration_seconds INTEGER DEFAULT 0;"))
        if "client_ip" not in er_cols:
            conn.execute(text("ALTER TABLE exam_results ADD COLUMN client_ip TEXT;"))
        if "user_agent" not in er_cols:
            conn.execute(text("ALTER TABLE exam_results ADD COLUMN user_agent TEXT;"))
        if "status" not in er_cols:
            conn.execute(text("ALTER TABLE exam_results ADD COLUMN status TEXT DEFAULT 'submitted';"))
            
        # Check users columns
        res_u = conn.execute(text("PRAGMA table_info(users);")).fetchall()
        user_cols = [r[1] for r in res_u]
        if "dob" not in user_cols:
            conn.execute(text("ALTER TABLE users ADD COLUMN dob VARCHAR;"))
        if "class_name" not in user_cols:
            conn.execute(text("ALTER TABLE users ADD COLUMN class_name VARCHAR;"))

        # Check exams columns
        res_e = conn.execute(text("PRAGMA table_info(exams);")).fetchall()
        exam_cols = [r[1] for r in res_e]
        if "mc_max_score" not in exam_cols:
            conn.execute(text("ALTER TABLE exams ADD COLUMN mc_max_score REAL DEFAULT 7.0;"))
        if "essay_max_score" not in exam_cols:
            conn.execute(text("ALTER TABLE exams ADD COLUMN essay_max_score REAL DEFAULT 3.0;"))

        # Deduplicate and ensure unique index on exam_results (user_id, exam_id)
        try:
            conn.execute(text("""
                DELETE FROM exam_results 
                WHERE id NOT IN (
                    SELECT MAX(id) FROM exam_results GROUP BY user_id, exam_id
                );
            """))
            conn.execute(text("CREATE UNIQUE INDEX IF NOT EXISTS uq_exam_results_user_exam ON exam_results(user_id, exam_id);"))
        except Exception as e:
            print(f"[migrate_database] Note on uq_exam_results_user_exam: {e}")

        conn.commit()

        # Check if any Exam exists, if not create default
        exams = conn.execute(text("SELECT id, title FROM exams;")).fetchall()
        default_exam_id = None
        if not exams:
            # Check legacy exam_config
            legacy_config = conn.execute(text("SELECT num_questions, duration_minutes, open_time, close_time FROM exam_config LIMIT 1;")).fetchone()
            num_q = 10
            dur = 30
            op_t = None
            cl_t = None
            if legacy_config:
                num_q = legacy_config[0] or 10
                dur = legacy_config[1] or 30
                op_t = legacy_config[2]
                cl_t = legacy_config[3]
                
            conn.execute(
                text("""
                    INSERT INTO exams (title, code, description, num_questions, duration_minutes, open_time, close_time, is_active, is_archived, allow_review, shuffle_questions, shuffle_options, created_at)
                    VALUES (:title, :code, :description, :num_q, :dur, :op_t, :cl_t, 1, 0, 1, 1, 0, :now);
                """),
                {
                    "title": "Kỳ thi Trắc nghiệm Chính thức",
                    "code": "EXAM-001",
                    "description": "Kỳ thi được khởi tạo tự động từ cấu hình hệ thống.",
                    "num_q": num_q,
                    "dur": dur,
                    "op_t": op_t,
                    "cl_t": cl_t,
                    "now": datetime.datetime.now()
                }
            )
            conn.commit()
            new_exam = conn.execute(text("SELECT id FROM exams ORDER BY id DESC LIMIT 1;")).fetchone()
            default_exam_id = new_exam[0]
        else:
            default_exam_id = exams[0][0]
            
        # Migrate any orphan questions and results to default_exam_id
        if default_exam_id:
            conn.execute(text("UPDATE questions SET exam_id = :eid WHERE exam_id IS NULL;"), {"eid": default_exam_id})
            conn.execute(text("UPDATE exam_results SET exam_id = :eid WHERE exam_id IS NULL;"), {"eid": default_exam_id})
            
            # Backfill any existing results that lack correct_count or duration_seconds
            existing_results = conn.execute(text("SELECT id, score, answers, questions, start_time, submit_time FROM exam_results;")).fetchall()
            for r in existing_results:
                r_id, r_score, r_ans, r_qs, r_st, r_sub = r
                try:
                    q_ids = json.loads(r_qs) if r_qs else []
                    tot_q = len(q_ids)
                    c_count = int(r_score) if r_score is not None else 0
                    dur_sec = 0
                    if r_st and r_sub:
                        try:
                            st_dt = datetime.datetime.fromisoformat(r_st)
                            sub_dt = datetime.datetime.fromisoformat(r_sub)
                            dur_sec = int((sub_dt - st_dt).total_seconds())
                        except:
                            pass
                    # Scale score out of 10 if total questions known
                    scaled_score = round((c_count / tot_q * 10.0), 2) if tot_q > 0 else (r_score or 0.0)
                    conn.execute(
                        text("""
                            UPDATE exam_results
                            SET correct_count = :cc, total_questions = :tq, score = :sc, duration_seconds = :dur, status = 'submitted'
                            WHERE id = :rid AND (total_questions = 0 OR total_questions IS NULL);
                        """),
                        {"cc": c_count, "tq": tot_q, "sc": scaled_score, "dur": dur_sec, "rid": r_id}
                    )
                except Exception as ex:
                    print(f"Migration result backfill error for id {r_id}: {ex}")
            conn.commit()
