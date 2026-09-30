import os
import re

with open("main.py", "r", encoding="utf-8") as f:
    content = f.read()

# Replace upload_questions
new_upload = """@app.post("/api/admin/upload_questions")
def upload_questions(file: UploadFile = File(...), db: Session = Depends(get_db), current_user: User = Depends(get_current_user)):
    if not current_user.is_admin:
        raise HTTPException(status_code=403, detail="Not authorized")
    
    import base64
    import re
    doc = docx.Document(file.file)
    db.query(Question).delete() # clear old questions
    
    q_content = ""
    opts = {}
    correct = ""
    
    def process_para(para, doc):
        html = ""
        for run in para.runs:
            text = run.text
            # Convert URL to link
            text = re.sub(r'(https?://\S+)', r'<a href="\\1" target="_blank">\\1</a>', text)
            html += text
            
            drawings = run._r.xpath('.//w:drawing')
            if drawings:
                for blip in run._r.xpath('.//a:blip/@r:embed'):
                    image_part = doc.part.related_parts[blip]
                    img_b64 = base64.b64encode(image_part.blob).decode('utf-8')
                    mime = image_part.content_type
                    html += f'<br><img src="data:{mime};base64,{img_b64}" style="max-width:100%;"><br>'
        return html
        
    for para in doc.paragraphs:
        para_html = process_para(para, doc)
        clean_text = para.text.strip()
        
        if clean_text.startswith("Q:"):
            if q_content and opts:
                db.add(Question(content=q_content, option_a=opts.get('A',''), option_b=opts.get('B',''), option_c=opts.get('C',''), option_d=opts.get('D',''), correct_option=correct))
            
            if para_html.startswith("Q:"):
                q_content = para_html[2:].strip()
            else:
                q_content = para_html
            opts = {}
            correct = ""
        elif clean_text.startswith("A."): opts['A'] = para_html[2:].strip() if para_html.startswith("A.") else para_html
        elif clean_text.startswith("B."): opts['B'] = para_html[2:].strip() if para_html.startswith("B.") else para_html
        elif clean_text.startswith("C."): opts['C'] = para_html[2:].strip() if para_html.startswith("C.") else para_html
        elif clean_text.startswith("D."): opts['D'] = para_html[2:].strip() if para_html.startswith("D.") else para_html
        elif clean_text.startswith("Correct:"): correct = clean_text[8:].strip()
        elif q_content and not clean_text.startswith("A.") and not clean_text.startswith("B.") and not clean_text.startswith("C.") and not clean_text.startswith("D.") and not clean_text.startswith("Correct:"):
            if para_html:
                q_content += "<br>" + para_html
                
    if q_content and opts:
        db.add(Question(content=q_content, option_a=opts.get('A',''), option_b=opts.get('B',''), option_c=opts.get('C',''), option_d=opts.get('D',''), correct_option=correct))
    
    db.commit()
    return {"detail": "Questions uploaded"}
"""

content = re.sub(r'@app\.post\("/api/admin/upload_questions"\).*?(?=@app\.post\("/api/admin/upload_students"\))', new_upload, content, flags=re.DOTALL)


# Replace config
new_config = """@app.post("/api/admin/config")
def set_config(num_questions: int = Form(...), duration_minutes: int = Form(...), open_time: str = Form(None), close_time: str = Form(None), db: Session = Depends(get_db), current_user: User = Depends(get_current_user)):
    if not current_user.is_admin:
        raise HTTPException(status_code=403, detail="Not authorized")
    
    config = db.query(ExamConfig).first()
    
    def parse_dt(dt_str):
        if not dt_str: return None
        try: return datetime.datetime.fromisoformat(dt_str)
        except: return None
        
    if not config:
        config = ExamConfig(num_questions=num_questions, duration_minutes=duration_minutes, open_time=parse_dt(open_time), close_time=parse_dt(close_time))
        db.add(config)
    else:
        config.num_questions = num_questions
        config.duration_minutes = duration_minutes
        config.open_time = parse_dt(open_time)
        config.close_time = parse_dt(close_time)
    db.commit()
    return {"detail": "Config updated"}
"""

content = re.sub(r'@app\.post\("/api/admin/config"\).*?(?=@app\.get\("/api/admin/results"\))', new_config, content, flags=re.DOTALL)

# Modify get_exam
new_get_exam = """@app.get("/api/exam")
def get_exam(db: Session = Depends(get_db), current_user: User = Depends(get_current_user)):
    if current_user.is_admin:
        raise HTTPException(status_code=400, detail="Admin cannot take exam")
        
    config = db.query(ExamConfig).first()
    if not config:
        raise HTTPException(status_code=400, detail="Exam not configured")
        
    now = datetime.datetime.utcnow()
    if config.open_time and now < config.open_time:
        raise HTTPException(status_code=400, detail="Exam has not opened yet")
    if config.close_time and now > config.close_time:
        raise HTTPException(status_code=400, detail="Exam is closed")
        
    result = db.query(ExamResult).filter(ExamResult.user_id == current_user.id).first()
    if result and result.submit_time:
        return {"status": "submitted", "score": result.score}
        
    if not result:
        questions = db.query(Question).all()
        if len(questions) < config.num_questions:
             raise HTTPException(status_code=400, detail="Not enough questions in bank")
        
        selected = random.sample(questions, config.num_questions)
        q_ids = [q.id for q in selected]
        
        result = ExamResult(
            user_id=current_user.id,
            start_time=datetime.datetime.utcnow(),
            questions=json.dumps(q_ids)
        )
        db.add(result)
        db.commit()
        db.refresh(result)
        
    q_ids = json.loads(result.questions)
    questions = db.query(Question).filter(Question.id.in_(q_ids)).all()
    
    q_list = []
    for q in questions:
        q_list.append({
            "id": q.id,
            "content": q.content,
            "option_a": q.option_a,
            "option_b": q.option_b,
            "option_c": q.option_c,
            "option_d": q.option_d,
        })
    
    # calc time left
    elapsed = (datetime.datetime.utcnow() - result.start_time).total_seconds()
    time_left = max(0, config.duration_minutes * 60 - elapsed)
    if time_left == 0:
        # Auto submit if expired
        pass # The frontend will submit, or they can submit later with 0 score/current score
    
    return {
        "status": "ongoing",
        "questions": q_list,
        "time_left": time_left,
        "answers": json.loads(result.answers) if result.answers else {}
    }
"""
content = re.sub(r'@app\.get\("/api/exam"\).*?(?=class AnswerPayload)', new_get_exam, content, flags=re.DOTALL)

with open("main.py", "w", encoding="utf-8") as f:
    f.write(content)
