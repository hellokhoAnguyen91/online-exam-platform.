import os
import re
import io
import json
import html
import base64
import zipfile
import datetime
from typing import List, Dict, Any, Optional, Tuple

from PIL import Image as PILImage
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, 
    Image as RLImage, KeepTogether, HRFlowable
)
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont


# --- Font Registration ---
FONTS_REGISTERED = False

def ensure_fonts():
    global FONTS_REGISTERED
    if FONTS_REGISTERED:
        return
    
    font_dir = "C:/Windows/Fonts"
    arial = os.path.join(font_dir, "arial.ttf")
    arial_bd = os.path.join(font_dir, "arialbd.ttf")
    arial_it = os.path.join(font_dir, "ariali.ttf")
    arial_bi = os.path.join(font_dir, "arialbi.ttf")

    if os.path.exists(arial):
        pdfmetrics.registerFont(TTFont("Arial", arial))
    if os.path.exists(arial_bd):
        pdfmetrics.registerFont(TTFont("Arial-Bold", arial_bd))
    if os.path.exists(arial_it):
        pdfmetrics.registerFont(TTFont("Arial-Italic", arial_it))
    if os.path.exists(arial_bi):
        pdfmetrics.registerFont(TTFont("Arial-BoldItalic", arial_bi))

    FONTS_REGISTERED = True


# --- Numbered Canvas for Page Numbers & Header/Footer ---
class NumberedCanvas(canvas.Canvas):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._saved_page_states = []

    def showPage(self):
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        num_pages = len(self._saved_page_states)
        for state in self._saved_page_states:
            self.__dict__.update(state)
            self.draw_page_decorations(num_pages)
            super().showPage()
        super().save()

    def draw_page_decorations(self, page_count):
        self.saveState()
        self.setFont("Arial", 8)
        self.setFillColor(colors.HexColor("#64748b"))
        self.setStrokeColor(colors.HexColor("#e2e8f0"))
        self.setLineWidth(0.6)

        # Footer divider line & text
        self.line(36, 32, 559, 32)
        self.drawString(36, 20, "Hệ thống Thi Trực Tuyến - Phiếu bài thi thí sinh (Lưu trữ đào tạo)")
        page_text = f"Trang {self._pageNumber} / {page_count}"
        self.drawRightString(559, 20, page_text)
        self.restoreState()


# --- HTML Sanitation for ReportLab XML ---
def prepare_rl_text(text: Optional[str]) -> Tuple[str, List[RLImage]]:
    """Clean text containing HTML and extract any embedded base64 images."""
    if not text:
        return "", []

    images = []

    # 1. Extract embedded base64 images
    img_pattern = re.compile(r'<img[^>]+src=[\'"]([^\'"]+)[\'"][^>]*>', re.IGNORECASE)
    for m in img_pattern.finditer(text):
        src = m.group(1)
        if "base64," in src:
            try:
                b64_part = src.split("base64,")[1]
                img_data = base64.b64decode(b64_part)
                pil_img = PILImage.open(io.BytesIO(img_data))
                orig_w, orig_h = pil_img.size
                
                # Max width in points for A4 (margin 36 pt left & right = 595 - 72 = 523 max)
                max_w = 460.0
                if orig_w > max_w:
                    scaled_h = orig_h * (max_w / float(orig_w))
                    scaled_w = max_w
                else:
                    scaled_w = float(orig_w)
                    scaled_h = float(orig_h)

                # Cap height if overly tall
                if scaled_h > 350.0:
                    scaled_w = scaled_w * (350.0 / scaled_h)
                    scaled_h = 350.0

                rl_img = RLImage(io.BytesIO(img_data), width=scaled_w, height=scaled_h)
                images.append(rl_img)
            except Exception as e:
                print(f"[PDF] Warning: could not parse image: {e}")

    # 2. Strip <img> tags from text
    text = img_pattern.sub("", text)

    # 3. Preserve breaks
    text = re.sub(r'<(?:br|p|div)[^>]*>', '__BR__', text, flags=re.IGNORECASE)
    text = re.sub(r'</(?:p|div)>', '__BR__', text, flags=re.IGNORECASE)

    # 4. Map formatting tags to placeholders
    text = re.sub(r'<strong[^>]*>', '__B_OPEN__', text, flags=re.IGNORECASE)
    text = re.sub(r'</strong>', '__B_CLOSE__', text, flags=re.IGNORECASE)
    text = re.sub(r'<b[^>]*>', '__B_OPEN__', text, flags=re.IGNORECASE)
    text = re.sub(r'</b>', '__B_CLOSE__', text, flags=re.IGNORECASE)

    text = re.sub(r'<em[^>]*>', '__I_OPEN__', text, flags=re.IGNORECASE)
    text = re.sub(r'</em>', '__I_CLOSE__', text, flags=re.IGNORECASE)
    text = re.sub(r'<i[^>]*>', '__I_OPEN__', text, flags=re.IGNORECASE)
    text = re.sub(r'</i>', '__I_CLOSE__', text, flags=re.IGNORECASE)

    text = re.sub(r'<u[^>]*>', '__U_OPEN__', text, flags=re.IGNORECASE)
    text = re.sub(r'</u>', '__U_CLOSE__', text, flags=re.IGNORECASE)

    text = re.sub(r'<sub[^>]*>', '__SUB_OPEN__', text, flags=re.IGNORECASE)
    text = re.sub(r'</sub>', '__SUB_CLOSE__', text, flags=re.IGNORECASE)

    text = re.sub(r'<sup[^>]*>', '__SUP_OPEN__', text, flags=re.IGNORECASE)
    text = re.sub(r'</sup>', '__SUP_CLOSE__', text, flags=re.IGNORECASE)

    # Strip remaining HTML tags
    text = re.sub(r'<[^>]+>', ' ', text)

    # Unescape any preexisting entities
    text = html.unescape(text)

    # Re-escape special XML characters
    text = html.escape(text)

    # Restore allowed tags
    text = text.replace('__B_OPEN__', '<b>').replace('__B_CLOSE__', '</b>')
    text = text.replace('__I_OPEN__', '<i>').replace('__I_CLOSE__', '</i>')
    text = text.replace('__U_OPEN__', '<u>').replace('__U_CLOSE__', '</u>')
    text = text.replace('__SUB_OPEN__', '<sub>').replace('__SUB_CLOSE__', '</sub>')
    text = text.replace('__SUP_OPEN__', '<super>').replace('__SUP_CLOSE__', '</super>')
    text = text.replace('__BR__', '<br/>')

    # Clean redundant spaces
    text = re.sub(r'\s+', ' ', text)
    text = re.sub(r'(<br/>\s*)+', '<br/>', text)
    text = balance_xml_tags(text)
    return text.strip(), images


def balance_xml_tags(s: str) -> str:
    """Ensure all <b>, <i>, <u>, <sub>, <super> tags are properly matched and closed."""
    if not s:
        return ""
    for tag in ['b', 'i', 'u', 'sub', 'super']:
        open_tag = f'<{tag}>'
        close_tag = f'</{tag}>'
        opens = s.count(open_tag)
        closes = s.count(close_tag)
        if opens > closes:
            s += close_tag * (opens - closes)
        elif closes > opens:
            excess = closes - opens
            # Remove excess close_tags from the beginning
            for _ in range(excess):
                pos = s.find(close_tag)
                if pos != -1:
                    s = s[:pos] + s[pos + len(close_tag):]
    return s


def safe_str(val: Any) -> str:
    if val is None:
        return ""
    return str(val).strip()


def sanitize_filename(name: str) -> str:
    """Make filename safe for ZIP entry across Windows, Linux, and Mac."""
    # Transliterate basic Vietnamese diacritics for clean filename
    v_map = {
        'à':'a','á':'a','ả':'a','ã':'a','ạ':'a','ă':'a','ằ':'a','ắ':'a','ẳ':'a','ẵ':'a','ặ':'a','â':'a','ầ':'a','ấ':'a','ẩ':'a','ẫ':'a','ậ':'a',
        'đ':'d',
        'è':'e','é':'e','ẻ':'e','ẽ':'e','ẹ':'e','ê':'e','ề':'e','ế':'e','ể':'e','ễ':'e','ệ':'e',
        'ì':'i','í':'i','ỉ':'i','ĩ':'i','ị':'i',
        'ò':'o','ó':'o','ỏ':'o','õ':'o','ọ':'o','ô':'o','ồ':'o','ố':'o','ổ':'o','ỗ':'o','ộ':'o','ơ':'o','ờ':'o','ớ':'o','ở':'o','ỡ':'o','ợ':'o',
        'ù':'u','ú':'u','ủ':'u','ũ':'u','ụ':'u','ư':'u','ừ':'u','ứ':'u','ử':'u','ữ':'u','ự':'u',
        'ỳ':'y','ý':'y','ỷ':'y','ỹ':'y','ỵ':'y',
        'À':'A','Á':'A','Ả':'A','Ã':'A','Ạ':'A','Ă':'A','Ằ':'A','Ắ':'A','Ẳ':'A','Ẵ':'A','Ặ':'A','Â':'A','Ầ':'A','Ấ':'A','Ẩ':'A','Ẫ':'A','Ậ':'A',
        'Đ':'D',
        'È':'E','É':'E','Ẻ':'E','Ẽ':'E','Ẹ':'E','Ê':'E','Ề':'E','Ế':'E','Ể':'E','Ễ':'E','Ệ':'E',
        'Ì':'I','Í':'I','Ỉ':'I','Ĩ':'I','Ị':'I',
        'Ò':'O','Ó':'O','Ỏ':'O','Õ':'O','Ọ':'O','Ô':'O','Ồ':'O','Ố':'O','Ổ':'O','Ỗ':'O','Ộ':'O','Ơ':'O','Ờ':'O','Ớ':'O','Ở':'O','Ỡ':'O','Ợ':'O',
        'Ù':'U','Ú':'U','Ủ':'U','Ũ':'U','Ụ':'U','Ư':'U','Ừ':'U','Ứ':'U','Ử':'U','Ữ':'U','Ự':'U',
        'Ỳ':'Y','Ý':'Y','Ỷ':'Y','Ỹ':'Y','Ỵ':'Y'
    }
    res = "".join(v_map.get(c, c) for c in name)
    res = re.sub(r'[^a-zA-Z0-9_-]', '_', res)
    res = re.sub(r'_+', '_', res)
    return res.strip('_')


# --- Core Single Candidate PDF Generator ---
def generate_candidate_pdf(exam_info: Dict[str, Any], candidate_info: Dict[str, Any], audit_questions: List[Dict[str, Any]]) -> bytes:
    """Generate a high-grade, official-style PDF of a candidate's complete exam."""
    ensure_fonts()

    buffer = io.BytesIO()
    # A4: 595.27 x 841.89 points. Margins: 36 pt (0.5 inch)
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        leftMargin=36,
        rightMargin=36,
        topMargin=36,
        bottomMargin=42
    )

    styles = getSampleStyleSheet()

    # Custom styles
    style_univ = ParagraphStyle('Univ', parent=styles['Normal'], fontName='Arial-Bold', fontSize=10, leading=13, textColor=colors.HexColor('#1e3a8a'))
    style_sub = ParagraphStyle('Sub', parent=styles['Normal'], fontName='Arial', fontSize=8.5, leading=11, textColor=colors.HexColor('#475569'))
    style_title = ParagraphStyle('ExamTitle', parent=styles['Normal'], fontName='Arial-Bold', fontSize=13, leading=17, alignment=1, textColor=colors.HexColor('#0f172a'))
    style_subtitle = ParagraphStyle('ExamSub', parent=styles['Normal'], fontName='Arial-Italic', fontSize=9, leading=12, alignment=1, textColor=colors.HexColor('#334155'))

    style_info_lbl = ParagraphStyle('InfoLbl', parent=styles['Normal'], fontName='Arial', fontSize=9, leading=12, textColor=colors.HexColor('#475569'))
    style_info_val = ParagraphStyle('InfoVal', parent=styles['Normal'], fontName='Arial-Bold', fontSize=9, leading=12, textColor=colors.HexColor('#0f172a'))

    style_score_val = ParagraphStyle('ScoreVal', parent=styles['Normal'], fontName='Arial-Bold', fontSize=18, leading=20, alignment=1, textColor=colors.HexColor('#1d4ed8'))
    style_score_lbl = ParagraphStyle('ScoreLbl', parent=styles['Normal'], fontName='Arial', fontSize=8, leading=10, alignment=1, textColor=colors.HexColor('#64748b'))

    style_sec_header = ParagraphStyle('SecH', parent=styles['Normal'], fontName='Arial-Bold', fontSize=10.5, leading=14, textColor=colors.HexColor('#1e3a8a'))

    style_q_num = ParagraphStyle('QNum', parent=styles['Normal'], fontName='Arial-Bold', fontSize=9.5, leading=13, textColor=colors.HexColor('#1e40af'))
    style_q_text = ParagraphStyle('QText', parent=styles['Normal'], fontName='Arial', fontSize=9.5, leading=13.5, textColor=colors.HexColor('#0f172a'))

    style_opt_normal = ParagraphStyle('OptNorm', parent=styles['Normal'], fontName='Arial', fontSize=9, leading=12.5, textColor=colors.HexColor('#334155'))
    style_opt_correct = ParagraphStyle('OptCorr', parent=styles['Normal'], fontName='Arial-Bold', fontSize=9, leading=12.5, textColor=colors.HexColor('#15803d'))
    style_opt_wrong = ParagraphStyle('OptWrong', parent=styles['Normal'], fontName='Arial-Bold', fontSize=9, leading=12.5, textColor=colors.HexColor('#b91c1c'))

    style_ans_summary = ParagraphStyle('AnsSumm', parent=styles['Normal'], fontName='Arial-Bold', fontSize=8.5, leading=11)

    story = []

    # 1. Header Table (Left: Institution, Right: Country/Exam metadata)
    left_cell = [
        Paragraph("<b>TRƯỜNG ĐẠI HỌC CÔNG NGHỆ KỸ THUẬT TP. HỒ CHÍ MINH (HCMUTE)</b>", style_univ),
        Paragraph("HỆ THỐNG THI TRẮC NGHIỆM TRỰC TUYẾN", style_sub),
        Paragraph(f"Khoa / Bộ môn Đào tạo phụ trách", style_sub)
    ]

    exam_date_str = candidate_info.get("start_time_str") or datetime.datetime.now().strftime("%d/%m/%Y")
    right_cell = [
        Paragraph("<b>CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM</b>", ParagraphStyle('R1', parent=style_univ, alignment=1, fontSize=9)),
        Paragraph("<b>Độc lập - Tự do - Hạnh phúc</b>", ParagraphStyle('R2', parent=style_sub, alignment=1)),
        Spacer(1, 2),
        Paragraph(f"<i>Ngày thi: {exam_date_str}</i>", ParagraphStyle('R3', parent=style_sub, alignment=1))
    ]

    header_table = Table([[left_cell, right_cell]], colWidths=[310, 213])
    header_table.setStyle(TableStyle([
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('BOTTOMPADDING', (0,0), (-1,-1), 0),
        ('TOPPADDING', (0,0), (-1,-1), 0),
    ]))
    story.append(header_table)
    story.append(Spacer(1, 8))
    story.append(HRFlowable(width="100%", thickness=1, color=colors.HexColor('#cbd5e1'), spaceBefore=2, spaceAfter=8))

    # 2. Main Title
    story.append(Paragraph(f"PHIẾU KẾT QUẢ BÀI THI TRỰC TUYẾN", style_title))
    exam_title = safe_str(exam_info.get("title", "Kỳ thi trắc nghiệm"))
    exam_code = safe_str(exam_info.get("code", ""))
    sub_title_text = f"Kỳ thi: <b>{exam_title}</b>"
    if exam_code:
        sub_title_text += f" | Mã môn/Mã ca: <b>{exam_code}</b>"
    story.append(Paragraph(sub_title_text, style_subtitle))
    story.append(Spacer(1, 10))

    # 3. Candidate & Score Summary Card
    fullname = safe_str(candidate_info.get("fullname", "N/A"))
    username = safe_str(candidate_info.get("username", "N/A"))
    dob = safe_str(candidate_info.get("dob", "-"))
    score = candidate_info.get("score")
    max_score = candidate_info.get("max_score") or 10.0
    correct_cnt = candidate_info.get("correct_count", 0)
    total_q = candidate_info.get("total_questions", len(audit_questions))
    duration_str = safe_str(candidate_info.get("duration_str", "-"))
    submit_str = safe_str(candidate_info.get("submit_time_str", "-"))
    status_str = safe_str(candidate_info.get("status_text", "Đã nộp bài"))

    # Col 1: Identity info
    col1_content = [
        Paragraph(f"Họ và tên thí sinh: <b>{fullname}</b>", style_info_val),
        Spacer(1, 2),
        Paragraph(f"Mã số sinh viên (MSSV): <b>{username}</b>", style_info_lbl),
        Spacer(1, 2),
        Paragraph(f"Ngày sinh: <b>{dob}</b>", style_info_lbl),
    ]

    # Col 2: Exam log info
    col2_content = [
        Paragraph(f"Thời gian bắt đầu: <b>{safe_str(candidate_info.get('start_time_str', '-'))}</b>", style_info_lbl),
        Spacer(1, 2),
        Paragraph(f"Thời gian nộp bài: <b>{submit_str}</b>", style_info_lbl),
        Spacer(1, 2),
        Paragraph(f"Thời lượng làm bài: <b>{duration_str}</b>", style_info_lbl),
        Spacer(1, 2),
        Paragraph(f"Trạng thái: <b>{status_str}</b>", style_info_lbl),
    ]

    # Partition questions into Part I (Trắc nghiệm) and Part II (Tự luận)
    mc_questions = [q for q in audit_questions if q.get("question_type") != "essay"]
    essay_questions = [q for q in audit_questions if q.get("question_type") == "essay"]

    mc_weight = sum((q.get("score_weight") or 1.0) for q in mc_questions)
    mc_earned = sum((q.get("earned_score") or 0.0) for q in mc_questions)
    mc_correct = sum(1 for q in mc_questions if q.get("is_correct") or q.get("status") == "correct")

    essay_weight = sum((q.get("score_weight") or 1.0) for q in essay_questions)
    essay_scores = [q.get("essay_score") for q in essay_questions if "essay_score" in q and q.get("essay_score") is not None]
    essay_graded = (len(essay_scores) == len(essay_questions)) if essay_questions else True
    essay_earned = sum((s or 0.0) for s in essay_scores) if essay_scores else 0.0

    # Col 3: Score Box
    if score is None:
        score_val_str = "<font size=12 color='#d97706'><b>Chờ chấm</b></font>"
    else:
        score_val_str = f"{score:.2f} <font size=9 color='#64748b'>/ {max_score:.1f}</font>"

    if mc_questions and essay_questions:
        tl_str = f"{essay_earned:.2f}/{essay_weight:.1f}đ" if essay_graded else "<font color='#d97706'>Chờ chấm</font>"
        correct_lbl_str = f"TN: <b>{mc_correct}/{len(mc_questions)}</b> ({mc_earned:.1f}đ)<br/>TL: <b>{tl_str}</b>"
    elif essay_questions:
        tl_str = f"{essay_earned:.2f}/{essay_weight:.1f}đ" if essay_graded else "<font color='#d97706'>Chờ chấm</font>"
        correct_lbl_str = f"Tự luận: <b>{tl_str}</b>"
    else:
        correct_lbl_str = f"Đúng: <b>{correct_cnt} / {total_q}</b> câu"

    score_content = [
        Paragraph("ĐIỂM SỐ", style_score_lbl),
        Spacer(1, 2),
        Paragraph(score_val_str, style_score_val),
        Spacer(1, 2),
        Paragraph(correct_lbl_str, style_score_lbl),
    ]

    info_card_data = [[col1_content, col2_content, score_content]]
    info_table = Table(info_card_data, colWidths=[200, 203, 120])
    info_table.setStyle(TableStyle([
        ('BOX', (0,0), (-1,-1), 1, colors.HexColor('#93c5fd')),
        ('BACKGROUND', (0,0), (1,0), colors.HexColor('#f8fafc')),
        ('BACKGROUND', (2,0), (2,0), colors.HexColor('#eff6ff')),
        ('LINEBEFORE', (2,0), (2,0), 1, colors.HexColor('#bfdbfe')),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), 8),
        ('BOTTOMPADDING', (0,0), (-1,-1), 8),
        ('LEFTPADDING', (0,0), (-1,-1), 10),
        ('RIGHTPADDING', (0,0), (-1,-1), 10),
    ]))
    story.append(info_table)
    story.append(Spacer(1, 14))

    # Helper function to render a single question block
    def render_question_block(q):
        q_flowables = []

        q_idx = q.get("q_idx", 1)
        raw_content = q.get("content", "")
        q_text_clean, q_images = prepare_rl_text(raw_content)
        q_text_clean = re.sub(r'^\s*(?:<p[^>]*>)?\s*(?:<(?:strong|b|span|em|i)[^>]*>)?\s*(?:(?:câu|cau|question|q|bài|bai)\s*\d+|\d+)(?:\s*\([^)]*(?:điểm|diem|pt|point)[^)]*\))?[\s.:)\-–—]*(?:<\/(?:strong|b|span|em|i)>\s*)*[\s.:)\-–—]*', '', q_text_clean, flags=re.IGNORECASE)

        # Question Header & Content
        if q.get("question_type") == "essay":
            q_label = f"<b>Câu {q_idx} [Tự luận]: </b>"
        else:
            q_label = f"<b>Câu {q_idx}: </b>"
        q_flowables.append(Paragraph(f"{q_label}{q_text_clean}", style_q_text))

        # Render Question Images if present
        for q_img in q_images:
            q_flowables.append(Spacer(1, 4))
            q_flowables.append(q_img)
            q_flowables.append(Spacer(1, 4))

        q_flowables.append(Spacer(1, 4))

        selected_raw = q.get("selected")
        correct_raw = q.get("correct")
        q_type = q.get("question_type", "multiple_choice")
        is_correct = q.get("is_correct", False)
        q_status = q.get("status", "")

        # Normalize selected & correct for matching options
        if isinstance(selected_raw, list):
            selected_keys = {str(k).strip().upper() for k in selected_raw}
            selected_disp = ", ".join(sorted(list(selected_keys)))
        elif isinstance(selected_raw, str):
            if q_type == "multi_select" or "," in selected_raw:
                selected_keys = set(re.findall(r'[A-Fa-f]', selected_raw.upper()))
                selected_disp = ", ".join(sorted(list(selected_keys))) if selected_keys else selected_raw.strip()
            else:
                selected_keys = {selected_raw.strip().upper()} if selected_raw.strip() else set()
                selected_disp = selected_raw.strip()
        else:
            selected_keys = {str(selected_raw).strip().upper()} if selected_raw is not None else set()
            selected_disp = str(selected_raw) if selected_raw is not None else ""

        if isinstance(correct_raw, list):
            correct_keys = {str(k).strip().upper() for k in correct_raw}
            correct_disp = ", ".join(sorted(list(correct_keys)))
        elif isinstance(correct_raw, str):
            if q_type == "multi_select" or "," in correct_raw:
                correct_keys = set(re.findall(r'[A-Fa-f]', correct_raw.upper()))
                correct_disp = ", ".join(sorted(list(correct_keys))) if correct_keys else correct_raw.strip()
            else:
                correct_keys = {correct_raw.strip().upper()} if correct_raw.strip() else set()
                correct_disp = correct_raw.strip()
        else:
            correct_keys = {str(correct_raw).strip().upper()} if correct_raw is not None else set()
            correct_disp = str(correct_raw) if correct_raw is not None else ""

        options = [
            ("A", q.get("option_a", "")),
            ("B", q.get("option_b", "")),
            ("C", q.get("option_c", "")),
            ("D", q.get("option_d", "")),
            ("E", q.get("option_e", "")),
            ("F", q.get("option_f", ""))
        ]

        opt_rows = []
        for opt_key, opt_val in options:
            if not opt_val and opt_val != 0:
                continue
            opt_clean, opt_imgs = prepare_rl_text(str(opt_val))
            opt_clean = re.sub(r'^(?:<b>)?\s*[A-Fa-f][\.:\)]\s*(?:</b>)?\s*', '', opt_clean, flags=re.IGNORECASE)

            is_selected = (opt_key in selected_keys)
            is_this_correct = (opt_key in correct_keys)

            if is_selected and is_this_correct:
                prefix = f"<b>[X] {opt_key}.</b> "
                suffix = " <font color='#16a34a' size=8><b>(Lựa chọn của thí sinh - ĐÚNG)</b></font>"
                style_use = style_opt_correct
            elif is_selected and not is_this_correct:
                prefix = f"<b>[X] {opt_key}.</b> "
                suffix = " <font color='#dc2626' size=8><b>(Lựa chọn của thí sinh - SAI)</b></font>"
                style_use = style_opt_wrong
            elif is_this_correct:
                prefix = f"<b>[&nbsp;&nbsp;] {opt_key}.</b> "
                suffix = " <font color='#15803d' size=8><b>(Đáp án đúng)</b></font>"
                style_use = style_opt_correct
            else:
                prefix = f"<b>[&nbsp;&nbsp;] {opt_key}.</b> "
                suffix = ""
                style_use = style_opt_normal

            row_text = balance_xml_tags(f"{prefix}{opt_clean}{suffix}")
            row_p = Paragraph(row_text, style_use)

            cell_items = [row_p]
            for oi in opt_imgs:
                cell_items.append(oi)
            opt_rows.append(cell_items)

        if opt_rows:
            opt_table = Table([[item] for item in opt_rows], colWidths=[510])
            opt_table.setStyle(TableStyle([
                ('VALIGN', (0,0), (-1,-1), 'TOP'),
                ('TOPPADDING', (0,0), (-1,-1), 1.5),
                ('BOTTOMPADDING', (0,0), (-1,-1), 1.5),
                ('LEFTPADDING', (0,0), (-1,-1), 12),
                ('RIGHTPADDING', (0,0), (-1,-1), 0),
            ]))
            q_flowables.append(opt_table)
        elif q_type == "essay":
            essay_ans = str(selected_raw or "").strip()
            if not essay_ans:
                ans_items = [Paragraph("<i>(Thí sinh không làm bài / để trống)</i>", style_opt_normal)]
            else:
                essay_clean, essay_imgs = prepare_rl_text(essay_ans)
                ans_items = [Paragraph(f"<b>Bài làm của thí sinh:</b><br/>{essay_clean}", style_opt_normal)]
                for ei in essay_imgs:
                    ans_items.append(ei)
            
            ans_box = Table([[ans_items]], colWidths=[510])
            ans_box.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#f8fafc')),
                ('BOX', (0,0), (-1,-1), 0.8, colors.HexColor('#cbd5e1')),
                ('VALIGN', (0,0), (-1,-1), 'TOP'),
                ('TOPPADDING', (0,0), (-1,-1), 5),
                ('BOTTOMPADDING', (0,0), (-1,-1), 5),
                ('LEFTPADDING', (0,0), (-1,-1), 10),
                ('RIGHTPADDING', (0,0), (-1,-1), 10),
            ]))
            q_flowables.append(ans_box)
            q_flowables.append(Spacer(1, 5))

            # Teacher grading box (ô chấm điểm của giảng viên)
            essay_sc = q.get("essay_score")
            weight = q.get("score_weight", 1.0)
            if essay_sc is not None:
                score_disp = f"<font color='#16a34a'><b>{essay_sc:.2f} / {weight:.2f} điểm</b></font>"
            elif q_status == "unanswered" or not essay_ans:
                score_disp = f"<font color='#dc2626'><b>0.00 / {weight:.2f} điểm (Chưa làm bài)</b></font>"
            else:
                score_disp = f"<font color='#d97706'><b>..... / {weight:.2f} điểm (Chờ chấm)</b></font>"

            teacher_box_data = [
                [
                    Paragraph("<b>Ô CHẤM ĐIỂM VÀ NHẬN XÉT CỦA GIẢNG VIÊN</b>", ParagraphStyle('TBHead', parent=style_sub, fontName='Arial-Bold', fontSize=8)),
                    Paragraph(f"Điểm số: {score_disp}", ParagraphStyle('TBScore', parent=style_sub, fontSize=8, alignment=2))
                ],
                [
                    Paragraph("<b>Nhận xét:</b> ....................................................................................................................................................................................................<br/>........................................................................................................................................................................................................................................", ParagraphStyle('TBNote', parent=style_sub, fontSize=7.5)),
                    Paragraph("<b>Chữ ký CB chấm:</b><br/><br/>....................................", ParagraphStyle('TBSign', parent=style_sub, fontSize=7.5, alignment=1))
                ]
            ]
            teacher_table = Table(teacher_box_data, colWidths=[370, 140])
            teacher_table.setStyle(TableStyle([
                ('BOX', (0,0), (-1,-1), 0.8, colors.HexColor('#f59e0b')),
                ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#fef3c7')),
                ('BACKGROUND', (0,1), (-1,1), colors.HexColor('#fffdf7')),
                ('VALIGN', (0,0), (-1,-1), 'TOP'),
                ('TOPPADDING', (0,0), (-1,-1), 4),
                ('BOTTOMPADDING', (0,0), (-1,-1), 4),
                ('LEFTPADDING', (0,0), (-1,-1), 8),
                ('RIGHTPADDING', (0,0), (-1,-1), 8),
                ('LINEBELOW', (0,0), (-1,0), 0.5, colors.HexColor('#fde68a')),
                ('LINEBEFORE', (1,0), (1,-1), 0.5, colors.HexColor('#fde68a')),
            ]))
            q_flowables.append(teacher_table)

        # Question Verdict Summary Line
        q_flowables.append(Spacer(1, 3))
        earned_sc = q.get("earned_score", 0.0)
        weight = q.get("score_weight", 1.0)
        if q_type == "essay":
            essay_sc = q.get("essay_score")
            if essay_sc is not None:
                summ_html = f"<font color='#16a34a'><b>=> Điểm tự luận: {essay_sc:.2f} / {weight:.2f} điểm</b></font>"
            elif q_status == "unanswered" or not str(selected_raw or "").strip():
                summ_html = f"<font color='#dc2626'><b>=> Điểm tự luận: 0.0 / {weight:.2f} điểm (Chưa trả lời)</b></font>"
            else:
                summ_html = f"<font color='#d97706'><b>=> Điểm tự luận: Chờ chấm (Thang điểm: {weight:.2f})</b></font>"
        elif q_type == "multi_select":
            if q_status == "correct" or is_correct:
                summ_html = f"<font color='#16a34a'><b>=> Kết quả: ĐÚNG (+{earned_sc:.2f}đ)</b> (Thí sinh chọn: [{selected_disp}])</font>"
            elif q_status == "partial":
                summ_html = f"<font color='#d97706'><b>=> Kết quả: ĐÚNG MỘT PHẦN (+{earned_sc:.2f}đ)</b> (Thí sinh chọn: [{selected_disp}] | Đáp án đúng: <b>[{correct_disp}]</b>)</font>"
            elif selected_disp:
                summ_html = f"<font color='#dc2626'><b>=> Kết quả: SAI (0.0đ)</b> (Thí sinh chọn: [{selected_disp}] | Đáp án đúng: <b>[{correct_disp}]</b>)</font>"
            else:
                summ_html = f"<font color='#d97706'><b>=> Kết quả: CHƯA TRẢ LỜI</b> (Đáp án đúng: <b>[{correct_disp}]</b>)</font>"
        else:
            if is_correct:
                summ_html = f"<font color='#16a34a'><b>=> Kết quả: ĐÚNG</b> (Thí sinh chọn: [{selected_disp}])</font>"
            elif selected_disp:
                summ_html = f"<font color='#dc2626'><b>=> Kết quả: SAI</b> (Thí sinh chọn: [{selected_disp}] | Đáp án đúng: <b>[{correct_disp}]</b>)</font>"
            else:
                summ_html = f"<font color='#d97706'><b>=> Kết quả: CHƯA TRẢ LỜI</b> (Đáp án đúng: <b>[{correct_disp}]</b>)</font>"
        
        q_flowables.append(Paragraph(summ_html, style_ans_summary))
        q_flowables.append(Spacer(1, 4))
        q_flowables.append(HRFlowable(width="100%", thickness=0.4, color=colors.HexColor('#e2e8f0'), spaceBefore=2, spaceAfter=6))

        return KeepTogether(q_flowables)

    # 4. Section: Detailed Questions Breakdown
    story.append(Paragraph("<b>CHI TIẾT BÀI LÀM CỦA THÍ SINH</b>", style_sec_header))
    story.append(HRFlowable(width="100%", thickness=0.8, color=colors.HexColor('#94a3b8'), spaceBefore=3, spaceAfter=8))

    if not audit_questions:
        story.append(Paragraph("<i>Chưa có dữ liệu bài làm chi tiết cho thí sinh này.</i>", style_sub))
    else:
        has_both_sections = bool(mc_questions and essay_questions)

        # Render Part I: Trắc nghiệm
        if mc_questions:
            mc_title_text = "<b>PHẦN I: CÂU HỎI TRẮC NGHIỆM</b>" if has_both_sections else "<b>CÂU HỎI TRẮC NGHIỆM</b>"
            sec_title1 = Paragraph(mc_title_text, ParagraphStyle('SecT1', parent=style_sec_header, textColor=colors.HexColor('#1e40af'), fontSize=10.5))
            sec_info1 = Paragraph(f"<i>(Tổng: {len(mc_questions)} câu | Điểm trắc nghiệm: {mc_earned:.2f} / {mc_weight:.2f} điểm | Đúng: {mc_correct}/{len(mc_questions)} câu)</i>", ParagraphStyle('SecI1', parent=style_sub, fontSize=8, alignment=2))
            banner_table1 = Table([[sec_title1, sec_info1]], colWidths=[260, 253])
            banner_table1.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#eff6ff')),
                ('BOX', (0,0), (-1,-1), 1, colors.HexColor('#bfdbfe')),
                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                ('TOPPADDING', (0,0), (-1,-1), 5),
                ('BOTTOMPADDING', (0,0), (-1,-1), 5),
                ('LEFTPADDING', (0,0), (-1,-1), 8),
                ('RIGHTPADDING', (0,0), (-1,-1), 8),
            ]))
            story.append(banner_table1)
            story.append(Spacer(1, 6))

            for q in mc_questions:
                story.append(render_question_block(q))

        # Render Part II: Tự luận
        if essay_questions:
            if mc_questions:
                story.append(Spacer(1, 10))
            essay_title_text = "<b>PHẦN II: CÂU HỎI TỰ LUẬN</b>" if has_both_sections else "<b>CÂU HỎI TỰ LUẬN</b>"
            sec_title2 = Paragraph(essay_title_text, ParagraphStyle('SecT2', parent=style_sec_header, textColor=colors.HexColor('#92400e'), fontSize=10.5))
            tl_status = f"{essay_earned:.2f} / {essay_weight:.2f} điểm" if essay_graded else "Chờ chấm điểm"
            sec_info2 = Paragraph(f"<i>(Tổng: {len(essay_questions)} câu | Điểm tự luận: {tl_status})</i>", ParagraphStyle('SecI2', parent=style_sub, fontSize=8, alignment=2))
            banner_table2 = Table([[sec_title2, sec_info2]], colWidths=[260, 253])
            banner_table2.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#fffbeb')),
                ('BOX', (0,0), (-1,-1), 1, colors.HexColor('#fde68a')),
                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                ('TOPPADDING', (0,0), (-1,-1), 5),
                ('BOTTOMPADDING', (0,0), (-1,-1), 5),
                ('LEFTPADDING', (0,0), (-1,-1), 8),
                ('RIGHTPADDING', (0,0), (-1,-1), 8),
            ]))
            story.append(banner_table2)
            story.append(Spacer(1, 6))

            for q in essay_questions:
                story.append(render_question_block(q))

    # 5. Sign-off Footer block
    story.append(Spacer(1, 15))
    sign_left = [
        Paragraph("<b>CÁN BỘ COI THI / GIÁM THỊ</b>", ParagraphStyle('SignL', parent=style_sub, alignment=1, fontName='Arial-Bold')),
        Paragraph("<i>(Ký và ghi rõ họ tên)</i>", ParagraphStyle('SignL2', parent=style_sub, alignment=1)),
        Spacer(1, 35)
    ]
    sign_right = [
        Paragraph(f"<i>TP. Hồ Chí Minh, ngày {datetime.datetime.now().strftime('%d')} tháng {datetime.datetime.now().strftime('%m')} năm {datetime.datetime.now().strftime('%Y')}</i>", ParagraphStyle('SignD', parent=style_sub, alignment=1)),
        Paragraph("<b>GIẢNG VIÊN CHẤM THI / PHỤ TRÁCH</b>", ParagraphStyle('SignR', parent=style_sub, alignment=1, fontName='Arial-Bold')),
        Paragraph("<i>(Xác nhận hồ sơ kỳ thi)</i>", ParagraphStyle('SignR2', parent=style_sub, alignment=1)),
        Spacer(1, 35)
    ]
    sign_table = Table([[sign_left, sign_right]], colWidths=[260, 263])
    sign_table.setStyle(TableStyle([
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
    ]))
    story.append(KeepTogether([sign_table]))

    # Build PDF with custom NumberedCanvas
    doc.build(story, canvasmaker=NumberedCanvas)
    return buffer.getvalue()


# --- Batch ZIP Export Generator ---
def generate_batch_exam_zip(exam: Any, results: List[Any], db: Any) -> Tuple[bytes, int]:
    """
    Generate an in-memory ZIP archive containing individual PDF exam submissions 
    for all candidates in the exam, plus an executive summary report.
    """
    from models import User, Question

    zip_buffer = io.BytesIO()
    exported_count = 0
    summary_lines = []

    exam_title = exam.title if exam else "Kỳ thi trắc nghiệm"
    exam_code = getattr(exam, 'code', '') or ''
    now_str = datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")

    summary_lines.append("=" * 80)
    summary_lines.append(f" BÁO CÁO TỔNG HỢP KẾT QUẢ KỲ THI TRỰC TUYẾN")
    summary_lines.append(f" Kỳ thi: {exam_title} (Mã: {exam_code})")
    summary_lines.append(f" Thời gian xuất báo cáo: {now_str}")
    summary_lines.append(f" Tổng số thí sinh dự thi: {len(results)}")
    summary_lines.append("=" * 80)
    summary_lines.append(f"{'STT':<5} | {'MSSV':<15} | {'HỌ VÀ TÊN':<28} | {'ĐIỂM':<7} | {'SỐ CÂU ĐÚNG':<12} | {'THỜI LƯỢNG':<10} | {'TRẠNG THÁI'}")
    summary_lines.append("-" * 105)

    exam_info = {
        "id": exam.id if exam else 1,
        "title": exam_title,
        "code": exam_code,
        "duration_minutes": getattr(exam, 'duration_minutes', 30),
        "num_questions": getattr(exam, 'num_questions', 10)
    }

    with zipfile.ZipFile(zip_buffer, 'w', compression=zipfile.ZIP_DEFLATED) as zip_file:
        for idx, res in enumerate(results):
            student = db.query(User).filter(User.id == res.user_id).first()
            username = student.username if student else f"sv_{res.user_id}"
            fullname = student.fullname if student else "Thí sinh"
            dob = getattr(student, 'dob', None) or ""

            # Extract audit details
            audit_questions = []
            if res.answers_detail:
                try:
                    audit_questions = json.loads(res.answers_detail)
                except Exception:
                    audit_questions = []

            # Fallback reconstruction if answers_detail is empty
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
                            "option_e": getattr(q, 'option_e', '') or '',
                            "option_f": getattr(q, 'option_f', '') or '',
                            "question_type": getattr(q, 'question_type', 'multiple_choice'),
                            "score_weight": getattr(q, 'score_weight', 1.0),
                            "selected": selected,
                            "correct": q.correct_option,
                            "is_correct": is_corr,
                            "status": st
                        })
                except Exception as e:
                    print(f"[PDF] Error rebuilding questions for result {res.id}: {e}")

            # Duration formatting
            dur_sec = res.duration_seconds or 0
            dur_min = dur_sec // 60
            dur_rem = dur_sec % 60
            duration_str = f"{dur_min}p {dur_rem:02d}s" if dur_sec > 0 else "-"

            start_str = res.start_time.strftime("%d/%m/%Y %H:%M:%S") if res.start_time else "-"
            submit_str = res.submit_time.strftime("%d/%m/%Y %H:%M:%S") if res.submit_time else "-"

            status_text = "Đã nộp bài" if res.status == "submitted" else "Đang làm dở"
            if res.status == "auto_submitted":
                status_text = "Nộp tự động (Hết giờ)"

            candidate_info = {
                "username": username,
                "fullname": fullname,
                "dob": dob,
                "score": res.score,
                "max_score": res.max_score or 10.0,
                "correct_count": res.correct_count or 0,
                "total_questions": res.total_questions or len(audit_questions),
                "duration_str": duration_str,
                "start_time_str": start_str,
                "submit_time_str": submit_str,
                "status_text": status_text
            }

            try:
                pdf_bytes = generate_candidate_pdf(exam_info, candidate_info, audit_questions)
                safe_name = sanitize_filename(f"BaiThi_{username}_{fullname}")
                zip_filename = f"{safe_name}.pdf"
                zip_file.writestr(zip_filename, pdf_bytes)
                exported_count += 1
            except Exception as e:
                print(f"[PDF] Error generating PDF for candidate {username}: {e}")
                import traceback
                traceback.print_exc()

            # Record in summary
            score_str = f"{res.score:.2f}" if res.score is not None else "Chờ chấm"
            correct_str = f"{res.correct_count or 0}/{res.total_questions or len(audit_questions)}" if res.score is not None else f"-/{res.total_questions or len(audit_questions)}"
            summary_lines.append(
                f"{idx+1:<5} | {username:<15} | {fullname:<28} | {score_str:<10} | {correct_str:<12} | {duration_str:<10} | {status_text}"
            )

        summary_lines.append("=" * 105)
        summary_lines.append(f"Tổng cộng đã xuất thành công: {exported_count} file PDF bài thi.")
        summary_text = "\n".join(summary_lines)
        zip_file.writestr("00_BANG_TONG_HOP_DIEM.txt", summary_text.encode('utf-8'))

    zip_bytes = zip_buffer.getvalue()
    return zip_bytes, exported_count
