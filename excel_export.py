import io
import re
import html
from typing import List, Dict, Any, Optional

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from docx_parser import strip_question_prefix


def clean_html_to_text(html_text: Optional[str]) -> str:
    """Converts HTML text to clean, readable plain text for Excel cells."""
    if not html_text:
        return ""
    text = str(html_text)
    # Replace image tags
    text = re.sub(r'<img[^>]*>', ' [Hình ảnh] ', text, flags=re.IGNORECASE)
    # Replace breaks and paragraphs with newlines
    text = re.sub(r'<(?:br|p|div)[^>]*>', '\n', text, flags=re.IGNORECASE)
    text = re.sub(r'</(?:p|div)>', '\n', text, flags=re.IGNORECASE)
    # Strip any remaining tags
    text = re.sub(r'<[^>]+>', ' ', text)
    # Decode HTML entities (&nbsp;, &amp;, etc.)
    text = html.unescape(text)
    # Normalize whitespace per line
    lines = [re.sub(r'[ \t]+', ' ', line).strip() for line in text.split('\n')]
    # Remove redundant blank lines
    cleaned_lines = []
    for line in lines:
        if line:
            cleaned_lines.append(line)
        elif cleaned_lines and cleaned_lines[-1] != '':
            cleaned_lines.append('')
    return "\n".join(cleaned_lines).strip()


def format_option_text(letter: str, raw_text: Optional[str]) -> str:
    """Formats an option with letter prefix, e.g. 'A. [Nội dung phương án A]'."""
    clean_txt = clean_html_to_text(raw_text)
    if not clean_txt:
        return letter.upper()
    # If the text already starts with 'A.' or 'A:' or 'A -', don't duplicate
    if re.match(r'^[A-F][\s.:)\-–—]', clean_txt, re.IGNORECASE):
        return clean_txt
    return f"{letter.upper()}. {clean_txt}"


def format_answers_text(val: Any, options_map: Dict[str, str], is_essay: bool, is_correct_guide: bool = False, explanation: str = "") -> str:
    """Formats correct answers or student answers for Excel cells."""
    if is_essay:
        if is_correct_guide:
            exp_clean = clean_html_to_text(explanation or val or "")
            if exp_clean:
                return f"Gợi ý đáp án / Rubric chấm:\n{exp_clean}"
            return "Theo thang điểm & hướng dẫn chấm tự luận của giảng viên"
        else:
            txt = clean_html_to_text(val)
            return txt if txt else "(Chưa làm bài tự luận)"

    if val is None or (isinstance(val, str) and not val.strip()):
        return "(Chưa trả lời)" if not is_correct_guide else "(Chưa có đáp án mẫu)"

    # Extract chosen letters
    if isinstance(val, (list, tuple, set)):
        letters = [str(x).strip().upper() for x in val if str(x).strip()]
    elif isinstance(val, str):
        # Extract letters A, B, C, D, E, F
        letters = [c.upper() for c in re.split(r'[,;\s]+', val.strip()) if c]
    else:
        letters = [str(val).strip().upper()]

    if not letters:
        return clean_html_to_text(str(val))

    parts = []
    for let in letters:
        opt_text = options_map.get(let.upper(), "")
        parts.append(format_option_text(let, opt_text))

    return "\n".join(parts) if parts else str(val)


def generate_candidate_excel(
    exam_info: Dict[str, Any],
    candidate_info: Dict[str, Any],
    audit_questions: List[Dict[str, Any]]
) -> bytes:
    """
    Generates an official candidate submission review spreadsheet (.xlsx)
    containing:
      1. STT
      2. Câu hỏi
      3. Câu trả lời đúng
      4. Câu trả lời của thí sinh
      5. Làm đúng (x)
      6. Điểm số (với hàm =SUM ở dòng tổng cộng)
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Chi Tiết Bài Làm"
    ws.views.sheetView[0].showGridLines = True

    # Palette styles
    title_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    total_fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
    alt_row_fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
    
    font_org = Font(name="Calibri", size=10, bold=True, color="1E3A8A")
    font_main_title = Font(name="Calibri", size=14, bold=True, color="0F172A")
    font_meta = Font(name="Calibri", size=10, color="475569")
    font_candidate = Font(name="Calibri", size=10, bold=True, color="1E293B")
    
    font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    font_total_label = Font(name="Calibri", size=11, bold=True, color="0F172A")
    font_total_count = Font(name="Calibri", size=11, bold=True, color="059669")
    font_total_sum = Font(name="Calibri", size=12, bold=True, color="1E3A8A")

    thin_border = Border(
        left=Side(style="thin", color="CBD5E1"),
        right=Side(style="thin", color="CBD5E1"),
        top=Side(style="thin", color="CBD5E1"),
        bottom=Side(style="thin", color="CBD5E1")
    )
    total_border = Border(
        top=Side(style="medium", color="1E3A8A"),
        bottom=Side(style="double", color="1E3A8A"),
        left=Side(style="thin", color="CBD5E1"),
        right=Side(style="thin", color="CBD5E1")
    )

    # 1. Title Banner
    ws.merge_cells("A1:F1")
    ws["A1"] = "TRƯỜNG ĐH SƯ PHẠM KỸ THUẬT TP.HCM — KHOA ĐÀO TẠO & KHẢO THÍ"
    ws["A1"].font = font_org
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")

    ws.merge_cells("A2:F2")
    ws["A2"] = "BẢNG ĐỐI CHIẾU CHI TIẾT BÀI LÀM THÍ SINH"
    ws["A2"].font = font_main_title
    ws["A2"].alignment = Alignment(horizontal="center", vertical="center")

    mc_qs = [q for q in audit_questions if q.get("question_type") != "essay"]
    essay_qs = [q for q in audit_questions if q.get("question_type") == "essay"]

    if not essay_qs:
        eff_mc_max = 10.0
        eff_essay_max = 0.0
    elif not mc_qs:
        eff_mc_max = 0.0
        eff_essay_max = 10.0
    else:
        cfg_mc = float(exam_info.get("mc_max_score") if exam_info.get("mc_max_score") is not None else 7.0)
        cfg_essay = float(exam_info.get("essay_max_score") if exam_info.get("essay_max_score") is not None else 3.0)
        tot_cfg = cfg_mc + cfg_essay
        if tot_cfg <= 0:
            cfg_mc = 7.0
            cfg_essay = 3.0
            tot_cfg = 10.0
        eff_mc_max = round((cfg_mc / tot_cfg) * 10.0, 2)
        eff_essay_max = round(10.0 - eff_mc_max, 2)

    total_mc_w = sum(float(q.get("score_weight") or 1.0) for q in mc_qs)
    mc_scale = (eff_mc_max / total_mc_w) if total_mc_w > 0 else 0.0

    total_essay_w = sum(float(q.get("score_weight") or 1.0) for q in essay_qs)
    essay_scale = (eff_essay_max / total_essay_w) if total_essay_w > 0 else 0.0

    exam_title = exam_info.get("title", "Kỳ thi trắc nghiệm")
    exam_code = exam_info.get("code") or "CK-2026"
    split_info = f" | Thang điểm: Trắc nghiệm {eff_mc_max:g}đ" + (f" + Tự luận {eff_essay_max:g}đ" if essay_qs else "")
    ws.merge_cells("A3:F3")
    ws["A3"] = f"Kỳ thi: {exam_title} | Mã ca thi: {exam_code}{split_info}"
    ws["A3"].font = font_meta
    ws["A3"].alignment = Alignment(horizontal="center", vertical="center")

    fullname = candidate_info.get("fullname", "Thí sinh")
    mssv = candidate_info.get("username", "-")
    class_name = candidate_info.get("class_name", "-") or "-"
    submit_str = candidate_info.get("submit_time_str", "-")
    score_display = candidate_info.get("score")
    score_str = f"{score_display}/10.0" if score_display is not None else "Chờ chấm"

    ws.merge_cells("A4:F4")
    ws["A4"] = f"Thí sinh: {fullname} | MSSV: {mssv} | Lớp: {class_name} | Thời gian nộp: {submit_str} | Điểm số: {score_str}"
    ws["A4"].font = font_candidate
    ws["A4"].alignment = Alignment(horizontal="center", vertical="center")

    ws.row_dimensions[1].height = 20
    ws.row_dimensions[2].height = 26
    ws.row_dimensions[3].height = 18
    ws.row_dimensions[4].height = 20
    ws.row_dimensions[5].height = 10

    # 2. Table Headers (Row 6)
    headers = [
        "STT",
        "Nội dung câu hỏi",
        "Câu trả lời đúng",
        "Câu trả lời của thí sinh",
        "Làm đúng (x)",
        "Điểm số"
    ]
    header_row = 6
    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=header_row, column=col_idx, value=h)
        cell.fill = title_fill
        cell.font = font_header
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border
    ws.row_dimensions[header_row].height = 28

    # 4. Populate Question Rows
    start_data_row = 7
    current_row = start_data_row

    for idx, q in enumerate(audit_questions):
        is_essay = (q.get("question_type") == "essay")
        options_map = {
            "A": q.get("option_a", ""),
            "B": q.get("option_b", ""),
            "C": q.get("option_c", ""),
            "D": q.get("option_d", ""),
            "E": q.get("option_e", ""),
            "F": q.get("option_f", "")
        }

        # Col A: STT
        stt_val = idx + 1
        
        # Col B: Question clean text without 'Câu X.' prefix
        raw_content = q.get("content", "")
        q_cleaned = clean_html_to_text(strip_question_prefix(raw_content))

        # Col C: Correct answer text
        correct_val = q.get("correct") or q.get("correct_option") or ""
        correct_text = format_answers_text(
            val=correct_val,
            options_map=options_map,
            is_essay=is_essay,
            is_correct_guide=True,
            explanation=q.get("explanation", "")
        )

        # Col D: Student answer text
        selected_val = q.get("selected")
        student_text = format_answers_text(
            val=selected_val,
            options_map=options_map,
            is_essay=is_essay,
            is_correct_guide=False
        )

        # Col E: Correct mark 'x'
        # Col F: Numeric points
        earned_pts = 0.0
        mark_val = ""

        if is_essay:
            if q.get("essay_score") is not None:
                essay_sc = float(q.get("essay_score"))
                earned_pts = round(essay_sc * essay_scale, 2)
                mark_val = "x" if essay_sc > 0 else ""
            elif q.get("status") == "pending_grading":
                mark_val = "Chờ chấm"
                earned_pts = 0.0
            else:
                mark_val = ""
                earned_pts = 0.0
        else:
            is_corr = q.get("is_correct") or (q.get("status") == "correct")
            if is_corr:
                mark_val = "x"
                raw_sc = float(q.get("earned_score") if q.get("earned_score") is not None else (q.get("score_weight") or 1.0))
                earned_pts = round(raw_sc * mc_scale, 2)
            elif q.get("status") == "partial":
                mark_val = "x (1 phần)"
                raw_sc = float(q.get("earned_score") or 0.0)
                earned_pts = round(raw_sc * mc_scale, 2)
            else:
                mark_val = ""
                earned_pts = 0.0

        # Write cells
        c_stt = ws.cell(row=current_row, column=1, value=stt_val)
        c_stt.alignment = Alignment(horizontal="center", vertical="top")

        c_q = ws.cell(row=current_row, column=2, value=q_cleaned)
        c_q.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)

        c_corr = ws.cell(row=current_row, column=3, value=correct_text)
        c_corr.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)

        c_stu = ws.cell(row=current_row, column=4, value=student_text)
        c_stu.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)

        c_mark = ws.cell(row=current_row, column=5, value=mark_val)
        c_mark.alignment = Alignment(horizontal="center", vertical="top")
        if mark_val == "x":
            c_mark.font = Font(name="Calibri", size=12, bold=True, color="059669")
        elif "x" in mark_val:
            c_mark.font = Font(name="Calibri", size=11, bold=True, color="0284C7")
        elif mark_val == "Chờ chấm":
            c_mark.font = Font(name="Calibri", size=10, italic=True, color="D97706")

        c_pts = ws.cell(row=current_row, column=6, value=earned_pts)
        c_pts.alignment = Alignment(horizontal="right", vertical="top")
        c_pts.number_format = "0.00"

        # Borders & Zebra striping
        is_even = ((idx + 1) % 2 == 0)
        for col_idx in range(1, 7):
            cell = ws.cell(row=current_row, column=col_idx)
            cell.border = thin_border
            if is_even and col_idx != 5:
                cell.fill = alt_row_fill

        current_row += 1

    last_data_row = current_row - 1
    if last_data_row < start_data_row:
        last_data_row = start_data_row

    # 5. Bottom Total Row with SUM and COUNTIF
    total_row = current_row
    ws.cell(row=total_row, column=1, value="")
    ws.cell(row=total_row, column=2, value="")
    ws.cell(row=total_row, column=3, value="")
    
    lbl_cell = ws.cell(row=total_row, column=4, value="TỔNG CỘNG:")
    lbl_cell.font = font_total_label
    lbl_cell.alignment = Alignment(horizontal="right", vertical="center")

    # Count of correct questions
    cnt_cell = ws.cell(row=total_row, column=5, value=f'=COUNTIF(E{start_data_row}:E{last_data_row}, "x*")')
    cnt_cell.font = font_total_count
    cnt_cell.alignment = Alignment(horizontal="center", vertical="center")

    # SUM formula for total points
    sum_cell = ws.cell(row=total_row, column=6, value=f"=SUM(F{start_data_row}:F{last_data_row})")
    sum_cell.font = font_total_sum
    sum_cell.alignment = Alignment(horizontal="right", vertical="center")
    sum_cell.number_format = "0.00"

    for col_idx in range(1, 7):
        cell = ws.cell(row=total_row, column=col_idx)
        cell.border = total_border
        cell.fill = total_fill
    ws.row_dimensions[total_row].height = 26

    # 6. Column Widths
    ws.column_dimensions["A"].width = 7
    ws.column_dimensions["B"].width = 46
    ws.column_dimensions["C"].width = 36
    ws.column_dimensions["D"].width = 36
    ws.column_dimensions["E"].width = 16
    ws.column_dimensions["F"].width = 14

    stream = io.BytesIO()
    wb.save(stream)
    stream.seek(0)
    return stream.getvalue()
