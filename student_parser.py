"""
Smart Student List Parser for Vietnamese University & School Exam Systems
Supports: .xlsx, .xls, .csv, .docx
Intelligently locates table headers (skipping preamble/school titles),
maps relevant columns (MSSV, Full Name, DOB/Password),
combines split name columns (Họ lót + Tên),
safely ignores unrelated columns (STT, Lớp, Phòng thi, Ghi chú, Giới tính, etc.),
and normalizes Vietnamese date formats.
"""

import io
import re
import datetime
import unicodedata
from typing import List, Dict, Any, Tuple, Optional
import pandas as pd
import docx

def remove_accents(input_str: str) -> str:
    """Removes Vietnamese accents and lowercases for fuzzy header matching."""
    if not input_str:
        return ""
    # Normalize unicode to decomposed form
    nfkd_form = unicodedata.normalize('NFKD', str(input_str))
    no_accent = "".join([c for c in nfkd_form if not unicodedata.combining(c)])
    # Replace đ, Đ
    no_accent = no_accent.replace('đ', 'd').replace('Đ', 'D')
    return no_accent.lower().strip()

def clean_cell_text(val: Any) -> str:
    """Cleans a single cell's text value."""
    if val is None or pd.isna(val):
        return ""
    s = str(val).strip()
    if s.endswith(".0"):
        # Excel float converted from int e.g. 21110001.0
        # Check if the preceding chars are digits
        if s[:-2].isdigit():
            s = s[:-2]
    return s

def parse_vietnamese_date(val: Any) -> str:
    """
    Normalizes diverse date inputs into 'YYYY-MM-DD'.
    Handles:
      - pd.Timestamp, datetime.date, datetime.datetime
      - Excel serial numbers (e.g. 37636 -> 2003-01-15)
      - Strings: DD/MM/YYYY, DD-MM-YYYY, DD.MM.YYYY, YYYY-MM-DD, D/M/YYYY
    """
    if val is None or pd.isna(val):
        return "2003-01-01"
        
    if isinstance(val, (datetime.datetime, datetime.date, pd.Timestamp)):
        return val.strftime("%Y-%m-%d")
        
    # Check if numeric excel serial (e.g. 37636 or 37636.0)
    try:
        if isinstance(val, (int, float)) or (isinstance(val, str) and val.strip().replace('.0', '').isdigit()):
            num_val = float(val)
            # Excel serial dates typically range from 20000 (1954) to 60000 (2064)
            if 15000 <= num_val <= 60000:
                base_date = datetime.date(1899, 12, 30)
                dt = base_date + datetime.timedelta(days=int(num_val))
                return dt.strftime("%Y-%m-%d")
    except Exception:
        pass
        
    s = str(val).strip()
    if not s or s.lower() == "nan":
        return "2003-01-01"
        
    # Replace separators
    s_clean = s.replace(".", "/").replace("-", "/").replace(" ", "")
    
    # Check format DD/MM/YYYY or D/M/YYYY
    m = re.match(r"^(\d{1,2})/(\d{1,2})/(\d{4})$", s_clean)
    if m:
        day, month, year = int(m.group(1)), int(m.group(2)), int(m.group(3))
        # Swap if day > 12 and month <= 12
        try:
            return datetime.date(year, month, day).strftime("%Y-%m-%d")
        except ValueError:
            try:
                return datetime.date(year, day, month).strftime("%Y-%m-%d")
            except ValueError:
                return f"{year:04d}-{month:02d}-{day:02d}"
                
    # Check format YYYY/MM/DD
    m2 = re.match(r"^(\d{4})/(\d{1,2})/(\d{1,2})$", s_clean)
    if m2:
        year, month, day = int(m2.group(1)), int(m2.group(2)), int(m2.group(3))
        try:
            return datetime.date(year, month, day).strftime("%Y-%m-%d")
        except ValueError:
            return f"{year:04d}-{month:02d}-{day:02d}"
            
    # Check format DD/MM/YY (2 digits year)
    m3 = re.match(r"^(\d{1,2})/(\d{1,2})/(\d{2})$", s_clean)
    if m3:
        day, month, short_year = int(m3.group(1)), int(m3.group(2)), int(m3.group(3))
        year = 1900 + short_year if short_year > 40 else 2000 + short_year
        try:
            return datetime.date(year, month, day).strftime("%Y-%m-%d")
        except ValueError:
            return f"{year:04d}-{month:02d}-{day:02d}"

    # Only year e.g. "2003"
    if re.match(r"^\d{4}$", s):
        return f"{s}-01-01"
        
    return s

def identify_columns(headers: List[str]) -> Tuple[Optional[int], Optional[int], Optional[int], Optional[int], Optional[int], Optional[int], List[str]]:
    """
    Identifies column indices for:
      - mssv_idx: Student code
      - name_idx: Combined full name column
      - ho_idx: Split "Họ và tên đệm / Họ lót" column
      - ten_idx: Split "Tên" column
      - dob_idx: Date of birth / Password
      - class_idx: Class name
      - ignored_headers: List of column names that are safely skipped
    """
    mssv_idx = None
    name_idx = None
    ho_idx = None
    ten_idx = None
    dob_idx = None
    class_idx = None
    ignored_headers = []
    
    # Priority keyword patterns
    mssv_keywords = [
        "mssv", "ma sv", "masv", "ma_sv", "ma sinh vien", "ma so sinh vien", 
        "ma so sv", "ma hoc vien", "mahocvien", "student id", "student_id", 
        "id sinh vien", "sbd", "so bao danh", "ma dinh danh", "username", 
        "tai khoan", "ma_sinh_vien"
    ]
    
    dob_keywords = [
        "ngay sinh", "ngay_sinh", "ngaysinh", "ngay thang nam sinh", "dob", 
        "date of birth", "birthdate", "sinh ngay", "nam sinh", "ng.sinh", "ng. sinh"
    ]
    
    combined_name_keywords = [
        "ho va ten", "ho va ten thi sinh", "ho va ten sinh vien", "ho ten", 
        "ho_ten", "fullname", "full name", "ten sinh vien", "ten thi sinh"
    ]
    
    ho_keywords = [
        "ho va chu lot", "ho va ten dem", "ho lot", "ho_lot", "ho dem", "ho va ten lot",
        "ho", "last name", "surname", "family name"
    ]
    
    ten_keywords = [
        "ten", "first name", "given name"
    ]
    
    class_keywords = [
        "lop", "ten lop", "class", "class name", "class_name"
    ]

    used_indices = set()
    
    # 1. Detect MSSV
    for idx, col in enumerate(headers):
        clean = remove_accents(col)
        if any(kw == clean or f" {kw} " in f" {clean} " or clean.startswith(f"{kw} ") or clean.endswith(f" {kw}") for kw in mssv_keywords):
            mssv_idx = idx
            used_indices.add(idx)
            break
            
    # Fallback for MSSV: contains "mssv" or "ma sv"
    if mssv_idx is None:
        for idx, col in enumerate(headers):
            clean = remove_accents(col)
            if "mssv" in clean or "ma sv" in clean or "masv" in clean:
                mssv_idx = idx
                used_indices.add(idx)
                break

    # 2. Detect DOB
    for idx, col in enumerate(headers):
        if idx in used_indices:
            continue
        clean = remove_accents(col)
        if any(kw == clean or kw in clean for kw in dob_keywords):
            dob_idx = idx
            used_indices.add(idx)
            break

    # Detect Class
    for idx, col in enumerate(headers):
        if idx in used_indices:
            continue
        clean = remove_accents(col)
        if any(kw == clean or f" {kw} " in f" {clean} " or clean.startswith(f"{kw} ") or clean.endswith(f" {kw}") for kw in class_keywords):
            class_idx = idx
            used_indices.add(idx)
            break

    # 3. Detect Name (Single or Split)
    # Check if a dedicated "Tên" column exists first!
    # In Vietnamese student lists, if there is a separate "Tên" column,
    # the name is ALWAYS split into (Họ lót / Họ và tên đệm) + (Tên).
    for idx, col in enumerate(headers):
        if idx in used_indices:
            continue
        clean = remove_accents(col)
        if any(kw == clean or clean == kw or f" {kw} " in f" {clean} " or clean.endswith(f" {kw}") for kw in ten_keywords):
            # Make sure it's not a combined name like "ho va ten" or "ten sinh vien"
            if not any(comb in clean for comb in ["ho va ten", "ho ten", "ten sinh vien", "ten thi sinh", "full name", "fullname"]):
                ten_idx = idx
                used_indices.add(idx)
                break

    if ten_idx is not None:
        # We have a split name! Find the corresponding Họ / Họ lót column
        for idx, col in enumerate(headers):
            if idx in used_indices:
                continue
            clean = remove_accents(col)
            if any(kw == clean or kw in clean for kw in ho_keywords) or "lot" in clean or "dem" in clean:
                ho_idx = idx
                used_indices.add(idx)
                break
        if ho_idx is None:
            for idx, col in enumerate(headers):
                if idx in used_indices:
                    continue
                clean = remove_accents(col)
                if any(kw in clean for kw in combined_name_keywords) or "ho" in clean:
                    ho_idx = idx
                    used_indices.add(idx)
                    break
        if ho_idx is None and ten_idx > 0 and (ten_idx - 1) not in used_indices:
            ho_idx = ten_idx - 1
            used_indices.add(ho_idx)
    else:
        # No separate "Tên" column, check for combined full name
        for idx, col in enumerate(headers):
            if idx in used_indices:
                continue
            clean = remove_accents(col)
            if "lot" in clean or "dem" in clean:
                continue
            if any(kw == clean or f" {kw} " in f" {clean} " or clean.startswith(kw) for kw in combined_name_keywords):
                name_idx = idx
                used_indices.add(idx)
                break
                
        # If no combined name, check for split "Họ" and "Tên"
        if name_idx is None:
            for idx, col in enumerate(headers):
                if idx in used_indices:
                    continue
                clean = remove_accents(col)
                if any(kw == clean for kw in ho_keywords) or "ho lot" in clean or "chu lot" in clean or "ho va ten lot" in clean:
                    ho_idx = idx
                    used_indices.add(idx)
                    break
                    
            for idx, col in enumerate(headers):
                if idx in used_indices:
                    continue
                clean = remove_accents(col)
                if any(kw == clean for kw in ten_keywords):
                    ten_idx = idx
                    used_indices.add(idx)
                    break

        # If still no name, check if any column contains "ten" or "name"
        if name_idx is None and ho_idx is None and ten_idx is None:
            for idx, col in enumerate(headers):
                if idx in used_indices:
                    continue
                clean = remove_accents(col)
                if "ten" in clean or "name" in clean:
                    name_idx = idx
                    used_indices.add(idx)
                    break

    # Detect STT (Order index)
    stt_idx = None
    stt_keywords = ["stt", "tt", "so thu tu", "sothutu", "no", "order"]
    for idx, col in enumerate(headers):
        if idx in used_indices:
            continue
        clean = remove_accents(col)
        if any(kw == clean or f" {kw} " in f" {clean} " or clean == kw for kw in stt_keywords):
            stt_idx = idx
            break

    # Record all ignored columns
    for idx, col in enumerate(headers):
        if idx not in used_indices:
            ignored_headers.append(str(col).strip())

    return mssv_idx, name_idx, ho_idx, ten_idx, dob_idx, class_idx, stt_idx, ignored_headers

def find_header_row_in_matrix(matrix: List[List[Any]], max_scan_rows: int = 25) -> Tuple[int, List[str]]:
    """
    Scans the matrix to locate the most probable header row.
    Returns (header_row_index, list_of_cleaned_header_strings).
    """
    best_row = 0
    best_score = -1
    best_headers = []
    
    score_keywords = [
        "mssv", "ma sv", "masv", "ma sinh vien", "so bao danh", "sbd", 
        "ho ten", "ho va ten", "ten", "ho lot", "ngay sinh", "dob", "birth"
    ]
    
    limit = min(len(matrix), max_scan_rows)
    for r_idx in range(limit):
        row = matrix[r_idx]
        score = 0
        cleaned_row = []
        for cell in row:
            text = remove_accents(str(cell) if cell is not None and not pd.isna(cell) else "")
            cleaned_row.append(str(cell).strip() if cell is not None and not pd.isna(cell) else f"Col_{len(cleaned_row)}")
            for kw in score_keywords:
                if kw in text:
                    score += 2
                    
        # Check if contains both MSSV and (Name or DOB)
        has_mssv = any("mssv" in remove_accents(str(c)) or "ma sv" in remove_accents(str(c)) or "masv" in remove_accents(str(c)) for c in row if c is not None)
        has_dob = any("sinh" in remove_accents(str(c)) or "dob" in remove_accents(str(c)) for c in row if c is not None)
        has_name = any("ten" in remove_accents(str(c)) or "name" in remove_accents(str(c)) or "ho" in remove_accents(str(c)) for c in row if c is not None)
        
        if has_mssv:
            score += 5
        if has_dob:
            score += 3
        if has_name:
            score += 3
            
        if score > best_score:
            best_score = score
            best_row = r_idx
            best_headers = cleaned_row
            
    # If no row matched well, default to row 0
    if best_score <= 1 and limit > 0:
        best_row = 0
        best_headers = [str(c).strip() if c is not None and not pd.isna(c) else f"Col_{i}" for i, c in enumerate(matrix[0])]
        
    # Check if there is a 2-row merged header (sub-header row immediately following best_row)
    if best_row + 1 < len(matrix):
        next_row = matrix[best_row + 1]
        sub_keywords = ["ten", "ho", "ho lot", "chu lot", "ho dem", "mssv", "ma sv", "sinh", "lop", "stt", "phong thi", "sbd"]
        sub_matches = sum(1 for c in next_row if any(sk == remove_accents(str(c)) or sk in remove_accents(str(c)) for sk in sub_keywords))
        
        has_sub_header = False
        if sub_matches >= 1:
            mssv_cand_idx = None
            for idx, h in enumerate(best_headers):
                h_clean = remove_accents(h)
                if "mssv" in h_clean or "ma sv" in h_clean or "masv" in h_clean:
                    mssv_cand_idx = idx
                    break
            
            if mssv_cand_idx is not None and mssv_cand_idx < len(next_row):
                next_mssv_val = str(next_row[mssv_cand_idx]).strip() if next_row[mssv_cand_idx] is not None and not pd.isna(next_row[mssv_cand_idx]) else ""
                is_student_id = bool(re.search(r'\d{4,}', next_mssv_val))
                if not is_student_id:
                    has_sub_header = True
            elif sub_matches >= 2:
                has_sub_header = True
                
        if has_sub_header:
            merged_headers = []
            max_cols = max(len(best_headers), len(next_row))
            for c_idx in range(max_cols):
                h1 = best_headers[c_idx] if c_idx < len(best_headers) else ""
                val2 = next_row[c_idx] if c_idx < len(next_row) else None
                h2 = str(val2).strip() if val2 is not None and not pd.isna(val2) else ""
                
                h1_is_col = h1.startswith("Col_") or not h1
                h2_clean = remove_accents(h2)
                
                if h1_is_col and h2:
                    merged_headers.append(h2)
                elif h2 and h2_clean in ["ten", "ho", "ho lot", "chu lot", "ho va chu lot", "ho va ten dem", "ho dem"]:
                    merged_headers.append(f"{h1} - {h2}" if h1 and not h1_is_col else h2)
                elif h2 and h1 != h2 and not h1_is_col:
                    merged_headers.append(f"{h1} - {h2}")
                elif h1 and not h1_is_col:
                    merged_headers.append(h1)
                else:
                    merged_headers.append(h2 if h2 else f"Col_{c_idx}")
            best_headers = merged_headers
            best_row = best_row + 1

    return best_row, best_headers

def parse_student_data_matrix(matrix: List[List[Any]], header_row_idx: int, headers: List[str]) -> Dict[str, Any]:
    """
    Extracts student records from a 2D matrix starting after header_row_idx.
    """
    mssv_idx, name_idx, ho_idx, ten_idx, dob_idx, class_idx, stt_idx, ignored_headers = identify_columns(headers)
    
    if mssv_idx is None:
        raise ValueError("Không thể tự động nhận diện cột Mã số sinh viên (MSSV / Mã SV). Vui lòng kiểm tra lại tiêu đề cột trong file.")
        
    students = []
    seen_mssv = set()
    
    # Process data rows
    for r_idx in range(header_row_idx + 1, len(matrix)):
        row = matrix[r_idx]
        if not row or len(row) <= mssv_idx:
            continue
            
        raw_mssv = clean_cell_text(row[mssv_idx])
        if not raw_mssv or raw_mssv.lower() in ("nan", "none", "", "tổng cộng", "tong cong", "cán bộ", "can bo"):
            continue
            
        # Skip if MSSV doesn't look like a student code (e.g. repeated header or remark row)
        if "mssv" in remove_accents(raw_mssv) or "ma sv" in remove_accents(raw_mssv):
            continue
            
        # Avoid duplicate MSSVs within the same upload batch
        if raw_mssv in seen_mssv:
            continue
            
        # Parse Full Name
        fullname = ""
        if name_idx is not None and name_idx < len(row):
            fullname = clean_cell_text(row[name_idx])
        elif ho_idx is not None and ten_idx is not None:
            ho = clean_cell_text(row[ho_idx]) if ho_idx < len(row) else ""
            ten = clean_cell_text(row[ten_idx]) if ten_idx < len(row) else ""
            fullname = f"{ho} {ten}".strip()
        elif ho_idx is not None and ho_idx < len(row):
            fullname = clean_cell_text(row[ho_idx])
        elif ten_idx is not None and ten_idx < len(row):
            fullname = clean_cell_text(row[ten_idx])
            
        if not fullname or fullname.lower() == "nan":
            fullname = raw_mssv
            
        # Parse DOB
        dob_val = ""
        if dob_idx is not None and dob_idx < len(row):
            dob_val = parse_vietnamese_date(row[dob_idx])
        else:
            # Default password/DOB if column completely missing
            dob_val = "2003-01-01"

        # Parse Class Name
        class_name = ""
        if class_idx is not None and class_idx < len(row):
            class_name = clean_cell_text(row[class_idx])

        # Parse STT / Order Index
        order_num = len(students) + 1
        if stt_idx is not None and stt_idx < len(row):
            raw_stt = clean_cell_text(row[stt_idx])
            try:
                m_num = re.search(r'\d+', raw_stt)
                if m_num:
                    order_num = int(m_num.group(0))
            except Exception:
                pass
            
        students.append({
            "mssv": raw_mssv,
            "fullname": fullname,
            "dob": dob_val,
            "class_name": class_name,
            "order_index": order_num
        })
        seen_mssv.add(raw_mssv)
        
    detected_cols = {
        "mssv": headers[mssv_idx] if mssv_idx < len(headers) else "MSSV",
        "fullname": headers[name_idx] if name_idx is not None and name_idx < len(headers) else (
            f"{headers[ho_idx]} + {headers[ten_idx]}" if ho_idx is not None and ten_idx is not None else "Họ và tên"
        ),
        "dob": headers[dob_idx] if dob_idx is not None and dob_idx < len(headers) else "Mặc định (2003-01-01)",
        "class_name": headers[class_idx] if class_idx is not None and class_idx < len(headers) else "Không có"
    }
    
    return {
        "success": True,
        "students": students,
        "detected_columns": detected_cols,
        "ignored_columns": ignored_headers,
        "total_parsed": len(students),
        "header_row_index": header_row_idx + 1
    }

def parse_student_file(contents: bytes, filename: str) -> Dict[str, Any]:
    """
    Main entry point for parsing any student list file.
    Supports: .xlsx, .xls, .csv, .docx
    """
    fn_lower = filename.lower()
    
    # 1. EXCEL (.xlsx, .xls)
    if fn_lower.endswith(".xlsx") or fn_lower.endswith(".xls"):
        df_raw = None
        try:
            df_raw = pd.read_excel(io.BytesIO(contents), header=None, dtype=str)
        except Exception:
            # Fallback 1: Many Vietnamese university portals export HTML tables with .xls extension
            try:
                tables = pd.read_html(io.BytesIO(contents), header=None)
                if tables:
                    df_raw = max(tables, key=lambda t: t.shape[0] * t.shape[1]).astype(str)
            except Exception:
                pass
            
            # Fallback 2: Tab-separated or CSV with .xls extension
            if df_raw is None:
                for enc in ["utf-8-sig", "utf-8", "cp1258", "cp1252", "latin1"]:
                    for sep in ["\t", ",", ";"]:
                        try:
                            df_try = pd.read_csv(io.BytesIO(contents), header=None, encoding=enc, sep=sep, dtype=str)
                            if df_try.shape[1] > 1:
                                df_raw = df_try
                                break
                        except Exception:
                            continue
                    if df_raw is not None:
                        break
                        
        if df_raw is None:
            raise ValueError(f"Không thể đọc file Excel '{filename}'. Vui lòng kiểm tra lại định dạng tệp.")
            
        matrix = df_raw.values.tolist()
        header_row_idx, headers = find_header_row_in_matrix(matrix)
        return parse_student_data_matrix(matrix, header_row_idx, headers)
        
    # 2. CSV (.csv)
    elif fn_lower.endswith(".csv"):
        # Try multiple encodings
        encodings = ["utf-8-sig", "utf-8", "cp1258", "cp1252", "latin1"]
        df_raw = None
        for enc in encodings:
            for sep in [",", ";", "\t"]:
                try:
                    df_try = pd.read_csv(io.BytesIO(contents), header=None, encoding=enc, sep=sep, dtype=str)
                    if df_try.shape[1] > 1:
                        df_raw = df_try
                        break
                except Exception:
                    continue
            if df_raw is not None:
                break
                
        if df_raw is None:
            raise ValueError("Không thể đọc file CSV. Vui lòng kiểm tra định dạng hoặc bảng mã tiếng Việt.")
            
        matrix = df_raw.values.tolist()
        header_row_idx, headers = find_header_row_in_matrix(matrix)
        return parse_student_data_matrix(matrix, header_row_idx, headers)
        
    # 3. WORD (.docx)
    elif fn_lower.endswith(".docx"):
        doc = docx.Document(io.BytesIO(contents))
        # Find first table in docx
        if not doc.tables:
            raise ValueError("File Word không chứa bảng danh sách sinh viên nào.")
            
        table = doc.tables[0]
        matrix = []
        for r in table.rows:
            row_cells = [cell.text.strip() for cell in r.cells]
            matrix.append(row_cells)
            
        header_row_idx, headers = find_header_row_in_matrix(matrix)
        return parse_student_data_matrix(matrix, header_row_idx, headers)
        
    else:
        raise ValueError(f"Định dạng file '{filename}' không được hỗ trợ. Vui lòng tải lên file Excel (.xlsx, .xls), CSV (.csv), hoặc Word (.docx).")
