"""
Exam Question Parser for Word documents (.docx)
Features:
- Robust image extraction across all OpenXML structures:
  Inline drawings (wp:inline), Floating/Anchor drawings (wp:anchor),
  Legacy VML shapes/imagedata (v:imagedata, w:pict), Group shapes (wpg:wgp, wps:wsp),
  mc:AlternateContent with deduplication (Choice priority, Fallback safe),
  Table cell images, Option images, and duplicate image reuse across multiple questions.
- Word Equation (OMML / MathML) support:
  Converts <m:oMath> equations (fractions, powers, subscripts, radicals, delimiters)
  into clean, readable HTML so formulas never disappear.
- Hyperlink support: preserves text and URLs from <w:hyperlink> nodes.
- Multi-format question detection:
  'Câu 1:', 'Câu 01.', 'Câu hỏi 1:', 'CH 1:', '1.', '1)', '1/', 'Question 1:', 'Q1:', 'Bài 1:', etc.
- Multi-format option detection:
  Separate lines (A. \n B. \n ...), single line (A. ... B. ... C. ... D. ...), table cells.
- Flexible correct answer detection:
  1. Explicit answer lines ('Đáp án: A', 'Key: B', 'Đ/A: C', 'Hướng dẫn: Chọn D', etc.)
  2. Asterisk markings ('*A.', '[x] A.')
  3. Underlined options (single line or exact character-range overlap in multi-option lines)
  4. Bolded options (single line or exact character-range overlap in multi-option lines)
- HTML rich text preservation (bold, italic, underline, strikethrough, subscript, superscript, tables).
- Zero image loss: images embedded as base64 data URIs.
"""

import base64
import html
import io
import re
from typing import List, Dict, Any, Optional, Set, Tuple
import docx
from docx.oxml.ns import nsmap

# Ensure all needed namespaces are present in nsmap
for prefix, uri in [
    ('v', 'urn:schemas-microsoft-com:vml'),
    ('o', 'urn:schemas-microsoft-com:office:office'),
    ('wps', 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape'),
    ('wpg', 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup'),
    ('mc', 'http://schemas.openxmlformats.org/markup-compatibility/2006'),
    ('m', 'http://schemas.openxmlformats.org/officeDocument/2006/math'),
    ('wp', 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'),
    ('a', 'http://schemas.openxmlformats.org/drawingml/2006/main'),
    ('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'),
]:
    if prefix not in nsmap:
        nsmap[prefix] = uri


def get_image_data_uri(image_part) -> Optional[str]:
    """Convert a docx ImagePart to a base64 data URI string."""
    try:
        blob = image_part.blob
        content_type = getattr(image_part, 'content_type', None)
        
        # Detect content type from bytes if missing or generic
        if not content_type or content_type == 'application/octet-stream':
            if blob.startswith(b'\x89PNG\r\n\x1a\n'):
                content_type = 'image/png'
            elif blob.startswith(b'\xff\xd8\xff'):
                content_type = 'image/jpeg'
            elif blob.startswith(b'GIF87a') or blob.startswith(b'GIF89a'):
                content_type = 'image/gif'
            elif blob.startswith(b'RIFF') and blob[8:12] == b'WEBP':
                content_type = 'image/webp'
            elif blob.startswith(b'BM'):
                content_type = 'image/bmp'
            elif blob.startswith(b'<svg') or blob.startswith(b'<?xml') or b'<svg' in blob[:100]:
                content_type = 'image/svg+xml'
            else:
                content_type = 'image/png'
                
        b64 = base64.b64encode(blob).decode('ascii')
        return f"data:{content_type};base64,{b64}"
    except Exception as e:
        print(f"Error encoding image: {e}")
        return None


def format_img_tags(img_uris: List[str]) -> str:
    """Format image URIs as clean HTML img tags with responsive styling."""
    html_out = ""
    for uri in img_uris:
        html_out += (
            f'<div class="exam-img-container" style="margin: 10px 0; text-align: center;">'
            f'<img src="{uri}" class="exam-question-img" alt="Hình ảnh minh họa" '
            f'style="max-width: 100%; height: auto; max-height: 450px; border-radius: 8px; '
            f'border: 1px solid #cbd5e1; box-shadow: 0 2px 6px rgba(0,0,0,0.06); display: inline-block;" />'
            f'</div>'
        )
    return html_out


def strip_question_prefix(text: str) -> str:
    """
    Loại bỏ tiền tố định danh câu hỏi (ví dụ: 'Câu 1.', 'Câu 01:', '1.', 'Q1:', '<strong>Câu 1. </strong>', v.v.)
    khỏi nội dung câu hỏi để câu hỏi lưu trong ngân hàng chỉ chứa nội dung thuần túy.
    Điều này giúp khi hệ thống xáo ngẫu nhiên câu hỏi cho thí sinh thì thứ tự câu hỏi không bị lộn xộn.
    """
    if not text:
        return ""
    pattern = re.compile(
        r'^\s*(?:<p[^>]*>)?\s*(?:<(?:strong|b|span|em|i)[^>]*>)?\s*'
        r'(?:(?:câu|cau|question|q|bài|bai)\s*\d+|\d+)'
        r'(?:\s*\([^)]*(?:điểm|diem|pt|point)[^)]*\))?'
        r'[\s.:)\-–—]*'
        r'(?:</(?:strong|b|span|em|i)>\s*)*'
        r'[\s.:)\-–—]*',
        re.IGNORECASE
    )
    cleaned = pattern.sub('', text, count=1).strip()
    cleaned = re.sub(r'^(?:<br\s*/?>|\s)+', '', cleaned, flags=re.IGNORECASE)
    return cleaned if cleaned else text


def extract_images_from_element(element, doc, seen_rids: Optional[Set[str]] = None) -> List[str]:
    """
    Extract image data URIs from an element.
    Uses mc:Choice first, ignoring mc:Fallback if Choice contains images.
    seen_rids deduplicates within the scope of the caller (e.g. current paragraph).
    """
    if seen_rids is None:
        seen_rids = set()
    img_uris = []
    
    # 1. Standard DrawingML: w:drawing -> a:blip (embed or link)
    # Exclude any blip inside mc:Fallback to prevent duplicate extraction
    blips = element.xpath('.//a:blip[not(ancestor::mc:Fallback)]/@r:embed') + \
            element.xpath('.//a:blip[not(ancestor::mc:Fallback)]/@r:link')
            
    for r_id in blips:
        if r_id and r_id not in seen_rids:
            seen_rids.add(r_id)
            if r_id in doc.part.related_parts:
                part = doc.part.related_parts[r_id]
                uri = get_image_data_uri(part)
                if uri:
                    img_uris.append(uri)

    # 2. Legacy VML: w:pict -> v:imagedata (r:id or r:href)
    vml_rids = element.xpath('.//v:imagedata[not(ancestor::mc:Fallback)]/@r:id') + \
               element.xpath('.//v:imagedata[not(ancestor::mc:Fallback)]/@r:href')
    for r_id in vml_rids:
        if r_id and r_id not in seen_rids:
            seen_rids.add(r_id)
            if r_id in doc.part.related_parts:
                part = doc.part.related_parts[r_id]
                uri = get_image_data_uri(part)
                if uri:
                    img_uris.append(uri)

    # 3. If no images found yet, check mc:Fallback as a fallback
    if not img_uris:
        fallback_blips = element.xpath('.//mc:Fallback//a:blip/@r:embed') + \
                         element.xpath('.//mc:Fallback//v:imagedata/@r:id')
        for r_id in fallback_blips:
            if r_id and r_id not in seen_rids:
                seen_rids.add(r_id)
                if r_id in doc.part.related_parts:
                    part = doc.part.related_parts[r_id]
                    uri = get_image_data_uri(part)
                    if uri:
                        img_uris.append(uri)

    return img_uris


def omml_to_html(node) -> str:
    """
    Recursively converts an Office OpenXML Math element (<m:oMath>) to clean HTML.
    Supports fractions, powers, subscripts, sub-sup, radicals, and delimiters.
    """
    if node is None:
        return ""
    tag = node.tag.split('}')[-1]
    
    if tag == 't':
        return html.escape(node.text or '')
    elif tag == 'f':
        # Fraction: numerator / denominator
        num_el = node.find('{http://schemas.openxmlformats.org/officeDocument/2006/math}num')
        den_el = node.find('{http://schemas.openxmlformats.org/officeDocument/2006/math}den')
        num_html = omml_to_html(num_el) if num_el is not None else ''
        den_html = omml_to_html(den_el) if den_el is not None else ''
        return f'<span class="math-frac"><sup>{num_html}</sup>/<sub>{den_html}</sub></span>'
    elif tag == 'sSup':
        e_el = node.find('{http://schemas.openxmlformats.org/officeDocument/2006/math}e')
        sup_el = node.find('{http://schemas.openxmlformats.org/officeDocument/2006/math}sup')
        e_html = omml_to_html(e_el) if e_el is not None else ''
        sup_html = omml_to_html(sup_el) if sup_el is not None else ''
        return f'{e_html}<sup>{sup_html}</sup>'
    elif tag == 'sSub':
        e_el = node.find('{http://schemas.openxmlformats.org/officeDocument/2006/math}e')
        sub_el = node.find('{http://schemas.openxmlformats.org/officeDocument/2006/math}sub')
        e_html = omml_to_html(e_el) if e_el is not None else ''
        sub_html = omml_to_html(sub_el) if sub_el is not None else ''
        return f'{e_html}<sub>{sub_html}</sub>'
    elif tag == 'sSubSup':
        e_el = node.find('{http://schemas.openxmlformats.org/officeDocument/2006/math}e')
        sub_el = node.find('{http://schemas.openxmlformats.org/officeDocument/2006/math}sub')
        sup_el = node.find('{http://schemas.openxmlformats.org/officeDocument/2006/math}sup')
        e_html = omml_to_html(e_el) if e_el is not None else ''
        sub_html = omml_to_html(sub_el) if sub_el is not None else ''
        sup_html = omml_to_html(sup_el) if sup_el is not None else ''
        return f'{e_html}<sub>{sub_html}</sub><sup>{sup_html}</sup>'
    elif tag == 'rad':
        deg_el = node.find('.//{http://schemas.openxmlformats.org/officeDocument/2006/math}deg')
        e_el = node.find('{http://schemas.openxmlformats.org/officeDocument/2006/math}e')
        deg_html = omml_to_html(deg_el) if deg_el is not None else ''
        e_html = omml_to_html(e_el) if e_el is not None else ''
        prefix = f'<sup>{deg_html}</sup>' if deg_html else ''
        return f'{prefix}&radic;({e_html})'
    elif tag == 'd':
        beg, end = '(', ')'
        dPr = node.find('{http://schemas.openxmlformats.org/officeDocument/2006/math}dPr')
        if dPr is not None:
            begChr = dPr.find('{http://schemas.openxmlformats.org/officeDocument/2006/math}begChr')
            endChr = dPr.find('{http://schemas.openxmlformats.org/officeDocument/2006/math}endChr')
            if begChr is not None:
                beg = begChr.get('{http://schemas.openxmlformats.org/officeDocument/2006/math}val', beg)
            if endChr is not None:
                end = endChr.get('{http://schemas.openxmlformats.org/officeDocument/2006/math}val', end)
        inner = ''.join(omml_to_html(child) for child in node if not child.tag.endswith('dPr'))
        return f'{beg}{inner}{end}'
    else:
        return ''.join(omml_to_html(child) for child in node)


def run_to_html(run, doc, seen_rids: Set[str]) -> str:
    """Convert a docx Run to HTML, including inline formatting and images."""
    text = run.text
    run_imgs = extract_images_from_element(run._r, doc, seen_rids)
    
    if not text and not run_imgs:
        return ""
        
    escaped_text = html.escape(text) if text else ""
    
    # Auto-linkify URLs in plain text
    if escaped_text:
        escaped_text = re.sub(
            r'(https?://[^\s<]+)',
            r'<a href="\1" target="_blank" rel="noopener noreferrer" style="color:#2563eb; text-decoration:underline;">\1</a>',
            escaped_text
        )
        
    # Check bold, italic, underline, strike, sub, sup
    is_bold = bool(run.bold or run._r.xpath('.//w:b[not(@w:val="0" or @w:val="false")]'))
    is_italic = bool(run.italic or run._r.xpath('.//w:i[not(@w:val="0" or @w:val="false")]'))
    
    is_underlined = False
    if run.underline is True or (run.underline is not None and run.underline != 0):
        is_underlined = True
    else:
        u_tags = run._r.xpath('.//w:u')
        if u_tags:
            val = u_tags[0].get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', '')
            if val and val != 'none':
                is_underlined = True
                
    is_strike = bool(run._r.xpath('.//w:strike'))
    is_sub = bool(run._r.xpath('.//w:vertAlign[@w:val="subscript"]'))
    is_sup = bool(run._r.xpath('.//w:vertAlign[@w:val="superscript"]'))
    
    res = escaped_text
    if is_bold and res:
        res = f"<strong>{res}</strong>"
    if is_italic and res:
        res = f"<em>{res}</em>"
    if is_underlined and res:
        res = f"<u>{res}</u>"
    if is_strike and res:
        res = f"<s>{res}</s>"
    if is_sub and res:
        res = f"<sub>{res}</sub>"
    if is_sup and res:
        res = f"<sup>{res}</sup>"
        
    if run_imgs:
        res += format_img_tags(run_imgs)
        
    return res


def paragraph_to_html(para, doc, seen_rids: Optional[Set[str]] = None) -> str:
    """
    Convert a paragraph to rich HTML by traversing its XML children in true visual order.
    Handles runs, hyperlinks, OMML equations, and paragraph-level images.
    """
    if seen_rids is None:
        seen_rids = set()
        
    html_parts = []
    
    for child in para._p:
        tag = child.tag.split('}')[-1]
        
        if tag == 'r':
            run = docx.text.run.Run(child, para)
            r_html = run_to_html(run, doc, seen_rids)
            if r_html:
                html_parts.append(r_html)
                
        elif tag == 'hyperlink':
            # Word hyperlink: wraps runs in <a href="...">
            h_runs_html = []
            for r_child in child.xpath('.//w:r'):
                r = docx.text.run.Run(r_child, para)
                h_runs_html.append(run_to_html(r, doc, seen_rids))
            link_text = "".join(h_runs_html)
            
            # Lookup target URL from relationships
            r_id = child.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
            url = ""
            if r_id and r_id in doc.part.rels:
                url = doc.part.rels[r_id].target_ref
                
            if url and link_text:
                html_parts.append(f'<a href="{html.escape(url)}" target="_blank" rel="noopener noreferrer" style="color:#2563eb; text-decoration:underline;">{link_text}</a>')
            elif link_text:
                html_parts.append(link_text)
                
        elif tag in ('oMath', 'oMathPara'):
            # Math equation
            m_html = omml_to_html(child)
            if m_html:
                html_parts.append(f'<span class="math-expr" style="font-family: Cambria Math, Times New Roman, serif;">{m_html}</span>')
                
        elif tag in ('drawing', 'pict'):
            # Paragraph-level floating/anchor drawings
            p_imgs = extract_images_from_element(child, doc, seen_rids)
            if p_imgs:
                html_parts.append(format_img_tags(p_imgs))
                
        elif tag == 'AlternateContent':
            choice = child.find('{http://schemas.openxmlformats.org/markup-compatibility/2006}Choice')
            if choice is not None:
                ac_imgs = extract_images_from_element(choice, doc, seen_rids)
                if ac_imgs:
                    html_parts.append(format_img_tags(ac_imgs))
            else:
                fallback = child.find('{http://schemas.openxmlformats.org/markup-compatibility/2006}Fallback')
                if fallback is not None:
                    ac_imgs = extract_images_from_element(fallback, doc, seen_rids)
                    if ac_imgs:
                        html_parts.append(format_img_tags(ac_imgs))

    out = "".join(html_parts).strip()
    return out


def table_to_html(tbl, doc, seen_rids: Optional[Set[str]] = None) -> str:
    """Convert a docx table to styled responsive HTML table."""
    if seen_rids is None:
        seen_rids = set()
        
    rows_html = []
    for row in tbl.rows:
        cells_html = []
        for cell in row.cells:
            cell_content = []
            for p in cell.paragraphs:
                p_html = paragraph_to_html(p, doc, seen_rids)
                if p_html:
                    cell_content.append(p_html)
            cell_imgs = extract_images_from_element(cell._tc, doc, seen_rids)
            if cell_imgs:
                cell_content.append(format_img_tags(cell_imgs))
                
            cell_text = "<br>".join(cell_content) if cell_content else "&nbsp;"
            cells_html.append(f'<td style="border:1px solid #cbd5e1; padding:8px 12px; vertical-align:top;">{cell_text}</td>')
        rows_html.append(f"<tr>{''.join(cells_html)}</tr>")
        
    return f'<div class="exam-table-wrapper" style="overflow-x:auto; margin:12px 0;"><table style="border-collapse:collapse; width:100%; border:1px solid #cbd5e1; background:white; font-size:0.95em;">{"".join(rows_html)}</table></div>'


def clean_option_prefix_html(h: str, opt_char: str) -> str:
    """
    Remove option label (e.g. 'A.', '<strong>A.</strong>', '*B.', '<u>C.</u>')
    without leaving broken/unclosed dangling HTML tags.
    """
    if not h:
        return ""
    # Strip nested tags wrapping option prefix e.g. <strong><em>B.</em></strong>
    for _ in range(3):
        h = re.sub(
            r'^\s*<([a-z0-9]+)[^>]*>\s*[\*\[\(]?\s*' + re.escape(opt_char) + r'[\.\:\)\/\-\]]\s*<\/\1>\s*',
            '', h, flags=re.IGNORECASE
        )
    # Strip prefix inside an open tag e.g. <strong>B. Content</strong> -> <strong>Content</strong>
    h = re.sub(
        r'(<[a-z0-9]+[^>]*>)\s*[\*\[\(]?\s*' + re.escape(opt_char) + r'[\.\:\)\/\-\]]\s*',
        r'\1', h, flags=re.IGNORECASE
    )
    # Strip plain prefix at start
    h = re.sub(
        r'^\s*[\*\[\(]?\s*' + re.escape(opt_char) + r'[\.\:\)\/\-\]]\s*',
        '', h, flags=re.IGNORECASE
    )
    # Remove empty tags like <strong></strong>
    while re.search(r'<([a-z0-9]+)[^>]*>\s*<\/\1>', h, flags=re.IGNORECASE):
        h = re.sub(r'<([a-z0-9]+)[^>]*>\s*<\/\1>', '', h, flags=re.IGNORECASE)
        
    return h.strip()


def check_para_for_underline(para) -> bool:
    """Check if any run in paragraph has underline."""
    for run in para.runs:
        if run.text.strip():
            if run.underline is True or (run.underline is not None and run.underline != 0):
                return True
            u_tags = run._r.xpath('.//w:u')
            if u_tags:
                val = u_tags[0].get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', '')
                if val and val != 'none':
                    return True
    return False


def check_para_for_bold(para) -> bool:
    """Check if any run in paragraph is bold."""
    for run in para.runs:
        if run.text.strip():
            if run.bold or run._r.xpath('.//w:b[not(@w:val="0" or @w:val="false")]'):
                return True
    return False


def check_opt_content_bold(para, opt_char: str) -> bool:
    """
    Check if the option content (beyond just the prefix label A.) is bold.
    This prevents false positives when a teacher styles all option prefixes in bold.
    """
    runs = [r for r in para.runs if r.text.strip()]
    if not runs:
        return False
    # If first run is just the prefix (e.g. "A." or "A)"):
    if len(runs) > 1 and re.match(r'^\s*[\*\[\(]?\s*' + re.escape(opt_char) + r'[\.\:\)\/\-\]]\s*$', runs[0].text.strip(), re.IGNORECASE):
        for r in runs[1:]:
            if r.bold or r._r.xpath('.//w:b[not(@w:val="0" or @w:val="false")]'):
                return True
        return False
    else:
        # First run has prefix and potentially content
        first_r = runs[0]
        is_first_bold = bool(first_r.bold or first_r._r.xpath('.//w:b[not(@w:val="0" or @w:val="false")]'))
        if is_first_bold:
            txt = first_r.text.strip()
            txt_no_prefix = re.sub(r'^\s*[\*\[\(]?\s*' + re.escape(opt_char) + r'[\.\:\)\/\-\]]\s*', '', txt, flags=re.IGNORECASE)
            if txt_no_prefix.strip():
                return True
        for r in runs[1:]:
            if r.bold or r._r.xpath('.//w:b[not(@w:val="0" or @w:val="false")]'):
                return True
        return False


def check_opt_content_underline(para, opt_char: str) -> bool:
    """
    Check if the option content (beyond just the prefix label A.) is underlined.
    """
    runs = [r for r in para.runs if r.text.strip()]
    if not runs:
        return False
    if len(runs) > 1 and re.match(r'^\s*[\*\[\(]?\s*' + re.escape(opt_char) + r'[\.\:\)\/\-\]]\s*$', runs[0].text.strip(), re.IGNORECASE):
        for r in runs[1:]:
            if r.underline is True or (r.underline is not None and r.underline != 0) or r._r.xpath('.//w:u[not(@w:val="none")]'):
                return True
        return False
    else:
        first_r = runs[0]
        is_first_u = bool(first_r.underline is True or (first_r.underline is not None and first_r.underline != 0) or first_r._r.xpath('.//w:u[not(@w:val="none")]'))
        if is_first_u:
            txt = first_r.text.strip()
            txt_no_prefix = re.sub(r'^\s*[\*\[\(]?\s*' + re.escape(opt_char) + r'[\.\:\)\/\-\]]\s*', '', txt, flags=re.IGNORECASE)
            if txt_no_prefix.strip():
                return True
        for r in runs[1:]:
            if r.underline is True or (r.underline is not None and r.underline != 0) or r._r.xpath('.//w:u[not(@w:val="none")]'):
                return True
        return False


def para_has_media(para) -> bool:
    """Check if paragraph has any drawings, blips, math, or picts."""
    blips = para._p.xpath('.//a:blip')
    vmls = para._p.xpath('.//v:imagedata')
    drawings = para._p.xpath('.//w:drawing')
    picts = para._p.xpath('.//w:pict')
    maths = para._p.xpath('.//m:oMath')
    return bool(blips or vmls or drawings or picts or maths)


# Regex patterns for Vietnamese & English questions, options, and answers
QUESTION_REGEX = re.compile(
    r'^(?:câu\s*hỏi|cau\s*hoi|ch|câu|cau|question|q|bài|bai)\s*(\d*)[\.\:\)\/\-\s]\s*(.*)',
    re.IGNORECASE
)
NUMBERED_Q_REGEX = re.compile(
    r'^(\d+)[\.\:\)\/\-]\s+(.*)',
    re.IGNORECASE
)
SCORE_REGEX = re.compile(
    r'\((\d+(?:\.\d+)?)\s*(?:điểm|diem|pts?|points?)\)', 
    re.IGNORECASE
)
PLUS_OPTION_REGEX = re.compile(
    r'^\s*\+\s*(.+)', 
    re.IGNORECASE
)

# Matches answer indicators:
# Đáp án: A, Đ/A: B, Key: A, B, C, E, Hướng dẫn giải: Chọn A, C
ANSWER_REGEX = re.compile(
    r'(?:đáp\s*án(?:\s*đúng)?|dap\s*an(?:\s*dung)?|đ\/a|d\/a|đa|da|correct(?:\s*answer)?|answer|key|hướng\s*dẫn(?:\s*giải)?\s*:\s*chọn|chọn)\s*[\:\-\.\s]+\s*([A-Ha-h](?:[\s,\;\.\-\+]+[A-Ha-h])*)\b',
    re.IGNORECASE
)

# Single option line: 'A. ...' or 'A) ...' or 'E. ...' or '*A. ...'
OPTION_LINE_REGEX = re.compile(
    r'^\s*[\*\[\(]?\s*([A-H])[\.\:\)\/\-\]]\s*(.*)',
    re.IGNORECASE
)

# Multiple options on the same line: 'A. Nội dung A    B. Nội dung B ...'
MULTI_OPTION_REGEX = re.compile(
    r'(?:^|\s+)[\*\[\(]?\s*([A-H])[\.\:\)\/\-\]]\s*(.*?)(?=(?:\s+[\*\[\(]?\s*[A-H][\.\:\)\/\-\]]|$))',
    re.IGNORECASE
)

# Hint for multi-select question in question title
MULTI_SELECT_HINT_REGEX = re.compile(
    r'(?:chọn\s+nhiều|nhiều\s+đáp\s*án|nhiều\s+lựa\s*chọn|lựa\s*chọn\s+các|chọn\s+các|multi-?select|multiple\s+choice|nhiều\s+hơn\s+1\s+đáp\s*án)',
    re.IGNORECASE
)


def detect_section_header(text: str) -> Optional[str]:
    """
    Detect if a paragraph is a section header:
    e.g. 'PHẦN I: TRẮC NGHIỆM', 'PHẦN II: TỰ LUẬN', 'PART I: MULTIPLE CHOICE', 'PART II: ESSAY'
    Returns 'essay', 'multiple_choice', or None.
    """
    if not text:
        return None
    t = text.strip()
    if len(t) > 150:
        return None
    t_clean = re.sub(r'<[^>]+>', '', t).strip()

    is_part_prefix = bool(re.match(r'^(?:(?:phần|phan|part|mục|muc|chương|chuong)\s*(?:[0-9ivx]+|[a-z])?|(?:[0-9ivx]+|[a-z])[\.\:\-\–\—\s]+(?:phần|phan|part|mục|muc))\b', t_clean, re.IGNORECASE))
    is_standalone_sec = bool(re.match(r'^(?:phần|phan|part|mục|muc|bài\s*tập|câu\s*hỏi|các\s*câu\s*hỏi)?\s*(?:trắc|trac|tự|tu|essay|multiple|nhiều\s*lựa|nhieu\s*lua)', t_clean, re.IGNORECASE))
    is_heading_type = bool(re.match(r'^(?:[0-9ivx]+|[a-z])[\.\:\-\–\—\s]+(?:câu\s*hỏi|bài\s*tập|phần|mục)?\s*(?:trắc|trac|tự|tu|essay|multiple|nhiều\s*lựa|nhieu\s*lua)', t_clean, re.IGNORECASE))

    is_essay_keyword = bool(re.search(r'(?:tự\s*luận|tu\s*luan|\bessay\b)', t_clean, re.IGNORECASE))
    is_mc_keyword = bool(re.search(r'(?:trắc\s*nghiệm|trac\s*nghiem|multiple\s*choice|nhiều\s*lựa\s*chọn|nhieu\s*lua\s*chon)', t_clean, re.IGNORECASE))

    if (is_part_prefix or is_standalone_sec or is_heading_type) and (is_essay_keyword or is_mc_keyword):
        if is_essay_keyword and is_mc_keyword:
            m_pos = re.search(r'(?:trắc\s*nghiệm|trac\s*nghiem|multiple\s*choice|nhiều\s*lựa\s*chọn|nhieu\s*lua\s*chon)', t_clean, re.IGNORECASE).start()
            e_pos = re.search(r'(?:tự\s*luận|tu\s*luan|\bessay\b)', t_clean, re.IGNORECASE).start()
            return "multiple_choice" if m_pos < e_pos else "essay"
        if is_essay_keyword:
            return "essay"
        if is_mc_keyword:
            return "multiple_choice"

    return None


def parse_docx_questions(doc_file) -> List[Dict[str, Any]]:
    """
    Parses a docx file (file path or file-like object) and returns a list of questions:
    [
        {
            "content": "HTML string containing question text and any embedded images/tables/math",
            "option_a": "HTML or text of option A",
            "option_b": "HTML or text of option B",
            "option_c": "HTML or text of option C",
            "option_d": "HTML or text of option D",
            "correct_option": "A" | "B" | "C" | "D"
        }
    ]
    """
    doc = docx.Document(doc_file)
    
    # Extract body items in true visual order (paragraphs and tables)
    body_items = []
    for child in doc.element.body:
        tag = child.tag.split('}')[-1]
        if tag == 'p':
            p = docx.text.paragraph.Paragraph(child, doc)
            body_items.append(('p', p))
        elif tag == 'tbl':
            tbl = docx.table.Table(child, doc)
            body_items.append(('tbl', tbl))
            
    questions: List[Dict[str, Any]] = []
    curr_q: Optional[Dict[str, Any]] = None
    last_opt_key: Optional[str] = None
    curr_section: Optional[str] = None
    has_section_headers: bool = False
    
    def finalize_current_q():
        nonlocal curr_q
        if not curr_q:
            return
            
        opts = curr_q.get('opts', {})
        score_weight = 1.0
        m_score = SCORE_REGEX.search(curr_q['content'])
        if m_score:
            score_weight = float(m_score.group(1))
            
        plus_opts = curr_q.get('plus_opts', [])
        
        # Strip question number prefix (e.g. 'Câu 1.', 'Câu 20.', etc.) so question content contains pure text
        curr_q['content'] = strip_question_prefix(curr_q['content'])
        
        # If question is explicitly under Essay section, it is ALWAYS an essay question
        if curr_section == "essay":
            if opts:
                sub_items_html = "<br>" + "<br>".join([f"<b>{k}.</b> {v}" for k, v in sorted(opts.items())])
                curr_q['content'] += sub_items_html
                opts = {}
            if plus_opts:
                bullets_html = "<br>" + "<br>".join([f"+ {p}" for p in plus_opts])
                curr_q['content'] += bullets_html
                plus_opts = []
            questions.append({
                "content": curr_q['content'],
                "question_type": "essay",
                "score_weight": score_weight,
                "option_a": "",
                "option_b": "",
                "option_c": "",
                "option_d": "",
                "option_e": "",
                "option_f": "",
                "correct_option": ""
            })
            curr_q = None
            return

        # Only activate plus_opts if question does NOT have standard options A, B, C, D (len(opts) < 2)
        # and has at least 2 plus_opts!
        if len(opts) < 2 and len(plus_opts) >= 2:
            opt_chars = ['A', 'B', 'C', 'D', 'E', 'F']
            mapped_opts = {}
            for i, opt in enumerate(plus_opts):
                if i < len(opt_chars):
                    mapped_opts[opt_chars[i]] = opt
            
            questions.append({
                "content": curr_q['content'],
                "question_type": "multi_select",
                "score_weight": score_weight,
                "option_a": mapped_opts.get('A', ''),
                "option_b": mapped_opts.get('B', ''),
                "option_c": mapped_opts.get('C', ''),
                "option_d": mapped_opts.get('D', ''),
                "option_e": mapped_opts.get('E', ''),
                "option_f": mapped_opts.get('F', ''),
                "correct_option": ",".join(opt_chars[:min(len(plus_opts), len(opt_chars))])
            })
        elif len(opts) >= 2:
            # If there were plus_opts collected before the options (e.g. from bullet points in question description),
            # restore them into content so their text isn't lost!
            if plus_opts:
                bullets_html = "<br>" + "<br>".join([f"+ {p}" for p in plus_opts])
                curr_q['content'] += bullets_html

            # 0. Check if question content explicitly hints at multi-select
            has_multi_hint = bool(MULTI_SELECT_HINT_REGEX.search(curr_q['content']))
            
            correct_opts_set = set()
            
            # Check explicit answer line: "Đáp án: A, B, C, E"
            if curr_q.get('correct'):
                for ch in re.findall(r'[A-Ha-h]', curr_q['correct']):
                    ch_u = ch.upper()
                    if ch_u in opts:
                        correct_opts_set.add(ch_u)
            
            # Check asterisk marks
            if not correct_opts_set and curr_q.get('asterisk_opts'):
                correct_opts_set = set(curr_q['asterisk_opts']).intersection(opts.keys())
                
            # Check underline marks (if not ALL options are underlined)
            if not correct_opts_set and curr_q.get('underlined_opts'):
                if len(curr_q['underlined_opts']) < len(opts):
                    correct_opts_set = set(curr_q['underlined_opts']).intersection(opts.keys())
                    
            # Check bold marks (if not ALL options are bolded)
            if not correct_opts_set and curr_q.get('bolded_opts'):
                if len(curr_q['bolded_opts']) < len(opts):
                    correct_opts_set = set(curr_q['bolded_opts']).intersection(opts.keys())
                    
            # Do NOT fallback to 'A' if no answer detected! If the test has no answers,
            # correct_opts_set remains empty, so the teacher can see and set answers later.
            is_multi = has_multi_hint or (len(correct_opts_set) > 1)
            q_type = "multi_select" if is_multi else "multiple_choice"
            correct_str = ",".join(sorted(list(correct_opts_set))) if correct_opts_set else ""
            
            questions.append({
                "content": curr_q['content'],
                "question_type": q_type,
                "score_weight": score_weight,
                "option_a": opts.get('A', ''),
                "option_b": opts.get('B', ''),
                "option_c": opts.get('C', ''),
                "option_d": opts.get('D', ''),
                "option_e": opts.get('E', ''),
                "option_f": opts.get('F', ''),
                "correct_option": correct_str
            })
        else:
            if plus_opts:
                bullets_html = "<br>" + "<br>".join([f"+ {p}" for p in plus_opts])
                curr_q['content'] += bullets_html
            questions.append({
                "content": curr_q['content'],
                "question_type": "essay",
                "score_weight": score_weight,
                "option_a": "",
                "option_b": "",
                "option_c": "",
                "option_d": "",
                "option_e": "",
                "option_f": "",
                "correct_option": ""
            })
        curr_q = None

    for item_type, item in body_items:
        if item_type == 'p':
            para = item
            clean_text = para.text.strip()
            has_media = para_has_media(para)
            
            # If completely empty of both text and media, skip
            if not clean_text and not has_media:
                continue
                
            # If paragraph has NO text but HAS media (drawings, VML, math):
            if not clean_text and has_media:
                para_seen_rids: Set[str] = set()
                para_html = paragraph_to_html(para, doc, para_seen_rids)
                if curr_q and para_html:
                    if curr_q['opts'] and last_opt_key:
                        curr_q['opts'][last_opt_key] += "<br>" + para_html
                    else:
                        curr_q['content'] += "<br>" + para_html
                continue

            # -1. Check for Section Header (PHẦN I: TRẮC NGHIỆM / PHẦN II: TỰ LUẬN)
            detected_sec = detect_section_header(clean_text)
            if detected_sec:
                has_section_headers = True
                finalize_current_q()
                curr_section = detected_sec
                curr_q = None
                last_opt_key = None
                continue

            # 0. Check for Multi-select (+) option line FIRST (prevents ANSWER_REGEX eating "+ Đáp án A")
            m_plus = PLUS_OPTION_REGEX.match(clean_text)
            if m_plus and curr_q:
                opt_text = m_plus.group(1).strip()
                opt_seen_rids: Set[str] = set()
                para_html = paragraph_to_html(para, doc, opt_seen_rids)
                # Remove the plus sign for cleaner HTML
                para_html = re.sub(r'^\s*\+\s*', '', para_html)
                curr_q['plus_opts'].append(para_html if para_html.strip() else opt_text)
                continue

            # 1. Check for explicit Answer Line:
            m_ans = ANSWER_REGEX.search(clean_text)
            if m_ans and curr_q:
                ans_str = m_ans.group(1).strip()
                ans_letters = [ch.upper() for ch in re.findall(r'[A-Ha-h]', ans_str)]
                if ans_letters:
                    curr_q['correct'] = ",".join(sorted(list(set(ans_letters))))
                last_opt_key = None
                continue
                
            # 2. Check for Multiple Options on single line (e.g. A. ... B. ... C. ... D. ...)
            multi_matches = list(MULTI_OPTION_REGEX.finditer(clean_text))
            if len(multi_matches) >= 2 and curr_q:
                opt_spans = []
                for m in multi_matches:
                    opt_char_upper = m.group(1).upper()
                    opt_val = m.group(2).strip()
                    if opt_val.startswith('*') or opt_val.endswith('*'):
                        curr_q['asterisk_opts'].add(opt_char_upper)
                        opt_val = opt_val.strip('* ')
                    curr_q['opts'][opt_char_upper] = opt_val
                    last_opt_key = opt_char_upper
                    opt_spans.append({
                        'char': opt_char_upper,
                        'start': m.start(),
                        'end': m.end()
                    })
                    
                # Character-offset overlap analysis for underline and bold
                # Accurately checks which run corresponds to which option without false positives
                offset = 0
                for r in para.runs:
                    r_len = len(r.text)
                    r_start = offset
                    r_end = offset + r_len
                    offset = r_end
                    
                    if not r.text.strip():
                        continue
                        
                    r_u = bool(r.underline is True or (r.underline is not None and r.underline != 0) or r._r.xpath('.//w:u[not(@w:val="none")]'))
                    r_b = bool(r.bold or r._r.xpath('.//w:b[not(@w:val="0" or @w:val="false")]'))
                    
                    if r_u or r_b:
                        for span in opt_spans:
                            # Option content starts after prefix label
                            content_start = span['start'] + 2
                            if max(r_start, content_start) < min(r_end, span['end']):
                                if r_u:
                                    curr_q['underlined_opts'].add(span['char'])
                                if r_b:
                                    curr_q['bolded_opts'].add(span['char'])
                continue
                
            # 3. Check for Single Option line (e.g. A. ... or A) ... or *A. ...)
            m_opt = OPTION_LINE_REGEX.match(clean_text)
            if m_opt and curr_q:
                opt_char = m_opt.group(1).upper()
                opt_text = m_opt.group(2).strip()
                
                # Check for asterisk marker
                if '*' in clean_text or '[x]' in clean_text.lower():
                    curr_q['asterisk_opts'].add(opt_char)
                    opt_text = opt_text.replace('*', '').strip()
                    
                # Check for underline in option content (excluding prefix label)
                if check_opt_content_underline(para, opt_char):
                    curr_q['underlined_opts'].add(opt_char)
                    
                # Check for bold in option content (excluding prefix label)
                if check_opt_content_bold(para, opt_char):
                    curr_q['bolded_opts'].add(opt_char)
                    
                # Convert option runs and media into HTML using scoped seen_rids
                opt_seen_rids: Set[str] = set()
                para_html = paragraph_to_html(para, doc, opt_seen_rids)
                clean_opt_html = clean_option_prefix_html(para_html, opt_char)
                
                curr_q['opts'][opt_char] = clean_opt_html if clean_opt_html.strip() else opt_text
                last_opt_key = opt_char
                continue


            # 4. Check for New Question line
            m_q = QUESTION_REGEX.match(clean_text)
            if not m_q and not OPTION_LINE_REGEX.match(clean_text) and not PLUS_OPTION_REGEX.match(clean_text):
                m_num = NUMBERED_Q_REGEX.match(clean_text)
                if m_num:
                    has_score = bool(SCORE_REGEX.search(clean_text))
                    is_prev_complete = (curr_q is None or len(curr_q['opts']) >= 2 or len(curr_q.get('plus_opts', [])) >= 2)
                    is_prev_essay = (curr_q is not None and not curr_q['opts'] and not curr_q.get('plus_opts') and bool(curr_q['content'].strip()))
                    if is_prev_complete or has_score or is_prev_essay or (curr_section == 'essay'):
                        m_q = m_num
                    
            if m_q and not OPTION_LINE_REGEX.match(clean_text) and not PLUS_OPTION_REGEX.match(clean_text):
                # New question! Finalize previous
                finalize_current_q()
                
                q_seen_rids: Set[str] = set()
                para_html = paragraph_to_html(para, doc, q_seen_rids)
                curr_q = {
                    'content': para_html,
                    'opts': {},
                    'plus_opts': [],
                    'correct': None,
                    'underlined_opts': set(),
                    'bolded_opts': set(),
                    'asterisk_opts': set()
                }
                last_opt_key = None
                continue
                
            # 5. Continuation text of question or option
            if curr_q is not None:
                cont_seen_rids: Set[str] = set()
                para_html = paragraph_to_html(para, doc, cont_seen_rids)
                if not curr_q['opts']:
                    curr_q['content'] += "<br>" + para_html
                elif last_opt_key:
                    curr_q['opts'][last_opt_key] += "<br>" + para_html

        elif item_type == 'tbl':
            tbl = item
            # Check if this table contains questions or options
            table_has_options = False
            extracted_opts = {}
            tbl_underlined = set()
            tbl_bolded = set()
            tbl_asterisk = set()
            
            for row in tbl.rows:
                for cell in row.cells:
                    c_text = cell.text.strip()
                    m_cell_opt = OPTION_LINE_REGEX.match(c_text)
                    if m_cell_opt:
                        table_has_options = True
                        c_char = m_cell_opt.group(1).upper()
                        c_val = m_cell_opt.group(2).strip()
                        if '*' in c_text or '[x]' in c_text.lower():
                            tbl_asterisk.add(c_char)
                            c_val = c_val.replace('*', '').strip()
                        # Check cell underline / bold
                        for p in cell.paragraphs:
                            if check_para_for_underline(p):
                                tbl_underlined.add(c_char)
                            if check_para_for_bold(p):
                                tbl_bolded.add(c_char)
                                
                        cell_seen: Set[str] = set()
                        c_imgs = extract_images_from_element(cell._tc, doc, cell_seen)
                        if c_imgs:
                            c_val += "<br>" + format_img_tags(c_imgs)
                        extracted_opts[c_char] = c_val
                        
            if curr_q is not None:
                if table_has_options and len(extracted_opts) >= 2:
                    curr_q['opts'].update(extracted_opts)
                    curr_q['underlined_opts'].update(tbl_underlined)
                    curr_q['bolded_opts'].update(tbl_bolded)
                    curr_q['asterisk_opts'].update(tbl_asterisk)
                else:
                    tbl_seen: Set[str] = set()
                    tbl_html = table_to_html(tbl, doc, tbl_seen)
                    if not curr_q['opts']:
                        curr_q['content'] += "<br>" + tbl_html
                    elif last_opt_key:
                        curr_q['opts'][last_opt_key] += "<br>" + tbl_html

    finalize_current_q()

    # If section headers (e.g. PHẦN I: TRẮC NGHIỆM, PHẦN II: TỰ LUẬN) were present,
    # guarantee strict separation: Part I (Multiple Choice & Multi Select) first, Part II (Essay) second
    if has_section_headers:
        mc_part = [q for q in questions if q.get("question_type") != "essay"]
        essay_part = [q for q in questions if q.get("question_type") == "essay"]
        return mc_part + essay_part
    return questions
