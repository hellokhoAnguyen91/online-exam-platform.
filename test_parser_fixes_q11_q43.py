import sys
sys.stdout.reconfigure(encoding='utf-8')
import os
import docx
from docx_parser import parse_docx_questions

def test_regression_q11_q43():
    docx_path = r"E:\New folder (2)\111\thông tư 01\Ngan hang trac nghiem - 50 cau.docx"
    if not os.path.exists(docx_path):
        print(f"Skipping file test, {docx_path} not found")
        return

    qs = parse_docx_questions(docx_path)
    assert len(qs) == 50, f"Expected 50 questions, got {len(qs)}"

    # Test Q11 (MSS question)
    q11 = None
    for q in qs:
        if 'MSS' in q['content']:
            q11 = q
            break
    assert q11 is not None, "Q11 (MSS question) not found"
    assert q11['option_a'], "Option A of Q11 must not be empty"
    assert "Mọi tiêu chuẩn MSS loại B" in q11['option_a']
    assert q11['option_b'], "Option B of Q11 must not be empty"
    assert "giống như loại A." in q11['option_b'] or "giống như loại A" in q11['option_b']
    assert q11['option_c'], "Option C of Q11 must not be empty"
    assert "xem là loại A." in q11['option_c'] or "xem là loại A" in q11['option_c']
    assert q11['option_d'], "Option D of Q11 must not be empty"
    assert "tiêu chuẩn loại A." in q11['option_d'] or "tiêu chuẩn loại A" in q11['option_d']
    assert q11['option_e'], "Option E of Q11 must not be empty"
    print("✓ Test Q11 (MSS 5 options preservation) passed!")

    # Test Q43 (Corrosive hazard images)
    q43 = None
    for q in qs:
        if 'ăn mòn' in q['content'] or 'vận chuyển hóa chất nguy hiểm' in q['content']:
            q43 = q
            break
    assert q43 is not None, "Q43 (Corrosive hazard) not found"
    for opt in ['a', 'b', 'c', 'd', 'e']:
        val = q43.get(f'option_{opt}')
        assert val, f"Option {opt.upper()} of Q43 must not be empty"
        assert '<img' in val, f"Option {opt.upper()} of Q43 must contain an image tag"
        assert 'class="exam-question-img"' in val or 'exam-img-container' in val
    print("✓ Test Q43 (Corrosive hazard table images in options) passed!")

def test_regression_de1():
    docx_path = r"E:\New folder (2)\111\thông tư 01\DE 1 - CK (50 TN + 5 TL).docx"
    if not os.path.exists(docx_path):
        print(f"Skipping file test, {docx_path} not found")
        return

    qs = parse_docx_questions(docx_path)
    assert len(qs) == 55, f"Expected 55 questions, got {len(qs)}"
    mc_part = [q for q in qs if q['question_type'] != 'essay']
    essay_part = [q for q in qs if q['question_type'] == 'essay']
    assert len(mc_part) == 50, f"Expected 50 MC questions, got {len(mc_part)}"
    assert len(essay_part) == 5, f"Expected 5 Essay questions, got {len(essay_part)}"

    # Test Q23 in DE 1 (MSS question)
    q23 = None
    for q in qs:
        if 'MSS' in q['content']:
            q23 = q
            break
    assert q23 is not None, "Q23 in DE 1 not found"
    assert q23['option_a'], "Option A of Q23 must not be empty"
    assert q23['option_e'], "Option E of Q23 must not be empty"
    print("✓ Test DE 1 (Word numbering resolution & 50 MC + 5 Essay) passed!")

if __name__ == "__main__":
    test_regression_q11_q43()
    test_regression_de1()
    print("\n🎉 ALL REGRESSION TESTS PASSED 100%!")
