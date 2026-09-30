import sys
import os
import docx
from docx.oxml import parse_xml
from docx_parser import parse_docx_questions

sys.stdout.reconfigure(encoding='utf-8')

def test_existing_files():
    print("=== Testing sample_questions.docx ===")
    if os.path.exists("sample_questions.docx"):
        qs = parse_docx_questions("sample_questions.docx")
        print(f"Parsed {len(qs)} questions from sample_questions.docx")
        for i, q in enumerate(qs):
            has_img = "data:image" in q["content"]
            print(f"  Q{i+1}: {q['content'][:40]}... | Ans: {q['correct_option']} | Img: {has_img}")
        assert len(qs) >= 3, "Expected at least 3 questions in sample_questions.docx"
        # Check that question 3 has the image!
        q3_has_img = any("data:image" in q["content"] for q in qs)
        assert q3_has_img, "Expected image in at least one question of sample_questions.docx"
        print("  ✓ sample_questions.docx passed!")

def test_advanced_cases():
    print("\n=== Generating and Testing Advanced Docx Test Cases ===")
    doc = docx.Document()
    
    # Case 1: Standard Vietnamese question with underline answer (no answer line)
    doc.add_paragraph("Câu 1: Thủ đô của Việt Nam là gì?")
    p_opt_a = doc.add_paragraph()
    r_a = p_opt_a.add_run("A. Hà Nội")
    r_a.underline = True  # Underlined correct answer!
    doc.add_paragraph("B. TP. Hồ Chí Minh")
    doc.add_paragraph("C. Đà Nẵng")
    doc.add_paragraph("D. Cần Thơ")
    
    # Case 2: Multi-option on single line with explicit answer
    doc.add_paragraph("Câu 2. Trong các số sau, số nào là số nguyên tố?")
    doc.add_paragraph("A. 4       B. 6       C. 7       D. 9")
    doc.add_paragraph("Đáp án đúng: C")
    
    # Case 3: Question with an image in an empty paragraph right below question text
    p_q3 = doc.add_paragraph("Câu 3: Quan sát hình ảnh bên dưới và cho biết đây là gì?")
    p_img = doc.add_paragraph() # EMPTY TEXT PARAGRAPH WITH IMAGE
    p_img.add_run().add_picture("test.png")
    
    doc.add_paragraph("A. Điốt")
    doc.add_paragraph("B. Transistor")
    doc.add_paragraph("C. Tụ điện")
    doc.add_paragraph("D. Cuộn cảm")
    doc.add_paragraph("Key: B")
    
    # Case 4: Asterisk marked correct option
    doc.add_paragraph("4. Phương tiện nào sau đây chạy bằng điện?")
    doc.add_paragraph("A. Xe máy xăng")
    doc.add_paragraph("*B. Xe điện VinFast")
    doc.add_paragraph("C. Xe tải diesel")
    doc.add_paragraph("D. Tàu hỏa hơi nước")
    
    # Case 5: Table inside question
    doc.add_paragraph("Câu 5: Dựa vào bảng thông số kỹ thuật sau:")
    tbl = doc.add_table(rows=2, cols=2)
    tbl.rows[0].cells[0].text = "Điện áp"
    tbl.rows[0].cells[1].text = "220V"
    tbl.rows[1].cells[0].text = "Công suất"
    tbl.rows[1].cells[1].text = "1000W"
    
    p_sub = doc.add_paragraph("Hỏi cường độ dòng điện định mức xấp xỉ bao nhiêu?")
    doc.add_paragraph("A. 1.5 A")
    doc.add_paragraph("B. 4.55 A")
    doc.add_paragraph("C. 10 A")
    doc.add_paragraph("D. 22 A")
    doc.add_paragraph("Hướng dẫn giải: chọn B")

    # Case 6: Bold option detection (single line, Option B bold, no answer line)
    doc.add_paragraph("Câu 6: Đơn vị của điện trở là gì?")
    doc.add_paragraph("A. Vôn (V)")
    p6_b = doc.add_paragraph()
    r6_b = p6_b.add_run("B. Ôm (Ω)")
    r6_b.bold = True
    doc.add_paragraph("C. Ampe (A)")
    doc.add_paragraph("D. Oát (W)")

    # Case 7: Single line multi-option underline with tricky text ('C' in 'Cần Thơ')
    doc.add_paragraph("Câu 7: Tỉnh nào sau đây thuộc miền Tây Nam Bộ?")
    p7_opt = doc.add_paragraph()
    p7_opt.add_run("A. Hà Nội    B. Hải Phòng    C. Đà Nẵng    ")
    r7_d = p7_opt.add_run("D. Cần Thơ")
    r7_d.underline = True

    # Case 8: Image reuse across questions (uses test.png again)
    doc.add_paragraph("Câu 8: Hình bên dưới có cùng biểu tượng với câu 3 không?")
    p8_img = doc.add_paragraph()
    p8_img.add_run().add_picture("test.png")
    doc.add_paragraph("A. Có")
    doc.add_paragraph("B. Không")
    doc.add_paragraph("Đáp án: A")

    # Case 9: Word Equation (<m:oMath>)
    p9 = doc.add_paragraph("Câu 9: Nghiệm của phương trình ")
    omath_xml = '<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:sSup><m:e><m:r><m:t>x</m:t></m:r></m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup><m:r><m:t> - 4 = 0</m:t></m:r></m:oMath>'
    p9._p.append(parse_xml(omath_xml))
    p9.add_run(" là bao nhiêu?")
    doc.add_paragraph("A. x = ±2")
    doc.add_paragraph("B. x = 2")
    doc.add_paragraph("C. x = 4")
    doc.add_paragraph("D. Vô nghiệm")
    doc.add_paragraph("Đáp án: A")

    # Case 10: Hyperlink (<w:hyperlink>)
    p10 = doc.add_paragraph("Câu 10: Hãy truy cập ")
    hlink_xml = '<w:hyperlink xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rIdH"><w:r><w:t>https://fit.hcmute.edu.vn</w:t></w:r></w:hyperlink>'
    p10._p.append(parse_xml(hlink_xml))
    p10.add_run(" và trả lời câu hỏi.")
    doc.add_paragraph("A. Đáp án 1")
    doc.add_paragraph("B. Đáp án 2")
    doc.add_paragraph("Đáp án: A")
    
    # Case 11: Question description containing bullet points (- and •) followed by A, B, C, D
    doc.add_paragraph("Câu 11: Cho các nhận định sau trong an toàn điện:")
    doc.add_paragraph("- Trường hợp 1: dòng điện quá tải gây nóng dây dẫn")
    doc.add_paragraph("• Trường hợp 2: ngắn mạch gây phóng hồ quang")
    doc.add_paragraph("A. Nhận định 1 đúng")
    doc.add_paragraph("B. Nhận định 2 đúng")
    doc.add_paragraph("C. Cả hai nhận định đều đúng")
    doc.add_paragraph("D. Cả hai nhận định đều sai")
    doc.add_paragraph("Đáp án: C")

    doc.save("test_advanced.docx")
        
    qs = parse_docx_questions("test_advanced.docx")
    print(f"Parsed {len(qs)} questions from test_advanced.docx")
    assert len(qs) == 11, f"Expected 11 questions, got {len(qs)}"
    
    # Verify Case 1: Underlined option A
    assert qs[0]["correct_option"] == "A", f"Q1 expected A, got {qs[0]['correct_option']}"
    print("  ✓ Case 1 (Underline detection): Correct ->", qs[0]["correct_option"])
    
    # Verify Case 2: Multi-option on one line
    assert qs[1]["correct_option"] == "C", f"Q2 expected C, got {qs[1]['correct_option']}"
    assert qs[1]["option_a"] == "4" and qs[1]["option_c"] == "7"
    print("  ✓ Case 2 (Multi-option line): Correct ->", qs[1]["correct_option"])
    
    # Verify Case 3: Empty paragraph with image attached to question
    assert qs[2]["correct_option"] == "B", f"Q3 expected B, got {qs[2]['correct_option']}"
    assert "data:image" in qs[2]["content"], "Q3 expected image in content"
    print("  ✓ Case 3 (Image in empty paragraph): Correct & Image preserved ->", qs[2]["correct_option"])
    
    # Verify Case 4: Asterisk marked option
    assert qs[3]["correct_option"] == "B", f"Q4 expected B, got {qs[3]['correct_option']}"
    print("  ✓ Case 4 (Asterisk *B. detection): Correct ->", qs[3]["correct_option"])
    
    # Verify Case 5: Table inside question
    assert qs[4]["correct_option"] == "B", f"Q5 expected B, got {qs[4]['correct_option']}"
    assert "<table" in qs[4]["content"], "Q5 expected table in content"
    print("  ✓ Case 5 (Table in question): Correct & Table preserved ->", qs[4]["correct_option"])

    # Verify Case 6: Bold option detection
    assert qs[5]["correct_option"] == "B", f"Q6 expected B, got {qs[5]['correct_option']}"
    assert "<strong>" in qs[5]["option_b"], "Q6 option B should contain bold HTML"
    print("  ✓ Case 6 (Bold option detection): Correct ->", qs[5]["correct_option"])

    # Verify Case 7: Tricky multi-option line with D underlined
    assert qs[6]["correct_option"] == "D", f"Q7 expected D, got {qs[6]['correct_option']}"
    print("  ✓ Case 7 (Multi-option underline offset match): Correct ->", qs[6]["correct_option"])

    # Verify Case 8: Image reuse across questions
    assert "data:image" in qs[7]["content"], "Q8 expected reused image in content"
    assert qs[7]["correct_option"] == "A", f"Q8 expected A, got {qs[7]['correct_option']}"
    print("  ✓ Case 8 (Image reuse across questions): Correct & Reused image preserved ->", qs[7]["correct_option"])

    # Verify Case 9: Word Equation (<m:oMath>)
    assert "x<sup>2</sup>" in qs[8]["content"] or "math-expr" in qs[8]["content"], f"Q9 math equation missing in: {qs[8]['content']}"
    assert qs[8]["correct_option"] == "A", f"Q9 expected A, got {qs[8]['correct_option']}"
    print("  ✓ Case 9 (OMML Math Equation): Correct & Equation HTML preserved ->", qs[8]["correct_option"])

    # Verify Case 10: Hyperlink (<w:hyperlink>)
    assert "fit.hcmute.edu.vn" in qs[9]["content"], f"Q10 hyperlink missing in: {qs[9]['content']}"
    assert qs[9]["correct_option"] == "A", f"Q10 expected A, got {qs[9]['correct_option']}"
    print("  ✓ Case 10 (Hyperlink preservation): Correct & Link preserved ->", qs[9]["correct_option"])

    # Verify Case 11: Bullet points in description must NOT overwrite options A-D!
    assert qs[10]["question_type"] == "multiple_choice", f"Q11 must be multiple_choice, got {qs[10]['question_type']}"
    assert qs[10]["option_a"] == "Nhận định 1 đúng", f"Q11 option A wrong: {qs[10]['option_a']}"
    assert qs[10]["option_b"] == "Nhận định 2 đúng", f"Q11 option B wrong: {qs[10]['option_b']}"
    assert qs[10]["option_c"] == "Cả hai nhận định đều đúng", f"Q11 option C wrong: {qs[10]['option_c']}"
    assert qs[10]["option_d"] == "Cả hai nhận định đều sai", f"Q11 option D wrong: {qs[10]['option_d']}"
    assert qs[10]["correct_option"] == "C", f"Q11 correct_option wrong: {qs[10]['correct_option']}"
    print("  ✓ Case 11 (Bullets in description do not overwrite options): Correct ->", qs[10]["correct_option"])
    
    print("\n🎉 ALL 11 ADVANCED PARSER TESTS PASSED 100%!")

if __name__ == "__main__":
    test_existing_files()
    test_advanced_cases()
