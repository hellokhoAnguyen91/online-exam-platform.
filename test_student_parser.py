import unittest
import io
import datetime
import pandas as pd
import docx
from student_parser import parse_student_file, parse_vietnamese_date

class TestStudentParser(unittest.TestCase):
    def test_standard_excel(self):
        df = pd.DataFrame({
            "MSSV": ["21110001", "21110002"],
            "Họ và tên": ["Nguyễn Văn An", "Trần Thị Bình"],
            "Ngày sinh": ["15/01/2003", "20/02/2003"]
        })
        buf = io.BytesIO()
        df.to_excel(buf, index=False)
        res = parse_student_file(buf.getvalue(), "students.xlsx")
        self.assertEqual(res["total_parsed"], 2)
        self.assertEqual(res["students"][0]["mssv"], "21110001")
        self.assertEqual(res["students"][0]["fullname"], "Nguyễn Văn An")
        self.assertEqual(res["students"][0]["dob"], "2003-01-15")

    def test_preamble_rows_and_ignored_columns(self):
        # 3 rows of preamble header, then headers, then extra columns STT, Lớp, Phòng thi, Ghi chú
        rows = [
            ["TRƯỜNG ĐẠI HỌC SƯ PHẠM KỸ THUẬT TP.HCM", "", "", "", "", "", ""],
            ["DANH SÁCH THÍ SINH DỰ THI CUỐI KỲ", "", "", "", "", "", ""],
            ["Môn thi: An toàn điện", "", "", "", "", "", ""],
            ["STT", "Mã SV", "Họ và chữ lót", "Tên", "Ngày sinh", "Lớp", "Phòng thi"],
            [1, "21110005", "Vũ Thị", "Hoa", "25/12/2003", "21110CLC", "A1-302"],
            [2, "21110006", "Đặng Quang", "Huy", "05-08-2003", "21110CLC", "A1-302"],
            ["Tổng cộng: 2 thí sinh", "", "", "", "", "", ""]
        ]
        df = pd.DataFrame(rows)
        buf = io.BytesIO()
        df.to_excel(buf, index=False, header=False)
        res = parse_student_file(buf.getvalue(), "danh_sach_phong_thi.xlsx")
        
        self.assertEqual(res["total_parsed"], 2)
        self.assertEqual(res["students"][0]["mssv"], "21110005")
        self.assertEqual(res["students"][0]["fullname"], "Vũ Thị Hoa")
        self.assertEqual(res["students"][0]["dob"], "2003-12-25")
        self.assertEqual(res["students"][1]["mssv"], "21110006")
        self.assertEqual(res["students"][1]["fullname"], "Đặng Quang Huy")
        self.assertEqual(res["students"][1]["dob"], "2003-08-05")
        
        # Verify ignored columns
        self.assertIn("STT", res["ignored_columns"])
        self.assertEqual(res["students"][0]["class_name"], "21110CLC")
        self.assertIn("Phòng thi", res["ignored_columns"])

    def test_excel_serial_date(self):
        # Excel date serial 37636 is 2003-01-15
        df = pd.DataFrame({
            "Mã số sinh viên": ["21110001.0"],
            "Họ tên sinh viên": ["Nguyễn Văn An"],
            "DOB": [37636]
        })
        buf = io.BytesIO()
        df.to_excel(buf, index=False)
        res = parse_student_file(buf.getvalue(), "students.xlsx")
        self.assertEqual(res["students"][0]["mssv"], "21110001")
        self.assertEqual(res["students"][0]["dob"], "2003-01-15")

    def test_csv_file(self):
        csv_data = "STT,Mã sinh viên,Họ và tên,Ngày tháng năm sinh,Ghi chú\n1,21110099,Hoàng Văn Em,19/09/2003,Đủ điều kiện\n"
        res = parse_student_file(csv_data.encode("utf-8-sig"), "students.csv")
        self.assertEqual(res["total_parsed"], 1)
        self.assertEqual(res["students"][0]["mssv"], "21110099")
        self.assertEqual(res["students"][0]["fullname"], "Hoàng Văn Em")
        self.assertEqual(res["students"][0]["dob"], "2003-09-19")
        self.assertIn("Ghi chú", res["ignored_columns"])

    def test_docx_table_file(self):
        doc = docx.Document()
        doc.add_heading("Danh sách sinh viên", 0)
        table = doc.add_table(rows=1, cols=4)
        hdr_cells = table.rows[0].cells
        hdr_cells[0].text = "STT"
        hdr_cells[1].text = "MSSV"
        hdr_cells[2].text = "Họ và tên"
        hdr_cells[3].text = "Ngày sinh"
        
        row_cells = table.add_row().cells
        row_cells[0].text = "1"
        row_cells[1].text = "21110088"
        row_cells[2].text = "Phạm Thị Lan"
        row_cells[3].text = "01/01/2003"
        
        buf = io.BytesIO()
        doc.save(buf)
        res = parse_student_file(buf.getvalue(), "students.docx")
        self.assertEqual(res["total_parsed"], 1)
        self.assertEqual(res["students"][0]["mssv"], "21110088")
        self.assertEqual(res["students"][0]["fullname"], "Phạm Thị Lan")
        self.assertEqual(res["students"][0]["dob"], "2003-01-01")

    def test_merged_header_split_name(self):
        # 2-row merged header common in HCMUTE
        # Row 1: STT, Mã SV, Họ và tên, (merged/empty), Ngày sinh, Lớp
        # Row 2: (empty), (empty), Họ lót, Tên, (empty), (empty)
        # Row 3: 1, 09123456, Nguyễn Văn, An, 15/01/2003, 21110CLC
        rows = [
            ["STT", "Mã SV", "Họ và tên", "", "Ngày sinh", "Lớp"],
            ["", "", "Họ và chữ lót", "Tên", "", ""],
            ["1", "09123456", "Nguyễn Văn", "An", "15/01/2003", "21110CLC"],
            ["2", "08123457", "Trần Thị", "Bình", "20/02/2003", "21110CLC"]
        ]
        df = pd.DataFrame(rows)
        buf = io.BytesIO()
        df.to_excel(buf, index=False, header=False)
        res = parse_student_file(buf.getvalue(), "students_merged.xlsx")
        self.assertEqual(res["total_parsed"], 2)
        # Leading zero must be preserved!
        self.assertEqual(res["students"][0]["mssv"], "09123456")
        # Full name must combine Họ and Tên without truncation!
        self.assertEqual(res["students"][0]["fullname"], "Nguyễn Văn An")
        self.assertEqual(res["students"][1]["mssv"], "08123457")
        self.assertEqual(res["students"][1]["fullname"], "Trần Thị Bình")

    def test_leading_zero_mssv_csv(self):
        csv_data = "MSSV,Họ và tên,Ngày sinh\n01234567,Lê Văn Cường,25/08/2003\n"
        res = parse_student_file(csv_data.encode("utf-8-sig"), "students_zero.csv")
        self.assertEqual(res["students"][0]["mssv"], "01234567")
        self.assertEqual(res["students"][0]["fullname"], "Lê Văn Cường")

    def test_empty_or_invalid_dob(self):
        # Empty DOB or invalid format should NOT be filled with 2003-01-01
        csv_data = "MSSV,Họ và tên,Ngày sinh\n21110001,Trần Văn An,\n21110002,Lê Thị Bình,invalid_date\n"
        res = parse_student_file(csv_data.encode("utf-8-sig"), "students_empty_dob.csv")
        self.assertEqual(res["students"][0]["dob"], "")
    def test_email_column_as_mssv(self):
        # When file only has STT, Họ và tên, Email (no dedicated MSSV column)
        rows = [
            ["STT", "Họ và tên", "Email"],
            ["1", "Huỳnh Võ Phúc An", "23150014@student.hcmute.edu.vn"],
            ["2", "Bùi Nguyễn Ngọc Anh", "23150015@student.hcmute.edu.vn"]
        ]
        df = pd.DataFrame(rows)
        buf = io.BytesIO()
        df.to_excel(buf, index=False, header=False)
        res = parse_student_file(buf.getvalue(), "students_email.xlsx")
        self.assertEqual(res["total_parsed"], 2)
        self.assertEqual(res["students"][0]["mssv"], "23150014")
        self.assertEqual(res["students"][0]["fullname"], "Huỳnh Võ Phúc An")
        self.assertEqual(res["students"][1]["mssv"], "23150015")
        self.assertEqual(res["students"][1]["fullname"], "Bùi Nguyễn Ngọc Anh")

    def test_sv_prefix_normalization(self):
        csv_data = "STT,Username,Họ và tên,Ngày sinh\n1,sv23150014,Huỳnh Võ Phúc An,26/05/2005\n"
        res = parse_student_file(csv_data.encode("utf-8-sig"), "students_vip.csv")
        self.assertEqual(res["students"][0]["mssv"], "23150014")
        self.assertEqual(res["students"][0]["dob"], "2005-05-26")

    def test_data_driven_heuristic_detection(self):
        # File has non-standard header names: Col_A, Col_B
        rows = [
            ["Mã số định danh", "Danh tính thí sinh"],
            ["23150014", "Huỳnh Võ Phúc An"],
            ["23150015", "Bùi Nguyễn Ngọc Anh"]
        ]
        df = pd.DataFrame(rows)
        buf = io.BytesIO()
        df.to_excel(buf, index=False, header=False)
        res = parse_student_file(buf.getvalue(), "students_heuristic.xlsx")
        self.assertEqual(res["total_parsed"], 2)
        self.assertEqual(res["students"][0]["mssv"], "23150014")
        self.assertEqual(res["students"][0]["fullname"], "Huỳnh Võ Phúc An")

if __name__ == "__main__":
    unittest.main()

