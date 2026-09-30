import docx
import pandas as pd

def create_sample_docx():
    doc = docx.Document()
    
    doc.add_paragraph("Q: What is the capital of France?")
    doc.add_paragraph("A. London")
    doc.add_paragraph("B. Berlin")
    doc.add_paragraph("C. Paris")
    doc.add_paragraph("D. Madrid")
    doc.add_paragraph("Correct: C")
    
    doc.add_paragraph("Q: What is 2 + 2? Here is a helpful video: https://youtube.com/watch?v=dQw4w9WgXcQ")
    doc.add_paragraph("A. 3")
    doc.add_paragraph("B. 4")
    doc.add_paragraph("C. 5")
    doc.add_paragraph("D. 6")
    doc.add_paragraph("Correct: B")
    
    import base64
    img_data = base64.b64decode('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII=')
    with open('temp.png', 'wb') as f:
        f.write(img_data)
        
    p = doc.add_paragraph("Q: What color is this pixel?")
    p.add_run().add_picture('temp.png')
    doc.add_paragraph("A. Red")
    doc.add_paragraph("B. Black")
    doc.add_paragraph("C. Transparent")
    doc.add_paragraph("D. White")
    doc.add_paragraph("Correct: B")
    
    doc.save("sample_questions.docx")
    import os
    if os.path.exists('temp.png'):
        os.remove('temp.png')

def create_sample_xlsx():
    data = {
        "MSSV": ["101", "102"],
        "FullName": ["Nguyen Van A", "Tran Thi B"],
        "DOB": ["2000-01-01", "2000-12-31"]
    }
    df = pd.DataFrame(data)
    df.to_excel("sample_students.xlsx", index=False)

if __name__ == "__main__":
    create_sample_docx()
    create_sample_xlsx()
