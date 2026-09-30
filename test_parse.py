import docx

def test():
    doc = docx.Document("sample_questions.docx")
    for p in doc.paragraphs:
        print(f"Para: {p.text}")
        for run in p.runs:
            drawings = run._r.xpath('w:drawing')
            if drawings:
                print("Found drawing!")
test()
