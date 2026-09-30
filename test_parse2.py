import docx

def test():
    doc = docx.Document("test_with_img.docx")
    for p in doc.paragraphs:
        print(f"Para: {p.text}")
        for run in p.runs:
            print("Run:", run.text)
            drawings = run._r.xpath('.//w:drawing')
            if drawings:
                print("Found drawing!")
                for blip in run._r.xpath('.//a:blip/@r:embed'):
                    print("Blip:", blip)
                    image_part = doc.part.related_parts[blip]
                    print("Image part:", type(image_part), len(image_part.blob))
test()
