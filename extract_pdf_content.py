from pypdf import PdfReader

reader = PdfReader("018-6.pdf")
text = ""
for page in reader.pages:
    text += page.extract_text() + "\n"

print(text)
