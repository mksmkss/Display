from pdf2docx.main import parse

for i in range(0, 11):
    print(i)
    parse(
        f"/Users/masataka/Desktop/Data/Description PDF/description_{i}.pdf",
        f"/Users/masataka/Desktop/Data/Description PDF/description_{i}.docx",
    )
