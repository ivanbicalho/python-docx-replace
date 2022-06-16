from docx import Document
from python_docx_replace import docx_replace


def manual_test():
    doc = Document("hello.docx")

    docx_replace(doc, name="Ivan")

    doc.save("hello2.docx")


if __name__ == "__main__":
    manual_test()
