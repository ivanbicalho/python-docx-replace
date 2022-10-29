from docx import Document
from python_docx_replace import docx_replace, docx_handle_blocks


def manual_test():
    doc = Document("hello.docx")

    #docx_replace(doc, name="IVAN RIBEIRO BICALHO")
    docx_handle_blocks(doc, block=True)

    doc.save("hello2.docx")


if __name__ == "__main__":
    manual_test()
