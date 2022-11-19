from docx import Document
from python_docx_replace import docx_remove_table, docx_replace, docx_blocks


def manual_test():
    doc = Document("test/hello.docx")

    docx_replace(doc, name="Ivan Bicalho")
    docx_blocks(doc, block=True)
    docx_remove_table(doc, 0)

    doc.save("test/hello2.docx")

if __name__ == "__main__":
    manual_test()