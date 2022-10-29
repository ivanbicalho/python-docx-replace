from typing import Any, Dict, List
from python_docx_replace.blocks import Blocks

from python_docx_replace.exceptions import MaxRetriesReached
from python_docx_replace.replacer import Replacer

__all__ = ["docx_replace", "docx_handle_blocks"]


def docx_replace(doc, **kwargs: Dict[str, str]) -> None:
    """
    Replace all the keys in the word document with the values in the kwargs

    ATTENTION: The required format for the keys inside the Word document is: ${key}

    Example usage:
        Word content = "Hello ${name}, your phone is ${phone}, is that okay?"

        doc = Document("document.docx")  # python-docx dependency

        docx_replace(doc, name="Ivan", phone="+55123456789")

    More information: https://github.com/ivanbicalho/python-docx-replace
    """
    replacer = Replacer()
    for key, value in kwargs.items():
        key = f"${{{key}}}"
        for p in _get_all_paragraphs(doc):
            if key in p.text:
                replacer.simple_replace(p, key, value)
                if key in p.text:
                    replacer.complex_replace(p, key, value)


def docx_handle_blocks(doc, **kwargs: Dict[str, bool]) -> None:
    """
    Keep or remove blocks in the word document

    ATTENTION: The required format for the block keys inside the Word document are: ${i:key} and ${e:key}
        ${i:key} stands for initial block and ${e:key} stands for end block
        Everything inside the block will be removed or not, depending on the kwargs config

    Example usage:
        Word content = "Hello${i:name} Ivan${e:name}, are you okay?"

        doc = Document("document.docx")  # python-docx dependency

        docx_handle_blocks(doc, name=False)
        result = "Hello, are you okay?"

        docx_handle_blocks(doc, name=True)
        result = "Hello Ivan, are you okay?"

    More information: https://github.com/ivanbicalho/python-docx-replace
    """
    blocks = Blocks()
    for key, keep_block in kwargs.items():
        initial = f"${{i:{key}}}"
        end = f"${{e:{key}}}"
        for p in _get_all_paragraphs(doc):
            if initial in p.text:
                blocks.replace_blocks(p, initial, end, keep_block)


def _get_all_paragraphs(doc) -> List[Any]:
    paragraphs = list(doc.paragraphs)
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    paragraphs.append(paragraph)
    return paragraphs
