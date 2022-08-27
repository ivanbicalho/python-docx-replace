from typing import Dict
from common import get_all_paragraphs
from docx_handle_blocks import _replace_blocks
from docx_replace import _simple_replace, _complex_replace


def docx_replace(doc, **kwargs: Dict[str, str]):
    """
    Replace all the keys in the word document with the values in the kwargs

    ATTENTION: The required format for the keys inside the Word document is: ${key}

    Example usage:
        Word content = "Hello ${name}, your phone is ${phone}, is that okay?"

        doc = Document("document.docx")  # python-docx dependency

        docx_replace(doc, name="Ivan", phone="+55123456789")

    More information: https://github.com/ivanbicalho/python-docx-replace
    """
    for key, value in kwargs.items():
        key = f"${{{key}}}"
        for p in get_all_paragraphs(doc):
            if key in p.text:
                _simple_replace(p, key, value)
                if key in p.text:
                    _complex_replace(p, key, value)


def docx_handle_blocks(doc, **kwargs: Dict[str, bool]):
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
    for key, keep_block in kwargs.items():
        initial = f"${{i:{key}}}"
        end = f"${{e:{key}}}"
        for p in get_all_paragraphs(doc):
            if initial in p.text:
                _replace_blocks(p, initial, end, keep_block)


