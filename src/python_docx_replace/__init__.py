from typing import Any

from python_docx_replace.exceptions import EndTagNotFound, InitialTagNotFound
from python_docx_replace.paragraph import Paragraph

__all__ = ["docx_replace", "docx_handle_blocks"]


def docx_replace(doc, **kwargs: str) -> None:
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
        for p in Paragraph.get_all(doc):
            paragraph = Paragraph(p)
            paragraph.replace_key(key, value)


def docx_handle_blocks(doc: Any, **kwargs: bool) -> None:
    """
    Keep or remove blocks in the word document

    ATTENTION: The required format for the block keys inside the Word document are: <key> and </key>
        <key> stands for initial block
        </key> stands for end block
        Everything inside the block will be removed or not, depending on the configuration

    Example usage:
        Word content = "Hello<name> Ivan</name>, are you okay?"

        doc = Document("document.docx")  # python-docx dependency

        docx_handle_blocks(doc, name=False)
        result = "Hello, are you okay?"

        docx_handle_blocks(doc, name=True)
        result = "Hello Ivan, are you okay?"

    More information: https://github.com/ivanbicalho/python-docx-replace
    """
    for key, keep_block in kwargs.items():
        initial = f"<{key}>"
        end = f"</{key}>"

        result = _handle_blocks(doc, initial, end, keep_block)
        while result:  # if the keys appear more than once, it will replace all
            result = _handle_blocks(doc, initial, end, keep_block)

        _search_for_lost_end_tag(doc, initial, end)


def _handle_blocks(doc: Any, initial: str, end: str, keep_block: bool) -> bool:
    look_for_initial = True
    for p in Paragraph.get_all(doc):
        paragraph = Paragraph(p)
        if look_for_initial:
            if paragraph.contains(initial):
                look_for_initial = False
                if paragraph.contains(end):
                    paragraph.replace_block(initial, end, keep_block)
                    return True
                else:
                    if paragraph.startswith(initial):
                        paragraph.delete()
                        continue
                    else:
                        paragraph.clear_tag_and_after(initial, keep_block)
                        continue
        else:
            if paragraph.contains(end):
                if paragraph.endswith(end):
                    paragraph.delete()
                    return True
                else:
                    paragraph.clear_tag_and_before(end, keep_block)
                    return True
            else:
                if not keep_block:
                    paragraph.delete()
                continue
    if look_for_initial:
        return False
    else:
        raise EndTagNotFound(initial, end)


def _search_for_lost_end_tag(doc: Any, initial: str, end: str) -> None:
    for p in Paragraph.get_all(doc):
        paragraph = Paragraph(p)
        if paragraph.contains(end):
            raise InitialTagNotFound(initial, end)
