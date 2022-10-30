from typing import Any

from python_docx_replace.exceptions import EndTagNotFound, InitialTagNotFound
from python_docx_replace.paragraph import Paragraph

__all__ = ["docx_replace", "docx_blocks"]


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


def docx_blocks(doc: Any, **kwargs: bool) -> None:
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

        # end tags can exists alone in the document
        # this function just make sure that if it's the case, raise an error
        _search_for_lost_end_tag(doc, initial, end)


def _handle_blocks(doc: Any, initial: str, end: str, keep_block: bool) -> bool:
    look_for_initial = True
    for p in Paragraph.get_all(doc):
        paragraph = Paragraph(p)
        if look_for_initial:
            if paragraph.contains(initial):
                look_for_initial = False  # initial tag found, next search will be for end tag
                if paragraph.contains(end):
                    # if the initial and end tag are in the same paragraph, treat them together
                    paragraph.replace_block(initial, end, keep_block)
                    return True  # block completed, returns
                else:
                    # the current paragraph doesn't have the end tag
                    if paragraph.startswith(initial):
                        # if the paragraph starts with the initial tag, we can delete the entire paragraph,
                        # because the end tag is not here
                        paragraph.delete()
                        continue
                    else:
                        # if the paragraph doesn't start with the initial tag, we cannot delete the entire
                        # paragraph, we have to clear the tag and remove the content right after (if not keep_block)
                        paragraph.clear_tag_and_after(initial, keep_block)
                        continue
        else:
            # we are looking for the end tag as the initial tag was found and treated before
            if paragraph.contains(end):
                # end tag found in this paragraph
                if paragraph.endswith(end):
                    # if the paragraph ends with the end tag, we can delete the entire paragraph
                    paragraph.delete()
                    return True  # block completed, returns
                else:
                    # if the paragraph doesn't end with the end tag, we cannot delete the entire
                    # paragraph, we have to clear the tag and remove the content right before (if not keep_block)
                    paragraph.clear_tag_and_before(end, keep_block)
                    return True  # block completed, returns
            else:
                # paragraph doesn't have the end key, that means there is no tags here. In this case,
                # we can remove the entire paragraph if not keep_block, otherwise do nothing
                if not keep_block:
                    paragraph.delete()
                continue
    if look_for_initial:
        # if the initial tag wasn't found, the block doesn't exist
        return False  # block completed, returns
    else:
        # if the initial tag was found, but not end tag, raise an error
        raise EndTagNotFound(initial, end)


def _search_for_lost_end_tag(doc: Any, initial: str, end: str) -> None:
    for p in Paragraph.get_all(doc):
        paragraph = Paragraph(p)
        if paragraph.contains(end):
            raise InitialTagNotFound(initial, end)
