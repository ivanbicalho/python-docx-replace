import re
from typing import Any, List
from python_docx_replace.exceptions import EndTagNotFound, InitialTagNotFound, TableIndexNotFound
from python_docx_replace.paragraph import Paragraph

__all__ = ["docx_replace", "docx_blocks", "docx_remove_table"]


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
            paragraph.replace_key(key, str(value))


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


def docx_remove_table(doc: Any, index: int) -> None:
    """
    Remove a table from your Word document by index

    Example usage:
        docx_remove_table(doc, 0)  # it will remove the first table

    If the table index wasn't found, an error will be raised

    ATTENTION:
        The table index works exactly like any indexing property. It means if you
        remove an index, it will affect the other indexes. For example, if you want
        to remove the first two tables, you can't do like this:
            - docx_remove_table(doc, 0)
            - docx_remove_table(doc, 1)
        You should instead do like this:
            - docx_remove_table(doc, 0)
            - docx_remove_table(doc, 0)
    """
    try:
        table = doc.tables[index]
        table._element.getparent().remove(table._element)
    except IndexError:
        raise TableIndexNotFound(index, len(doc.tables))


def docx_get_keys(doc: Any) -> List[str]:
    """
    Search for all keys in the Word document and return a list of unique elements

    ATTENTION: The required format for the keys inside the Word document is: ${key}

    For a document with the following content: "Hello ${name}, is your phone ${phone}?"
    Result example: ["name", "phone"]
    """
    result = set()  # unique items
    for p in Paragraph.get_all(doc):
        paragraph = Paragraph(p)
        matches = re.finditer(r"\$\{([^{}]+)\}", paragraph.get_text())
        for match in matches:
            result.add(match.groups()[0])
    return list(result)


def _handle_blocks(doc: Any, initial: str, end: str, keep_block: bool) -> bool:
    # The below process is a little bit complex, so I decided to comment each step
    look_for_initial = True
    for p in Paragraph.get_all(doc):
        paragraph = Paragraph(p)
        if look_for_initial:
            if paragraph.contains(initial):
                look_for_initial = False  # initial tag found, next search will be for the end tag
                if paragraph.contains(end):
                    # if the initial and end tag are in the same paragraph, treat them together
                    paragraph.replace_block(initial, end, keep_block)
                    return True  # block completed, returns
                else:
                    # the current paragraph doesn't have the end tag
                    if paragraph.startswith(initial):
                        # if the paragraph starts with the initial tag, we can clear the tag if is to keep_block
                        # otherwise we can delete the entire paragraph because the end tag is not here
                        if keep_block:
                            paragraph.clear_tag_and_before(initial, keep_block)
                        else:
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
                    # if the paragraph ends with the end tag we can clear the tag if is to keep_block
                    # otherwise we can delete the entire paragraph
                    if keep_block:
                        paragraph.clear_tag_and_after(end, keep_block)
                    else:
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
        # if the initial tag wasn't found, the block doesn't exist in the Word document
        return False  # block completed, returns
    else:
        # if the initial tag was found, but not end tag, raise an error
        raise EndTagNotFound(initial, end)


def _search_for_lost_end_tag(doc: Any, initial: str, end: str) -> None:
    for p in Paragraph.get_all(doc):
        paragraph = Paragraph(p)
        if paragraph.contains(end):
            raise InitialTagNotFound(initial, end)
