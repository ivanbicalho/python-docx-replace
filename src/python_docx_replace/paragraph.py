from typing import Any, List

from python_docx_replace.block_handler import BlockHandler
from python_docx_replace.key_changer import KeyChanger


class Paragraph:
    @staticmethod
    def get_all(doc) -> List[Any]:
        paragraphs = list()
        paragraphs.extend(Paragraph._get_paragraphs(doc))

        for section in doc.sections:
            paragraphs.extend(Paragraph._get_paragraphs(section.header))
            paragraphs.extend(Paragraph._get_paragraphs(section.footer))

        return paragraphs

    @staticmethod
    def _get_paragraphs(item: Any) -> Any:
        yield from item.paragraphs

        # get paragraphs from tables
        for table in item.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        yield paragraph

    def __init__(self, p) -> None:
        self.p = p

    def delete(self) -> None:
        paragraph = self.p._element
        paragraph.getparent().remove(paragraph)
        paragraph._p = paragraph._element = None

    def contains(self, key) -> bool:
        return key in self.p.text

    def startswith(self, key) -> bool:
        return str(self.p.text).strip().startswith(key)

    def endswith(self, key) -> bool:
        return str(self.p.text).strip().endswith(key)

    def replace_key(self, key, value) -> None:
        if key in self.p.text:
            self._simple_replace_key(key, value)
            if key in self.p.text:
                self._complex_replace_key(key, value)

    def replace_block(self, initial, end, keep_block) -> None:
        block_handler = BlockHandler(self.p)
        block_handler.replace(initial, end, keep_block)

    def clear_tag_and_before(self, key, keep_block) -> None:
        block_handler = BlockHandler(self.p)
        block_handler.clear_key_and_before(key, keep_block)

    def clear_tag_and_after(self, key, keep_block) -> None:
        block_handler = BlockHandler(self.p)
        block_handler.clear_key_and_after(key, keep_block)

    def get_text(self) -> str:
        return self.p.text

    def _simple_replace_key(self, key, value) -> None:
        # try to replace a key in the paragraph runs, simpler alternative
        for run in self.p.runs:
            if key in run.text:
                run.text = run.text.replace(key, value)

    def _complex_replace_key(self, key, value) -> None:
        # complex alternative, which check all broken items inside the runs
        while key in self.p.text:
            # if the key appears more than once in the paragraph, it will replaced all
            key_changer = KeyChanger(self.p, key, value)
            key_changer.replace()
