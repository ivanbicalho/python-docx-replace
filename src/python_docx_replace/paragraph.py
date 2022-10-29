from typing import Any, List
from python_docx_replace.exceptions import EndBlockNotFound, InverseInitialEndBlock, MaxRetriesReached


MAX_RETRIES_REPLACE_A_KEY = 100  # to avoid infinite loop, this value is set


class Paragraph:
    @staticmethod
    def get_all(doc) -> List[Any]:
        paragraphs = list(doc.paragraphs)
        for t in doc.tables:
            for row in t.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        paragraphs.append(paragraph)
        return paragraphs

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
        key = f"${{{key}}}"
        if key in self.p.text:
            self._simple_replace(key, value)
            if key in self.p.text:
                self._complex_replace(key, value)

    def replace_block(self, initial, end, keep_block) -> None:
        current = 0
        while initial in self.p.text:  # if the key appears more than once in the paragraph, it will replaced all
            if end not in self.p.text:
                raise EndBlockNotFound(initial, end)
            if current >= MAX_RETRIES_REPLACE_A_KEY:
                raise MaxRetriesReached(MAX_RETRIES_REPLACE_A_KEY, initial)

            changer = RunBlocksRemoval(self.p, initial, end, keep_block)
            changer.replace()
            current += 1

    def _simple_replace(self, key, value) -> None:
        """
        Try to replace a key in the paragraph runs, simpler alternative
        """
        for run in self.p.runs:
            if key in run.text:
                run.text = run.text.replace(key, value)

    def _complex_replace(self, key, value) -> None:
        """
        Complex alternative, which check all broken items inside the runs
        """
        current = 0

        while key in self.p.text:  # if the key appears more than once in the paragraph, it will replaced all
            if current >= MAX_RETRIES_REPLACE_A_KEY:
                raise MaxRetriesReached(MAX_RETRIES_REPLACE_A_KEY, key)

            changer = RunTextChanger(self.p, key, value)
            changer.replace()
            current += 1


class RunTextChanger:
    def __init__(self, p, key, value) -> None:
        self.p = p
        self.key = key
        self.value = value
        self.run_text = ""
        self.runs_indexes = []
        self.run_char_indexes = []
        self.runs_to_change = {}

    def _initialize(self) -> None:
        run_index = 0
        for run in self.p.runs:
            self.run_text += run.text
            self.runs_indexes += [run_index for _ in run.text]
            self.run_char_indexes += [char_index for char_index, char in enumerate(run.text)]
            run_index += 1

    def replace(self) -> None:
        self._initialize()
        parsed_key_length = len(self.key)
        index_to_replace = self.run_text.find(self.key)

        for i in range(parsed_key_length):
            index = index_to_replace + i
            run_index = self.runs_indexes[index]
            run = self.p.runs[run_index]
            run_char_index = self.run_char_indexes[index]

            if not self.runs_to_change.get(run_index):
                self.runs_to_change[run_index] = [char for char_index, char in enumerate(run.text)]

            run_to_change = self.runs_to_change.get(run_index)
            if index == index_to_replace:
                run_to_change[run_char_index] = self.value
            else:
                run_to_change[run_char_index] = ""

        # make the real replace
        for index, text in self.runs_to_change.items():
            run = self.p.runs[index]
            run.text = "".join(text)


class RunBlocksRemoval:
    def __init__(self, p, initial, end, keep_block) -> None:
        self.p = p
        self.initial = initial
        self.end = end
        self.keep_block = keep_block
        self.run_text = ""
        self.runs_indexes = []
        self.run_char_indexes = []
        self.runs_to_change = {}

    def _initialize(self) -> None:
        run_index = 0
        for run in self.p.runs:
            self.run_text += run.text
            self.runs_indexes += [run_index for _ in run.text]
            self.run_char_indexes += [char_index for char_index, char in enumerate(run.text)]
            run_index += 1

    def replace(self) -> None:
        self._initialize()
        key_length = len(self.initial)  # initial and end have the same length

        initial_index = self.run_text.find(self.initial)
        end_index = self.run_text.find(self.end)

        if end_index < initial_index:
            raise InverseInitialEndBlock(self.initial, self.end)

        initial_index_plus_key_length = initial_index + key_length
        end_index_plus_key_length = end_index + key_length

        for index in range(initial_index, end_index_plus_key_length):
            run_index = self.runs_indexes[index]
            run = self.p.runs[run_index]
            run_char_index = self.run_char_indexes[index]

            if not self.runs_to_change.get(run_index):
                self.runs_to_change[run_index] = [char for char_index, char in enumerate(run.text)]

            run_to_change = self.runs_to_change.get(run_index)
            if (
                (not self.keep_block)
                or (index in range(initial_index, initial_index_plus_key_length))
                or (index in range(end_index, end_index_plus_key_length))
            ):
                run_to_change[run_char_index] = ""

            if index > end_index_plus_key_length:
                break

        # make the real replace
        for index, text in self.runs_to_change.items():
            run = self.p.runs[index]
            run.text = "".join(text)
