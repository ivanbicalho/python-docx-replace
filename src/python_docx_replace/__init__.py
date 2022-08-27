from typing import Dict

__all__ = ["docx_replace", "docx_handle_blocks"]


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
        for p in _get_all_paragraphs(doc):
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
        for p in _get_all_paragraphs(doc):
            if initial in p.text:
                _replace_blocks(p, initial, end, keep_block)


def _replace_blocks(p, initial, end, keep_block):
    max_retries_replace_a_key = 100  # to avoid infinite loop, this value is set
    current = 0

    while initial in p.text:  # if the key appears more than once in the paragraph, it will replaced all
        if end not in p.text:
            raise EndBlockNotFound(initial, end)
        if current >= max_retries_replace_a_key:
            raise MaxRetriesReached(max_retries_replace_a_key, initial)

        changer = RunBlocksRemoval(p, initial, end, keep_block)
        changer.replace()
        current += 1


def _get_all_paragraphs(doc):
    paragraphs = list(doc.paragraphs)
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    paragraphs.append(paragraph)
    return paragraphs


def _simple_replace(p, key, value):
    """Try to replace a key in the paragraph runs, simpler alternative"""
    for run in p.runs:
        if key in run.text:
            run.text = run.text.replace(f"${{{key}}}", value)


def _complex_replace(p, key, value):
    """Complex alternative, which check all broken items inside the runs"""
    max_retries_replace_a_key = 100  # to avoid infinite loop, this value is set
    current = 0

    while key in p.text:  # if the key appears more than once in the paragraph, it will replaced all
        if current >= max_retries_replace_a_key:
            raise MaxRetriesReached(max_retries_replace_a_key, key)

        changer = RunTextChanger(p, key, value)
        changer.replace()
        current += 1


class RunTextChanger:
    def __init__(self, p, key, value):
        self.p = p
        self.key = key
        self.value = value
        self.run_text = ""
        self.runs_indexes = []
        self.run_char_indexes = []
        self.runs_to_change = {}

    def _initialize(self):
        run_index = 0
        for run in self.p.runs:
            self.run_text += run.text
            self.runs_indexes += [run_index for _ in run.text]
            self.run_char_indexes += [char_index for char_index, char in enumerate(run.text)]
            run_index += 1

    def replace(self):
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
    def __init__(self, p, initial, end, keep_block):
        self.p = p
        self.initial = initial
        self.end = end
        self.keep_block = keep_block
        self.run_text = ""
        self.runs_indexes = []
        self.run_char_indexes = []
        self.runs_to_change = {}

    def _initialize(self):
        run_index = 0
        for run in self.p.runs:
            self.run_text += run.text
            self.runs_indexes += [run_index for _ in run.text]
            self.run_char_indexes += [char_index for char_index, char in enumerate(run.text)]
            run_index += 1

    def replace(self):
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


class MaxRetriesReached(Exception):
    def __init__(self, max, key):
        super().__init__(
            f"Max of {max} retries was reached when replacing the key '{key}' in the same paragraph. It can indicates that the system was in loop or you have more than {max} keys '{key}' in the same paragraph."
        )


class EndBlockNotFound(Exception):
    def __init__(self, initial, end):
        super().__init__(
            f"The initial block key {initial} was found, but the end key {end} wasn't found IN THE SAME PARAGRAPH. In this version, this replacer can only handle blocks in the same paragraph. Check your word document and make sure you put the initial and end block keys in the same paragraph."
        )


class InverseInitialEndBlock(Exception):
    def __init__(self, initial, end):
        super().__init__(
            f"The end block {end} appeared before the initial block {initial}. Make sure you put the initial block first."
        )
