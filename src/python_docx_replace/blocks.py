from python_docx_replace.exceptions import EndBlockNotFound, InverseInitialEndBlock, MaxRetriesReached
from python_docx_replace.docx_extensions import delete_paragraph, get_all_paragraphs
from python_docx_replace.replacer import Replacer

MAX_RETRIES_REPLACE_A_KEY = 100  # to avoid infinite loop, this value is set


class Blocks:
    def __init__(self) -> None:
        self.replacer = Replacer()

    def replace_blocks(self, doc, initial, end, keep_block) -> None:
        current = 0
        for p in get_all_paragraphs(doc):
            if initial in p.text:
                if end in p.text:
                    changer = RunBlocksRemoval(p, initial, end, keep_block)
                    changer.replace()
                else:
                    if str(p.text).startswith(initial):
                        p.clear()
                    else:
                        # clear only the initial key in the paragraph
                        self.replacer.complex_replace(p, initial, "")
                        

    def replace(self, doc, initial, end, keep_block) -> bool:
        look_for_initial = True
        for p in get_all_paragraphs(doc):
            if look_for_initial:
                if initial in p.text:
                    look_for_initial = False
                    if end in p.text:
                        changer = RunBlocksRemoval(p, initial, end, keep_block)
                        changer.replace()
                        return True
                    else:
                        if str(p.text).startswith(initial):
                            delete_paragraph(p)
                            continue
                        else:
                            # replace key initial by "" + clear everything until the end of the paragraph
                            continue
            else:
                if end in p.text:
                    if str(p.text).endswith(end):
                        delete_paragraph(p)
                        return True
                    else:
                        # replace key end by "" + clear everything before the key
                        return True
        if look_for_initial:
            return False
        else:
            raise EndBlockNotFound(initial, end)


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
