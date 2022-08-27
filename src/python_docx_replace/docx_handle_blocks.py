from common import MaxRetriesReached, EndBlockNotFound, InverseInitialEndBlock


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
