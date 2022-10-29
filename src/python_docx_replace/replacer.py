from python_docx_replace.exceptions import MaxRetriesReached

MAX_RETRIES_REPLACE_A_KEY = 100  # to avoid infinite loop, this value is set


class Replacer:
    def simple_replace(self, p, key, value) -> None:
        """
        Try to replace a key in the paragraph runs, simpler alternative
        """
        for run in p.runs:
            if key in run.text:
                run.text = run.text.replace(f"${{{key}}}", value)

    def complex_replace(self, p, key, value) -> None:
        """
        Complex alternative, which check all broken items inside the runs
        """
        current = 0

        while key in p.text:  # if the key appears more than once in the paragraph, it will replaced all
            if current >= MAX_RETRIES_REPLACE_A_KEY:
                raise MaxRetriesReached(MAX_RETRIES_REPLACE_A_KEY, key)

            changer = RunTextChanger(p, key, value)
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
