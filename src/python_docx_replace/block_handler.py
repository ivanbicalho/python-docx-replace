from typing import Dict

from python_docx_replace.exceptions import ReversedInitialEndTags


class BlockHandler:
    def __init__(self, p) -> None:
        self.p = p
        self.run_text = ""
        self.runs_indexes = []
        self.run_char_indexes = []

        run_index = 0
        for run in self.p.runs:
            self.run_text += run.text
            self.runs_indexes += [run_index for _ in run.text]
            self.run_char_indexes += [char_index for char_index, char in enumerate(run.text)]
            run_index += 1

    def replace(self, initial: str, end: str, keep_block: bool) -> None:
        initial_index = self.run_text.find(initial)
        initial_length = len(initial)
        initial_index_plus_key_length = initial_index + initial_length
        end_index = self.run_text.find(end)
        end_length = len(end)
        end_index_plus_key_length = end_index + end_length
        runs_to_change: Dict = {}

        if end_index < initial_index:
            raise ReversedInitialEndTags(initial, end)

        for index in range(initial_index, end_index_plus_key_length):
            run_index = self.runs_indexes[index]
            run = self.p.runs[run_index]
            run_char_index = self.run_char_indexes[index]

            if not runs_to_change.get(run_index):
                runs_to_change[run_index] = [char for char_index, char in enumerate(run.text)]

            run_to_change: Dict = runs_to_change.get(run_index)  # type: ignore[assignment]
            if (
                (not keep_block)
                or (index in range(initial_index, initial_index_plus_key_length))
                or (index in range(end_index, end_index_plus_key_length))
            ):
                run_to_change[run_char_index] = ""

        self._real_replace(runs_to_change)

    def clear_key_and_after(self, key: str, keep_block: bool) -> None:
        key_index = self.run_text.find(key)
        key_length = len(key)
        key_index_plus_key_length = key_index + key_length
        runs_to_change: Dict = {}

        for index in range(key_index, len(self.run_text)):
            run_index = self.runs_indexes[index]
            run = self.p.runs[run_index]
            run_char_index = self.run_char_indexes[index]

            if not runs_to_change.get(run_index):
                runs_to_change[run_index] = [char for char_index, char in enumerate(run.text)]

            run_to_change: Dict = runs_to_change.get(run_index)  # type: ignore[assignment]
            if (not keep_block) or (index in range(key_index, key_index_plus_key_length)):
                run_to_change[run_char_index] = ""

        self._real_replace(runs_to_change)

    def clear_key_and_before(self, key: str, keep_block: bool) -> None:
        key_index = self.run_text.find(key)
        key_length = len(key)
        key_index_plus_key_length = key_index + key_length
        runs_to_change: Dict = {}

        for index in range(0, key_index_plus_key_length):
            run_index = self.runs_indexes[index]
            run = self.p.runs[run_index]
            run_char_index = self.run_char_indexes[index]

            if not runs_to_change.get(run_index):
                runs_to_change[run_index] = [char for char_index, char in enumerate(run.text)]

            run_to_change: Dict = runs_to_change.get(run_index)  # type: ignore[assignment]
            if (not keep_block) or (index in range(key_index, key_index_plus_key_length)):
                run_to_change[run_char_index] = ""

        self._real_replace(runs_to_change)

    def _real_replace(self, runs_to_change: dict):
        # make the real replace
        for index, text in runs_to_change.items():
            run = self.p.runs[index]
            run.text = "".join(text)
