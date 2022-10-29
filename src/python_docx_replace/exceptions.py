class MaxRetriesReached(Exception):
    def __init__(self, max, key) -> None:
        super().__init__(
            f"Max of {max} retries was reached when replacing the key '{key}' in the same paragraph. It can indicates that the system was in loop or you have more than {max} keys '{key}' in the same paragraph."
        )


class EndBlockNotFound(Exception):
    def __init__(self, initial, end) -> None:
        super().__init__(
            f"The initial block key {initial} was found, but the end key {end} wasn't found IN THE SAME PARAGRAPH. In this version, this replacer can only handle blocks in the same paragraph. Check your word document and make sure you put the initial and end block keys in the same paragraph."
        )


class InverseInitialEndBlock(Exception):
    def __init__(self, initial, end) -> None:
        super().__init__(
            f"The end block {end} appeared before the initial block {initial}. Make sure you put the initial block first."
        )
