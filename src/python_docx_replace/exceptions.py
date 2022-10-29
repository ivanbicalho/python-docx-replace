class EndBlockNotFound(Exception):
    def __init__(self, initial, end) -> None:
        super().__init__(
            f"The initial tag '{initial}' was found, but the end tag '{end}' wasn't found. Check your Word document and make sure you set the initial and end tags correctly."
        )


class InverseInitialEndBlock(Exception):
    def __init__(self, initial, end) -> None:
        super().__init__(
            f"The end tag '{end}' appeared before the initial tag '{initial}'. Make sure you put the initial tag first."
        )