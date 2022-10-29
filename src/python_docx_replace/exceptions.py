class InitialTagNotFound(Exception):
    def __init__(self, initial, end) -> None:
        super().__init__(
            f"The end tag '{end}' was found, but the initial tag '{initial}' wasn't found. "
            "Check your Word document and make sure you set the initial and end tags correctly."
        )


class EndTagNotFound(Exception):
    def __init__(self, initial, end) -> None:
        super().__init__(
            f"The initial tag '{initial}' was found, but the end tag '{end}' wasn't found. "
            "Check your Word document and make sure you set the initial and end tags correctly."
        )


class ReversedInitialEndTags(Exception):
    def __init__(self, initial, end) -> None:
        super().__init__(
            f"The end tag '{end}' appeared before the initial tag '{initial}'. "
            "Make sure you put the initial tag first."
        )
