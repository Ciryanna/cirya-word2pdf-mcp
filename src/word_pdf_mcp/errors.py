class ConversionError(Exception):
    """An actionable error safe to return to an MCP client."""

    def __init__(self, code: str, message: str):
        self.code = code
        self.message = message
        super().__init__(f"{code}: {message}")
