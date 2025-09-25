from typing import BinaryIO, Protocol


class PasswordUnlocker(Protocol):
    def unlock(self, path: str, password: str) -> BinaryIO: ...
