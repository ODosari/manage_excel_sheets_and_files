from typing import Protocol, BinaryIO

class PasswordUnlocker(Protocol):
    def unlock(self, path: str, password: str) -> BinaryIO: ...
