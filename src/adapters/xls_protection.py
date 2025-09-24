from typing import BinaryIO
from io import BytesIO
from ._utils import import_optional
from src.core.errors import DecryptionError

def unlock_to_stream(path: str, password: str) -> BinaryIO:
    msoffcrypto = import_optional("msoffcrypto")
    if msoffcrypto is None:
        raise DecryptionError("msoffcrypto-tool is required to open password-protected files.")
    try:
        bio = BytesIO()
        with open(path, "rb") as f:
            office = msoffcrypto.OfficeFile(f)
            office.load_key(password=password)
            office.decrypt(bio)
        bio.seek(0)
        return bio
    except Exception as e:
        raise DecryptionError(f"Failed to decrypt: {e}") from e
