def import_optional(name: str):
    try:
        return __import__(name)
    except Exception:
        return None
