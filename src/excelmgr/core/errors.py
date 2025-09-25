class ExcelMgrError(Exception):
    """Base exception for excelmgr."""

class DecryptionError(ExcelMgrError):
    pass

class SheetNotFound(ExcelMgrError):
    pass

class InvalidTargetPattern(ExcelMgrError):
    pass

class MissingColumnsError(ExcelMgrError):
    pass

class MacroLossWarning(Warning):
    pass
